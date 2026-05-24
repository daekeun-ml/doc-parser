"""Document Parser CLI — PDF / Office(docx, pptx, xlsx) unified parsing.

Usage:
    # Single PDF
    uv run python run.py sample.pdf -o output

    # Single Office file
    uv run python run.py report.docx -o output --to-markdown

    # Batch folder (PDF + Office mixed)
    uv run python run.py ./docs/ -o output --workers 4
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

DEFAULT_MODEL_ID = os.getenv("MODEL_ID", "gemini-2.5-flash")

logger = logging.getLogger("doc_parser")

PDF_EXTENSIONS = {".pdf"}
OFFICE_EXTENSIONS = {".docx", ".pptx", ".xlsx", ".odt", ".odp", ".ods", ".rtf"}
ALL_EXTENSIONS = PDF_EXTENSIONS | OFFICE_EXTENSIONS


def _setup_logging(verbose: bool = False):
    fmt = "%(asctime)s %(levelname)-5s %(name)s — %(message)s"
    logging.basicConfig(level=logging.WARNING, format=fmt, datefmt="%H:%M:%S")
    level = logging.DEBUG if verbose else logging.INFO

    # 파일 핸들러: log/YYYYMMDD_HHMMSS.log
    from datetime import datetime
    log_dir = Path("log")
    log_dir.mkdir(exist_ok=True)
    fh = logging.FileHandler(
        log_dir / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.log", encoding="utf-8"
    )
    fh.setLevel(level)
    fh.setFormatter(logging.Formatter(fmt, datefmt="%H:%M:%S"))

    for name in ("doc_parser", "pdf_parser", "office_parser"):
        lg = logging.getLogger(name)
        lg.setLevel(level)
        lg.addHandler(fh)


# ── PDF parsing ───────────────────────────────────────────────────────
def parse_pdf(
    pdf_path: Path,
    output_dir: Path,
    model_id: str,
    no_summary: bool,
    table_mode: str,
    verbose: bool = False,
) -> Path:
    """Parse a single PDF → save final markdown."""
    _setup_logging(verbose)
    from pdf_parser.converter import DoclingConverter
    from pdf_parser.summarizer import GeminiSummarizer
    from pdf_parser.markdown_builder import MarkdownBuilder

    pdf_logger = logging.getLogger("pdf_parser")
    pdf_path = Path(pdf_path)
    name = pdf_path.name
    doc_output = output_dir / pdf_path.stem
    doc_output.mkdir(parents=True, exist_ok=True)
    t0 = time.time()

    pdf_logger.info("📄 [%s] Docling conversion started", name)
    converter = DoclingConverter(table_mode=table_mode)
    parsed = converter.convert(pdf_path)
    n_pages = len(parsed.doc.pages)
    n_figs = len(parsed.get_figures())
    n_tbls = len(parsed.get_tables())
    pdf_logger.info(
        "✅ [%s] Conversion done — %d pages, %d figures, %d tables (%.1fs)",
        name, n_pages, n_figs, n_tbls, time.time() - t0,
    )

    parsed.save_assets(doc_output)

    text_path = doc_output / f"{parsed.doc_name}_text.md"
    text_path.write_text(parsed.doc.export_to_markdown(), encoding="utf-8")
    pdf_logger.info("📃 [%s] Raw text saved → %s", name, text_path)

    page_summaries, image_summaries, table_summaries = {}, {}, {}
    if not no_summary:
        summarizer = GeminiSummarizer(model_id=model_id)

        pdf_logger.info("🔍 [%s] Summarizing pages... (%d pages)", name, n_pages)
        t1 = time.time()
        page_summaries = summarizer.summarize_pages(parsed)
        pdf_logger.info("✅ [%s] Page summaries done (%.1fs)", name, time.time() - t1)

        pdf_logger.info("🖼️  [%s] Summarizing figures... (%d)", name, n_figs)
        t1 = time.time()
        image_summaries = summarizer.summarize_figures(parsed, page_summaries)
        pdf_logger.info("✅ [%s] Figure summaries done (%.1fs)", name, time.time() - t1)

        pdf_logger.info("📊 [%s] Summarizing tables... (%d)", name, n_tbls)
        t1 = time.time()
        table_summaries = summarizer.summarize_tables(parsed, page_summaries)
        pdf_logger.info("✅ [%s] Table summaries done (%.1fs)", name, time.time() - t1)

    builder = MarkdownBuilder(parsed, doc_output)
    final_md = builder.build(page_summaries, image_summaries, table_summaries)

    out_path = doc_output / f"{parsed.doc_name}_final.md"
    out_path.write_text(final_md, encoding="utf-8")
    pdf_logger.info("🎉 [%s] Done → %s (%.1fs total)", name, out_path, time.time() - t0)
    return out_path


# ── Office parsing ────────────────────────────────────────────────────
def parse_office(
    file_path: Path,
    output_dir: Path,
    model_id: str,
    no_summary: bool,
    output_format: str = "markdown",
    reconstruct: bool = False,
    verbose: bool = False,
) -> Path:
    """Parse a single Office file."""
    _setup_logging(verbose)
    from office_parser.worker import parse_single
    from office_parser import OfficeParserConfig

    config = OfficeParserConfig(
        summarize=not no_summary,
        gemini_model_id=model_id,
        reconstruct=reconstruct,
    )
    return parse_single(str(file_path), config, output_format, str(output_dir))


# ── Unified dispatcher ────────────────────────────────────────────────
def parse_single(
    file_path: Path,
    output_dir: Path,
    model_id: str,
    no_summary: bool,
    table_mode: str,
    output_format: str,
    reconstruct: bool = False,
    verbose: bool = False,
) -> Path:
    """Dispatch to PDF / Office parser based on file extension."""
    file_path = Path(file_path)
    ext = file_path.suffix.lower()
    if ext in PDF_EXTENSIONS:
        return parse_pdf(file_path, output_dir, model_id, no_summary, table_mode, verbose)
    elif ext in OFFICE_EXTENSIONS:
        return parse_office(file_path, output_dir, model_id, no_summary, output_format, reconstruct, verbose)
    else:
        raise ValueError(f"Unsupported file format: {ext}")


def main():
    parser = argparse.ArgumentParser(description="Document Parser — PDF + Office(docx/pptx/xlsx) unified CLI")
    parser.add_argument("input", type=Path, help="File or folder path")
    parser.add_argument("-o", "--output", type=Path, default=Path("output"), help="Output directory (default: output)")
    parser.add_argument("--workers", type=int, default=2, help="Parallel workers for folder mode (default: 2)")
    parser.add_argument("--no-summary", action="store_true", help="Disable LLM summarization")
    parser.add_argument("--model-id", default=DEFAULT_MODEL_ID, help="Gemini model ID (default: .env MODEL_ID)")
    parser.add_argument("--table-mode", choices=["accurate", "fast"], default="accurate", help="PDF TableFormer mode")
    parser.add_argument("--to-markdown", action="store_true", help="Office output format: markdown (default)")
    parser.add_argument("--to-html", action="store_true", help="Office output format: html")
    parser.add_argument("--to-text", action="store_true", help="Office output format: text")
    parser.add_argument("--to-json", action="store_true", help="Office output format: compact JSON (RAG optimized)")
    parser.add_argument("--reconstruct", action="store_true", help="Gemini로 JSON→clean MD/HTML 재구성")
    parser.add_argument("-v", "--verbose", action="store_true", help="Enable DEBUG logging")
    args = parser.parse_args()

    _setup_logging(args.verbose)
    args.output.mkdir(parents=True, exist_ok=True)
    logger.info("🤖 Using model: %s", args.model_id)

    output_format = "json" if args.to_json else "html" if args.to_html else "text" if args.to_text else "markdown"

    # Single file
    if args.input.is_file():
        parse_single(
            args.input, args.output, args.model_id, args.no_summary,
            args.table_mode, output_format, args.reconstruct, args.verbose,
        )
        return

    # Batch folder
    if args.input.is_dir():
        files = sorted(f for f in args.input.iterdir() if f.suffix.lower() in ALL_EXTENSIONS)
        if not files:
            logger.error("❌ No supported files found in: %s", args.input)
            sys.exit(1)

        logger.info("📂 %d files found, starting parallel processing with %d workers", len(files), args.workers)
        t0 = time.time()
        success, fail = 0, 0

        with ProcessPoolExecutor(max_workers=args.workers) as ex:
            futs = {
                ex.submit(
                    parse_single, f, args.output, args.model_id, args.no_summary,
                    args.table_mode, output_format, args.verbose,
                ): f
                for f in files
            }
            for fut in as_completed(futs):
                f = futs[fut]
                try:
                    fut.result()
                    success += 1
                except Exception as e:
                    fail += 1
                    logger.error("❌ [%s] Error: %s", f.name, e)

        logger.info(
            "🏁 All done: %d/%d succeeded, %d failed (%.1fs total)",
            success, len(files), fail, time.time() - t0,
        )
        return

    logger.error("❌ Input path does not exist: %s", args.input)
    sys.exit(1)


if __name__ == "__main__":
    main()
