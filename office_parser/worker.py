"""Single file parsing worker — separate module for ProcessPoolExecutor pickling."""
import json
import logging
import time
from pathlib import Path

from office_parser import OfficeParser, OfficeParserConfig

logger = logging.getLogger("office_parser")


def parse_single(
    file_path: str,
    config: OfficeParserConfig,
    output_format: str,
    output_dir: str | None = None,
) -> Path:
    """Parse a single file → save result. Returns: output path."""
    import sys
    logging.basicConfig(level=logging.INFO, format="%(message)s", stream=sys.stderr)
    logging.getLogger("botocore").setLevel(logging.WARNING)

    file_path = Path(file_path)
    name = file_path.name
    t0 = time.time()
    logger.info("📄 [%s] Parsing started", name)

    ast = OfficeParser.parse_office(str(file_path), config)

    # output/{stem}/ 디렉토리 생성 (PDF와 동일한 구조)
    doc_output = Path(output_dir) / file_path.stem if output_dir else file_path.parent / file_path.stem
    doc_output.mkdir(parents=True, exist_ok=True)

    # 첨부파일 → pictures/ 폴더에 저장
    pictures_dir = doc_output / "pictures"
    image_dir = None
    if config.extract_attachments and ast.attachments:
        pictures_dir.mkdir(parents=True, exist_ok=True)
        for att in ast.attachments:
            (pictures_dir / att.filename).write_bytes(att.data)
        # md 내 상대경로용: "pictures"
        image_dir = "pictures"

    # 출력 생성
    ext_map = {"html": ".html", "markdown": ".md", "text": ".txt", "json": ".json"}
    out_ext = ext_map.get(output_format, ".json")
    out_path = doc_output / f"{file_path.stem}{out_ext}"

    if output_format == "html":
        output = ast.to_html(image_dir=image_dir)
    elif output_format == "markdown":
        output = ast.to_markdown(image_dir=image_dir)
    elif output_format == "text":
        output = ast.to_text()
    elif output_format == "json":
        output = ast.to_json_compact()
    else:
        output = json.dumps(ast.__dict__, default=str, indent=2, ensure_ascii=False)

    out_path.write_text(output, encoding="utf-8")

    # Gemini 기반 reconstruct (JSON → clean MD)
    if config.reconstruct:
        from office_parser.reconstructor import reconstruct_all_sheets
        model_id = config.gemini_model_id

        logger.info("🔄 [%s] Reconstructing to MD...", name)
        try:
            rc_md = reconstruct_all_sheets(ast, "md", model_id)
            rc_md_path = doc_output / f"{file_path.stem}_reconstructed.md"
            rc_md_path.write_text(rc_md, encoding="utf-8")
            logger.info("✅ [%s] Reconstructed MD → %s", name, rc_md_path.name)
        except Exception as e:
            logger.error("❌ [%s] MD reconstruct failed: %s", name, e)

    logger.info("✅ [%s] Done → %s (%.1fs)", name, out_path, time.time() - t0)
    return out_path
