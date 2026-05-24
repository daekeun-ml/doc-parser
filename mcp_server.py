"""Document Parser MCP Server — FastMCP 기반 stdio / streamable-http 지원."""

from __future__ import annotations

import asyncio
import logging
import os
import queue
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from dotenv import load_dotenv
from fastmcp import FastMCP
from fastmcp.server.context import Context

load_dotenv()

DEFAULT_MODEL_ID = os.getenv("MODEL_ID", "gemini-2.5-flash")

mcp = FastMCP(
    name="doc-parser",
    instructions=(
        "PDF 및 Office(docx, pptx, xlsx) 문서에서 텍스트, 테이블, 이미지를 추출하고 "
        "Google Gemini LLM으로 요약/엔티티를 생성하는 도구입니다."
    ),
)

PDF_EXTENSIONS = {".pdf"}
OFFICE_EXTENSIONS = {".docx", ".pptx", ".xlsx", ".odt", ".odp", ".ods", ".rtf"}
ALL_EXTENSIONS = PDF_EXTENSIONS | OFFICE_EXTENSIONS


class QueueHandler(logging.Handler):
    """로그 메시지를 큐로 전달하는 핸들러."""

    def __init__(self, log_queue: queue.Queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        self.log_queue.put(self.format(record))


async def _run_with_logging(ctx: Context, func, *args):
    """블로킹 함수를 별도 스레드에서 실행하면서 로그를 MCP 클라이언트로 실시간 전달."""
    log_queue: queue.Queue[str] = queue.Queue()

    # 파서 로거들에 큐 핸들러 부착
    handler = QueueHandler(log_queue)
    handler.setFormatter(logging.Formatter("%(name)s — %(message)s"))
    loggers = [logging.getLogger(n) for n in ("doc_parser", "pdf_parser", "office_parser")]
    for lg in loggers:
        lg.addHandler(handler)
        lg.setLevel(logging.INFO)

    loop = asyncio.get_event_loop()
    future = loop.run_in_executor(ThreadPoolExecutor(max_workers=1), func, *args)

    try:
        step = 0
        tick = 0
        while not future.done():
            # 큐에 쌓인 로그 메시지를 클라이언트로 전송
            while not log_queue.empty():
                msg = log_queue.get_nowait()
                await ctx.info(msg)
            # 10초마다 progress 전송 (타임아웃 방지)
            tick += 1
            if tick % 10 == 0:
                step = min(step + 1, 99)
                await ctx.report_progress(step, 100)
            await asyncio.sleep(1)
        # 완료 후 남은 로그 flush
        while not log_queue.empty():
            msg = log_queue.get_nowait()
            await ctx.info(msg)
        await ctx.report_progress(100, 100)

        return await future
    finally:
        for lg in loggers:
            lg.removeHandler(handler)


# ── Tools ─────────────────────────────────────────────────────────────

@mcp.tool(
    name="parse_document",
    description=(
        "PDF 또는 Office(docx, pptx, xlsx, odt, odp, ods, rtf) 문서를 파싱하여 "
        "텍스트, 테이블, 이미지를 추출하고 선택적으로 LLM 요약을 생성합니다. "
        "결과는 output_dir/{파일명}/ 디렉토리에 마크다운/HTML/텍스트로 저장됩니다."
    ),
    tags={"parsing", "document"},
    timeout=600,
)
async def parse_document(
    file_path: str,
    ctx: Context,
    output_dir: str = "output",
    model_id: str = DEFAULT_MODEL_ID,
    no_summary: bool = False,
    table_mode: str = "accurate",
    output_format: str = "markdown",
) -> str:
    """단일 문서를 파싱합니다.

    Args:
        file_path: 파싱할 문서 경로 (PDF 또는 Office 파일)
        output_dir: 결과 저장 디렉토리 (기본: output)
        model_id: Gemini 모델 ID
        no_summary: True이면 LLM 요약을 비활성화
        table_mode: PDF 테이블 모드 (accurate 또는 fast)
        output_format: Office 출력 형식 (markdown, html, text)
    """
    from run import parse_single

    fp = Path(file_path)
    if not fp.exists():
        return f"❌ 파일을 찾을 수 없습니다: {file_path}"

    ext = fp.suffix.lower()
    if ext not in ALL_EXTENSIONS:
        return f"❌ 지원하지 않는 형식입니다: {ext} (지원: {', '.join(sorted(ALL_EXTENSIONS))})"

    await ctx.info(f"📄 파싱 시작: {fp.name}")
    t0 = time.time()

    try:
        out_path = await _run_with_logging(
            ctx,
            parse_single,
            fp, Path(output_dir), model_id, no_summary,
            table_mode, output_format,
        )
        elapsed = time.time() - t0
        await ctx.info(f"🎉 파싱 완료: {fp.name} ({elapsed:.1f}초)")
        return f"✅ 파싱 완료 → {out_path} ({elapsed:.1f}초)"
    except Exception as e:
        return f"❌ 파싱 실패: {e}"


@mcp.tool(
    name="parse_directory",
    description=(
        "폴더 내 모든 문서(PDF + Office)를 병렬로 일괄 파싱합니다. "
        "지원 형식: pdf, docx, pptx, xlsx, odt, odp, ods, rtf"
    ),
    tags={"parsing", "batch"},
    timeout=1800,
)
async def parse_directory(
    dir_path: str,
    ctx: Context,
    output_dir: str = "output",
    workers: int = 2,
    model_id: str = DEFAULT_MODEL_ID,
    no_summary: bool = False,
    table_mode: str = "accurate",
    output_format: str = "markdown",
) -> str:
    """폴더 내 문서를 일괄 파싱합니다.

    Args:
        dir_path: 문서가 들어있는 폴더 경로
        output_dir: 결과 저장 디렉토리 (기본: output)
        workers: 병렬 처리 워커 수 (기본: 2)
        model_id: Gemini 모델 ID
        no_summary: True이면 LLM 요약을 비활성화
        table_mode: PDF 테이블 모드 (accurate 또는 fast)
        output_format: Office 출력 형식 (markdown, html, text)
    """
    from concurrent.futures import ThreadPoolExecutor, as_completed
    from run import parse_single

    dp = Path(dir_path)
    if not dp.is_dir():
        return f"❌ 디렉토리를 찾을 수 없습니다: {dir_path}"

    files = sorted(f for f in dp.iterdir() if f.suffix.lower() in ALL_EXTENSIONS)
    if not files:
        return f"❌ 지원되는 문서가 없습니다: {dir_path}"

    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)

    await ctx.info(f"📂 총 {len(files)}개 파일 파싱 시작 (workers={workers})")
    t0 = time.time()

    def _run_batch():
        success, fail, errors = 0, 0, []
        with ThreadPoolExecutor(max_workers=workers) as ex:
            futs = {
                ex.submit(
                    parse_single, f, out, model_id, no_summary,
                    table_mode, output_format,
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
                    errors.append(f"  - {f.name}: {e}")
        return success, fail, errors

    success, fail, errors = await _run_with_logging(ctx, _run_batch)
    elapsed = time.time() - t0
    await ctx.info(f"🏁 완료: {success}/{len(files)} 성공, {fail} 실패 ({elapsed:.1f}초)")

    result = f"🏁 완료: {success}/{len(files)} 성공, {fail} 실패 ({elapsed:.1f}초)"
    if errors:
        result += "\n\n실패 목록:\n" + "\n".join(errors)
    return result


@mcp.tool(
    name="list_supported_formats",
    description="지원하는 문서 형식 목록을 반환합니다.",
    tags={"info"},
)
def list_supported_formats() -> str:
    """지원 파일 형식과 파서 정보를 반환합니다."""
    formats = [
        ("PDF", ".pdf", "pdf_parser (Docling)"),
        ("Word", ".docx", "office_parser"),
        ("PowerPoint", ".pptx", "office_parser"),
        ("Excel", ".xlsx", "office_parser"),
        ("OpenDocument Text", ".odt", "office_parser"),
        ("OpenDocument Presentation", ".odp", "office_parser"),
        ("OpenDocument Spreadsheet", ".ods", "office_parser"),
        ("RTF", ".rtf", "office_parser"),
    ]
    lines = ["| 형식 | 확장자 | 파서 |", "|------|--------|------|"]
    for name, ext, parser in formats:
        lines.append(f"| {name} | `{ext}` | {parser} |")
    return "\n".join(lines)


# ── Entry point ───────────────────────────────────────────────────────

if __name__ == "__main__":
    transport = sys.argv[1] if len(sys.argv) > 1 else "stdio"

    if transport == "http":
        host = os.getenv("MCP_HOST", "0.0.0.0")
        port = int(os.getenv("MCP_PORT", "8000"))
        mcp.run(transport="streamable-http", host=host, port=port, stateless_http=True)
    else:
        mcp.run(transport="stdio")
