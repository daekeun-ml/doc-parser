"""Compact JSON → Gemini 기반 Markdown/HTML 재구성 모듈."""
import json
import logging
import os
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import yaml

logger = logging.getLogger("office_parser")

_PROMPTS_PATH = Path(__file__).parent / "prompts.yaml"
_prompts_cache = None


def _load_prompts() -> dict:
    global _prompts_cache
    if _prompts_cache is None:
        with open(_PROMPTS_PATH, "r", encoding="utf-8") as f:
            _prompts_cache = yaml.safe_load(f)
    return _prompts_cache


def _get_gemini_client(model_id: str):
    from google import genai
    return genai.Client(api_key=os.getenv("GOOGLE_API_KEY"))


def _call_gemini(client, model_id: str, system: str, user: str) -> str:
    """Gemini API 호출. system instruction + user prompt."""
    response = client.models.generate_content(
        model=model_id,
        contents=user,
        config={"system_instruction": system},
    )
    return response.text


def _extract_images_from_json(sheet_json: dict) -> tuple:
    """sheet JSON에서 이미지/차트 노드를 분리.

    Returns:
        (cleaned_json, images) — 이미지 제거된 JSON과 이미지 노드 리스트
    """
    images = []
    cleaned_rows = []
    for item in sheet_json.get("rows", []):
        if isinstance(item, dict) and item.get("type") in ("image", "chart"):
            images.append(item)
        else:
            cleaned_rows.append(item)

    cleaned = dict(sheet_json)
    cleaned["rows"] = cleaned_rows
    return cleaned, images


def _insert_images_md(text: str, images: list) -> str:
    """Gemini 재구성 결과에 이미지를 삽입.

    이미지는 문서 summary 뒤, 본문 시작 전에 배치.
    """
    if not images:
        return text

    img_lines = []
    for img in images:
        if img.get("type") == "chart":
            chart_type = img.get("chart_type", "Chart")
            title = img.get("title", "")
            img_lines.append(f"**[{chart_type}]** {title}\n")
        else:
            filename = img.get("filename", "")
            summary = img.get("summary", "")
            if not filename:
                continue
            alt = summary or "이미지"
            img_lines.append(f"![{alt}]({filename})\n")

    if not img_lines:
        return text

    img_block = "\n".join(img_lines)

    # summary 블록(첫 번째 문단) 뒤에 삽입 시도, 없으면 문서 앞에
    lines = text.split("\n")
    insert_idx = 0
    for i, line in enumerate(lines):
        # 첫 번째 빈 줄(summary 끝)을 찾으면 그 뒤에 삽입
        if i > 0 and line.strip() == "" and lines[i - 1].strip():
            insert_idx = i + 1
            break
    else:
        insert_idx = 0

    lines.insert(insert_idx, "\n" + img_block)
    return "\n".join(lines)


def reconstruct_sheet(
    sheet_json: dict,
    output_format: str,
    model_id: str = "gemini-2.5-flash",
) -> str:
    """단일 시트 JSON을 Gemini로 재구성.

    이미지/차트는 Gemini에 보내지 않고, 후처리로 삽입.

    Args:
        sheet_json: _sheet_to_compact() 결과 dict
        output_format: "md" 또는 "html"
        model_id: Gemini 모델 ID

    Returns:
        재구성된 Markdown 또는 HTML 문자열
    """
    # 이미지/차트 분리 → Gemini는 테이블/텍스트 재구성에만 집중
    cleaned_json, images = _extract_images_from_json(sheet_json)

    prompts = _load_prompts()
    prompt_key = "reconstruct_md" if output_format == "md" else "reconstruct_html"
    prompt = prompts[prompt_key]

    sheet_name = cleaned_json.get("sheet_name", "Sheet")
    json_content = json.dumps(cleaned_json, ensure_ascii=False, indent=2)

    system = prompt["system"]
    user = prompt["user"].format(
        sheet_name=sheet_name,
        json_content=json_content,
    )

    client = _get_gemini_client(model_id)
    result = _call_gemini(client, model_id, system, user)

    if not result:
        return ""

    # 코드 블록 래핑 제거
    if result.startswith("```markdown"):
        result = result[len("```markdown"):].strip()
    if result.startswith("```html"):
        result = result[len("```html"):].strip()
    if result.startswith("```"):
        result = result[3:].strip()
    if result.endswith("```"):
        result = result[:-3].strip()

    # 이미지/차트 후처리 삽입
    if output_format == "md":
        result = _insert_images_md(result, images)

    return result


def reconstruct_all_sheets(
    ast,
    output_format: str,
    model_id: str = "gemini-2.5-flash",
) -> str:
    """AST의 모든 시트를 병렬로 재구성하여 하나의 문서로 합침.

    Args:
        ast: OfficeParserAST 인스턴스
        output_format: "md" 또는 "html"
        model_id: Gemini 모델 ID

    Returns:
        전체 재구성된 문서 문자열
    """
    # 시트별 compact JSON 생성
    sheet_jsons = []
    for node in ast.content:
        if node.type == "sheet":
            sheet_json = ast._sheet_to_compact(node)
            sheet_jsons.append(sheet_json)

    if not sheet_jsons:
        return ""

    # 병렬 재구성
    logger.info("🔄 Reconstructing %d sheets (%s, parallel)...", len(sheet_jsons), output_format)
    results = {}

    with ThreadPoolExecutor() as executor:
        futures = {}
        for i, sj in enumerate(sheet_jsons):
            f = executor.submit(reconstruct_sheet, sj, output_format, model_id)
            futures[f] = (i, sj.get("sheet_name", f"Sheet_{i}"))

        for f in as_completed(futures):
            idx, name = futures[f]
            try:
                results[idx] = f.result()
                logger.info("✅ Reconstructed sheet '%s'", name)
            except Exception as e:
                logger.error("❌ Reconstruct failed for '%s': %s", name, e)
                results[idx] = f"<!-- Reconstruct failed: {name} -->"

    # 순서대로 합치기
    ordered = [results[i] for i in sorted(results.keys())]

    if output_format == "md":
        return "\n\n---\n\n".join(ordered)
    else:
        # HTML: wrap in full document
        body = "\n<hr />\n".join(ordered)
        title = ast.metadata.title or "Reconstructed Document"
        return f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{title}</title>
<style>
body {{ font-family: 'Pretendard', -apple-system, sans-serif; background: #f8f9fa; color: #374151; line-height: 1.6; padding: 2rem; max-width: 1400px; margin: 0 auto; }}
table {{ border-collapse: collapse; margin: 1rem 0; font-size: 0.85rem; }}
th, td {{ padding: 0.5rem 0.7rem; border: 1px solid #d1d5db; text-align: left; vertical-align: top; }}
th {{ background: #f1f5f9; font-weight: 600; }}
h1, h2, h3 {{ color: #1f2937; margin: 1rem 0 0.5rem; }}
hr {{ border: none; height: 1px; background: #e2e8f0; margin: 2rem 0; }}
</style>
</head>
<body>
{body}
</body>
</html>"""
