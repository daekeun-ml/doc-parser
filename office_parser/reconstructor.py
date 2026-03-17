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


def reconstruct_sheet(
    sheet_json: dict,
    output_format: str,
    model_id: str = "gemini-2.5-flash",
) -> str:
    """단일 시트 JSON을 Gemini로 재구성.

    Args:
        sheet_json: _sheet_to_compact() 결과 dict
        output_format: "md" 또는 "html"
        model_id: Gemini 모델 ID

    Returns:
        재구성된 Markdown 또는 HTML 문자열
    """
    prompts = _load_prompts()
    prompt_key = "reconstruct_md" if output_format == "md" else "reconstruct_html"
    prompt = prompts[prompt_key]

    sheet_name = sheet_json.get("sheet_name", "Sheet")
    json_content = json.dumps(sheet_json, ensure_ascii=False, indent=2)

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

    # 이미지 summary 후처리: Gemini 출력에 의존하지 않고 확정적으로 반영
    if output_format == "md":
        result = _ensure_image_summaries_md(result, sheet_json)

    return result


def _ensure_image_summaries_md(text: str, sheet_json: dict) -> str:
    """Gemini 재구성 결과에 이미지 summary가 정확히 반영되도록 후처리.

    1) Gemini가 이미지를 출력했으면 → alt text를 summary로 교체
    2) Gemini가 이미지를 누락했으면 → 문서 끝에 추가
    """
    # sheet JSON에서 이미지 노드 추출
    images = []
    for item in sheet_json.get("rows", []):
        if isinstance(item, dict) and item.get("type") == "image":
            images.append(item)

    if not images:
        return text

    for img in images:
        filename = img.get("filename", "")
        summary = img.get("summary", "")
        if not filename:
            continue

        # 파일명이 결과에 있는지 찾기
        # ![아무텍스트](filename) 또는 ![아무텍스트](경로/filename) 패턴
        escaped = re.escape(filename)
        pattern = re.compile(r'!\[([^\]]*)\]\(([^)]*' + escaped + r'[^)]*)\)')
        match = pattern.search(text)

        if match:
            if summary:
                # alt text를 summary로 교체
                old = match.group(0)
                new = f"![{summary}]({match.group(2)})"
                text = text.replace(old, new, 1)
        else:
            # Gemini가 이미지를 누락 → 문서 끝에 추가
            alt = summary or "이미지"
            text += f"\n\n![{alt}]({filename})\n"

    return text


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
