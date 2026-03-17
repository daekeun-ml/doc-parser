from dataclasses import dataclass, field
from typing import Optional, List, Dict, Literal, Union, TypedDict
from datetime import datetime


# ── Reconstruct 관련 타입 ──

class CompressedSheet(TypedDict):
    sheet: str
    meta: dict  # maxRow, maxCol, mergedCells, gap_rows, gap_cols, stats
    cols: list  # ["r","c","v","b","bg","fg","cs"]
    colDesc: dict
    cells: list  # list of [row, col, value, bold, bg, fg, colspan]


class TableRegion(TypedDict):
    id: str
    name: str
    range: list  # [min_row, min_col, max_row, max_col]
    cells: list


class ReconstructedTable(TypedDict):
    tableId: str
    name: str
    headers: Optional[list]
    cols: list  # ["r","c","v","action"]
    cells: list
    structure: Optional[str]  # LLM이 판단한 구조 설명

@dataclass
class OfficeParserConfig:
    output_error_to_console: bool = False
    newline_delimiter: str = "\n"
    ignore_notes: bool = False
    put_notes_at_last: bool = False
    extract_attachments: bool = True
    include_raw_content: bool = False
    ocr: bool = False
    ocr_language: str = "eng"
    summarize: bool = True
    min_image_size: int = 150  # 이미지 요약 최소 크기 (px). 가로/세로 모두 이 값 이상이어야 요약
    gemini_model_id: str = "gemini-2.5-flash"
    reconstruct: bool = False  # reconstruct 활성화
    reconstruct_model: str = "gemini-2.5-flash"  # reconstruct용 모델
    reconstruct_similarity_threshold: float = 0.7  # 파편 재결합 열 유사도 기준

@dataclass
class TextFormatting:
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    strikethrough: Optional[bool] = None
    color: Optional[str] = None
    background_color: Optional[str] = None
    size: Optional[str] = None
    font: Optional[str] = None
    subscript: Optional[bool] = None
    superscript: Optional[bool] = None
    alignment: Optional[Literal["left", "center", "right", "justify"]] = None

@dataclass
class ChartData:
    title: Optional[str] = None
    x_axis_title: Optional[str] = None
    y_axis_title: Optional[str] = None
    data_sets: List[Dict] = field(default_factory=list)
    labels: List[str] = field(default_factory=list)
    raw_texts: List[str] = field(default_factory=list)

@dataclass
class OfficeAttachment:
    type: Literal["image", "chart"]
    data: bytes = None
    filename: str = ""
    mime_type: str = ""
    extension: str = ""
    ocr_text: Optional[str] = None
    alt_text: Optional[str] = None
    chart_data: Optional[ChartData] = None

@dataclass
class OfficeContentNode:
    type: str
    text: Optional[str] = None
    children: Optional[List['OfficeContentNode']] = None
    formatting: Optional[TextFormatting] = None
    metadata: Optional[Dict] = None
    raw_content: Optional[str] = None

@dataclass
class OfficeMetadata:
    title: Optional[str] = None
    author: Optional[str] = None
    last_modified_by: Optional[str] = None
    created: Optional[datetime] = None
    modified: Optional[datetime] = None
    description: Optional[str] = None
    subject: Optional[str] = None
    pages: Optional[int] = None
    formatting: Optional[TextFormatting] = None
    style_map: Optional[Dict[str, TextFormatting]] = None
    document_summary: Optional[str] = None

@dataclass
class OfficeParserAST:
    type: str
    metadata: OfficeMetadata
    content: List[OfficeContentNode]
    attachments: List[OfficeAttachment] = field(default_factory=list)
    
    def to_json_compact(self) -> str:
        """RAG용 압축 JSON 변환. 빈 셀 제거, 헤더-값 매핑, 구조 보존."""
        import json
        result = {
            "type": self.type,
            "metadata": {},
            "sheets": [],
        }
        if self.metadata.title:
            result["metadata"]["title"] = self.metadata.title
        if self.metadata.document_summary:
            result["metadata"]["document_summary"] = self.metadata.document_summary

        for node in self.content:
            if node.type == "sheet":
                result["sheets"].append(self._sheet_to_compact(node))
            elif node.type == "slide":
                result.setdefault("slides", []).append(self._slide_to_compact(node))
            elif node.type == "page":
                result.setdefault("pages", []).append(self._page_to_compact(node))
            elif node.type == "section":
                result.setdefault("sections", []).append(self._section_to_compact(node))
            else:
                result.setdefault("content", []).append(self._node_to_compact(node))

        # 빈 리스트 제거
        result = {k: v for k, v in result.items() if v}
        return json.dumps(result, ensure_ascii=False, indent=2)

    def _sheet_to_compact(self, sheet: 'OfficeContentNode') -> dict:
        meta = sheet.metadata or {}
        out = {"sheet_name": meta.get("sheetName", "Sheet")}
        if meta.get("sheet_summary"):
            out["summary"] = meta["sheet_summary"]

        content = []
        for child in (sheet.children or []):
            if child.type == "row" and child.children:
                row_data = self._extract_row_compact(child)
                if row_data:
                    content.append(row_data)
            elif child.type == "chart":
                content.append(self._chart_to_compact(child))
            elif child.type == "image":
                content.append(self._image_to_compact(child))

        if content:
            out["rows"] = content
        return out

    def _extract_row_compact(self, row: 'OfficeContentNode') -> dict:
        """행을 col 위치 기반 compact dict로 변환. 빈 셀 제거."""
        row_meta = row.metadata or {}
        row_num = row_meta.get("row")
        cells = {}
        bg = {}
        for cell in (row.children or []):
            text = (cell.text or "").strip()
            meta = cell.metadata or {}
            col = meta.get("col", 0)
            col_key = str(col)
            cell_bg = meta.get("style", {}).get("background-color") if meta.get("style") else None
            if text:
                cells[col_key] = text
            if cell_bg:
                bg[col_key] = cell_bg
        if not cells and not bg:
            return {}
        out = {}
        if row_num is not None:
            out["r"] = row_num
        if cells:
            out["cells"] = cells
        if bg:
            out["bg"] = bg
        return out

    def _chart_to_compact(self, node: 'OfficeContentNode') -> dict:
        meta = node.metadata or {}
        out = {"type": "chart", "chart_type": meta.get("chartType", "")}
        if meta.get("title"):
            out["title"] = meta["title"]
        return out

    def _image_to_compact(self, node: 'OfficeContentNode') -> dict:
        meta = node.metadata or {}
        out = {"type": "image"}
        if meta.get("filename"):
            out["filename"] = meta["filename"]
        if meta.get("image_summary"):
            out["summary"] = meta["image_summary"]
        return out

    def _slide_to_compact(self, node: 'OfficeContentNode') -> dict:
        meta = node.metadata or {}
        out = {"slide": meta.get("slideNumber", 1)}
        if meta.get("slideTitle"):
            out["title"] = meta["slideTitle"]
        if meta.get("slide_summary"):
            out["summary"] = meta["slide_summary"]
        children = []
        for child in (node.children or []):
            compact = self._node_to_compact(child)
            if compact:
                children.append(compact)
        if children:
            out["content"] = children
        return out

    def _page_to_compact(self, node: 'OfficeContentNode') -> dict:
        out = {"page": (node.metadata or {}).get("pageNumber", 1)}
        if node.text:
            out["text"] = node.text
        return out

    def _section_to_compact(self, node: 'OfficeContentNode') -> dict:
        meta = node.metadata or {}
        out = {}
        if meta.get("sectionTitle"):
            out["title"] = meta["sectionTitle"]
        if meta.get("section_summary"):
            out["summary"] = meta["section_summary"]
        children = []
        for child in (node.children or []):
            compact = self._node_to_compact(child)
            if compact:
                children.append(compact)
        if children:
            out["content"] = children
        return out

    def _node_to_compact(self, node: 'OfficeContentNode') -> dict:
        if node.type in ("paragraph", "heading", "list", "notes"):
            if not node.text:
                return {}
            out = {"type": node.type, "text": node.text}
            if node.type == "heading" and node.metadata:
                out["level"] = node.metadata.get("level", 1)
            return out
        elif node.type == "table":
            rows = []
            for row in (node.children or []):
                if row.type == "row" and row.children:
                    row_data = self._extract_row_compact(row)
                    if row_data:
                        rows.append(row_data)
            return {"type": "table", "rows": rows} if rows else {}
        elif node.type == "image":
            return self._image_to_compact(node)
        elif node.type == "chart":
            return self._chart_to_compact(node)
        elif node.text:
            return {"text": node.text}
        return {}

    def to_text(self, delimiter: str = "\n") -> str:
        def extract_text(nodes: List[OfficeContentNode]) -> str:
            texts = []
            for node in nodes:
                if node.text:
                    texts.append(node.text)
                if node.children:
                    texts.append(extract_text(node.children))
            return delimiter.join(filter(None, texts))
        return extract_text(self.content)
    
    def to_compressed_json(self, sheet_node: 'OfficeContentNode') -> CompressedSheet:
        """시트 노드를 Schema+Values 압축 JSON으로 변환.

        셀 배열 형식: [row, col, value, bold(1/0), bg_hex_or_0, fg_hex_or_0, colspan_or_0]
        스타일 없음 = 0, 빈 값 = ""
        """
        meta = sheet_node.metadata or {}
        sheet_name = meta.get("sheetName", "Sheet")
        merged_cells = []

        cells = []
        for child in (sheet_node.children or []):
            if child.type != "row" or not child.children:
                continue
            for cell_node in child.children:
                cm = cell_node.metadata or {}
                r = cm.get("row", 0)
                c = cm.get("col", 0)
                v = cell_node.text or ""
                style = cm.get("style", {})
                b = 1 if style.get("font-weight") == "bold" else 0
                bg = style.get("background-color", 0) or 0
                fg = style.get("color", 0) or 0
                cs = cm.get("colspan", 0) or 0
                if cs == 1:
                    cs = 0
                cells.append([r, c, v, b, bg, fg, cs])

        return CompressedSheet(
            sheet=sheet_name,
            meta={
                "maxRow": meta.get("maxRow", 0),
                "maxCol": meta.get("maxColumn", 0),
                "mergedCells": merged_cells,
            },
            cols=["r", "c", "v", "b", "bg", "fg", "cs"],
            colDesc={
                "r": "row", "c": "col", "v": "value",
                "b": "bold(1/0)", "bg": "background hex or 0",
                "fg": "font color hex or 0", "cs": "colspan or 0",
            },
            cells=cells,
        )

    def to_markdown(self, image_dir: str = None) -> str:
        self._image_dir = image_dir
        self._heading_offset = 2 if self.metadata.document_summary else 0
        lines = []
        
        if self.metadata.title:
            lines.append(f"# {self.metadata.title}\n")
        
        if self.metadata.document_summary:
            label = "Document Summary" if self.type == "docx" else "Deck Summary"
            lines.append(f"## {label}\n\n{self.metadata.document_summary}\n")
            lines.append("## Content\n")
        
        for i, node in enumerate(self.content):
            if i > 0 and node.type == "sheet":
                lines.append("---\n")
            lines.append(self._node_to_markdown(node))
        
        return "\n".join(filter(None, lines))
    
    def to_html(self, image_dir: str = None) -> str:
        """HTML 테이블 형태로 변환 (메타데이터 포함)"""
        self._image_dir = image_dir
        parts = []

        if self.metadata.title:
            parts.append(f"<h1>{self.metadata.title}</h1>")

        if self.metadata.document_summary:
            label = "Document Summary" if self.type == "docx" else "Deck Summary"
            parts.append(f"<h2>{label}</h2>\n<p>{self.metadata.document_summary}</p>")
            parts.append("<h2>Content</h2>")

        for i, node in enumerate(self.content):
            if i > 0 and node.type == "sheet":
                parts.append("<hr />")
            parts.append(self._node_to_html(node))

        body = "\n".join(filter(None, parts))
        return self._wrap_html(body)

    def _wrap_html(self, body: str) -> str:
        """완성된 HTML 문서로 래핑 (CSS 테마 포함)"""
        title = self.metadata.title or "Document"
        return f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{title}</title>
<style>
:root {{
  --bg: #f8f9fa;
  --surface: #ffffff;
  --border: #e2e8f0;
  --text: #374151;
  --text-secondary: #6b7280;
  --accent: #2563eb;
  --accent-light: #eff6ff;
  --header-bg: #f1f5f9;
  --header-text: #1f2937;
  --summary-bg: #f0fdf4;
  --summary-border: #86efac;
  --image-summary-bg: #fefce8;
  --image-summary-border: #fde047;
  --table-summary-bg: #f5f3ff;
  --table-summary-border: #c4b5fd;
}}
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{
  font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
  background: var(--bg);
  color: var(--text);
  line-height: 1.6;
  padding: 2rem;
}}
.sheet, .slide {{
  background: var(--surface);
  border-radius: 12px;
  box-shadow: 0 1px 3px rgba(0,0,0,0.08), 0 1px 2px rgba(0,0,0,0.06);
  padding: 2rem;
  margin-bottom: 2rem;
}}
h1 {{
  font-size: 1.5rem;
  font-weight: 700;
  color: var(--text);
  margin-bottom: 1rem;
  padding-bottom: 0.75rem;
  border-bottom: 2px solid var(--accent);
}}
h2 {{
  font-size: 1.25rem;
  font-weight: 600;
  color: var(--text);
  margin: 1.5rem 0 0.75rem;
}}
h3 {{
  font-size: 1.1rem;
  font-weight: 600;
  color: var(--text);
  margin: 1rem 0 0.5rem;
}}
hr {{
  border: none;
  height: 1px;
  background: var(--border);
  margin: 2rem 0;
}}
p {{
  margin: 0.5rem 0;
  color: var(--text);
}}
.sheet-summary {{
  background: var(--summary-bg);
  border-left: 4px solid var(--summary-border);
  border-radius: 0 8px 8px 0;
  padding: 1rem 1.25rem;
  margin: 1rem 0 1.5rem;
  font-size: 0.9rem;
  color: var(--text-secondary);
  line-height: 1.7;
}}
table {{
  width: 100%;
  border-collapse: collapse;
  margin: 1rem 0;
  font-size: 0.85rem;
}}
table th, table td {{
  padding: 0.5rem 0.75rem;
  border: 1px solid var(--border);
  text-align: left;
  vertical-align: top;
}}
table th {{
  background: var(--header-bg);
  color: var(--header-text);
  font-weight: 600;
}}
table tr:nth-child(even) td:not([style*="background"]) {{
  background: #f8fafc;
}}
table tr:hover td {{
  background: var(--accent-light) !important;
  transition: background 0.15s;
}}
table.image-meta,
table.section-meta,
table.page-meta,
table.table-meta,
table.chart-meta {{
  border: none;
  margin: 0.75rem 0;
  font-size: 0.85rem;
}}
table.image-meta td,
table.section-meta td,
table.page-meta td,
table.table-meta td,
table.chart-meta td {{
  border: none;
  padding: 0.4rem 0.75rem;
}}
table.image-meta {{ background: var(--image-summary-bg); border-left: 4px solid var(--image-summary-border); border-radius: 0 8px 8px 0; }}
table.section-meta {{ background: var(--summary-bg); border-left: 4px solid var(--summary-border); border-radius: 0 8px 8px 0; }}
table.page-meta {{ background: var(--accent-light); border-left: 4px solid var(--accent); border-radius: 0 8px 8px 0; }}
table.table-meta {{ background: var(--table-summary-bg); border-left: 4px solid var(--table-summary-border); border-radius: 0 8px 8px 0; }}
.image {{
  margin: 1rem 0;
  text-align: center;
}}
.image img {{
  max-width: 100%;
  height: auto;
  border-radius: 8px;
  box-shadow: 0 2px 8px rgba(0,0,0,0.1);
}}
.image-summary {{
  font-size: 0.85rem;
  color: var(--text-secondary);
  margin-top: 0.5rem;
  font-style: italic;
}}
.chart {{
  background: #fefce8;
  border: 1px dashed #eab308;
  border-radius: 8px;
  padding: 0.75rem 1rem;
  margin: 0.75rem 0;
  font-size: 0.9rem;
}}
blockquote.notes {{
  background: #f1f5f9;
  border-left: 4px solid var(--accent);
  border-radius: 0 8px 8px 0;
  padding: 0.75rem 1rem;
  margin: 0.5rem 0;
  color: var(--text-secondary);
  font-style: italic;
}}
img {{
  max-width: 100%;
  height: auto;
}}
</style>
</head>
<body>
{body}
</body>
</html>"""
    
    def _node_to_html(self, node: OfficeContentNode) -> str:
        if node.type == "sheet":
            return self._sheet_to_html(node)
        elif node.type == "slide":
            return self._slide_to_html(node)
        elif node.type == "paragraph":
            indent = node.metadata.get("indent_level", 0) if node.metadata else 0
            style = f' style="margin-left:{indent * 2}em"' if indent > 0 else ""
            return f"<p{style}>{node.text}</p>"
        elif node.type == "heading":
            h_level = node.metadata.get("level", 1) if node.metadata else 1
            return f"<h{h_level}>{node.text}</h{h_level}>"
        elif node.type == "list":
            indent = node.metadata.get("indent_level", 0) if node.metadata else 0
            style = f' style="margin-left:{indent * 2}em"' if indent > 0 else ""
            return f"<li{style}>{node.text}</li>"
        elif node.type == "page":
            page_num = node.metadata.get("pageNumber", 1) if node.metadata else 1
            page_summary = node.metadata.get("page_summary", "") if node.metadata else ""
            parts = [f"<h2>Page {page_num}</h2>"]
            if page_summary:
                parts.append(f'<table class="page-meta"><tr><td>page_summary</td><td>{page_summary}</td></tr></table>')
            if node.text:
                parts.append(f"<p>{node.text}</p>")
            if node.children:
                for child in node.children:
                    parts.append(self._node_to_html(child))
            return "\n".join(parts)
        elif node.type == "section":
            meta = node.metadata or {}
            section_summary = meta.get("section_summary", "")
            parts = []
            if section_summary:
                parts.append(f'<table class="section-meta"><tr><td>section_summary</td><td>{section_summary}</td></tr></table>')
            if node.children:
                for child in node.children:
                    parts.append(self._node_to_html(child))
            return "\n".join(parts)
        elif node.type == "image":
            m = node.metadata or {}
            filename = m.get("filename", "")
            img_summary = m.get("image_summary", "")
            parts = ['<table class="image-meta">']
            if self._image_dir and filename:
                path = f"{self._image_dir}/{filename}"
                parts.append(f'<tr><td>image</td><td><img src="{path}" alt="{img_summary or "image"}" /></td></tr>')
            if img_summary:
                parts.append(f'<tr><td>image_summary</td><td>{img_summary}</td></tr>')
            parts.append('</table>')
            return "\n".join(parts)
        elif node.type == "table":
            return self._table_to_html_generic(node)
        elif node.text:
            return f"<p>{node.text}</p>"
        return ""

    def _slide_to_html(self, slide: OfficeContentNode) -> str:
        meta = slide.metadata or {}
        slide_num = meta.get("slideNumber", 1)
        slide_title = meta.get("slideTitle", "")
        slide_summary = meta.get("slide_summary", "")

        parts = [f'<div class="slide" data-slide="{slide_num}">']
        parts.append(f"<h2>Slide {slide_num}</h2>")

        # page-meta 테이블
        parts.append('<table class="page-meta">')
        if slide_title:
            parts.append(f'<tr><td>page_title</td><td>{slide_title}</td></tr>')
        slide_image = meta.get("slide_image", "")
        if self._image_dir and slide_image:
            path = f"{self._image_dir}/{slide_image}"
            parts.append(f'<tr><td>slide_image</td><td><img src="{path}" alt="Slide {slide_num}" /></td></tr>')
        if slide_summary:
            parts.append(f'<tr><td>page_summary</td><td>{slide_summary}</td></tr>')
        parts.append('</table>')

        if not slide.children:
            parts.append("</div>")
            return "\n".join(parts)

        for child in slide.children:
            if child.type == "paragraph":
                indent = child.metadata.get("indent_level", 0) if child.metadata else 0
                if indent > 0:
                    parts.append(f'<p style="margin-left:{indent * 2}em">{child.text}</p>')
                else:
                    fmt = child.formatting
                    text = child.text
                    if fmt:
                        if fmt.bold:
                            text = f"<strong>{text}</strong>"
                        if fmt.italic:
                            text = f"<em>{text}</em>"
                    parts.append(f"<p>{text}</p>")
            elif child.type == "table" and child.children:
                parts.append(self._rows_to_html_table(
                    [(r.metadata.get("row", ""), [c.text or "" for c in (r.children or [])], None)
                     for r in child.children],
                    child.metadata.get("cols", 0) if child.metadata else 0
                ))
            elif child.type == "image":
                m = child.metadata or {}
                filename = m.get("filename", "")
                img_summary = m.get("image_summary", "")
                bbox = m.get("bbox")
                parts.append('<table class="image-meta">')
                if self._image_dir and filename:
                    path = f"{self._image_dir}/{filename}"
                    parts.append(f'<tr><td>image</td><td><img src="{path}" alt="{img_summary or "image"}" /></td></tr>')
                if img_summary:
                    parts.append(f'<tr><td>image_summary</td><td>{img_summary}</td></tr>')
                if bbox:
                    parts.append(f'<tr><td>bbox</td><td>{bbox}</td></tr>')
                parts.append('</table>')
            elif child.type == "notes":
                parts.append(f'<h3>Note</h3>\n<blockquote class="notes">{child.text}</blockquote>')

        parts.append("</div>")
        return "\n".join(parts)
    
    def _sheet_to_html(self, sheet: OfficeContentNode) -> str:
        if not sheet.children:
            return ""
        
        meta = sheet.metadata or {}
        sheet_name = meta.get("sheetName", "Sheet")
        summary = meta.get("sheet_summary", "")
        
        # 시트 메타정보 헤더
        parts = [f'<div class="sheet" data-sheet-name="{sheet_name}">']
        parts.append(f"<h1>{sheet_name}</h1>")
        if summary:
            parts.append(f'<p class="sheet-summary">{summary}</p>')
        
        # children 순회
        rows = []
        max_cols = 0
        for child in sheet.children:
            if child.type == "row" and child.children:
                cells = []
                styles = []
                for cell in child.children:
                    cells.append(cell.text or "")
                    styles.append(cell.metadata.get("style") if cell.metadata else None)
                    colspan = cell.metadata.get("colspan", 1) if cell.metadata else 1
                    for _ in range(colspan - 1):
                        cells.append("")
                        styles.append(None)
                if len(cells) > max_cols:
                    max_cols = len(cells)
                row_num = child.metadata.get("row", "") if child.metadata else ""
                rows.append((row_num, cells, styles))
            elif child.type == "chart":
                # 테이블 앞에 쌓인 행이 있으면 먼저 출력
                if rows:
                    parts.append(self._rows_to_html_table(rows, max_cols))
                    rows, max_cols = [], 0
                ct = child.metadata.get("chartType", "Chart") if child.metadata else "Chart"
                title = child.metadata.get("title", "") if child.metadata else ""
                row_num = child.metadata.get("row", "") if child.metadata else ""
                parts.append(f'<div class="chart" data-row="{row_num}" data-type="{ct}">')
                parts.append(f"<strong>[{ct}]</strong> {title}")
                parts.append("</div>")
            elif child.type == "image":
                if rows:
                    parts.append(self._rows_to_html_table(rows, max_cols))
                    rows, max_cols = [], 0
                fmt = child.metadata.get("format", "png") if child.metadata else "png"
                row_num = child.metadata.get("row", "") if child.metadata else ""
                img_summary = child.metadata.get("image_summary", "") if child.metadata else ""
                filename = child.metadata.get("filename", "") if child.metadata else ""
                if self._image_dir and filename:
                    path = f"{self._image_dir}/{filename}"
                    parts.append(f'<div class="image" data-row="{row_num}">')
                    parts.append(f'<img src="{path}" alt="{img_summary or "image"}" />')
                    if img_summary:
                        parts.append(f'<p class="image-summary">{img_summary}</p>')
                    parts.append("</div>")
                else:
                    parts.append(f'<div class="image" data-row="{row_num}"><span>[Image]</span></div>')
        
        if rows:
            parts.append(self._rows_to_html_table(rows, max_cols))
        
        parts.append("</div>")
        return "\n".join(parts)
    
    def _table_to_html_generic(self, table: OfficeContentNode) -> str:
        """범용 table 노드 → HTML 변환"""
        if not table.children:
            return ""
        lines = ["<table>"]
        for i, row in enumerate(table.children):
            if row.type == "row" and row.children:
                tag = "th" if i == 0 else "td"
                cells = "".join(f"<{tag}>{c.text or ''}</{tag}>" for c in row.children)
                lines.append(f"<tr>{cells}</tr>")
        lines.append("</table>")
        summary = table.metadata.get("table_summary", "") if table.metadata else ""
        if summary:
            lines.append(f'<table class="table-meta"><tr><td>table_summary</td><td>{summary}</td></tr></table>')
        return "\n".join(lines)

    def _rows_to_html_table(self, rows: list, max_cols: int) -> str:
        lines = ["<table>"]
        for i, row_data in enumerate(rows):
            row_num = row_data[0]
            cells = row_data[1]
            styles = row_data[2] if len(row_data) > 2 else [None] * len(cells)
            padded = cells + [""] * (max_cols - len(cells))
            padded_styles = (styles or []) + [None] * (max_cols - len(styles or []))
            tag = "th" if i == 0 else "td"
            cells_html = []
            for j, c in enumerate(padded):
                s = padded_styles[j] if j < len(padded_styles) else None
                if s:
                    css = "; ".join(f"{k}: {v}" for k, v in s.items())
                    cells_html.append(f'<{tag} style="{css}">{c}</{tag}>')
                else:
                    cells_html.append(f"<{tag}>{c}</{tag}>")
            lines.append(f'<tr data-row="{row_num}">{"".join(cells_html)}</tr>')
        lines.append("</table>")
        return "\n".join(lines)
    
    def _node_to_markdown(self, node: OfficeContentNode, level: int = 0) -> str:
        if node.type == "heading":
            h_level = node.metadata.get("level", 1) if node.metadata else 1
            h_level = min(h_level + getattr(self, '_heading_offset', 0), 6)
            return f"{'#' * h_level} {node.text}\n"
        
        elif node.type == "paragraph":
            indent = node.metadata.get("indent_level", 0) if node.metadata else 0
            prefix = "&nbsp;" * (indent * 4) if indent > 0 else ""
            return f"{prefix}{node.text}\n"
        
        elif node.type == "list":
            indent_level = node.metadata.get("indent_level", 0) if node.metadata else 0
            indent = "  " * indent_level
            marker = "1." if node.metadata and node.metadata.get("listType") == "ordered" else "-"
            return f"{indent}{marker} {node.text}"
        
        elif node.type == "table":
            return self._table_to_markdown(node)
        
        elif node.type == "chart":
            chart_type = node.metadata.get("chartType", "Chart") if node.metadata else "Chart"
            title = node.metadata.get("title", "") if node.metadata else ""
            return f"**[{chart_type}]** {title}\n" if title else f"**[{chart_type}]**\n"
        
        elif node.type == "image":
            m = node.metadata or {}
            filename = m.get("filename", "")
            img_summary = m.get("image_summary", "")
            lines = ['<table class="image-meta">']
            if self._image_dir and filename:
                path = f"{self._image_dir}/{filename}"
                lines.append(f'<tr><td>image</td><td><img src="{path}" /></td></tr>')
            if img_summary:
                lines.append(f'<tr><td>image_summary</td><td>{img_summary}</td></tr>')
            lines.append('</table>\n')
            return "\n".join(lines)
        
        elif node.type == "sheet":
            meta = node.metadata or {}
            sheet_name = meta.get("sheetName", "Sheet")
            summary = meta.get("sheet_summary", "")
            
            lines = [f"## {sheet_name}\n"]
            if summary:
                lines.append(f"**Sheet Summary:** {summary}\n")
            lines.append(self._sheet_to_markdown(node))
            return "\n".join(lines)
        
        elif node.type == "slide":
            meta = node.metadata or {}
            slide_num = meta.get("slideNumber", 1)
            slide_title = meta.get("slideTitle", "")
            slide_summary = meta.get("slide_summary", "")

            lines = [f"### Slide {slide_num}\n"]
            if slide_title:
                lines.append(f"**{slide_title}**\n")
            slide_image = meta.get("slide_image", "")
            if self._image_dir and slide_image:
                path = f"{self._image_dir}/{slide_image}"
                lines.append(f"![Slide {slide_num}]({path})\n")
            if slide_summary:
                lines.append(f"**Slide Summary:** {slide_summary}\n")

            if node.children:
                for child in node.children:
                    if child.type == "paragraph":
                        indent = child.metadata.get("indent_level", 0) if child.metadata else 0
                        prefix = "  " * indent + "- " if indent > 0 else ""
                        text = child.text
                        fmt = child.formatting
                        if fmt:
                            if fmt.bold:
                                text = f"**{text}**"
                            if fmt.italic:
                                text = f"*{text}*"
                        lines.append(f"{prefix}{text}\n")
                    elif child.type == "table":
                        lines.append(self._table_to_markdown(child))
                    elif child.type == "image":
                        m = child.metadata or {}
                        filename = m.get("filename", "")
                        img_summary = m.get("image_summary", "")
                        lines.append('<table class="image-meta">')
                        if self._image_dir and filename:
                            path = f"{self._image_dir}/{filename}"
                            lines.append(f'<tr><td>image</td><td><img src="{path}" /></td></tr>')
                        if img_summary:
                            lines.append(f'<tr><td>image_summary</td><td>{img_summary}</td></tr>')
                        lines.append('</table>\n')
                    elif child.type == "notes":
                        lines.append(f"**Note:**\n\n> {child.text}\n")
                    else:
                        lines.append(self._node_to_markdown(child, level))
            return "\n".join(lines)
        
        elif node.type == "page":
            page_num = node.metadata.get("pageNumber", 1) if node.metadata else 1
            page_summary = node.metadata.get("page_summary", "") if node.metadata else ""
            lines = [f"### Page {page_num}\n"]
            if page_summary:
                lines.append(f"**Page Summary:** {page_summary}\n")
            if node.text:
                lines.append(f"{node.text}\n")
            if node.children:
                for child in node.children:
                    lines.append(self._node_to_markdown(child, level))
            return "\n".join(filter(None, lines))

        elif node.type == "section":
            meta = node.metadata or {}
            section_summary = meta.get("section_summary", "")
            lines = []
            if section_summary:
                lines.append(f"**Section Summary:** {section_summary}\n")
            if node.children:
                for child in node.children:
                    lines.append(self._node_to_markdown(child, level))
            return "\n".join(filter(None, lines))
        
        elif node.text:
            return node.text
        
        return ""
    
    def _table_to_markdown(self, table: OfficeContentNode) -> str:
        if not table.children:
            return ""
        
        rows = []
        max_cols = 0
        for row in table.children:
            if row.type == "row" and row.children:
                cells = [(cell.text or "").replace("|", "\\|").replace("\n", " ") for cell in row.children]
                if len(cells) > max_cols:
                    max_cols = len(cells)
                rows.append(cells)
        
        if not rows:
            return ""
        
        # 열 수 맞추기
        for r in rows:
            while len(r) < max_cols:
                r.append("")
        
        lines = []
        lines.append("")  # 테이블 앞 빈 줄
        lines.append("| " + " | ".join(rows[0]) + " |")
        lines.append("| " + " | ".join(["---"] * max_cols) + " |")
        for r in rows[1:]:
            lines.append("| " + " | ".join(r) + " |")
        
        # table_summary
        summary = table.metadata.get("table_summary", "") if table.metadata else ""
        if summary:
            lines.append("")
            lines.append(f'<table class="table-meta"><tr><td>table_summary</td><td>{summary}</td></tr></table>')
        
        return "\n".join(lines) + "\n"
    
    def _sheet_to_markdown(self, sheet: OfficeContentNode) -> str:
        if not sheet.children:
            return ""
        
        lines = []
        raw_rows = []
        
        def flush_rows():
            nonlocal raw_rows
            if not raw_rows:
                return
            # 비어있지 않은 셀 수 기준으로 그룹 분리
            from collections import Counter
            def nonempty(r): return sum(1 for c in r if c)
            counts = [nonempty(r) for r in raw_rows]
            mode_cols = Counter(counts).most_common(1)[0][0]
            threshold = max(mode_cols // 3, 1)

            groups = []
            current = [raw_rows[0]]
            for i, row in enumerate(raw_rows[1:], 1):
                prev_small = max(nonempty(r) for r in current) < threshold
                cur_small = counts[i] < threshold
                if prev_small != cur_small:
                    groups.append(current)
                    current = [row]
                else:
                    current.append(row)
            groups.append(current)
            
            for group in groups:
                max_cols = max(len(r) for r in group)
                padded = [r + [""] * (max_cols - len(r)) for r in group]
                header = "| " + " | ".join(padded[0]) + " |"
                separator = "| " + " | ".join(["---"] * max_cols) + " |"
                data = [("| " + " | ".join(r) + " |") for r in padded[1:]]
                lines.append("\n".join([header, separator] + data) + "\n")
            raw_rows = []
        
        for child in sheet.children:
            if child.type == "row" and child.children:
                cells = []
                for cell in child.children:
                    text = (cell.text or "").replace("|", "\\|").replace("\n", "<br>")
                    colspan = cell.metadata.get("colspan", 1) if cell.metadata else 1
                    cells.append(text)
                    for _ in range(colspan - 1):
                        cells.append("")
                raw_rows.append(cells)
            elif child.type == "chart":
                flush_rows()
                meta = child.metadata or {}
                ct = meta.get("chartType", "Chart")
                title = meta.get("title", "")
                row_num = meta.get("row", "")
                lines.append(f'<table class="chart-meta"><tr><td>type</td><td>{ct}</td><td>row</td><td>{row_num}</td></tr></table>')
                lines.append(f"**[{ct}]** {title}\n" if title else f"**[{ct}]**\n")
            elif child.type == "image":
                flush_rows()
                meta = child.metadata or {}
                row_num = meta.get("row", "")
                img_summary = meta.get("image_summary", "")
                filename = meta.get("filename", "")
                meta_parts = []
                if img_summary:
                    meta_parts.append(f'<tr><td>image_summary</td><td>{img_summary}</td></tr>')
                if meta_parts:
                    lines.append(f'<table class="image-meta">{"".join(meta_parts)}</table>')
                if self._image_dir and filename:
                    lines.append(f"![Image]({self._image_dir}/{filename})\n")
                else:
                    lines.append("![Image]\n")
        
        flush_rows()
        return "\n".join(lines)
