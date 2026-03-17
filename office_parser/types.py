from dataclasses import dataclass, field
from typing import Optional, List, Dict, Literal, Union
from datetime import datetime

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
        
        return "\n".join(filter(None, parts))
    
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
