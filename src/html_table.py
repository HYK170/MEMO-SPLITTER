from __future__ import annotations

from dataclasses import dataclass, field
from html import escape
from html.parser import HTMLParser
from typing import Literal

VOID_ELEMENTS = frozenset(
    {
        "area",
        "base",
        "br",
        "col",
        "embed",
        "hr",
        "img",
        "input",
        "link",
        "meta",
        "param",
        "source",
        "track",
        "wbr",
    }
)


@dataclass
class Cell:
    text: str
    html: str


@dataclass
class ParsedTable:
    headers: list[str]
    header_cells_html: list[str]
    rows: list[list[Cell]] = field(default_factory=list)


def _attrs_to_str(attrs: list[tuple[str, str | None]]) -> str:
    parts: list[str] = []
    for key, value in attrs:
        if value is None:
            parts.append(f" {key}")
        else:
            parts.append(f' {key}="{escape(value, quote=True)}"')
    return "".join(parts)


def _format_start_tag(tag: str, attrs: list[tuple[str, str | None]]) -> str:
    attr_str = _attrs_to_str(attrs)
    return f"<{tag}{attr_str}>"


class _FirstTableParser(HTMLParser):
    """첫 번째 table을 파싱한다. </td></tr> 생략(HTML optional end tags)도 허용한다."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.depth = 0
        self.in_table = False
        self.done = False

        self.in_thead = False
        self.in_tr = False
        self.in_cell = False
        self.cell_tag: Literal["th", "td"] | None = None

        self.current_cell_text: list[str] = []
        self.current_cell_html: list[str] = []
        self.current_row_cells: list[Cell] = []
        self.current_row_has_th = False

        self.thead_rows: list[list[Cell]] = []
        self.body_rows: list[list[Cell]] = []
        self.all_rows: list[tuple[list[Cell], bool]] = []

    def _finish_cell(self) -> None:
        if not self.in_cell:
            return
        cell = Cell(
            text="".join(self.current_cell_text).strip(),
            html="".join(self.current_cell_html),
        )
        self.current_row_cells.append(cell)
        self.in_cell = False
        self.cell_tag = None
        self.current_cell_text = []
        self.current_cell_html = []

    def _finish_row(self) -> None:
        self._finish_cell()
        if not self.in_tr:
            return
        row = self.current_row_cells
        if self.in_thead:
            self.thead_rows.append(row)
        else:
            self.body_rows.append(row)
        self.all_rows.append((row, self.current_row_has_th))
        self.in_tr = False
        self.current_row_cells = []
        self.current_row_has_th = False

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if self.done:
            return

        if tag == "table":
            if not self.in_table:
                self.in_table = True
                self.depth = 1
            else:
                self.depth += 1
                if self.in_cell:
                    self.current_cell_html.append(_format_start_tag(tag, attrs))
            return

        if not self.in_table:
            return

        if self.depth > 1:
            if self.in_cell:
                self.current_cell_html.append(_format_start_tag(tag, attrs))
                if tag == "br":
                    self.current_cell_text.append("\n")
            return

        if tag == "thead":
            self._finish_row()
            self.in_thead = True
            return

        if tag in ("tbody", "tfoot"):
            self._finish_row()
            self.in_thead = False
            return

        if tag == "tr":
            self._finish_row()
            self.in_tr = True
            self.current_row_cells = []
            self.current_row_has_th = False
            return

        if tag in ("th", "td"):
            if not self.in_tr:
                self.in_tr = True
                self.current_row_cells = []
                self.current_row_has_th = False
            self._finish_cell()
            self.in_cell = True
            self.cell_tag = tag  # type: ignore[assignment]
            self.current_cell_text = []
            self.current_cell_html = []
            if tag == "th":
                self.current_row_has_th = True
            return

        if self.in_cell:
            self.current_cell_html.append(_format_start_tag(tag, attrs))
            if tag == "br":
                self.current_cell_text.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if self.done or not self.in_table:
            return

        if tag == "table":
            if self.depth == 1:
                self._finish_row()
            self.depth -= 1
            if self.depth == 0:
                self.in_table = False
                self.done = True
            elif self.in_cell:
                self.current_cell_html.append(f"</{tag}>")
            return

        if self.depth > 1:
            if self.in_cell and tag not in VOID_ELEMENTS:
                self.current_cell_html.append(f"</{tag}>")
            return

        if tag == "thead":
            self._finish_row()
            self.in_thead = False
            return

        if tag in ("tbody", "tfoot"):
            self._finish_row()
            return

        if tag in ("th", "td"):
            self._finish_cell()
            return

        if tag == "tr":
            self._finish_row()
            return

        if self.in_cell and tag not in VOID_ELEMENTS:
            self.current_cell_html.append(f"</{tag}>")

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if self.done or not self.in_table:
            return
        if self.in_cell:
            self.current_cell_html.append(_format_start_tag(tag, attrs))
            if tag == "br":
                self.current_cell_text.append("\n")

    def handle_data(self, data: str) -> None:
        if self.done or not self.in_table or not self.in_cell:
            return
        self.current_cell_text.append(data)
        self.current_cell_html.append(escape(data))


def parse_first_table(html_text: str) -> ParsedTable:
    parser = _FirstTableParser()
    parser.feed(html_text)
    parser.close()
    if parser.in_table:
        parser._finish_row()

    if not parser.all_rows and not parser.thead_rows:
        raise ValueError("HTML에서 <table>을 찾을 수 없습니다.")

    header_row: list[Cell] | None = None
    data_rows: list[list[Cell]] = []

    if parser.thead_rows:
        header_row = parser.thead_rows[0]
        data_rows = list(parser.body_rows)
    else:
        header_idx = None
        for idx, (row, has_th) in enumerate(parser.all_rows):
            if has_th:
                header_idx = idx
                break
        if header_idx is None:
            # Excel 등: 첫 행이 <td>만 있는 헤더인 경우
            if not parser.all_rows:
                raise ValueError("테이블 헤더 행(<th> 또는 <thead>)을 찾을 수 없습니다.")
            header_idx = 0
        header_row = parser.all_rows[header_idx][0]
        data_rows = [row for row, _ in parser.all_rows[header_idx + 1 :]]

    return ParsedTable(
        headers=[cell.text for cell in header_row],
        header_cells_html=[cell.html for cell in header_row],
        rows=data_rows,
    )


def is_row_empty(cells: list[Cell]) -> bool:
    return all(not cell.text.strip() for cell in cells)


def build_split_html(header_cells_html: list[str], data_cells_html: list[str]) -> str:
    header_tds = "".join(f"<th>{html}</th>" for html in header_cells_html)
    data_tds = "".join(f"<td>{html}</td>" for html in data_cells_html)
    return (
        "<!DOCTYPE html>\n"
        '<html lang="ko">\n'
        "<head>\n"
        '<meta charset="utf-8">\n'
        "<title>split</title>\n"
        "</head>\n"
        "<body>\n"
        "<table>\n"
        f"<tr>{header_tds}</tr>\n"
        f"<tr>{data_tds}</tr>\n"
        "</table>\n"
        "</body>\n"
        "</html>\n"
    )
