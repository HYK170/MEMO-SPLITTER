from __future__ import annotations

from copy import copy

from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


def is_row_empty(ws: Worksheet, row: int, min_col: int, max_col: int) -> bool:
    for col in range(min_col, max_col + 1):
        value = ws.cell(row=row, column=col).value
        if value not in (None, ""):
            return False
    return True


def copy_cell_style(source, target) -> None:
    if source.has_style:
        target.font = copy(source.font)
        target.fill = copy(source.fill)
        target.border = copy(source.border)
        target.alignment = copy(source.alignment)
        target.number_format = source.number_format
        target.protection = copy(source.protection)


def copy_row(
    source_ws: Worksheet,
    target_ws: Worksheet,
    source_row: int,
    target_row: int,
    min_col: int,
    max_col: int,
    strip_hyperlinks: bool = False,
) -> None:
    for col in range(min_col, max_col + 1):
        src_cell = source_ws.cell(row=source_row, column=col)
        if isinstance(src_cell, MergedCell):
            continue
        dst_cell = target_ws.cell(row=target_row, column=col)
        dst_cell.value = src_cell.value
        copy_cell_style(src_cell, dst_cell)
        if strip_hyperlinks:
            dst_cell.hyperlink = None
        elif src_cell.hyperlink is not None:
            dst_cell._hyperlink = copy(src_cell.hyperlink)


def copy_column_dimensions(source_ws: Worksheet, target_ws: Worksheet, min_col: int, max_col: int) -> None:
    for col in range(min_col, max_col + 1):
        letter = get_column_letter(col)
        source_dim = source_ws.column_dimensions.get(letter)
        if source_dim is not None and source_dim.width is not None:
            target_ws.column_dimensions[letter].width = source_dim.width


def copy_merged_cells_for_rows(
    source_ws: Worksheet,
    target_ws: Worksheet,
    rows: list[int],
) -> None:
    row_set = set(rows)
    for merged_range in source_ws.merged_cells.ranges:
        if merged_range.min_row in row_set and merged_range.max_row in row_set:
            target_ws.merge_cells(str(merged_range))


def create_split_workbook(
    source_ws: Worksheet,
    header_row: int,
    data_row: int,
    sheet_title: str | None = None,
) -> Workbook:
    wb = Workbook()
    target_ws = wb.active
    target_ws.title = sheet_title or source_ws.title

    min_col = source_ws.min_column or 1
    max_col = source_ws.max_column or 1

    copy_row(source_ws, target_ws, header_row, header_row, min_col, max_col, strip_hyperlinks=True)
    copy_row(source_ws, target_ws, data_row, data_row, min_col, max_col, strip_hyperlinks=True)
    copy_column_dimensions(source_ws, target_ws, min_col, max_col)
    copy_merged_cells_for_rows(source_ws, target_ws, [header_row, data_row])

    return wb
