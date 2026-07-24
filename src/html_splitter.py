from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from src.filename_builder import (
    build_attachments_folder_name,
    build_html_filename,
    build_row_folder_name,
)
from src.html_table import Cell, build_split_html, is_row_empty, parse_first_table
from src.multimedia_copier import SAVED_NAME_COLUMN, copy_multimedia_for_row
from src.splitter import REQUIRED_COLUMNS, SplitResult, build_memo_output_root

LogCallback = Callable[[str], None]
ProgressCallback = Callable[[int, int], None]

_HEADER_WHITESPACE = re.compile(r"\s+", re.UNICODE)


@dataclass
class HtmlSplitConfig:
    input_path: Path
    multimedia_root: Path


def validate_html_config(config: HtmlSplitConfig) -> None:
    if not config.input_path.is_file():
        raise FileNotFoundError(f"INPUT HTML을 찾을 수 없습니다: {config.input_path}")
    if config.input_path.suffix.lower() not in {".html", ".htm"}:
        raise ValueError("INPUT 파일은 .html 또는 .htm 확장자여야 합니다.")
    if not config.multimedia_root.is_dir():
        raise FileNotFoundError(f"Multimedia 폴더를 찾을 수 없습니다: {config.multimedia_root}")


def split_html(
    config: HtmlSplitConfig,
    on_log: LogCallback | None = None,
    on_progress: ProgressCallback | None = None,
) -> SplitResult:
    validate_html_config(config)

    def log(message: str) -> None:
        if on_log:
            on_log(message)

    output_root = build_memo_output_root(config.input_path)
    output_root.mkdir(parents=True, exist_ok=False)
    log(f"OUTPUT: {output_root}")
    log(f"Python: {sys.executable}")

    html_text = config.input_path.read_text(encoding="utf-8-sig")
    table = parse_first_table(html_text)
    column_map = _build_column_map(table.headers)
    _ensure_required_columns(column_map)

    data_rows = [row for row in table.rows if not is_row_empty(row)]
    result = SplitResult(output_root=output_root)
    result.rows_skipped = len(table.rows) - len(data_rows)
    total = len(data_rows)
    base_name = config.input_path.stem
    header_html = table.header_cells_html
    col_count = len(header_html)

    for row_index, cells in enumerate(data_rows, start=1):
        if on_progress:
            on_progress(row_index, total)

        folder_name = build_row_folder_name(base_name, row_index)
        row_folder = output_root / folder_name

        try:
            row_folder.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            message = f"행 {row_index}: 폴더 생성 실패 ({row_folder}) - {exc}"
            result.row_errors.append(message)
            log(message)
            continue

        padded = _pad_cells(cells, col_count)
        body_value = padded[column_map["본문"]].text
        saved_names = padded[column_map[SAVED_NAME_COLUMN]].text
        html_name = build_html_filename(base_name, row_index, body_value)
        html_path = row_folder / html_name

        attachments_folder = row_folder / build_attachments_folder_name(base_name, row_index)
        copy_result = copy_multimedia_for_row(
            saved_names,
            config.multimedia_root,
            attachments_folder,
        )
        result.attachments_copied += len(copy_result.copied)
        result.attachment_skips.extend(copy_result.skipped)

        try:
            data_html = [cell.html for cell in padded]
            content = build_split_html(header_html, data_html)
            html_path.write_text(content, encoding="utf-8")
        except OSError as exc:
            message = f"행 {row_index}: HTML 저장 실패 - {exc}"
            result.row_errors.append(message)
            log(message)
            continue

        result.folders_created += 1
        log(f"행 {row_index}: {row_folder.name} 생성 ({html_name})")
        for name in copy_result.copied:
            log(f"  첨부 복사: {name}")
        for skipped in copy_result.skipped:
            log(f"  첨부 스킵: {skipped}")

    return result


def _normalize_header(value: object) -> str:
    return _HEADER_WHITESPACE.sub(" ", str(value).strip())


def _build_column_map(headers: list[str]) -> dict[str, int]:
    column_map: dict[str, int] = {}
    for idx, header in enumerate(headers):
        if not header:
            continue
        column_map[_normalize_header(header)] = idx
    return column_map


def _find_column(column_map: dict[str, int], name: str) -> int | None:
    normalized = _normalize_header(name)
    if normalized in column_map:
        return column_map[normalized]
    compact = normalized.replace(" ", "")
    for key, col in column_map.items():
        if key.replace(" ", "") == compact:
            return col
    return None


def _ensure_required_columns(column_map: dict[str, int]) -> None:
    missing = [name for name in REQUIRED_COLUMNS if _find_column(column_map, name) is None]
    if missing:
        raise ValueError(f"필수 컬럼이 없습니다: {', '.join(missing)}")
    for name in REQUIRED_COLUMNS:
        col = _find_column(column_map, name)
        if col is not None:
            column_map[name] = col


def _pad_cells(cells: list[Cell], col_count: int) -> list[Cell]:
    padded = list(cells)
    while len(padded) < col_count:
        padded.append(Cell(text="", html=""))
    return padded[:col_count]
