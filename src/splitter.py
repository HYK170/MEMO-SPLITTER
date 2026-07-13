from __future__ import annotations

import gc
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

from openpyxl import load_workbook

from src.attachment_copier import copy_attachments_for_row
from src.filename_builder import build_row_folder_name, build_xlsx_filename
from src.drawing_image_loader import build_images_by_row
from src.drawing_transplant import count_all_source_anchors, transplant_drawings_for_row
from src.image_handler import (
    add_row_images,
    collect_objects_for_row,
    flatten_unique_objects,
    index_images_by_row,
    strip_hyperlinks_from_xlsx_file,
)
from src.sheet_copier import create_split_workbook, is_row_empty
from src.xlsx_hyperlink_index import XlsxHyperlinkIndex

LogCallback = Callable[[str], None]
ProgressCallback = Callable[[int, int], None]

REQUIRED_COLUMNS = ("App", "본문")


@dataclass
class SplitResult:
    folders_created: int = 0
    attachments_copied: int = 0
    rows_skipped: int = 0
    attachment_skips: list[str] = field(default_factory=list)
    row_errors: list[str] = field(default_factory=list)


@dataclass
class SplitConfig:
    input_path: Path
    output_root: Path
    sheet_name: str
    header_row: int


def validate_config(config: SplitConfig) -> None:
    if not config.input_path.is_file():
        raise FileNotFoundError(f"INPUT XLSX를 찾을 수 없습니다: {config.input_path}")
    if config.input_path.suffix.lower() != ".xlsx":
        raise ValueError("INPUT 파일은 .xlsx 확장자여야 합니다.")
    if not config.output_root.is_dir():
        raise FileNotFoundError(f"OUTPUT 폴더를 찾을 수 없습니다: {config.output_root}")
    if config.header_row < 1:
        raise ValueError("HEADER ROW는 1 이상이어야 합니다.")


def split_workbook(
    config: SplitConfig,
    on_log: LogCallback | None = None,
    on_progress: ProgressCallback | None = None,
) -> SplitResult:
    validate_config(config)

    def log(message: str) -> None:
        if on_log:
            on_log(message)

    wb = load_workbook(config.input_path, data_only=False)
    try:
        if config.sheet_name not in wb.sheetnames:
            raise ValueError(f"시트를 찾을 수 없습니다: {config.sheet_name}")

        ws = wb[config.sheet_name]
        if config.header_row > (ws.max_row or 0):
            raise ValueError("HEADER ROW가 시트 범위를 벗어났습니다.")

        column_map = _build_column_map(ws, config.header_row)
        _ensure_required_columns(column_map)

        from src.sheet_path import get_effective_max_row, resolve_sheet_path
        import zipfile

        with zipfile.ZipFile(config.input_path, "r") as archive:
            sheet_path = resolve_sheet_path(archive, ws, wb)
            effective_max_row = get_effective_max_row(ws, archive, sheet_path)

        images_by_row = build_images_by_row(config.input_path, ws, wb, config.header_row)
        all_objects = flatten_unique_objects(images_by_row)
        openpyxl_objects = list(getattr(ws, "_images", []))
        source_anchor_total = count_all_source_anchors(config.input_path, ws, wb)

        openpyxl_count = sum(
            len(v) for v in index_images_by_row(ws, config.header_row, effective_max_row).values()
        )
        drawing_count = len(all_objects) - len(openpyxl_objects)
        if drawing_count < 0:
            drawing_count = sum(len(v) for v in images_by_row.values())
        log(
            f"시각 객체 {len(all_objects)}개 인식 / drawing anchor {source_anchor_total}개 "
            f"(openpyxl {openpyxl_count}, drawing-loader {drawing_count})"
        )
        if len(all_objects) == 0 and source_anchor_total == 0:
            from src.cell_image_loader import diagnose_image_sources

            log("객체 진단:")
            for line in diagnose_image_sources(config.input_path, ws, wb):
                log(f"  {line}")
        image_index = XlsxHyperlinkIndex(config.input_path, config.header_row)
        base_name = config.input_path.stem
        min_col = ws.min_column or 1
        max_col = ws.max_column or 1
        first_data_row = config.header_row + 1
        last_data_row = effective_max_row

        data_rows = [
            row
            for row in range(first_data_row, last_data_row + 1)
            if not is_row_empty(ws, row, min_col, max_col)
        ]
        result = SplitResult()
        result.rows_skipped = (last_data_row - first_data_row + 1) - len(data_rows)
        total = len(data_rows)
        row_index = 0

        for data_row in data_rows:
            row_index += 1
            if on_progress:
                on_progress(row_index, total)

            folder_name = build_row_folder_name(base_name, row_index)
            row_folder = config.output_root / folder_name

            try:
                row_folder.mkdir(parents=True, exist_ok=True)
            except OSError as exc:
                message = f"행 {data_row}: 폴더 생성 실패 ({row_folder}) - {exc}"
                result.row_errors.append(message)
                log(message)
                continue

            app_value = ws.cell(row=data_row, column=column_map["App"]).value
            body_value = ws.cell(row=data_row, column=column_map["본문"]).value
            xlsx_name = build_xlsx_filename(base_name, app_value, row_index, body_value)
            xlsx_path = row_folder / xlsx_name

            try:
                out_wb = create_split_workbook(ws, config.header_row, data_row, config.sheet_name)
                out_ws = out_wb.active
                row_openpyxl = collect_objects_for_row(
                    openpyxl_objects,
                    data_row,
                    config.header_row,
                )
                if row_openpyxl:
                    add_row_images(out_ws, row_openpyxl)
                out_wb.save(xlsx_path)
                out_wb.close()
                del out_wb

                try:
                    transplanted = transplant_drawings_for_row(
                        config.input_path,
                        xlsx_path,
                        ws,
                        wb,
                        data_row,
                        config.header_row,
                    )
                except Exception as exc:
                    message = f"행 {data_row}: drawing 객체 이식 실패 - {exc}"
                    result.row_errors.append(message)
                    log(message)
                    transplanted = 0

                if transplanted:
                    log(f"  drawing 객체 {transplanted}개 이식")
                elif row_openpyxl:
                    log(f"  openpyxl 객체 {len(row_openpyxl)}개 포함")

                strip_hyperlinks_from_xlsx_file(xlsx_path)
            except Exception as exc:
                message = f"행 {data_row}: XLSX 저장 실패 - {exc}"
                result.row_errors.append(message)
                log(message)
                continue

            attachment_result = copy_attachments_for_row(
                ws,
                data_row,
                config.input_path,
                row_folder,
                image_index,
                config.header_row,
            )
            result.folders_created += 1
            result.attachments_copied += len(attachment_result.copied)
            result.attachment_skips.extend(attachment_result.skipped)

            log(f"행 {data_row}: {row_folder.name} 생성 ({xlsx_name})")
            for copied_name in attachment_result.copied:
                log(f"  첨부 복사: {copied_name}")
            for skipped in attachment_result.skipped:
                log(f"  첨부 스킵: {skipped}")

            if row_index % 100 == 0:
                gc.collect()

        return result
    finally:
        wb.close()
        gc.collect()


def _build_column_map(ws, header_row: int) -> dict[str, int]:
    column_map: dict[str, int] = {}
    for col in range(ws.min_column or 1, (ws.max_column or 1) + 1):
        value = ws.cell(row=header_row, column=col).value
        if value is None:
            continue
        column_map[str(value).strip()] = col
    return column_map


def _ensure_required_columns(column_map: dict[str, int]) -> None:
    missing = [name for name in REQUIRED_COLUMNS if name not in column_map]
    if missing:
        raise ValueError(f"필수 컬럼이 없습니다: {', '.join(missing)}")
