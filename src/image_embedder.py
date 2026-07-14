from __future__ import annotations

from pathlib import Path

from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.worksheet.worksheet import Worksheet

from src.sheet_copier import OUTPUT_DATA_ROW

# 셀에 맞추기 위한 최대 표시 크기 (픽셀)
MAX_DISPLAY_WIDTH = 160
MAX_DISPLAY_HEIGHT = 120


def embed_images_in_column(
    ws: Worksheet,
    image_paths: list[Path],
    column: int,
    row: int = OUTPUT_DATA_ROW,
) -> tuple[int, list[str]]:
    """이미지 파일을 지정 열 셀에 임베드한다.

    Returns:
        (성공 개수, 실패 메시지 목록)
    """
    if column < 1 or not image_paths:
        return 0, []

    added = 0
    failures: list[str] = []
    for offset, path in enumerate(image_paths):
        try:
            image = Image(str(path))
        except Exception as exc:
            failures.append(f"{path.name}: 열기 실패 ({exc})")
            continue

        try:
            width = int(image.width or MAX_DISPLAY_WIDTH)
            height = int(image.height or MAX_DISPLAY_HEIGHT)
            scale = min(MAX_DISPLAY_WIDTH / width, MAX_DISPLAY_HEIGHT / height, 1.0)
            display_w = max(1, int(width * scale))
            display_h = max(1, int(height * scale))
            image.width = display_w
            image.height = display_h

            marker = AnchorMarker(
                col=column - 1,
                row=row - 1,
                colOff=0,
                rowOff=pixels_to_EMU(offset * (display_h + 4)),
            )
            image.anchor = OneCellAnchor(
                _from=marker,
                ext=XDRPositiveSize2D(pixels_to_EMU(display_w), pixels_to_EMU(display_h)),
            )

            data = path.read_bytes()
            image._cached_bytes = data  # noqa: SLF001
            image._data = lambda data=data: data  # noqa: SLF001
            ws.add_image(image)
            added += 1
        except Exception as exc:
            failures.append(f"{path.name}: 삽입 실패 ({exc})")

    if added:
        # 이미지가 보이도록 행 높이 약간 확보
        current = ws.row_dimensions[row].height
        needed = min(180, 20 + added * 90)
        if current is None or current < needed:
            ws.row_dimensions[row].height = needed

    return added, failures
