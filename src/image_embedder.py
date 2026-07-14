from __future__ import annotations

import sys
from pathlib import Path

from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import (
    EMU_to_pixels,
    cm_to_EMU,
    pixels_to_EMU,
    pixels_to_points,
)
from openpyxl.worksheet.worksheet import Worksheet

from src.sheet_copier import OUTPUT_DATA_ROW

# 표시 최대 박스 (비율 유지, 박스보다 작으면 확대하지 않음)
MAX_WIDTH_CM = 2.54
MAX_HEIGHT_CM = 2.36
GAP_CM = 0.1  # 여러 장 세로 쌓을 때 간격
ROW_PADDING_CM = 0.15  # 행 높이 여유


def ensure_pillow() -> str:
    """openpyxl Image가 Pillow를 쓰도록 보장한다. 성공 시 버전 문자열 반환."""
    try:
        from PIL import Image as PILImage
    except ImportError as exc:
        raise RuntimeError(
            "이미지 임베드에 Pillow가 필요합니다. "
            f"현재 Python: {sys.executable}\n"
            '설치: python -m pip install "Pillow>=10.0.0"'
        ) from exc

    # openpyxl이 Pillow 없이 먼저 import되면 PILImage=False로 고정된다.
    import openpyxl.drawing.image as oxl_image

    if not oxl_image.PILImage:
        oxl_image.PILImage = PILImage

    version = getattr(PILImage, "__version__", None)
    if version is None:
        import PIL

        version = getattr(PIL, "__version__", "unknown")
    return str(version)


def _cm_to_pixels(cm: float) -> float:
    return float(EMU_to_pixels(cm_to_EMU(cm)))


def _fit_display_size(width: float, height: float) -> tuple[int, int]:
    """최대 박스 안에 비율 유지로 맞춘다. 작은 이미지는 키우지 않는다."""
    max_w = _cm_to_pixels(MAX_WIDTH_CM)
    max_h = _cm_to_pixels(MAX_HEIGHT_CM)
    if width <= 0 or height <= 0:
        return max(1, int(max_w)), max(1, int(max_h))
    scale = min(max_w / width, max_h / height, 1.0)
    return max(1, int(width * scale)), max(1, int(height * scale))


def embed_images_in_column(
    ws: Worksheet,
    image_paths: list[Path],
    column: int,
    row: int = OUTPUT_DATA_ROW,
) -> tuple[int, list[str]]:
    """이미지 파일을 지정 열 셀에 임베드한다.

    여러 장이면 같은 셀 기준 세로로 쌓고, 2행 높이를 총 높이에 맞춘다.

    Returns:
        (성공 개수, 실패 메시지 목록)
    """
    if column < 1 or not image_paths:
        return 0, []

    try:
        ensure_pillow()
    except RuntimeError as exc:
        return 0, [str(exc)]

    added = 0
    failures: list[str] = []
    gap_px = _cm_to_pixels(GAP_CM)
    cursor_y = 0.0
    total_height_px = 0.0

    for path in image_paths:
        try:
            image = Image(str(path))
        except Exception as exc:
            failures.append(f"{path.name}: 열기 실패 ({exc})")
            continue

        try:
            width = float(image.width or _cm_to_pixels(MAX_WIDTH_CM))
            height = float(image.height or _cm_to_pixels(MAX_HEIGHT_CM))
            display_w, display_h = _fit_display_size(width, height)
            image.width = display_w
            image.height = display_h

            marker = AnchorMarker(
                col=column - 1,
                row=row - 1,
                colOff=0,
                rowOff=int(pixels_to_EMU(cursor_y)),
            )
            image.anchor = OneCellAnchor(
                _from=marker,
                ext=XDRPositiveSize2D(
                    int(pixels_to_EMU(display_w)),
                    int(pixels_to_EMU(display_h)),
                ),
            )

            data = path.read_bytes()
            image._cached_bytes = data  # noqa: SLF001
            image._data = lambda data=data: data  # noqa: SLF001
            ws.add_image(image)

            total_height_px = cursor_y + display_h
            cursor_y += display_h + gap_px
            added += 1
        except Exception as exc:
            failures.append(f"{path.name}: 삽입 실패 ({exc})")

    if added:
        needed = pixels_to_points(total_height_px + _cm_to_pixels(ROW_PADDING_CM))
        # Excel 행 높이 상한은 409pt
        needed = min(409.0, max(15.0, float(needed)))
        current = ws.row_dimensions[row].height
        if current is None or current < needed:
            ws.row_dimensions[row].height = needed

    return added, failures
