from __future__ import annotations

from pathlib import Path

from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.worksheet.worksheet import Worksheet

from src.hyperlink_resolver import resolve_local_path
from src.xlsx_hyperlink_index import XlsxHyperlinkIndex

IMAGE_EXTENSIONS = {
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".bmp",
    ".webp",
    ".tif",
    ".tiff",
}


def is_image_file(path: Path) -> bool:
    return path.suffix.lower() in IMAGE_EXTENSIONS


def collect_image_paths_for_row(
    image_index: XlsxHyperlinkIndex,
    data_row: int,
    base_dir: Path,
) -> tuple[list[Path], list[str]]:
    """이미지 하이퍼링크에서 로컬 이미지 파일 경로를 수집한다.

    Returns:
        (found_image_paths, skip_messages)
    """
    found: list[Path] = []
    skipped: list[str] = []
    seen: set[str] = set()

    for target in image_index.targets_for_row(data_row):
        local_path = resolve_local_path(target, base_dir)
        if local_path is None:
            skipped.append(f"이미지 링크 해석 실패: {target}")
            continue
        if not local_path.is_file():
            skipped.append(f"이미지 파일 없음: {local_path}")
            continue
        if not is_image_file(local_path):
            # 첨부는 attachment_copier가 처리. 여기선 이미지 임베드만.
            continue
        key = str(local_path.resolve()).lower()
        if key in seen:
            continue
        seen.add(key)
        found.append(local_path)

    return found, skipped


def add_images_from_paths(
    target_ws: Worksheet,
    image_paths: list[Path],
    *,
    anchor_row: int,
    anchor_col: int = 1,
) -> int:
    """로컬 이미지 파일을 새 워크시트에 임베드한다.

    openpyxl/Pillow가 읽을 수 있는 형식만 성공한다.
    """
    added = 0
    for offset, path in enumerate(image_paths):
        try:
            image = Image(str(path))
        except Exception:
            continue
        # 동일 셀 근처에 세로로 살짝 어긋나게 배치
        marker = AnchorMarker(
            col=max(0, anchor_col - 1),
            row=max(0, anchor_row - 1),
            colOff=0,
            rowOff=pixels_to_EMU(offset * 8),
        )
        width = int(image.width or 120)
        height = int(image.height or 120)
        image.anchor = OneCellAnchor(
            _from=marker,
            ext=XDRPositiveSize2D(pixels_to_EMU(width), pixels_to_EMU(height)),
        )
        # 파일 핸들이 닫혀도 save 가능하도록 바이트 캐시
        data = path.read_bytes()
        image._cached_bytes = data  # noqa: SLF001
        image._data = lambda data=data: data  # noqa: SLF001
        target_ws.add_image(image)
        added += 1
    return added
