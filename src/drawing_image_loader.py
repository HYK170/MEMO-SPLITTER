from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET

from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor, TwoCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.worksheet.worksheet import Worksheet

from src.cell_image_loader import load_cell_images
from src.image_handler import get_assigned_rows


def _iter_local_name(root: ET.Element, local_name: str) -> list[ET.Element]:
    suffix = f"}}{local_name}"
    return [el for el in root.iter() if el.tag.endswith(suffix) or el.tag == local_name]

NS = {
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
VML_NS = {
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "x": "urn:schemas-microsoft-com:office:excel",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
REL_ID_ATTR = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
REL_EMBED_ATTR = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
DEFAULT_ROW_EMU = 190_500


@dataclass
class AnchorInfo:
    from_row: int
    to_row: int
    from_col: int
    col_off: int = 0
    row_off: int = 0
    to_col: int | None = None
    cx: int = 0
    cy: int = 0
    is_one_cell: bool = True


def load_drawing_images(
    xlsx_path: Path,
    ws: Worksheet,
    wb=None,
    header_row: int = 1,
) -> dict[int, list[Image]]:
    images_by_row: dict[int, list[Image]] = {}
    max_row = ws.max_row or 1

    with zipfile.ZipFile(xlsx_path, "r") as archive:
        sheet_path = _resolve_sheet_path(archive, ws, wb)
        if not sheet_path:
            return {}

        drawing_paths = _get_all_related_paths(archive, sheet_path, "drawing")
        if not drawing_paths:
            from src.sheet_path import discover_drawing_paths

            drawing_paths = discover_drawing_paths(archive, sheet_path, ws, wb)
        for drawing_path in drawing_paths:
            drawing_rels = _get_rels_map(archive, drawing_path)
            _load_from_drawing_xml(
                archive, drawing_path, drawing_rels, images_by_row, max_row, header_row
            )

        vml_paths = _get_all_related_paths(archive, sheet_path, "vmlDrawing")
        if not vml_paths:
            from src.sheet_path import discover_vml_paths

            vml_paths = discover_vml_paths(archive, sheet_path, ws, wb)
        for vml_path in vml_paths:
            vml_rels = _get_rels_map(archive, vml_path)
            _load_from_vml(archive, vml_path, vml_rels, images_by_row, max_row, header_row)

    return images_by_row


def _resolve_sheet_path(archive: zipfile.ZipFile, ws: Worksheet, wb=None) -> str | None:
    from src.sheet_path import resolve_sheet_path

    return resolve_sheet_path(archive, ws, wb)


def merge_images_by_row(
    openpyxl_index: dict[int, list[Image]],
    drawing_index: dict[int, list[Image]],
    max_row: int,
) -> dict[int, list[Image]]:
    merged: dict[int, list[Image]] = {}
    seen: set[str] = set()

    for row in range(1, max_row + 1):
        bucket: list[Image] = []
        for source in (drawing_index, openpyxl_index):
            for image in source.get(row, []):
                key = _image_dedupe_key(image)
                if key in seen:
                    continue
                seen.add(key)
                bucket.append(image)
        if bucket:
            merged[row] = bucket
    return merged


def build_images_by_row(
    xlsx_path: Path,
    ws: Worksheet,
    wb=None,
    header_row: int = 1,
) -> dict[int, list[Image]]:
    from src.image_handler import index_images_by_row

    max_row = ws.max_row or 1
    openpyxl_index = index_images_by_row(ws, header_row)
    drawing_index = load_drawing_images(xlsx_path, ws, wb, header_row)
    cell_index = load_cell_images(xlsx_path, ws, wb)
    merged = merge_images_by_row(openpyxl_index, drawing_index, max_row)
    return merge_images_by_row(merged, cell_index, max_row)


def _image_dedupe_key(image: Image) -> str:
    try:
        data = image._data()
        digest = hash(data)
    except Exception:
        digest = id(image)
    row_range = None
    anchor = getattr(image, "anchor", None)
    if anchor is not None:
        from_anchor = getattr(anchor, "_from", None)
        if from_anchor is not None:
            row_range = (int(from_anchor.row), int(getattr(from_anchor, "col", 0)))
    return f"{digest}:{row_range}"


def _normalize_zip_path(target: str) -> str:
    cleaned = target.replace("\\", "/").lstrip("/")
    if cleaned.startswith("xl/"):
        return cleaned
    if cleaned.startswith("drawings/"):
        return f"xl/{cleaned}"
    if cleaned.startswith("media/"):
        return f"xl/{cleaned}"
    return cleaned.replace("../", "")


def _resolve_target(base_path: str, target: str) -> str:
    normalized = _normalize_zip_path(target)
    if normalized.startswith("xl/"):
        return normalized
    base = PurePosixPath(base_path).parent
    return _normalize_zip_path(str(base / target))


def _get_all_related_paths(archive: zipfile.ZipFile, sheet_path: str, rel_keyword: str) -> list[str]:
    from src.sheet_path import find_sheet_related_paths

    return find_sheet_related_paths(archive, sheet_path, rel_keyword)


def _get_related_path(archive: zipfile.ZipFile, sheet_path: str, rel_keyword: str) -> str | None:
    paths = _get_all_related_paths(archive, sheet_path, rel_keyword)
    return paths[0] if paths else None


def _get_rels_map(archive: zipfile.ZipFile, resource_path: str) -> dict[str, str]:
    from src.sheet_path import get_rels_map

    return get_rels_map(archive, resource_path)


def _load_from_drawing_xml(
    archive: zipfile.ZipFile,
    drawing_path: str,
    drawing_rels: dict[str, str],
    images_by_row: dict[int, list[Image]],
    max_row: int,
    header_row: int = 1,
) -> None:
    root = ET.fromstring(archive.read(drawing_path))
    anchor_elements: list[tuple[ET.Element, str]] = []
    for local_name in ("oneCellAnchor", "twoCellAnchor", "absoluteAnchor"):
        for anchor in _iter_local_name(root, local_name):
            anchor_elements.append((anchor, local_name))

    for anchor, local_name in anchor_elements:
        anchor_info = _parse_drawing_anchor(anchor, local_name)
        if anchor_info is None:
            continue
        for blip_rel_id in _find_blip_rel_ids(anchor):
            media_path = drawing_rels.get(blip_rel_id)
            if not media_path or media_path not in archive.namelist():
                continue
            media_data = archive.read(media_path)
            _assign_image_bytes_to_rows(
                media_data, anchor_info, images_by_row, max_row, header_row
            )


def _parse_drawing_anchor(anchor: ET.Element, tag: str) -> AnchorInfo | None:
    if tag == "absoluteAnchor":
        return _parse_absolute_anchor(anchor)
    if tag == "twoCellAnchor":
        return _parse_two_cell_anchor(anchor)
    return _parse_one_cell_anchor(anchor)


def _parse_one_cell_anchor(anchor: ET.Element) -> AnchorInfo | None:
    from_node = _find_child(anchor, "from")
    if from_node is None:
        return None
    from_row, from_col, row_off, col_off = _read_marker(from_node)
    ext = _find_child(anchor, "ext")
    cx, cy = _read_ext(ext)
    return AnchorInfo(
        from_row=from_row,
        to_row=from_row,
        from_col=from_col,
        col_off=col_off,
        row_off=row_off,
        cx=cx,
        cy=cy,
        is_one_cell=True,
    )


def _find_child(parent: ET.Element, local_name: str) -> ET.Element | None:
    for child in parent:
        if child.tag.endswith(f"}}{local_name}") or child.tag == local_name:
            return child
    return None


def _find_child_text(parent: ET.Element, local_name: str) -> str | None:
    child = _find_child(parent, local_name)
    return child.text if child is not None else None


def _parse_two_cell_anchor(anchor: ET.Element) -> AnchorInfo | None:
    from_node = _find_child(anchor, "from")
    to_node = _find_child(anchor, "to")
    if from_node is None or to_node is None:
        return None
    from_row, from_col, row_off, col_off = _read_marker(from_node)
    to_row, to_col, _, _ = _read_marker(to_node)
    if to_row < from_row:
        to_row = from_row
    return AnchorInfo(
        from_row=from_row,
        to_row=to_row,
        from_col=from_col,
        to_col=to_col,
        col_off=col_off,
        row_off=row_off,
        is_one_cell=False,
    )


def _parse_absolute_anchor(anchor: ET.Element) -> AnchorInfo | None:
    pos = _find_child(anchor, "pos")
    ext = _find_child(anchor, "ext")
    if pos is None:
        return None
    x = int(pos.attrib.get("x", "0"))
    y = int(pos.attrib.get("y", "0"))
    cx, cy = _read_ext(ext)
    from_row = max(1, y // DEFAULT_ROW_EMU + 1)
    to_row = max(from_row, (y + max(cy, DEFAULT_ROW_EMU // 2)) // DEFAULT_ROW_EMU + 1)
    from_col = max(0, x // 125_000)
    return AnchorInfo(
        from_row=from_row,
        to_row=to_row,
        from_col=from_col,
        cx=cx,
        cy=cy,
        is_one_cell=False,
    )


def _read_marker(node: ET.Element) -> tuple[int, int, int, int]:
    row = int(_find_child_text(node, "row") or "0") + 1
    col = int(_find_child_text(node, "col") or "0")
    row_off = int(_find_child_text(node, "rowOff") or "0")
    col_off = int(_find_child_text(node, "colOff") or "0")
    return row, col, row_off, col_off


def _read_ext(ext: ET.Element | None) -> tuple[int, int]:
    if ext is None:
        return 0, 0
    return int(ext.attrib.get("cx", "0")), int(ext.attrib.get("cy", "0"))


def _find_blip_rel_ids(anchor: ET.Element) -> list[str]:
    rel_ids: list[str] = []
    seen: set[str] = set()
    for blip in _iter_local_name(anchor, "blip"):
        rel_id = blip.attrib.get(REL_EMBED_ATTR) or blip.attrib.get(REL_ID_ATTR)
        if rel_id and rel_id not in seen:
            seen.add(rel_id)
            rel_ids.append(rel_id)
    return rel_ids


def _create_image_from_bytes(data: bytes, anchor_info: AnchorInfo) -> Image | None:
    try:
        image = Image(io.BytesIO(data))
    except Exception:
        return None
    image._cached_bytes = data  # noqa: SLF001
    image._data = lambda: data  # noqa: SLF001
    _apply_anchor(image, anchor_info)
    return image


def _apply_anchor(image: Image, anchor_info: AnchorInfo) -> None:
    if anchor_info.is_one_cell or anchor_info.to_col is None:
        marker = AnchorMarker(
            col=anchor_info.from_col,
            row=anchor_info.from_row - 1,
            colOff=anchor_info.col_off,
            rowOff=anchor_info.row_off,
        )
        ext = XDRPositiveSize2D(anchor_info.cx or 1, anchor_info.cy or 1)
        image.anchor = OneCellAnchor(_from=marker, ext=ext)
    else:
        from_marker = AnchorMarker(
            col=anchor_info.from_col,
            row=anchor_info.from_row - 1,
            colOff=anchor_info.col_off,
            rowOff=anchor_info.row_off,
        )
        to_marker = AnchorMarker(
            col=anchor_info.to_col,
            row=anchor_info.to_row - 1,
        )
        image.anchor = TwoCellAnchor(_from=from_marker, to=to_marker)

    return image


def _assign_image_bytes_to_rows(
    media_data: bytes,
    anchor_info: AnchorInfo,
    images_by_row: dict[int, list[Image]],
    max_row: int,
    header_row: int = 1,
) -> None:
    for row in get_assigned_rows(anchor_info.from_row, anchor_info.to_row, header_row):
        if row > max_row:
            continue
        image = _create_image_from_bytes(media_data, anchor_info)
        if image is None:
            continue
        images_by_row.setdefault(row, []).append(image)


def _load_from_vml(
    archive: zipfile.ZipFile,
    vml_path: str,
    vml_rels: dict[str, str],
    images_by_row: dict[int, list[Image]],
    max_row: int,
    header_row: int = 1,
) -> None:
    root = ET.fromstring(archive.read(vml_path))
    for shape in root.findall(".//v:shape", VML_NS):
        imagedata = shape.find(".//v:imagedata", VML_NS)
        if imagedata is None:
            continue
        rel_id = (
            imagedata.attrib.get(REL_ID_ATTR)
            or imagedata.attrib.get("{urn:schemas-microsoft-com:office:office}relid")
        )
        if not rel_id:
            continue
        media_path = vml_rels.get(rel_id)
        if not media_path or media_path not in archive.namelist():
            continue

        media_data = archive.read(media_path)
        anchor_info = _parse_vml_anchor(shape)
        if anchor_info is None:
            continue
        _assign_image_bytes_to_rows(
            media_data, anchor_info, images_by_row, max_row, header_row
        )


def _parse_vml_anchor(shape: ET.Element) -> AnchorInfo | None:
    anchor_text = None
    for client_data in shape.findall(".//x:ClientData", VML_NS):
        anchor_node = client_data.find("x:Anchor", VML_NS)
        if anchor_node is not None and anchor_node.text:
            anchor_text = anchor_node.text.strip()
            break
    if not anchor_text:
        return None

    parts = [int(part.strip()) for part in anchor_text.split(",")]
    if len(parts) != 8:
        return None

    from_col, _, from_row, _, to_col, _, to_row, _ = parts
    from_row += 1
    to_row += 1
    if to_row < from_row:
        to_row = from_row
    return AnchorInfo(
        from_row=from_row,
        to_row=to_row,
        from_col=from_col,
        to_col=to_col,
        is_one_cell=from_row == to_row and from_col == to_col,
    )
