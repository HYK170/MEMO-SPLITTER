from __future__ import annotations

import io
import zipfile
from copy import copy
from xml.etree import ElementTree as ET

from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

HYPERLINK_TAGS = {
    "{http://schemas.openxmlformats.org/drawingml/2006/main}hlinkClick",
    "{http://schemas.openxmlformats.org/drawingml/2006/main}hlinkHover",
}


def get_image_row(image: Image) -> int | None:
    anchor = getattr(image, "anchor", None)
    if anchor is None:
        return None
    from_anchor = getattr(anchor, "_from", None)
    if from_anchor is None:
        return None
    return int(from_anchor.row) + 1


def index_images_by_row(ws: Worksheet) -> dict[int, list[Image]]:
    images_by_row: dict[int, list[Image]] = {}
    for image in getattr(ws, "_images", []):
        row = get_image_row(image)
        if row is None:
            continue
        images_by_row.setdefault(row, []).append(image)
    return images_by_row


def strip_image_hyperlinks(image: Image) -> None:
    anchor = getattr(image, "anchor", None)
    if anchor is None:
        return

    pic = getattr(anchor, "pic", None)
    if pic is not None:
        nv_pic_pr = getattr(pic, "nvPicPr", None)
        if nv_pic_pr is not None:
            c_nv_pr = getattr(nv_pic_pr, "cNvPr", None)
            if c_nv_pr is not None:
                if hasattr(c_nv_pr, "hlinkClick"):
                    c_nv_pr.hlinkClick = None
                if hasattr(c_nv_pr, "hlinkHover"):
                    c_nv_pr.hlinkHover = None


def copy_image_for_row(source_image: Image) -> Image:
    data = source_image._data()
    new_image = Image(io.BytesIO(data))
    new_image.width = source_image.width
    new_image.height = source_image.height
    new_image.anchor = copy(source_image.anchor)
    strip_image_hyperlinks(new_image)
    return new_image


def add_row_images(target_ws: Worksheet, images: list[Image]) -> None:
    for image in images:
        copied = copy_image_for_row(image)
        target_ws.add_image(copied)


def get_image_cell_coordinate(image: Image) -> str | None:
    row = get_image_row(image)
    anchor = getattr(image, "anchor", None)
    if row is None or anchor is None:
        return None
    from_anchor = getattr(anchor, "_from", None)
    if from_anchor is None:
        return None
    col = int(from_anchor.col) + 1
    return f"{get_column_letter(col)}{row}"


def strip_hyperlinks_from_xlsx_file(xlsx_path) -> None:
    buffer = io.BytesIO()
    removed_rel_ids: set[str] = set()

    with zipfile.ZipFile(xlsx_path, "r") as zin:
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith("xl/drawings/") and item.filename.endswith(".xml"):
                    data, rel_ids = _strip_drawing_hyperlinks(data)
                    removed_rel_ids.update(rel_ids)
                elif item.filename.startswith("xl/drawings/_rels/") and item.filename.endswith(".rels"):
                    if removed_rel_ids:
                        data = _remove_relationships(data, removed_rel_ids)
                zout.writestr(item, data)

    with open(xlsx_path, "wb") as out_file:
        out_file.write(buffer.getvalue())


def _strip_drawing_hyperlinks(data: bytes) -> tuple[bytes, set[str]]:
    removed_rel_ids: set[str] = set()
    root = ET.fromstring(data)
    for tag in HYPERLINK_TAGS:
        for element in root.iter(tag):
            rel_id = element.attrib.get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            )
            if rel_id:
                removed_rel_ids.add(rel_id)
            parent = _find_parent(root, element)
            if parent is not None:
                parent.remove(element)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), removed_rel_ids


def _find_parent(root: ET.Element, child: ET.Element) -> ET.Element | None:
    for parent in root.iter():
        if child in list(parent):
            return parent
    return None


def _remove_relationships(data: bytes, rel_ids: set[str]) -> bytes:
    root = ET.fromstring(data)
    rel_tag = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
    for rel in list(root.findall(rel_tag)):
        if rel.attrib.get("Id") in rel_ids:
            root.remove(rel)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)
