from __future__ import annotations

import io
import re
import zipfile
from copy import deepcopy
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET

from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from src.drawing_image_loader import (
    REL_EMBED_ATTR,
    REL_ID_ATTR,
    _find_blip_rel_ids,
    _get_rels_map,
    _iter_local_name,
    _parse_drawing_anchor,
    _resolve_media_path,
)
from src.image_handler import get_assigned_rows
from src.sheet_path import discover_drawing_paths, get_rels_map, resolve_sheet_path

WS_DR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DRAWING_REL_TYPE = f"{OFFICE_REL_NS}/drawing"
IMAGE_REL_TYPE = f"{OFFICE_REL_NS}/image"
REL_TAG = f"{{{REL_NS}}}Relationship"
SHEET_TAG = f"{{{SHEET_NS}}}worksheet"
DRAWING_TAG = f"{{{SHEET_NS}}}drawing"


def count_matching_anchors(
    source_xlsx: Path,
    ws: Worksheet,
    wb: Workbook,
    data_row: int,
    header_row: int,
) -> int:
    total = 0
    with zipfile.ZipFile(source_xlsx, "r") as archive:
        sheet_path = resolve_sheet_path(archive, ws, wb)
        if not sheet_path:
            return 0
        for drawing_path in discover_drawing_paths(archive, sheet_path, ws, wb):
            total += len(_select_anchors_for_row(archive, drawing_path, data_row, header_row))
    return total


def count_all_source_anchors(source_xlsx: Path, ws: Worksheet, wb: Workbook) -> int:
    total = 0
    with zipfile.ZipFile(source_xlsx, "r") as archive:
        sheet_path = resolve_sheet_path(archive, ws, wb)
        if not sheet_path:
            return 0
        for drawing_path in discover_drawing_paths(archive, sheet_path, ws, wb):
            root = ET.fromstring(archive.read(drawing_path))
            for local_name in ("oneCellAnchor", "twoCellAnchor", "absoluteAnchor"):
                total += len(_iter_local_name(root, local_name))
    return total


def transplant_drawings_for_row(
    source_xlsx: Path,
    output_xlsx: Path,
    ws: Worksheet,
    wb: Workbook,
    data_row: int,
    header_row: int,
) -> int:
    with zipfile.ZipFile(source_xlsx, "r") as source_archive:
        sheet_path = resolve_sheet_path(source_archive, ws, wb)
        if not sheet_path:
            return 0

        selected: list[tuple[str, ET.Element, dict[str, str]]] = []
        for drawing_path in discover_drawing_paths(source_archive, sheet_path, ws, wb):
            drawing_rels = _get_rels_map(source_archive, drawing_path)
            for anchor, _local_name in _select_anchors_for_row(
                source_archive, drawing_path, data_row, header_row
            ):
                selected.append((drawing_path, anchor, drawing_rels))

        if not selected:
            return 0

        buffer = io.BytesIO()
        with zipfile.ZipFile(output_xlsx, "r") as output_in:
            with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as output_out:
                existing_names = set(output_in.namelist())
                drawing_path, media_map, rel_id_map, rel_entries = _build_output_drawing(
                    source_archive,
                    selected,
                    existing_names,
                )
                replace_names = {
                    drawing_path,
                    "xl/drawings/_rels/drawing1.xml.rels",
                    "xl/worksheets/sheet1.xml",
                    "xl/worksheets/_rels/sheet1.xml.rels",
                    "[Content_Types].xml",
                    *media_map.keys(),
                }

                content_types_data = output_in.read("[Content_Types].xml")
                sheet_data = output_in.read("xl/worksheets/sheet1.xml")
                sheet_rels_data = None
                if "xl/worksheets/_rels/sheet1.xml.rels" in existing_names:
                    sheet_rels_data = output_in.read("xl/worksheets/_rels/sheet1.xml.rels")

                for item in output_in.infolist():
                    if item.filename in replace_names:
                        continue
                    output_out.writestr(item, output_in.read(item.filename))

                output_out.writestr(drawing_path, _serialize_drawing(source_archive, selected, rel_id_map))
                output_out.writestr(
                    "xl/drawings/_rels/drawing1.xml.rels",
                    _serialize_drawing_rels(rel_entries),
                )
                for media_path, media_bytes in media_map.items():
                    output_out.writestr(media_path, media_bytes)

                _ensure_content_types(output_out, media_map, content_types_data)
                _link_sheet_to_drawing(output_out, sheet_data, sheet_rels_data)

        output_xlsx.write_bytes(buffer.getvalue())
        return len(selected)


def _select_anchors_for_row(
    archive: zipfile.ZipFile,
    drawing_path: str,
    data_row: int,
    header_row: int,
) -> list[tuple[ET.Element, str]]:
    root = ET.fromstring(archive.read(drawing_path))
    selected: list[tuple[ET.Element, str]] = []
    for local_name in ("oneCellAnchor", "twoCellAnchor", "absoluteAnchor"):
        for anchor in _iter_local_name(root, local_name):
            anchor_info = _parse_drawing_anchor(anchor, local_name)
            if anchor_info is None:
                continue
            if data_row in get_assigned_rows(
                anchor_info.from_row,
                anchor_info.to_row,
                header_row,
            ):
                selected.append((anchor, local_name))
    return selected


def _build_output_drawing(
    source_archive: zipfile.ZipFile,
    selected: list[tuple[str, ET.Element, dict[str, str]]],
    existing_names: set[str],
) -> tuple[str, dict[str, bytes], dict[str, str], list[tuple[str, str]]]:
    media_map: dict[str, bytes] = {}
    rel_id_map: dict[str, str] = {}
    rel_entries: list[tuple[str, str]] = []
    media_by_source: dict[str, str] = {}
    media_index = 1

    for _drawing_path, anchor, drawing_rels in selected:
        for blip_rel_id in _find_blip_rel_ids(anchor):
            if blip_rel_id in rel_id_map:
                continue
            raw_target = drawing_rels.get(blip_rel_id)
            if not raw_target:
                continue
            resolved = _resolve_media_path(source_archive, raw_target)
            if not resolved:
                continue

            if resolved in media_by_source:
                rel_id_map[blip_rel_id] = media_by_source[resolved]
                continue

            ext = PurePosixPath(resolved).suffix.lower() or ".bin"
            while True:
                candidate = f"xl/media/image{media_index}{ext}"
                media_index += 1
                if candidate not in existing_names and candidate not in media_map:
                    break

            new_rel_id = f"rId{len(rel_entries) + 1}"
            media_map[candidate] = source_archive.read(resolved)
            rel_id_map[blip_rel_id] = new_rel_id
            media_by_source[resolved] = new_rel_id
            rel_entries.append((new_rel_id, PurePosixPath(candidate).name))

    return "xl/drawings/drawing1.xml", media_map, rel_id_map, rel_entries


def _serialize_drawing(
    source_archive: zipfile.ZipFile,
    selected: list[tuple[str, ET.Element, dict[str, str]]],
    rel_id_map: dict[str, str],
) -> bytes:
    template_path = selected[0][0]
    template_root = ET.fromstring(source_archive.read(template_path))
    root = ET.Element(template_root.tag, dict(template_root.attrib))

    for _drawing_path, anchor, _drawing_rels in selected:
        cloned = deepcopy(anchor)
        anchor_map = {
            blip_id: rel_id_map[blip_id]
            for blip_id in _find_blip_rel_ids(anchor)
            if blip_id in rel_id_map
        }
        _rewrite_blip_embeds(cloned, anchor_map)
        root.append(cloned)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _rewrite_blip_embeds(anchor: ET.Element, rel_map: dict[str, str]) -> None:
    for blip in _iter_local_name(anchor, "blip"):
        old_id = blip.attrib.get(REL_EMBED_ATTR) or blip.attrib.get(REL_ID_ATTR)
        if old_id and old_id in rel_map:
            new_id = rel_map[old_id]
            blip.attrib[REL_EMBED_ATTR] = new_id
            if REL_ID_ATTR in blip.attrib:
                blip.attrib[REL_ID_ATTR] = new_id


def _serialize_drawing_rels(rel_entries: list[tuple[str, str]]) -> bytes:
    root = ET.Element(f"{{{REL_NS}}}Relationships")
    for rel_id, media_name in rel_entries:
        rel = ET.SubElement(root, REL_TAG)
        rel.set("Id", rel_id)
        rel.set("Type", IMAGE_REL_TYPE)
        rel.set("Target", f"../media/{media_name}")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _ensure_content_types(
    archive: zipfile.ZipFile,
    media_map: dict[str, bytes],
    content_types_data: bytes,
) -> None:
    root = ET.fromstring(content_types_data)
    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    default_tag = f"{{{ct_ns}}}Default"
    override_tag = f"{{{ct_ns}}}Override"

    extensions = {PurePosixPath(name).suffix.lstrip(".") for name in media_map.keys()}
    existing_ext = {el.attrib.get("Extension") for el in root.findall(default_tag)}
    for ext in extensions:
        if ext and ext not in existing_ext:
            default = ET.SubElement(root, default_tag)
            default.set("Extension", ext)
            default.set(
                "ContentType",
                _guess_content_type(ext),
            )

    drawing_part = "/xl/drawings/drawing1.xml"
    if not any(el.attrib.get("PartName") == drawing_part for el in root.findall(override_tag)):
        override = ET.SubElement(root, override_tag)
        override.set("PartName", drawing_part)
        override.set(
            "ContentType",
            "application/vnd.openxmlformats-officedocument.drawing+xml",
        )

    archive.writestr("[Content_Types].xml", ET.tostring(root, encoding="utf-8", xml_declaration=True))


def _guess_content_type(ext: str) -> str:
    mapping = {
        "png": "image/png",
        "jpg": "image/jpeg",
        "jpeg": "image/jpeg",
        "gif": "image/gif",
        "bmp": "image/bmp",
        "emf": "image/x-emf",
        "wmf": "image/x-wmf",
        "tif": "image/tiff",
        "tiff": "image/tiff",
    }
    return mapping.get(ext.lower(), "application/octet-stream")


def _link_sheet_to_drawing(
    archive: zipfile.ZipFile,
    sheet_data: bytes,
    sheet_rels_data: bytes | None,
) -> None:
    root = ET.fromstring(sheet_data)
    drawing_rel_id = _ensure_sheet_drawing_rel(archive, sheet_rels_data)

    has_drawing = False
    for child in list(root):
        if child.tag == DRAWING_TAG or child.tag.endswith("}drawing"):
            child.set(REL_ID_ATTR, drawing_rel_id)
            has_drawing = True
            break
    if not has_drawing:
        drawing_el = ET.Element(DRAWING_TAG)
        drawing_el.set(REL_ID_ATTR, drawing_rel_id)
        root.append(drawing_el)

    archive.writestr("xl/worksheets/sheet1.xml", ET.tostring(root, encoding="utf-8", xml_declaration=True))


def _ensure_sheet_drawing_rel(
    archive: zipfile.ZipFile,
    sheet_rels_data: bytes | None,
) -> str:
    if sheet_rels_data:
        root = ET.fromstring(sheet_rels_data)
    else:
        root = ET.Element(f"{{{REL_NS}}}Relationships")

    for rel in root.findall(REL_TAG):
        if "drawing" in rel.attrib.get("Type", ""):
            return rel.attrib["Id"]

    rel_ids = [rel.attrib.get("Id", "") for rel in root.findall(REL_TAG)]
    numbers = [int(rid[3:]) for rid in rel_ids if rid.startswith("rId") and rid[3:].isdigit()]
    next_id = max(numbers, default=0) + 1
    rel_id = f"rId{next_id}"

    rel = ET.SubElement(root, REL_TAG)
    rel.set("Id", rel_id)
    rel.set("Type", DRAWING_REL_TYPE)
    rel.set("Target", "../drawings/drawing1.xml")
    archive.writestr("xl/worksheets/_rels/sheet1.xml.rels", ET.tostring(root, encoding="utf-8", xml_declaration=True))
    return rel_id
