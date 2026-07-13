from __future__ import annotations

import unicodedata
import zipfile
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET

from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

REL_ID_ATTR = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
SHEET_TAG = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet"
REL_TAG = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"


def normalize_sheet_name(name: str) -> str:
    return unicodedata.normalize("NFC", name.strip())


def resolve_sheet_path(
    archive: zipfile.ZipFile,
    ws: Worksheet,
    wb: Workbook | None = None,
) -> str | None:
    names = set(archive.namelist())

    sheet_path = getattr(ws, "path", None)
    if sheet_path:
        normalized = _normalize_zip_path(sheet_path)
        if normalized in names:
            return normalized

    by_name = _resolve_by_workbook_name(archive, ws.title, names)
    if by_name:
        return by_name

    if wb is not None:
        by_index = _resolve_by_workbook_index(archive, wb, ws, names)
        if by_index:
            return by_index

    return _resolve_single_sheet_fallback(names)


def list_workbook_sheet_names(archive: zipfile.ZipFile) -> list[str]:
    workbook_path = "xl/workbook.xml"
    if workbook_path not in archive.namelist():
        return []
    root = ET.fromstring(archive.read(workbook_path))
    return [sheet.attrib.get("name", "") for sheet in root.findall(f".//{SHEET_TAG}")]


def get_rels_map(archive: zipfile.ZipFile, resource_path: str) -> dict[str, str]:
    normalized = _normalize_zip_path(resource_path)
    rels_path = str(PurePosixPath(normalized).parent / "_rels" / f"{PurePosixPath(normalized).name}.rels")
    if rels_path not in archive.namelist():
        return {}
    root = ET.fromstring(archive.read(rels_path))
    return {
        rel.attrib["Id"]: _resolve_target(normalized, rel.attrib.get("Target", ""))
        for rel in root.findall(REL_TAG)
        if rel.attrib.get("Id")
    }


def find_sheet_related_paths(archive: zipfile.ZipFile, sheet_path: str, keyword: str) -> list[str]:
    names = set(archive.namelist())
    normalized_sheet = _normalize_zip_path(sheet_path)
    rels_path = normalized_sheet.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels"
    if rels_path not in archive.namelist():
        return []
    root = ET.fromstring(archive.read(rels_path))
    paths: list[str] = []
    seen: set[str] = set()
    for rel in root.findall(REL_TAG):
        rel_type = rel.attrib.get("Type", "").lower()
        target = _resolve_target(normalized_sheet, rel.attrib.get("Target", ""))
        if target not in names or target in seen:
            continue
        matched = keyword.lower() in rel_type
        if not matched and keyword.lower() == "drawing":
            matched = "/drawings/drawing" in target.lower()
        if not matched and keyword.lower() == "vmldrawing":
            matched = "vml" in target.lower() or "legacydrawing" in rel_type
        if matched:
            seen.add(target)
            paths.append(target)
    return paths


def discover_drawing_paths(
    archive: zipfile.ZipFile,
    sheet_path: str,
    ws: Worksheet | None = None,
    wb: Workbook | None = None,
) -> list[str]:
    names = set(archive.namelist())
    normalized_sheet = _normalize_zip_path(sheet_path)
    paths: list[str] = []
    seen: set[str] = set()

    def add(path: str | None) -> None:
        if not path:
            return
        normalized = _normalize_zip_path(path)
        if normalized in names and normalized not in seen:
            seen.add(normalized)
            paths.append(normalized)

    for path in find_sheet_related_paths(archive, normalized_sheet, "drawing"):
        add(path)

    if normalized_sheet in names:
        root = ET.fromstring(archive.read(normalized_sheet))
        rels_map = get_rels_map(archive, normalized_sheet)
        for element in root.iter():
            local = element.tag.rsplit("}", 1)[-1]
            if local not in ("drawing", "legacyDrawing"):
                continue
            rel_id = element.attrib.get(REL_ID_ATTR)
            if rel_id:
                add(rels_map.get(rel_id))

    for path in _drawing_paths_by_sheet_index(normalized_sheet, ws, wb, names):
        add(path)

    return paths


def discover_vml_paths(
    archive: zipfile.ZipFile,
    sheet_path: str,
    ws: Worksheet | None = None,
    wb: Workbook | None = None,
) -> list[str]:
    names = set(archive.namelist())
    normalized_sheet = _normalize_zip_path(sheet_path)
    paths: list[str] = []
    seen: set[str] = set()

    def add(path: str | None) -> None:
        if not path:
            return
        normalized = _normalize_zip_path(path)
        if normalized in names and normalized not in seen:
            seen.add(normalized)
            paths.append(normalized)

    for path in find_sheet_related_paths(archive, normalized_sheet, "vmlDrawing"):
        add(path)

    if normalized_sheet in names:
        root = ET.fromstring(archive.read(normalized_sheet))
        rels_map = get_rels_map(archive, normalized_sheet)
        for element in root.iter():
            local = element.tag.rsplit("}", 1)[-1]
            if local != "legacyDrawing":
                continue
            rel_id = element.attrib.get(REL_ID_ATTR)
            if rel_id:
                add(rels_map.get(rel_id))

    sheet_index = _sheet_index_from_path(normalized_sheet, ws, wb)
    if sheet_index is not None:
        for candidate in (
            f"xl/drawings/vmlDrawing{sheet_index}.vml",
            f"xl/drawings/VmlDrawing{sheet_index}.vml",
        ):
            add(candidate)

    return paths


def list_package_drawing_parts(archive: zipfile.ZipFile) -> list[str]:
    return sorted(
        name
        for name in archive.namelist()
        if name.startswith("xl/drawings/")
        and name.endswith(".xml")
        and "_rels" not in name
        and "vml" not in name.lower()
    )


def list_sheet_relationships(archive: zipfile.ZipFile, sheet_path: str) -> list[tuple[str, str]]:
    normalized_sheet = _normalize_zip_path(sheet_path)
    rels_path = normalized_sheet.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels"
    if rels_path not in archive.namelist():
        return []
    root = ET.fromstring(archive.read(rels_path))
    rows: list[tuple[str, str]] = []
    for rel in root.findall(REL_TAG):
        rel_type = rel.attrib.get("Type", "")
        short_type = rel_type.rsplit("/", 1)[-1] if rel_type else "unknown"
        target = _resolve_target(normalized_sheet, rel.attrib.get("Target", ""))
        rows.append((short_type, target))
    return rows


def _sheet_index_from_path(
    sheet_path: str,
    ws: Worksheet | None,
    wb: Workbook | None,
) -> int | None:
    import re

    match = re.search(r"sheet(\d+)\.xml$", sheet_path, re.IGNORECASE)
    if match:
        return int(match.group(1))
    if wb is not None and ws is not None:
        try:
            return wb.sheetnames.index(ws.title) + 1
        except ValueError:
            return None
    return None


def _drawing_paths_by_sheet_index(
    sheet_path: str,
    ws: Worksheet | None,
    wb: Workbook | None,
    names: set[str],
) -> list[str]:
    sheet_index = _sheet_index_from_path(sheet_path, ws, wb)
    if sheet_index is None:
        return []
    candidates = [
        f"xl/drawings/drawing{sheet_index}.xml",
        f"xl/drawings/Drawing{sheet_index}.xml",
    ]
    return [candidate for candidate in candidates if candidate in names]


def _collapse_posix(path: str) -> str:
    parts: list[str] = []
    for part in path.replace("\\", "/").split("/"):
        if not part or part == ".":
            continue
        if part == "..":
            if parts:
                parts.pop()
            continue
        parts.append(part)
    return "/".join(parts)


def _normalize_zip_path(path: str) -> str:
    cleaned = _collapse_posix(path.replace("\\", "/").lstrip("/"))
    if cleaned.startswith("xl/"):
        return cleaned
    if cleaned.startswith("worksheets/"):
        return f"xl/{cleaned}"
    if cleaned.startswith("drawings/") or cleaned.startswith("media/"):
        return f"xl/{cleaned}"
    return cleaned


def _resolve_target(base_path: str, target: str) -> str:
    normalized_target = _normalize_zip_path(target)
    if normalized_target.startswith("xl/"):
        return normalized_target
    base = _normalize_zip_path(base_path)
    base_dir = PurePosixPath(base).parent
    joined = _collapse_posix(str(base_dir / target.replace("\\", "/").lstrip("/")))
    return _normalize_zip_path(joined)


def _resolve_by_workbook_name(
    archive: zipfile.ZipFile,
    sheet_name: str,
    names: set[str],
) -> str | None:
    workbook_path = "xl/workbook.xml"
    if workbook_path not in archive.namelist():
        return None

    workbook_root = ET.fromstring(archive.read(workbook_path))
    workbook_rels = get_rels_map(archive, workbook_path)
    target_name = normalize_sheet_name(sheet_name)

    for sheet in workbook_root.findall(f".//{SHEET_TAG}"):
        xml_name = normalize_sheet_name(sheet.attrib.get("name", ""))
        if xml_name != target_name:
            continue
        rel_id = sheet.attrib.get(REL_ID_ATTR)
        if not rel_id:
            continue
        target = workbook_rels.get(rel_id)
        if target and target in names:
            return target
    return None


def _resolve_by_workbook_index(
    archive: zipfile.ZipFile,
    wb: Workbook,
    ws: Worksheet,
    names: set[str],
) -> str | None:
    try:
        idx = wb.sheetnames.index(ws.title) + 1
    except ValueError:
        return None

    for candidate in (
        f"xl/worksheets/sheet{idx}.xml",
        f"xl/worksheets/Sheet{idx}.xml",
    ):
        if candidate in names:
            return candidate

    sheet_files = sorted(
        name for name in names if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")
    )
    if 0 < idx <= len(sheet_files):
        return sheet_files[idx - 1]
    return None


def _resolve_single_sheet_fallback(names: set[str]) -> str | None:
    sheet_files = sorted(
        name for name in names if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")
    )
    if len(sheet_files) == 1:
        return sheet_files[0]
    return None
