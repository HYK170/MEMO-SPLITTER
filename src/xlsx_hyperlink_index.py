from __future__ import annotations

import zipfile
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET

from src.image_handler import get_assigned_rows

NS = {
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
}


class XlsxHyperlinkIndex:
    def __init__(self, xlsx_path: Path) -> None:
        self._image_targets_by_row: dict[int, list[str]] = {}
        self._build_index(xlsx_path)

    def targets_for_row(self, row: int) -> list[str]:
        return list(self._image_targets_by_row.get(row, []))

    def _build_index(self, xlsx_path: Path) -> None:
        with zipfile.ZipFile(xlsx_path, "r") as archive:
            sheet_paths = [
                name
                for name in archive.namelist()
                if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")
            ]
            for sheet_path in sheet_paths:
                drawing_path = self._get_drawing_path(archive, sheet_path)
                if not drawing_path:
                    continue
                drawing_rels = self._get_drawing_rels(archive, drawing_path)
                self._parse_drawing(archive, drawing_path, drawing_rels)

    def _normalize_zip_path(self, target: str) -> str:
        cleaned = target.replace("\\", "/").lstrip("/")
        if cleaned.startswith("xl/"):
            return cleaned
        if cleaned.startswith("drawings/"):
            return f"xl/{cleaned}"
        return cleaned.replace("../", "")

    def _get_drawing_path(self, archive: zipfile.ZipFile, sheet_path: str) -> str | None:
        rels_path = sheet_path.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels"
        if rels_path not in archive.namelist():
            return None
        root = ET.fromstring(archive.read(rels_path))
        for rel in root.findall("rel:Relationship", NS):
            if "drawing" in rel.attrib.get("Type", ""):
                target = self._normalize_zip_path(rel.attrib.get("Target", ""))
                if target in archive.namelist():
                    return target
        return None

    def _get_drawing_rels(self, archive: zipfile.ZipFile, drawing_path: str) -> dict[str, str]:
        rels_path = f"xl/drawings/_rels/{PurePosixPath(drawing_path).name}.rels"
        if rels_path not in archive.namelist():
            return {}
        root = ET.fromstring(archive.read(rels_path))
        return {
            rel.attrib["Id"]: rel.attrib.get("Target", "")
            for rel in root.findall("rel:Relationship", NS)
            if rel.attrib.get("Id")
        }

    def _parse_drawing(
        self,
        archive: zipfile.ZipFile,
        drawing_path: str,
        drawing_rels: dict[str, str],
    ) -> None:
        root = ET.fromstring(archive.read(drawing_path))
        anchors = root.findall(".//xdr:twoCellAnchor", NS) + root.findall(".//xdr:oneCellAnchor", NS)
        for anchor in anchors:
            from_node = anchor.find("xdr:from", NS)
            if from_node is None:
                continue
            row_node = from_node.find("xdr:row", NS)
            if row_node is None or row_node.text is None:
                continue
            from_row = int(row_node.text) + 1

            to_node = anchor.find("xdr:to", NS)
            if to_node is not None:
                to_row_node = to_node.find("xdr:row", NS)
                to_row = int(to_row_node.text) + 1 if to_row_node is not None and to_row_node.text else from_row
            else:
                to_row = from_row

            if to_row < from_row:
                to_row = from_row

            targets: list[str] = []
            for tag_local in ("hlinkClick", "hlinkHover"):
                for hlink in anchor.findall(f".//a:{tag_local}", NS):
                    rel_id = hlink.attrib.get(
                        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
                    )
                    if not rel_id:
                        continue
                    target = drawing_rels.get(rel_id)
                    if target:
                        targets.append(target)

            if not targets:
                continue

            for row in get_assigned_rows(from_row, to_row):
                bucket = self._image_targets_by_row.setdefault(row, [])
                for target in targets:
                    if target not in bucket:
                        bucket.append(target)
