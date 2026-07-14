from __future__ import annotations

import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from src.image_handler import get_assigned_rows
from src.sheet_path import discover_drawing_paths, get_rels_map, resolve_sheet_path

REL_ID_ATTR = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
REL_TAG = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"


class XlsxHyperlinkIndex:
    """drawing XML의 이미지/객체 하이퍼링크를 행 단위로 인덱싱한다."""

    def __init__(
        self,
        xlsx_path: Path,
        header_row: int = 1,
        ws: Worksheet | None = None,
        wb: Workbook | None = None,
    ) -> None:
        self._header_row = header_row
        self._image_targets_by_row: dict[int, list[str]] = {}
        self._entries: list[tuple[int, int, str]] = []  # from_row, to_row, target
        self._build_index(xlsx_path, ws, wb)

    def targets_for_row(self, row: int) -> list[str]:
        return list(self._image_targets_by_row.get(row, []))

    def total_targets(self) -> int:
        return sum(len(v) for v in self._image_targets_by_row.values())

    def unique_targets(self) -> list[str]:
        seen: set[str] = set()
        result: list[str] = []
        for targets in self._image_targets_by_row.values():
            for target in targets:
                if target not in seen:
                    seen.add(target)
                    result.append(target)
        return result

    def summarize(self) -> list[str]:
        unique = self.unique_targets()
        lines = [
            f"이미지 하이퍼링크: 유니크 {len(unique)}개 / 행배정 {self.total_targets()}건",
        ]
        for target in unique[:5]:
            lines.append(f"  예시: {target}")
        if len(unique) > 5:
            lines.append(f"  ... 외 {len(unique) - 5}개")
        row_counts = {
            row: len(targets)
            for row, targets in sorted(self._image_targets_by_row.items())
            if targets
        }
        if row_counts:
            preview = list(row_counts.items())[:8]
            lines.append(
                "  행별: "
                + ", ".join(f"{row}행={count}" for row, count in preview)
                + (f" ... 외 {len(row_counts) - 8}행" if len(row_counts) > 8 else "")
            )
        return lines

    def _build_index(
        self,
        xlsx_path: Path,
        ws: Worksheet | None,
        wb: Workbook | None,
    ) -> None:
        with zipfile.ZipFile(xlsx_path, "r") as archive:
            drawing_paths: list[str] = []
            if ws is not None:
                sheet_path = resolve_sheet_path(archive, ws, wb)
                if sheet_path:
                    drawing_paths = discover_drawing_paths(archive, sheet_path, ws, wb)

            if not drawing_paths:
                drawing_paths = sorted(
                    name
                    for name in archive.namelist()
                    if name.startswith("xl/drawings/")
                    and name.endswith(".xml")
                    and "_rels" not in name
                    and "vml" not in name.lower()
                )

            for drawing_path in drawing_paths:
                self._parse_drawing(archive, drawing_path)

    def _parse_drawing(self, archive: zipfile.ZipFile, drawing_path: str) -> None:
        drawing_rels = self._get_hyperlink_targets(archive, drawing_path)
        if not drawing_rels and not self._has_any_hlink(archive, drawing_path):
            # still parse in case Target is inline (rare)
            drawing_rels = {}

        root = ET.fromstring(archive.read(drawing_path))
        for local_name in ("oneCellAnchor", "twoCellAnchor", "absoluteAnchor"):
            for anchor in _iter_local_name(root, local_name):
                from_row, to_row = _anchor_row_range(anchor, local_name)
                if from_row is None or to_row is None:
                    continue

                targets = _extract_hlink_targets(anchor, drawing_rels)
                if not targets:
                    continue

                self._entries.append((from_row, to_row, targets[0]))
                for row in get_assigned_rows(from_row, to_row, self._header_row):
                    bucket = self._image_targets_by_row.setdefault(row, [])
                    for target in targets:
                        if target not in bucket:
                            bucket.append(target)

    def _get_hyperlink_targets(
        self,
        archive: zipfile.ZipFile,
        drawing_path: str,
    ) -> dict[str, str]:
        """hyperlink relationship Id -> Target (file path / URL)."""
        normalized = drawing_path.replace("\\", "/").lstrip("/")
        from pathlib import PurePosixPath

        rels_path = str(
            PurePosixPath(normalized).parent / "_rels" / f"{PurePosixPath(normalized).name}.rels"
        )
        if rels_path not in archive.namelist():
            return {}

        root = ET.fromstring(archive.read(rels_path))
        targets: dict[str, str] = {}
        for rel in root.findall(REL_TAG):
            rel_id = rel.attrib.get("Id")
            if not rel_id:
                continue
            rel_type = rel.attrib.get("Type", "").lower()
            target = rel.attrib.get("Target", "")
            if not target:
                continue
            # External hyperlink OR any Non-image TargetMode External
            if "hyperlink" in rel_type or rel.attrib.get("TargetMode", "").lower() == "external":
                targets[rel_id] = target
        return targets

    def _has_any_hlink(self, archive: zipfile.ZipFile, drawing_path: str) -> bool:
        root = ET.fromstring(archive.read(drawing_path))
        return bool(_iter_local_name(root, "hlinkClick") or _iter_local_name(root, "hlinkHover"))


def _iter_local_name(root: ET.Element, local_name: str) -> list[ET.Element]:
    suffix = f"}}{local_name}"
    return [el for el in root.iter() if el.tag.endswith(suffix) or el.tag == local_name]


def _anchor_row_range(anchor: ET.Element, local_name: str) -> tuple[int | None, int | None]:
    from_node = _find_child(anchor, "from")
    if from_node is None:
        return None, None
    from_row = int(_find_child_text(from_node, "row") or "0") + 1
    to_node = _find_child(anchor, "to")
    if to_node is not None:
        to_row = int(_find_child_text(to_node, "row") or str(from_row - 1)) + 1
    else:
        to_row = from_row
    if to_row < from_row:
        to_row = from_row
    return from_row, to_row


def _find_child(parent: ET.Element, local_name: str) -> ET.Element | None:
    for child in parent:
        if child.tag.endswith(f"}}{local_name}") or child.tag == local_name:
            return child
    return None


def _find_child_text(parent: ET.Element, local_name: str) -> str | None:
    child = _find_child(parent, local_name)
    return child.text if child is not None else None


def _extract_hlink_targets(anchor: ET.Element, drawing_rels: dict[str, str]) -> list[str]:
    targets: list[str] = []
    seen: set[str] = set()
    for tag in ("hlinkClick", "hlinkHover"):
        for hlink in _iter_local_name(anchor, tag):
            rel_id = hlink.attrib.get(REL_ID_ATTR)
            if rel_id and rel_id in drawing_rels:
                target = drawing_rels[rel_id]
            else:
                # rare: action/target embedded directly
                target = hlink.attrib.get("action") or hlink.attrib.get("tgtFrame") or ""
            if target and target not in seen:
                seen.add(target)
                targets.append(target)
    return targets
