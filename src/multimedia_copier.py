from __future__ import annotations

import shutil
from dataclasses import dataclass, field
from pathlib import Path

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tif", ".tiff"}
SAVED_NAME_COLUMN = "저장된 파일 이름"


@dataclass
class MultimediaCopyResult:
    copied: list[str] = field(default_factory=list)
    image_paths: list[Path] = field(default_factory=list)
    skipped: list[str] = field(default_factory=list)
    # 원본 상대경로(정규화) -> 복사된 파일명
    path_map: dict[str, str] = field(default_factory=dict)


def parse_saved_file_names(cell_value: object) -> list[str]:
    if cell_value is None:
        return []
    text = str(cell_value).replace("\r\n", "\n").replace("\r", "\n")
    paths: list[str] = []
    for line in text.split("\n"):
        cleaned = line.strip().strip('"').strip("'")
        if cleaned:
            paths.append(cleaned)
    return paths


def resolve_multimedia_path(multimedia_root: Path, relative: str) -> Path:
    cleaned = relative.replace("\\", "/").lstrip("/")
    return (multimedia_root / cleaned).resolve()


def is_image_file(path: Path) -> bool:
    return path.suffix.lower() in IMAGE_EXTENSIONS


def unique_dest_path(folder: Path, filename: str) -> Path:
    dest = folder / filename
    if not dest.exists():
        return dest
    stem = dest.stem
    suffix = dest.suffix
    counter = 2
    while True:
        candidate = folder / f"{stem}_{counter}{suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


def normalize_relative_path(relative: str) -> str:
    cleaned = relative.replace("\\", "/").strip()
    while cleaned.startswith("./"):
        cleaned = cleaned[2:]
    return cleaned.lstrip("/")


def copy_multimedia_for_row(
    cell_value: object,
    multimedia_root: Path,
    dest_folder: Path,
) -> MultimediaCopyResult:
    result = MultimediaCopyResult()
    root = multimedia_root.resolve()
    for relative in parse_saved_file_names(cell_value):
        key = normalize_relative_path(relative)
        if not key:
            continue
        source = resolve_multimedia_path(root, key)
        try:
            source.relative_to(root)
        except ValueError:
            result.skipped.append(f"Multimedia 밖 경로 스킵: {relative}")
            continue
        if not source.is_file():
            result.skipped.append(f"파일 없음: {relative}")
            continue

        dest_folder.mkdir(parents=True, exist_ok=True)
        dest = unique_dest_path(dest_folder, source.name)
        try:
            shutil.copy2(source, dest)
        except OSError as exc:
            result.skipped.append(f"복사 실패 ({source.name}): {exc}")
            continue

        result.copied.append(dest.name)
        result.path_map[key] = dest.name
        # basename만 있는 참조도 매칭되도록
        result.path_map.setdefault(source.name, dest.name)
        if is_image_file(dest):
            result.image_paths.append(dest)

    return result
