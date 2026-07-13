import re
from pathlib import Path

TITLE_PREFIX = "제목 : "
INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
MAX_FILENAME_LEN = 50


def parse_title(body: str | None) -> str:
    if body is None:
        return "제목없음"
    text = str(body).strip()
    if text.startswith(TITLE_PREFIX):
        title = text[len(TITLE_PREFIX) :].strip()
        return title or "제목없음"
    return "제목없음"


def sanitize_windows_filename(name: str) -> str:
    cleaned = INVALID_FILENAME_CHARS.sub("_", name).strip(" .")
    return cleaned or "output"


def build_row_folder_name(base_name: str, row_index: int) -> str:
    row_no = str(row_index).zfill(3)
    return sanitize_windows_filename(f"{base_name}_{row_no}")


def build_xlsx_filename(
    base_name: str,
    app: str | None,
    row_index: int,
    body: str | None,
    existing_names: set[str] | None = None,
) -> str:
    app_text = sanitize_windows_filename(str(app).strip() if app is not None else "")
    row_no = str(row_index).zfill(3)
    title = sanitize_windows_filename(parse_title(body))
    filename = f"{base_name}_{app_text}_{row_no}_{title}.xlsx"
    filename = sanitize_windows_filename(filename)

    if len(filename) > MAX_FILENAME_LEN:
        stem = filename[: MAX_FILENAME_LEN - 5]
        filename = f"{stem}.xlsx"

    if existing_names is not None:
        filename = _dedupe_filename(filename, existing_names)

    return filename


def _dedupe_filename(filename: str, existing_names: set[str]) -> str:
    if filename not in existing_names:
        existing_names.add(filename)
        return filename

    stem = Path(filename).stem
    suffix = Path(filename).suffix
    counter = 2
    while True:
        candidate = f"{stem}_{counter}{suffix}"
        if len(candidate) > MAX_FILENAME_LEN:
            trim = MAX_FILENAME_LEN - len(suffix) - len(f"_{counter}")
            candidate = f"{stem[:trim]}_{counter}{suffix}"
        if candidate not in existing_names:
            existing_names.add(candidate)
            return candidate
        counter += 1


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
