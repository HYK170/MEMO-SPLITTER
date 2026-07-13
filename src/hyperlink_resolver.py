from __future__ import annotations

import re
from pathlib import Path
from urllib.parse import unquote, urlparse

REMOTE_SCHEMES = {"http", "https", "mailto", "ftp", "ftps"}


def resolve_local_path(target: str | None, base_dir: Path) -> Path | None:
    if not target:
        return None

    raw = target.strip()
    if not raw:
        return None

    if raw.startswith("#"):
        return None

    if re.match(r"^[A-Za-z]:[\\/]", raw):
        candidate = Path(raw).resolve()
        return candidate if candidate.is_file() else None

    if raw.startswith("\\\\"):
        candidate = Path(raw).resolve()
        return candidate if candidate.is_file() else None

    parsed = urlparse(raw)
    if parsed.scheme.lower() in REMOTE_SCHEMES:
        return None

    if parsed.scheme.lower() == "file":
        path_text = unquote(parsed.path or "")
        if re.match(r"^/[A-Za-z]:", path_text):
            path_text = path_text[1:]
        candidate = Path(path_text)
    elif parsed.scheme and parsed.scheme.lower() not in {"", "file"}:
        return None
    else:
        candidate = Path(raw)

    if not candidate.is_absolute():
        candidate = (base_dir / candidate).resolve()
    else:
        candidate = candidate.resolve()

    if candidate.is_file():
        return candidate
    return None
