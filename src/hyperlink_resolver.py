from __future__ import annotations

import re
from pathlib import Path
from urllib.parse import unquote, urlparse

REMOTE_SCHEMES = {"http", "https", "mailto", "ftp", "ftps"}


def is_non_file_link(raw: str) -> bool:
    lowered = raw.lower()
    return raw.startswith("#") or lowered.startswith(("http://", "https://", "mailto:", "ftp://", "ftps://"))


def normalize_hyperlink_target(raw: str) -> str:
    text = unquote(raw.strip())
    if not text or is_non_file_link(text):
        return text

    parsed = urlparse(text)
    if parsed.scheme.lower() == "file":
        path_text = unquote(parsed.path or "")
        if re.match(r"^/[A-Za-z]:", path_text):
            path_text = path_text[1:]
        elif re.match(r"^/[A-Za-z]\|", path_text):
            path_text = path_text[1].replace("|", ":") + path_text[2:]
        text = path_text

    text = re.sub(r"^([A-Za-z])\|", r"\1:", text)
    return text


def _build_path_candidates(target: str | None, base_dir: Path) -> list[Path]:
    if not target:
        return []

    raw = target.strip()
    if not raw or is_non_file_link(raw):
        return []

    variants: list[str] = []
    seen: set[str] = set()

    def add_variant(value: str) -> None:
        cleaned = value.strip()
        if cleaned and cleaned not in seen:
            seen.add(cleaned)
            variants.append(cleaned)

    normalized = normalize_hyperlink_target(raw)
    add_variant(normalized)
    add_variant(unquote(raw))
    add_variant(raw.replace("%20", " "))
    add_variant(normalized.replace("/", "\\"))
    add_variant(normalized.replace("\\", "/"))

    candidates: list[Path] = []
    candidate_seen: set[str] = set()

    for variant in variants:
        path_obj = Path(variant)
        options = [path_obj]
        if not path_obj.is_absolute():
            options.append(base_dir / path_obj)

        for option in options:
            try:
                resolved = option.resolve()
            except OSError:
                continue
            key = str(resolved).lower()
            if key in candidate_seen:
                continue
            candidate_seen.add(key)
            candidates.append(resolved)

    return candidates


def resolve_local_path(target: str | None, base_dir: Path) -> Path | None:
    if not target:
        return None

    raw = target.strip()
    if not raw or is_non_file_link(raw):
        return None

    for candidate in _build_path_candidates(target, base_dir):
        if candidate.is_file():
            return candidate
    return None
