from __future__ import annotations

import re
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from src.filename_builder import (
    build_attachments_folder_name,
    build_html_filename,
    build_row_folder_name,
)
from src.html_table import Cell, build_split_html, is_row_empty, parse_first_table
from src.multimedia_copier import (
    copy_paths_from_base,
    normalize_relative_path,
    resolve_path_from_base,
    unique_dest_path,
)
from src.splitter import SplitResult, build_memo_output_root

LogCallback = Callable[[str], None]
ProgressCallback = Callable[[int, int], None]

HTML_REQUIRED_COLUMNS = ("App", "본문", "첨부파일")
_HEADER_WHITESPACE = re.compile(r"\s+", re.UNICODE)
_ATTR_URL_RE = re.compile(
    r"""(?P<attr>href|src)\s*=\s*(?:(?P<q>["'])(?P<quoted>.*?)(?P=q)|(?P<unquoted>[^\s>]+))""",
    re.IGNORECASE | re.DOTALL,
)
_LINK_TAG_RE = re.compile(r"<link\b[^>]*>", re.IGNORECASE)
_STYLE_TAG_RE = re.compile(r"<style\b[^>]*>(.*?)</style>", re.IGNORECASE | re.DOTALL)
_TAG_ATTR_RE = re.compile(
    r"""(?P<name>[^\s=]+)\s*=\s*(?:(["'])(?P<qval>.*?)\2|(?P<uval>[^\s>]+))""",
    re.IGNORECASE | re.DOTALL,
)
_CSS_URL_RE = re.compile(
    r"""url\(\s*(?:(["'])(?P<quoted>.*?)\1|(?P<unquoted>[^'")]+))\s*\)""",
    re.IGNORECASE | re.DOTALL,
)


@dataclass
class HtmlSplitConfig:
    input_path: Path


def validate_html_config(config: HtmlSplitConfig) -> None:
    if not config.input_path.is_file():
        raise FileNotFoundError(f"INPUT HTML을 찾을 수 없습니다: {config.input_path}")
    if config.input_path.suffix.lower() not in {".html", ".htm"}:
        raise ValueError("INPUT 파일은 .html 또는 .htm 확장자여야 합니다.")


def _read_html_text(path: Path) -> str:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "cp949", "euc-kr", "utf-16"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def split_html(
    config: HtmlSplitConfig,
    on_log: LogCallback | None = None,
    on_progress: ProgressCallback | None = None,
) -> SplitResult:
    validate_html_config(config)

    def log(message: str) -> None:
        if on_log:
            on_log(message)

    output_root = build_memo_output_root(config.input_path)
    output_root.mkdir(parents=True, exist_ok=False)
    log(f"OUTPUT: {output_root}")
    log(f"Python: {sys.executable}")

    base_dir = config.input_path.resolve().parent
    html_text = _read_html_text(config.input_path)
    table = parse_first_table(html_text)
    column_map = _build_column_map(table.headers)
    _ensure_required_columns(column_map)
    stylesheet_hrefs, style_blocks = extract_styles(html_text)
    log(f"헤더: {', '.join(table.headers)}")
    log(f"스타일시트: {len(stylesheet_hrefs)}개, inline style: {len(style_blocks)}개")
    log(f"데이터 행(빈 행 제외 전): {len(table.rows)}")

    data_rows = [row for row in table.rows if not is_row_empty(row)]
    result = SplitResult(output_root=output_root)
    result.rows_skipped = len(table.rows) - len(data_rows)
    total = len(data_rows)
    base_name = config.input_path.stem
    header_html = table.header_cells_html
    col_count = len(header_html)
    log(f"처리 대상 행: {total}")

    for row_index, cells in enumerate(data_rows, start=1):
        if on_progress:
            on_progress(row_index, total)

        folder_name = build_row_folder_name(base_name, row_index)
        row_folder = output_root / folder_name

        try:
            row_folder.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            message = f"행 {row_index}: 폴더 생성 실패 ({row_folder}) - {exc}"
            result.row_errors.append(message)
            log(message)
            continue

        padded = _pad_cells(cells, col_count)
        body_value = padded[column_map["본문"]].text
        attach_folder_name = build_attachments_folder_name(base_name, row_index)
        relative_paths = extract_local_paths("".join(cell.html for cell in padded))
        html_name = build_html_filename(base_name, row_index, body_value)
        html_path = row_folder / html_name

        attachments_folder = row_folder / attach_folder_name
        copy_result = copy_paths_from_base(relative_paths, base_dir, attachments_folder)
        result.attachments_copied += len(copy_result.copied)
        result.attachment_skips.extend(copy_result.skipped)

        head_extra, css_copied, css_skips = prepare_head_assets(
            stylesheet_hrefs,
            style_blocks,
            base_dir,
            row_folder,
            attach_folder_name,
        )
        result.attachments_copied += css_copied
        result.attachment_skips.extend(css_skips)

        try:
            data_html = [
                rewrite_local_urls(cell.html, attach_folder_name, copy_result.path_map)
                for cell in padded
            ]
            content = build_split_html(header_html, data_html, head_extra=head_extra)
            html_path.write_text(content, encoding="utf-8")
        except OSError as exc:
            message = f"행 {row_index}: HTML 저장 실패 ({html_path.name}) - {exc}"
            result.row_errors.append(message)
            log(message)
            continue

        result.folders_created += 1
        log(f"행 {row_index}: {row_folder.name} 생성 ({html_name})")
        for name in copy_result.copied:
            log(f"  첨부 복사: {name}")
        for skipped in copy_result.skipped:
            log(f"  첨부 스킵: {skipped}")
        for skipped in css_skips:
            log(f"  CSS 스킵: {skipped}")

    return result


def extract_styles(html_text: str) -> tuple[list[str], list[str]]:
    """원본 HTML에서 stylesheet href와 style 블록 내용을 추출한다."""
    hrefs: list[str] = []
    seen: set[str] = set()
    for tag in _LINK_TAG_RE.findall(html_text or ""):
        attrs = _parse_tag_attrs(tag)
        rel = attrs.get("rel", "").lower()
        href = attrs.get("href", "").strip()
        if "stylesheet" not in rel or not href:
            continue
        if _is_external_or_special_href(href):
            continue
        key = href
        if key in seen:
            continue
        seen.add(key)
        hrefs.append(href)

    styles = [m.group(1) for m in _STYLE_TAG_RE.finditer(html_text or "")]
    return hrefs, styles


def prepare_head_assets(
    stylesheet_hrefs: list[str],
    style_blocks: list[str],
    base_dir: Path,
    row_folder: Path,
    attach_folder_name: str,
) -> tuple[str, int, list[str]]:
    """CSS 파일을 행 폴더에 복사하고 head용 link/style 마크업을 만든다."""
    parts: list[str] = []
    copied = 0
    skips: list[str] = []
    attach_folder = row_folder / attach_folder_name

    for href in stylesheet_hrefs:
        source = resolve_path_from_base(base_dir, href)
        if not source.is_file():
            skips.append(f"CSS 파일 없음: {href}")
            continue
        try:
            attach_folder.mkdir(parents=True, exist_ok=True)
            dest = unique_dest_path(attach_folder, source.name)
            css_text = source.read_text(encoding="utf-8", errors="replace")
            rewritten, asset_copied, asset_skips = rewrite_css_urls(
                css_text,
                source.parent,
                attach_folder,
            )
            dest.write_text(rewritten, encoding="utf-8")
            copied += 1 + asset_copied
            skips.extend(asset_skips)
            rel = f"{attach_folder_name}/{dest.name}".replace("\\", "/")
            parts.append(f'<link rel="stylesheet" href="{rel}">')
        except OSError as exc:
            skips.append(f"CSS 복사 실패 ({href}): {exc}")

    for block in style_blocks:
        rewritten, asset_copied, asset_skips = rewrite_css_urls(
            block,
            base_dir,
            attach_folder,
        )
        copied += asset_copied
        skips.extend(asset_skips)
        # style 안 url()은 attach 폴더 기준 상대경로로 맞춤
        rewritten = _prefix_css_urls_with_attach(rewritten, attach_folder_name)
        parts.append(f"<style>\n{rewritten}\n</style>")

    return "\n".join(parts), copied, skips


def rewrite_css_urls(
    css_text: str,
    css_base_dir: Path,
    dest_folder: Path,
) -> tuple[str, int, list[str]]:
    """CSS url(...) 로컬 자산을 dest_folder로 복사하고 파일명만 남긴다."""
    copied = 0
    skips: list[str] = []
    path_map: dict[str, str] = {}

    def replacer(match: re.Match[str]) -> str:
        nonlocal copied
        raw = match.group("quoted")
        if raw is None:
            raw = (match.group("unquoted") or "").strip()
        original = raw.strip()
        if not original or _is_external_or_special_href(original):
            return match.group(0)
        if original in path_map:
            return f"url({path_map[original]})"

        source = resolve_path_from_base(css_base_dir, original)
        if not source.is_file():
            skips.append(f"CSS 자산 없음: {original}")
            return match.group(0)
        try:
            dest_folder.mkdir(parents=True, exist_ok=True)
            dest = unique_dest_path(dest_folder, source.name)
            shutil.copy2(source, dest)
        except OSError as exc:
            skips.append(f"CSS 자산 복사 실패 ({original}): {exc}")
            return match.group(0)

        path_map[original] = dest.name
        copied += 1
        quote = match.group(1) or ""
        if quote:
            return f"url({quote}{dest.name}{quote})"
        return f"url({dest.name})"

    return _CSS_URL_RE.sub(replacer, css_text), copied, skips


def _prefix_css_urls_with_attach(css_text: str, attach_folder_name: str) -> str:
    """inline style의 url(file)을 url(attach/file)로 바꾼다."""

    def replacer(match: re.Match[str]) -> str:
        raw = match.group("quoted")
        quote = match.group(1) or ""
        if raw is None:
            raw = (match.group("unquoted") or "").strip()
            quote = ""
        original = raw.strip()
        if not original or _is_external_or_special_href(original):
            return match.group(0)
        if "/" in original.replace("\\", "/") or original.startswith(attach_folder_name):
            return match.group(0)
        new_url = f"{attach_folder_name}/{original}".replace("\\", "/")
        if quote:
            return f"url({quote}{new_url}{quote})"
        return f"url({new_url})"

    return _CSS_URL_RE.sub(replacer, css_text)


def extract_local_paths(cell_html: str) -> list[str]:
    """셀 HTML의 href/src에서 로컬 경로를 추출한다."""
    paths: list[str] = []
    seen: set[str] = set()
    for match in _ATTR_URL_RE.finditer(cell_html or ""):
        raw = match.group("quoted")
        if raw is None:
            raw = match.group("unquoted") or ""
        href = raw.strip()
        if not href or _is_external_or_special_href(href):
            continue
        if Path(href).is_absolute():
            key = href.replace("\\", "/")
        else:
            key = normalize_relative_path(href)
        if not key or key in seen:
            continue
        seen.add(key)
        paths.append(key)
    return paths


def extract_href_paths(cell_html: str) -> list[str]:
    """하위 호환: href/src 로컬 경로 추출."""
    return extract_local_paths(cell_html)


def rewrite_local_urls(
    html: str,
    attach_folder_name: str,
    path_map: dict[str, str],
) -> str:
    """복사된 첨부 기준으로 href/src 상대경로를 행 폴더 기준으로 다시 쓴다."""
    if not html or not path_map:
        return html

    def replacer(match: re.Match[str]) -> str:
        raw = match.group("quoted")
        if raw is None:
            raw = match.group("unquoted") or ""
        original = raw.strip()
        if not original or _is_external_or_special_href(original):
            return match.group(0)

        candidates = [
            original,
            original.replace("\\", "/"),
            normalize_relative_path(original),
            Path(original).name,
        ]
        dest_name = None
        for key in candidates:
            if key in path_map:
                dest_name = path_map[key]
                break
        if dest_name is None:
            return match.group(0)

        new_url = f"{attach_folder_name}/{dest_name}".replace("\\", "/")
        attr = match.group("attr")
        quote = match.group("q")
        if quote:
            return f"{attr}={quote}{new_url}{quote}"
        return f"{attr}={new_url}"

    return _ATTR_URL_RE.sub(replacer, html)


def _parse_tag_attrs(tag: str) -> dict[str, str]:
    attrs: dict[str, str] = {}
    for match in _TAG_ATTR_RE.finditer(tag):
        name = match.group("name").lower()
        value = match.group("qval")
        if value is None:
            value = match.group("uval") or ""
        attrs[name] = value
    return attrs


def _is_external_or_special_href(href: str) -> bool:
    lower = href.strip().lower()
    if lower.startswith(("#", "mailto:", "tel:", "javascript:", "data:")):
        return True
    return "://" in lower


def _normalize_header(value: object) -> str:
    return _HEADER_WHITESPACE.sub(" ", str(value).strip())


def _build_column_map(headers: list[str]) -> dict[str, int]:
    column_map: dict[str, int] = {}
    for idx, header in enumerate(headers):
        if not header:
            continue
        column_map[_normalize_header(header)] = idx
    return column_map


def _find_column(column_map: dict[str, int], name: str) -> int | None:
    normalized = _normalize_header(name)
    if normalized in column_map:
        return column_map[normalized]
    compact = normalized.replace(" ", "")
    for key, col in column_map.items():
        if key.replace(" ", "") == compact:
            return col
    return None


def _ensure_required_columns(column_map: dict[str, int]) -> None:
    missing = [name for name in HTML_REQUIRED_COLUMNS if _find_column(column_map, name) is None]
    if missing:
        raise ValueError(f"필수 컬럼이 없습니다: {', '.join(missing)}")
    for name in HTML_REQUIRED_COLUMNS:
        col = _find_column(column_map, name)
        if col is not None:
            column_map[name] = col


def _pad_cells(cells: list[Cell], col_count: int) -> list[Cell]:
    padded = list(cells)
    while len(padded) < col_count:
        padded.append(Cell(text="", html=""))
    return padded[:col_count]
