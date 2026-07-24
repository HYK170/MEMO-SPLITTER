from __future__ import annotations

import os
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from src.filename_builder import (
    build_html_attachments_folder_name,
    build_html_filename,
)
from src.html_table import Cell, build_split_html, is_row_empty, parse_first_table
from src.multimedia_copier import (
    copy_paths_from_base,
    normalize_relative_path,
    resolve_path_from_base,
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


def path_relative_to(from_dir: Path, target: Path) -> str:
    """from_dir 기준 target 상대경로(posix)."""
    try:
        return Path(os.path.relpath(target.resolve(), start=from_dir.resolve())).as_posix()
    except ValueError:
        return target.resolve().as_posix()


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
    col_count = len(table.header_cells_html)
    existing_html_names: set[str] = set()
    log(f"처리 대상 행: {total}")

    for row_index, cells in enumerate(data_rows, start=1):
        if on_progress:
            on_progress(row_index, total)

        padded = _pad_cells(cells, col_count)
        body_value = padded[column_map["본문"]].text
        html_name = build_html_filename(
            base_name, row_index, body_value, existing_names=existing_html_names
        )
        attach_folder_name = build_html_attachments_folder_name(html_name)
        html_path = output_root / html_name
        # a href만 복사 (img src 썸네일은 복사하지 않음)
        href_paths = extract_href_only_paths("".join(cell.html for cell in padded))

        attachments_folder = output_root / attach_folder_name
        copy_result = copy_paths_from_base(href_paths, base_dir, attachments_folder)
        result.attachments_copied += len(copy_result.copied)
        result.attachment_skips.extend(copy_result.skipped)

        head_extra = prepare_head_assets(
            stylesheet_hrefs,
            style_blocks,
            base_dir,
            output_root,
        )

        try:
            data_html = [
                rewrite_cell_urls(
                    cell.html,
                    attach_folder_name=attach_folder_name,
                    path_map=copy_result.path_map,
                    base_dir=base_dir,
                    row_folder=output_root,
                )
                for cell in padded
            ]
            content = build_split_html(
                data_html,
                table_open_tag=table.table_open_tag,
                colgroup_html=table.colgroup_html,
                thead_html=table.thead_html,
                head_extra=head_extra,
            )
            html_path.write_text(content, encoding="utf-8")
        except OSError as exc:
            message = f"행 {row_index}: HTML 저장 실패 ({html_name}) - {exc}"
            result.row_errors.append(message)
            log(message)
            continue

        result.folders_created += 1
        log(f"행 {row_index}: {html_name} 생성")
        for name in copy_result.copied:
            log(f"  첨부 복사: {name}")
        for skipped in copy_result.skipped:
            log(f"  첨부 스킵: {skipped}")

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
        if href in seen:
            continue
        seen.add(href)
        hrefs.append(href)

    styles = [m.group(1) for m in _STYLE_TAG_RE.finditer(html_text or "")]
    return hrefs, styles


def prepare_head_assets(
    stylesheet_hrefs: list[str],
    style_blocks: list[str],
    base_dir: Path,
    row_folder: Path,
) -> str:
    """CSS는 복사하지 않고, split HTML 기준 원본 파일 상대경로만 넣는다."""
    parts: list[str] = []
    for href in stylesheet_hrefs:
        target = resolve_path_from_base(base_dir, href)
        rel = path_relative_to(row_folder, target)
        parts.append(f'<link rel="stylesheet" href="{rel}">')

    for block in style_blocks:
        rewritten = rewrite_css_urls_to_original(block, base_dir, row_folder)
        parts.append(f"<style>\n{rewritten}\n</style>")

    return "\n".join(parts)


def rewrite_css_urls_to_original(
    css_text: str,
    css_base_dir: Path,
    row_folder: Path,
) -> str:
    """CSS url(...)을 원본 자산 상대경로로만 다시 쓴다(복사 없음)."""

    def replacer(match: re.Match[str]) -> str:
        raw = match.group("quoted")
        quote = match.group(1) or ""
        if raw is None:
            raw = (match.group("unquoted") or "").strip()
            quote = ""
        original = raw.strip()
        if not original or _is_external_or_special_href(original):
            return match.group(0)
        target = resolve_path_from_base(css_base_dir, original)
        rel = path_relative_to(row_folder, target)
        if quote:
            return f"url({quote}{rel}{quote})"
        return f"url({rel})"

    return _CSS_URL_RE.sub(replacer, css_text)


def extract_href_only_paths(cell_html: str) -> list[str]:
    """셀 HTML의 a href 로컬 경로만 추출한다(img src 제외)."""
    paths: list[str] = []
    seen: set[str] = set()
    for match in _ATTR_URL_RE.finditer(cell_html or ""):
        if match.group("attr").lower() != "href":
            continue
        raw = match.group("quoted")
        if raw is None:
            raw = match.group("unquoted") or ""
        href = raw.strip()
        if not href or _is_external_or_special_href(href):
            continue
        key = _path_key(href)
        if not key or key in seen:
            continue
        seen.add(key)
        paths.append(key)
    return paths


def extract_local_paths(cell_html: str) -> list[str]:
    """하위 호환: href/src 로컬 경로 추출."""
    paths: list[str] = []
    seen: set[str] = set()
    for match in _ATTR_URL_RE.finditer(cell_html or ""):
        raw = match.group("quoted")
        if raw is None:
            raw = match.group("unquoted") or ""
        href = raw.strip()
        if not href or _is_external_or_special_href(href):
            continue
        key = _path_key(href)
        if not key or key in seen:
            continue
        seen.add(key)
        paths.append(key)
    return paths


def extract_href_paths(cell_html: str) -> list[str]:
    return extract_href_only_paths(cell_html)


def rewrite_cell_urls(
    html: str,
    *,
    attach_folder_name: str,
    path_map: dict[str, str],
    base_dir: Path,
    row_folder: Path,
) -> str:
    """
    href: 복사된 첨부는 attach 경로, 그 외/원본 참조는 INPUT 기준 상대경로.
    src(썸네일): 복사하지 않고 INPUT 기준 상대경로만 사용.
    """
    if not html:
        return html

    def replacer(match: re.Match[str]) -> str:
        attr = match.group("attr")
        raw = match.group("quoted")
        if raw is None:
            raw = match.group("unquoted") or ""
        original = raw.strip()
        if not original or _is_external_or_special_href(original):
            return match.group(0)

        quote = match.group("q")
        attr_lower = attr.lower()

        if attr_lower == "href":
            dest_name = _lookup_path_map(original, path_map)
            if dest_name is not None:
                new_url = f"{attach_folder_name}/{dest_name}".replace("\\", "/")
            else:
                new_url = path_relative_to(row_folder, resolve_path_from_base(base_dir, original))
        else:
            # img src 등: 원본 경로 참조
            new_url = path_relative_to(row_folder, resolve_path_from_base(base_dir, original))

        if quote:
            return f"{attr}={quote}{new_url}{quote}"
        return f"{attr}={new_url}"

    return _ATTR_URL_RE.sub(replacer, html)


def rewrite_local_urls(
    html: str,
    attach_folder_name: str,
    path_map: dict[str, str],
) -> str:
    """테스트 하위 호환용: 복사 맵 기준 href/src 재작성."""
    if not html or not path_map:
        return html

    def replacer(match: re.Match[str]) -> str:
        raw = match.group("quoted")
        if raw is None:
            raw = match.group("unquoted") or ""
        original = raw.strip()
        if not original or _is_external_or_special_href(original):
            return match.group(0)
        dest_name = _lookup_path_map(original, path_map)
        if dest_name is None:
            return match.group(0)
        new_url = f"{attach_folder_name}/{dest_name}".replace("\\", "/")
        attr = match.group("attr")
        quote = match.group("q")
        if quote:
            return f"{attr}={quote}{new_url}{quote}"
        return f"{attr}={new_url}"

    return _ATTR_URL_RE.sub(replacer, html)


def _path_key(path_str: str) -> str:
    if Path(path_str).is_absolute():
        return path_str.replace("\\", "/")
    return normalize_relative_path(path_str)


def _lookup_path_map(original: str, path_map: dict[str, str]) -> str | None:
    candidates = [
        original,
        original.replace("\\", "/"),
        normalize_relative_path(original),
        Path(original).name,
    ]
    for key in candidates:
        if key in path_map:
            return path_map[key]
    return None


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
