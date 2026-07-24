"""HTML split mode integration tests."""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from PIL import Image as PILImage

from src.filename_builder import build_html_filename
from src.html_splitter import (
    HtmlSplitConfig,
    extract_href_paths,
    extract_styles,
    rewrite_local_urls,
    split_html,
)
from src.html_table import build_split_html, is_row_empty, parse_first_table


def _tiny_png_bytes() -> bytes:
    from io import BytesIO

    buffer = BytesIO()
    PILImage.new("RGB", (1, 1), (255, 0, 0)).save(buffer, format="PNG")
    return buffer.getvalue()


TINY_PNG = _tiny_png_bytes()

SAMPLE_HTML = """<!DOCTYPE html>
<html>
<head>
  <link rel="stylesheet" href="css/app.css">
  <style>
    body { background: url(images/bg.png); }
  </style>
</head>
<body>
<table class="memo" border="1">
  <colgroup>
    <col style="width:10%">
    <col style="width:60%">
    <col style="width:30%">
  </colgroup>
  <thead>
    <tr>
      <th>App</th>
      <th>본문</th>
      <th>첨부파일</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>Kakao</td>
      <td>제목 : 회의록<br>본문</td>
      <td>
        <a href="images/shot.png"><img src="images/shot_thumb.png"></a><br>
        <a href="./docs/memo.txt">memo.txt</a>
      </td>
    </tr>
    <tr>
      <td></td>
      <td></td>
      <td></td>
    </tr>
    <tr>
      <td>Line</td>
      <td>본문만 있음</td>
      <td></td>
    </tr>
    <tr>
      <td>Slack</td>
      <td>제목 : 긴급 공지<br>상세</td>
      <td><a href="images/shot.png"><img src="images/shot_thumb.png"></a></td>
    </tr>
  </tbody>
</table>
</body>
</html>
"""


def test_parse_first_table_thead() -> None:
    table = parse_first_table(SAMPLE_HTML)
    assert table.headers == ["App", "본문", "첨부파일"]
    assert len(table.rows) == 4
    assert "회의록" in table.rows[0][1].text
    assert 'href="images/shot.png"' in table.rows[0][2].html
    assert "<colgroup>" in table.colgroup_html
    assert "<thead>" in table.thead_html
    assert 'class="memo"' in table.table_open_tag


def test_parse_omitted_end_tags() -> None:
    html = (
        "<table><tr><td>App<td>본문<td>첨부파일"
        "<tr><td>Kakao<td>제목 : 회의록<td><a href=\"images/shot.png\">f</a></table>"
    )
    table = parse_first_table(html)
    assert table.headers == ["App", "본문", "첨부파일"]
    assert len(table.rows) == 1
    assert table.rows[0][0].text == "Kakao"
    assert "images/shot.png" in table.rows[0][2].html


def test_parse_th_without_thead() -> None:
    html = """
    <table>
      <tr><th>App</th><th>본문</th><th>첨부파일</th></tr>
      <tr><td>A</td><td>제목 : T</td><td></td></tr>
    </table>
    """
    table = parse_first_table(html)
    assert table.headers == ["App", "본문", "첨부파일"]
    assert len(table.rows) == 1
    assert table.rows[0][0].text == "A"
    assert "<thead>" in table.thead_html


def test_extract_href_paths() -> None:
    html = (
        '<a href="images/shot.png"><img src="images/thumb.png"></a>'
        '<a href="./docs/memo.txt">b</a>'
        '<a href="https://example.com/x.png">c</a>'
        '<a href="#top">d</a>'
        '<a href="images/shot.png">dup</a>'
    )
    # img src는 제외, href만
    assert extract_href_paths(html) == ["images/shot.png", "docs/memo.txt"]


def test_extract_styles() -> None:
    hrefs, styles = extract_styles(SAMPLE_HTML)
    assert hrefs == ["css/app.css"]
    assert len(styles) == 1
    assert "url(images/bg.png)" in styles[0]


def test_is_row_empty() -> None:
    table = parse_first_table(SAMPLE_HTML)
    assert is_row_empty(table.rows[1])
    assert not is_row_empty(table.rows[0])


def test_build_html_filename() -> None:
    name = build_html_filename("memo", 1, "제목 : 회의록\n본문")
    assert name == "memo_001_회의록.html"


def test_rewrite_local_urls() -> None:
    html = '<a href="images/shot.png"><img src="./images/shot.png"></a>'
    out = rewrite_local_urls(html, "memo_001_attach", {"images/shot.png": "shot.png"})
    assert 'href="memo_001_attach/shot.png"' in out
    assert 'src="memo_001_attach/shot.png"' in out


def test_build_split_html_structure() -> None:
    out = build_split_html(
        ['<a href="a.png"><img src="t.png"></a>'],
        table_open_tag='<table class="memo">',
        colgroup_html="<colgroup><col></colgroup>",
        thead_html="<thead><tr><th>첨부파일</th></tr></thead>",
        head_extra='<link rel="stylesheet" href="../css/app.css">',
    )
    assert '<table class="memo">' in out
    assert "<colgroup><col></colgroup>" in out
    assert "<thead><tr><th>첨부파일</th></tr></thead>" in out
    assert "<tbody>" in out
    assert 'href="../css/app.css"' in out
    assert "<!DOCTYPE html>" in out


def test_split_html_integration() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        input_path = root / "memo.html"
        (root / "images").mkdir(parents=True)
        (root / "docs").mkdir(parents=True)
        (root / "css").mkdir(parents=True)
        (root / "images" / "shot.png").write_bytes(TINY_PNG)
        (root / "images" / "shot_thumb.png").write_bytes(TINY_PNG)
        (root / "images" / "bg.png").write_bytes(TINY_PNG)
        (root / "docs" / "memo.txt").write_text("hello", encoding="utf-8")
        (root / "css" / "app.css").write_text(
            "td { color: red; background: url(../images/bg.png); }",
            encoding="utf-8",
        )
        input_path.write_text(SAMPLE_HTML, encoding="utf-8")

        result = split_html(
            HtmlSplitConfig(input_path=input_path),
            on_log=print,
        )

        assert result.output_root is not None
        assert result.folders_created == 3
        assert result.rows_skipped == 1
        assert not result.row_errors

        attachments1 = result.output_root / "memo_001_회의록_attach"
        # href 대상만 복사, 썸네일/CSS는 복사하지 않음
        assert (attachments1 / "shot.png").is_file()
        assert (attachments1 / "memo.txt").is_file()
        assert not (attachments1 / "shot_thumb.png").exists()
        assert not (attachments1 / "app.css").exists()
        assert not (attachments1 / "bg.png").exists()
        assert not (result.output_root / "memo_001").is_dir()

        first_html = result.output_root / "memo_001_회의록.html"
        content = first_html.read_text(encoding="utf-8")
        assert "<colgroup>" in content
        assert "<thead>" in content
        assert "<tbody>" in content
        assert 'class="memo"' in content
        assert 'href="memo_001_회의록_attach/shot.png"' in content
        assert 'href="memo_001_회의록_attach/memo.txt"' in content
        # 썸네일/CSS는 원본 상대경로 참조
        assert "shot_thumb.png" in content
        assert "attach/shot_thumb" not in content
        assert "app.css" in content and "memo_001_회의록_attach/app.css" not in content
        assert "images/bg.png" in content or "bg.png" in content

        content3 = (result.output_root / "memo_003_긴급 공지.html").read_text(encoding="utf-8")
        assert 'href="memo_003_긴급 공지_attach/shot.png"' in content3
        assert "shot_thumb.png" in content3
        assert (result.output_root / "memo_003_긴급 공지_attach" / "shot.png").is_file()


def main() -> None:
    test_parse_first_table_thead()
    test_parse_omitted_end_tags()
    test_parse_th_without_thead()
    test_extract_href_paths()
    test_extract_styles()
    test_rewrite_local_urls()
    test_is_row_empty()
    test_build_html_filename()
    test_build_split_html_structure()
    test_split_html_integration()
    print("ALL HTML TESTS PASSED")


if __name__ == "__main__":
    main()
