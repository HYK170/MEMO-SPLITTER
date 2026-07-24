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
from src.html_splitter import HtmlSplitConfig, extract_href_paths, split_html
from src.html_table import build_split_html, is_row_empty, parse_first_table


def _tiny_png_bytes() -> bytes:
    from io import BytesIO

    buffer = BytesIO()
    PILImage.new("RGB", (1, 1), (255, 0, 0)).save(buffer, format="PNG")
    return buffer.getvalue()


TINY_PNG = _tiny_png_bytes()

SAMPLE_HTML = """<!DOCTYPE html>
<html>
<body>
<table>
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
        <a href="images/shot.png">shot.png</a><br>
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
      <td><a href="images/shot.png"><img src="images/shot.png"></a></td>
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


def test_extract_href_paths() -> None:
    html = (
        '<a href="images/shot.png">a</a>'
        '<a href="./docs/memo.txt">b</a>'
        '<a href="https://example.com/x.png">c</a>'
        '<a href="#top">d</a>'
        '<a href="images/shot.png">dup</a>'
    )
    assert extract_href_paths(html) == ["images/shot.png", "docs/memo.txt"]


def test_is_row_empty() -> None:
    table = parse_first_table(SAMPLE_HTML)
    assert is_row_empty(table.rows[1])
    assert not is_row_empty(table.rows[0])


def test_build_html_filename() -> None:
    name = build_html_filename("memo", 1, "제목 : 회의록\n본문")
    assert name == "memo_001_회의록.html"


def test_build_split_html_preserves_markup() -> None:
    out = build_split_html(["첨부파일"], ['<a href="a.png"><img src="a.png"></a>'])
    assert "<th>첨부파일</th>" in out
    assert '<td><a href="a.png"><img src="a.png"></a></td>' in out
    assert "<!DOCTYPE html>" in out


def test_split_html_integration() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        input_path = root / "memo.html"
        multimedia = root / "Multimedia"
        (multimedia / "images").mkdir(parents=True)
        (multimedia / "docs").mkdir(parents=True)
        (multimedia / "images" / "shot.png").write_bytes(TINY_PNG)
        (multimedia / "docs" / "memo.txt").write_text("hello", encoding="utf-8")
        input_path.write_text(SAMPLE_HTML, encoding="utf-8")

        result = split_html(
            HtmlSplitConfig(input_path=input_path, multimedia_root=multimedia),
            on_log=print,
        )

        assert result.output_root is not None
        assert result.output_root.parent == root
        assert result.output_root.name.startswith("memo_")
        assert result.folders_created == 3
        assert result.attachments_copied == 3
        assert result.images_embedded == 0
        assert result.rows_skipped == 1
        assert not result.attachment_skips
        assert not result.row_errors

        folder1 = result.output_root / "memo_001"
        attachments1 = folder1 / "memo_001_attach"
        assert (attachments1 / "shot.png").is_file()
        assert (attachments1 / "memo.txt").is_file()

        first_html = folder1 / "memo_001_회의록.html"
        assert first_html.is_file()
        content = first_html.read_text(encoding="utf-8")
        assert "<th>App</th>" in content
        assert "<td>Kakao</td>" in content
        assert 'href="images/shot.png"' in content
        assert content.count("<tr>") == 2

        folder2 = result.output_root / "memo_002"
        assert (folder2 / "memo_002_제목없음.html").is_file()


def main() -> None:
    test_parse_first_table_thead()
    test_parse_th_without_thead()
    test_extract_href_paths()
    test_is_row_empty()
    test_build_html_filename()
    test_build_split_html_preserves_markup()
    test_split_html_integration()
    print("ALL HTML TESTS PASSED")


if __name__ == "__main__":
    main()
