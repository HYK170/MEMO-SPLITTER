"""Manual integration test helper for MEMO SPLITTER."""

from __future__ import annotations

import sys
import tempfile
import zipfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from openpyxl import Workbook, load_workbook

from src.filename_builder import parse_title
from src.multimedia_copier import parse_saved_file_names
from src.splitter import SplitConfig, split_workbook

TINY_PNG = (
    b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01"
    b"\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDATx\x9cc``\x00\x00"
    b"\x00\x02\x00\x01\xe2!\xbc3\x00\x00\x00\x00IEND\xaeB`\x82"
)


def test_parse_title_first_line_only() -> None:
    title = parse_title("제목 : 회의록\n본문 내용이 길게 이어짐")
    assert title == "회의록", f"expected '회의록', got {title!r}"


def test_parse_saved_file_names() -> None:
    paths = parse_saved_file_names("a/photo.jpg\nb/note.txt\n\n c/doc.pdf ")
    assert paths == ["a/photo.jpg", "b/note.txt", "c/doc.pdf"]


def create_sample_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    headers = ["App", "본문", "저장된 파일 이름", "첨부파일"]
    for col, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col, value=header)

    rows = [
        ("Kakao", "제목 : 회의록\n본문", "images/shot.png\ndocs/memo.txt", ""),
        ("Line", "본문만 있음", "", ""),
        ("Slack", "제목 : 긴급 공지\n상세", "images/shot.png", ""),
    ]
    for idx, (app, body, saved, attach) in enumerate(rows, start=2):
        ws.cell(row=idx, column=1, value=app)
        ws.cell(row=idx, column=2, value=body)
        ws.cell(row=idx, column=3, value=saved)
        ws.cell(row=idx, column=4, value=attach)

    wb.save(path)
    wb.close()


def main() -> None:
    test_parse_title_first_line_only()
    test_parse_saved_file_names()

    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        input_path = root / "memo.xlsx"
        output_root = root / "output"
        multimedia = root / "Multimedia"
        output_root.mkdir()
        (multimedia / "images").mkdir(parents=True)
        (multimedia / "docs").mkdir(parents=True)
        (multimedia / "images" / "shot.png").write_bytes(TINY_PNG)
        (multimedia / "docs" / "memo.txt").write_text("hello", encoding="utf-8")

        create_sample_workbook(input_path)

        result = split_workbook(
            SplitConfig(
                input_path=input_path,
                output_root=output_root,
                multimedia_root=multimedia,
                sheet_name="Sheet1",
                header_row=1,
            ),
            on_log=print,
        )

        print("RESULT", result)
        assert result.folders_created == 3
        assert result.attachments_copied == 3
        assert result.images_embedded == 2
        assert not result.attachment_skips
        assert not result.row_errors

        folder1 = output_root / "memo_001"
        assert (folder1 / "shot.png").is_file()
        assert (folder1 / "memo.txt").is_file()

        first_xlsx = next(folder1.glob("*.xlsx"))
        check_wb = load_workbook(first_xlsx)
        check_ws = check_wb.active
        assert check_ws.cell(1, 1).value == "App"
        assert check_ws.cell(2, 1).value == "Kakao"
        assert check_ws.cell(3, 1).value is None
        assert len(getattr(check_ws, "_images", [])) == 1
        check_wb.close()

        with zipfile.ZipFile(first_xlsx) as z:
            assert any(n.startswith("xl/media/") for n in z.namelist())

        for folder in sorted(output_root.iterdir()):
            print("FOLDER", folder.name, [p.name for p in folder.iterdir()])

    print("ALL TESTS PASSED")


if __name__ == "__main__":
    main()
