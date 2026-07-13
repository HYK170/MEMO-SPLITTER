"""Manual integration test helper for MEMO SPLITTER."""

from __future__ import annotations

import sys
import tempfile
from io import BytesIO
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles import Font
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.worksheet.hyperlink import Hyperlink

from src.filename_builder import parse_title
from src.hyperlink_resolver import resolve_local_path
from src.image_handler import image_matches_row
from src.drawing_image_loader import build_images_by_row, load_drawing_images
from src.image_handler import index_images_by_row
from src.splitter import SplitConfig, split_workbook


def test_parse_title_first_line_only() -> None:
    title = parse_title("제목 : 회의록\n본문 내용이 길게 이어짐")
    assert title == "회의록", f"expected '회의록', got {title!r}"


def test_resolve_local_path_encoded_and_mixed_slashes(tmp: Path) -> None:
    folder = tmp / "My Folder"
    folder.mkdir()
    attachment = folder / "sample file.txt"
    attachment.write_text("data", encoding="utf-8")

    encoded = str(attachment).replace(" ", "%20").replace("\\", "/")
    resolved = resolve_local_path(encoded, tmp)
    assert resolved == attachment.resolve(), f"failed for {encoded}"

    mixed = encoded.replace("/", "\\")
    resolved_mixed = resolve_local_path(mixed, tmp)
    assert resolved_mixed == attachment.resolve(), f"failed for {mixed}"


def create_sample_workbook(path: Path, attachment: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    headers = ["App", "본문", "Link"]
    for col, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col, value=header)

    rows = [
        ("Kakao", "제목 : 회의록\n본문 첫 줄 이후 내용", attachment),
        ("Line", "본문만 있음", None),
        ("Slack", "제목 : 긴급 공지\n상세 내용", attachment),
    ]

    for idx, (app, body, link_target) in enumerate(rows, start=2):
        ws.cell(row=idx, column=1, value=app)
        ws.cell(row=idx, column=2, value=body)
        if link_target is not None:
            cell = ws.cell(row=idx, column=3, value="첨부열기")
            encoded_target = str(link_target).replace(" ", "%20").replace("\\", "/")
            cell.hyperlink = Hyperlink(ref=cell.coordinate, target=encoded_target)
            cell.font = Font(color="0563C1", underline="single")

    png_path = path.parent / "tiny.png"
    png_path.write_bytes(
        b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01"
        b"\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDATx\x9cc``\x00\x00"
        b"\x00\x02\x00\x01\xe2!\xbc3\x00\x00\x00\x00IEND\xaeB`\x82"
    )
    image = Image(str(png_path))
    marker = AnchorMarker(col=0, row=3, colOff=0, rowOff=0)
    image.anchor = OneCellAnchor(_from=marker, ext=XDRPositiveSize2D(pixels_to_EMU(32), pixels_to_EMU(32)))
    ws.add_image(image)

    wb.save(path)
    wb.close()


def test_image_matches_row_one_cell_above() -> None:
    png_bytes = (
        b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01"
        b"\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDATx\x9cc``\x00\x00"
        b"\x00\x02\x00\x01\xe2!\xbc3\x00\x00\x00\x00IEND\xaeB`\x82"
    )

    # Samsung Notes 패턴: anchor가 Excel 3행이면 데이터 2행에도 매칭
    image = Image(BytesIO(png_bytes))
    marker = AnchorMarker(col=0, row=2, colOff=0, rowOff=0)
    image.anchor = OneCellAnchor(_from=marker, ext=XDRPositiveSize2D(pixels_to_EMU(32), pixels_to_EMU(32)))
    assert image_matches_row(image, 2), "anchor row 3 should match data row 2"
    assert image_matches_row(image, 3), "anchor row 3 should match itself"

    # 일반 oneCell: anchor 행과 인접 행(±1)에 매칭
    image_below = Image(BytesIO(png_bytes))
    marker_below = AnchorMarker(col=0, row=3, colOff=0, rowOff=0)
    image_below.anchor = OneCellAnchor(
        _from=marker_below, ext=XDRPositiveSize2D(pixels_to_EMU(32), pixels_to_EMU(32))
    )
    assert image_matches_row(image_below, 4), "anchor row 4 should match data row 4"
    assert not image_matches_row(image_below, 3), "anchor row 4 should not match row 3"


def test_drawing_loader_finds_saved_images(tmp: Path) -> None:
    input_path = tmp / "memo.xlsx"
    attachment = tmp / "sample attachment.txt"
    attachment.write_text("attachment content", encoding="utf-8")
    create_sample_workbook(input_path, attachment)

    from openpyxl import load_workbook

    wb = load_workbook(input_path, data_only=False)
    ws = wb.active
    ws._images = []
    wb.close()

    drawing_index = load_drawing_images(input_path, load_workbook(input_path).active)
    load_workbook(input_path).close()
    assert drawing_index, "drawing XML loader should find images even when ws._images is empty"
    assert any(len(items) > 0 for items in drawing_index.values())


def main() -> None:
    test_parse_title_first_line_only()
    test_image_matches_row_one_cell_above()

    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        test_resolve_local_path_encoded_and_mixed_slashes(root)

        input_path = root / "memo.xlsx"
        attachment = root / "sample attachment.txt"
        attachment.write_text("attachment content", encoding="utf-8")
        output_root = root / "output"
        output_root.mkdir()

        create_sample_workbook(input_path, attachment)
        test_drawing_loader_finds_saved_images(root)

        result = split_workbook(
            SplitConfig(
                input_path=input_path,
                output_root=output_root,
                sheet_name="Sheet1",
                header_row=1,
            ),
            on_log=print,
        )

        print("RESULT", result)
        assert result.folders_created == 3
        assert result.attachments_copied == 2
        assert not result.attachment_skips

        folder_names = sorted(p.name for p in output_root.iterdir())
        assert folder_names == ["memo_001", "memo_002", "memo_003"]

        first_xlsx = next((output_root / "memo_001").glob("*.xlsx"))
        assert "회의록" in first_xlsx.name
        assert "본문" not in first_xlsx.name

        third_folder = output_root / "memo_003"
        assert any(p.suffix == ".xlsx" for p in third_folder.iterdir())
        assert any(p.name == "sample attachment.txt" for p in third_folder.iterdir())

        for folder in sorted(output_root.iterdir()):
            print("FOLDER", folder.name, [p.name for p in folder.iterdir()])

    print("ALL TESTS PASSED")


if __name__ == "__main__":
    main()
