"""Manual integration test helper for MEMO SPLITTER."""

from __future__ import annotations

import tempfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Font
from openpyxl.worksheet.hyperlink import Hyperlink

from src.splitter import SplitConfig, split_workbook


def create_sample_workbook(path: Path, attachment: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    headers = ["App", "본문", "Link"]
    for col, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col, value=header)

    rows = [
        ("Kakao", "제목 : 회의록", attachment),
        ("Line", "본문만 있음", None),
        ("Slack", "제목 : 긴급 공지", attachment),
    ]

    for idx, (app, body, link_target) in enumerate(rows, start=2):
        ws.cell(row=idx, column=1, value=app)
        ws.cell(row=idx, column=2, value=body)
        if link_target is not None:
            cell = ws.cell(row=idx, column=3, value="첨부열기")
            cell.hyperlink = Hyperlink(ref=cell.coordinate, target=str(link_target))
            cell.font = Font(color="0563C1", underline="single")

    png_path = path.parent / "tiny.png"
    png_path.write_bytes(
        b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01"
        b"\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDATx\x9cc``\x00\x00"
        b"\x00\x02\x00\x01\xe2!\xbc3\x00\x00\x00\x00IEND\xaeB`\x82"
    )
    image = Image(str(png_path))
    image.anchor = "A4"
    ws.add_image(image)

    wb.save(path)
    wb.close()


def main() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        input_path = root / "memo.xlsx"
        attachment = root / "sample_attachment.txt"
        attachment.write_text("attachment content", encoding="utf-8")
        output_root = root / "output"
        output_root.mkdir()

        create_sample_workbook(input_path, attachment)

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
        for folder in sorted(output_root.iterdir()):
            print("FOLDER", folder.name, [p.name for p in folder.iterdir()])


if __name__ == "__main__":
    main()
