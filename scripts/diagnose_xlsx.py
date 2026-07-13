"""XLSX 이미지 진단 스크립트.

사용법:
  python scripts/diagnose_xlsx.py "INPUT.xlsx" "시트명"
"""

from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from openpyxl import load_workbook

from src.cell_image_loader import diagnose_image_sources
from src.drawing_image_loader import build_images_by_row


def main() -> None:
    if len(sys.argv) < 3:
        print('사용법: python scripts/diagnose_xlsx.py "INPUT.xlsx" "시트명"')
        sys.exit(1)

    xlsx_path = Path(sys.argv[1])
    sheet_name = sys.argv[2]

    if not xlsx_path.is_file():
        print(f"파일 없음: {xlsx_path}")
        sys.exit(1)

    wb = load_workbook(xlsx_path, data_only=False)
    if sheet_name not in wb.sheetnames:
        print(f"시트 없음. 사용 가능: {wb.sheetnames}")
        sys.exit(1)

    ws = wb[sheet_name]
    print(f"파일: {xlsx_path}")
    print(f"시트: {sheet_name}")
    print("--- 진단 ---")
    for line in diagnose_image_sources(xlsx_path, ws, wb):
        print(line)

    images_by_row = build_images_by_row(xlsx_path, ws, wb)
    total = sum(len(v) for v in images_by_row.values())
    print(f"--- 결과: 이미지 {total}개 인식 ---")
    for row, images in sorted(images_by_row.items()):
        print(f"  행 {row}: {len(images)}개")

    wb.close()


if __name__ == "__main__":
    main()
