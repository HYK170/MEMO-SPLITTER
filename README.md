# MEMO SPLITTER

INPUT XLSX의 지정 시트를 **HEADER 1행 + 데이터 1행** 단위로 분할하고, Multimedia 폴더의 첨부파일을 함께 복사합니다.

## 요구 사항

- Python 3.11+
- Windows (CustomTkinter UI)

## 설치

```powershell
cd "c:\Users\rlagp\OneDrive\문서\작업\MEMO SPLITTER"
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## 실행

```powershell
python main.py
```

## 사용 방법

1. **INPUT XLSX** — 분할할 원본 Excel 파일
2. **Multimedia 폴더** — `저장된 파일 이름` 컬럼의 상대 경로 기준 루트
3. **SHEET** — 분할 대상 시트명
4. **HEADER ROW** — 헤더 행 번호 (1-based)
5. **SPLIT 실행** 클릭

결과는 INPUT XLSX와 **같은 경로**에 `Memo_{YYYYMMDDHHMMSS}` 폴더를 만들고 그 안에 저장합니다.

## INPUT 형식

필수 컬럼:

- `App`
- `본문` — `제목 : ` 접두어 이후 첫 줄을 파일명에 사용 (없으면 `제목없음`)
- `저장된 파일 이름` — Multimedia 하위 상대 경로. 여러 개는 줄바꿈으로 구분

선택 컬럼:

- `첨부 파일` — jpg/png/jpeg 등 이미지 첨부 시 이 열 셀에 이미지를 임베드

## 출력 구조

```
[INPUT XLSX 경로]/
├── memo.xlsx
└── Memo_20260714132400/
    ├── memo_001/
    │   ├── memo_Kakao_001_회의록.xlsx
    │   ├── shot.png
    │   └── memo.txt
    └── ...
```

- 출력 루트: `Memo_{timestamp}` (초 단위 시리얼, 충돌 시 `_2`, `_3` …)
- 행별 폴더명: `{원본명}_{행번호0패딩}`
- XLSX: 항상 **1행=HEADER, 2행=데이터**
- 이미지는 파일로 복사 + `첨부 파일` 열에 임베드
- 이미지가 아닌 첨부는 폴더에만 복사
