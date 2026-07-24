# MEMO SPLITTER

INPUT XLSX 또는 HTML의 테이블을 **HEADER 1행 + 데이터 1행** 단위로 분할합니다.

앱에서 **XLSX / HTML 모드**를 전환할 수 있습니다.

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

### 공통

1. 상단에서 **XLSX** 또는 **HTML** 모드 선택
2. **INPUT** — 분할할 원본 파일
3. **SPLIT 실행** 클릭

결과는 INPUT과 **같은 경로**에 `{원본파일명}_{YYYYMMDDHHMMSS}` 폴더를 만들고 그 안에 저장합니다.

### XLSX 모드 추가 입력

- **Multimedia 폴더** — `저장된 파일 이름` 컬럼의 상대 경로 기준 루트
- **SHEET** — 분할 대상 시트명
- **HEADER ROW** — 헤더 행 번호 (1-based)

### HTML 모드

- Multimedia / SHEET / HEADER ROW 지정 불필요
- INPUT HTML의 **첫 번째 `<table>`**만 사용
- 헤더: `<thead>` 첫 행, 없으면 `<th>`가 있는 첫 `<tr>`
- 첨부 `a href` 경로는 **INPUT HTML 파일 위치** 기준으로 해석·복사 (`Multimedia 밖 스킵` 없음)
- `a href` 안 `img src` 썸네일·`<link rel="stylesheet">` / `<style> url(...)` 은 복사하지 않고, split HTML 기준 **원본 상대경로만** 재작성
- 출력 테이블에 원본 `<colgroup>`, `<thead>`(및 table 속성) 포함 + 데이터 1행은 `<tbody>`로 출력

## INPUT 형식

### XLSX 필수 컬럼

- `App`
- `본문` — `제목 : ` 접두어 이후 첫 줄을 파일명에 사용 (없으면 `제목없음`)
- `저장된 파일 이름` — Multimedia 하위 상대 경로. 여러 개는 줄바꿈으로 구분

선택 컬럼 (XLSX):

- `첨부 파일` — jpg/png/jpeg 등 이미지 첨부 시 이 열 셀에 이미지를 임베드

### HTML 필수 컬럼

- `App`
- `본문` — `제목 : ` 접두어 이후 첫 줄을 파일명에 사용 (없으면 `제목없음`)
- `첨부파일` (`첨부 파일`도 동일 취급) — `<a href>` 로컬 파일만 attach로 복사. 하위 `<img src>` 썸네일·외부 링크는 복사하지 않음

## 출력 구조

### XLSX

```
[INPUT 경로]/
├── memo.xlsx
└── memo_20260714132400/
    ├── memo_001/
    │   ├── memo_001_회의록.xlsx
    │   └── memo_001_attach/
    │       ├── shot.png
    │       └── memo.txt
    └── ...
```

### HTML

```
[INPUT 경로]/
├── memo.html
├── css/app.css
└── memo_20260714132400/
    ├── memo_001/
    │   ├── memo_001_회의록.html   # CSS/썸네일은 ../../css, ../../images 참조
    │   └── memo_001_attach/
    │       ├── shot.png           # a href 원본만 복사
    │       └── memo.txt
    └── ...
```

- 출력 루트: `{원본파일명}_{timestamp}` (초 단위 시리얼, 충돌 시 `_2`, `_3` …)
- 행별 폴더명: `{원본파일명}_{행번호0패딩}`
- 파일명: `{원본파일명}_{행번호0패딩}_{제목}.xlsx` 또는 `.html`
- 첨부: 행 폴더 하위 `{원본파일명}_{행번호}_attach/` (`a href` 대상만)
- 분할 결과: **colgroup/thead 유지 + tbody 데이터 1행**
- XLSX: 이미지는 파일 복사 + `첨부 파일` 열에 임베드
- HTML: `a href`만 복사, CSS·썸네일(`img src`)은 원본 경로 참조로 재작성
