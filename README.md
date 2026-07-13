# MEMO SPLITTER

INPUT XLSX의 지정 시트를 **HEADER 1행 + 데이터 1행** 단위로 분할하고, 행별 폴더에 XLSX와 하이퍼링크 첨부파일을 함께 내보냅니다.

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

## 이미지 진단 (문제 발생 시)

프로젝트 폴더에서 아래 명령을 실행하세요. `PYTHONPATH` 설정은 필요 없습니다.

```powershell
python scripts/diagnose_xlsx.py "INPUT.xlsx" "메모"
```

## 사용 방법

1. **INPUT XLSX** — 분할할 원본 Excel 파일
2. **OUTPUT 폴더** — 행별 결과 폴더가 생성될 루트 경로
3. **SHEET** — 분할 대상 시트명
4. **HEADER ROW** — 헤더 행 번호 (1-based)
5. **SPLIT 실행** 클릭

## INPUT 형식

- HEADER 행에 `App`, `본문` 컬럼이 반드시 있어야 합니다.
- `본문` 값이 `제목 : `로 시작하면 이후 문자열을 XLSX 파일명에 사용합니다. 없으면 `제목없음`을 사용합니다.

## 출력 구조

```
[OUTPUT 폴더]/
├── memo_001/
│   ├── memo_Kakao_001_회의록.xlsx
│   ├── 첨부1.pdf
│   └── screenshot.png
├── memo_002/
│   └── ...
```

- 행별 폴더명: `{원본명}_{행번호0패딩}`
- XLSX 파일명: `{원본명}_{App}_{행번호0패딩}_{본문제목|제목없음}.xlsx` (50자 초과 시 잘림)
- 데이터 행의 **셀/이미지 하이퍼링크** 대상 로컬 파일을 같은 폴더로 복사
- OUTPUT XLSX에는 데이터 행 이미지는 유지하되 **하이퍼링크는 제거**
- HEADER 행 이미지는 OUTPUT XLSX에 포함하지 않음

## 수동 검증 체크리스트

1. HEADER 1행 + 데이터 3행 샘플 XLSX (셀 hyperlink + 이미지 hyperlink 포함) 준비
2. SPLIT 후 `memo_001/`, `memo_002/`, `memo_003/` 폴더 생성 확인
3. 각 폴더에 XLSX 1개 + 연결된 로컬 첨부파일 복사 확인
4. OUTPUT XLSX에서 이미지 표시 유지, 클릭 시 hyperlink 없음 확인
5. App, 행번호, 본문 제목 파싱 및 50자 truncate 확인
6. HEADER 행 이미지가 OUTPUT XLSX에 없음 확인
7. 존재하지 않는 첨부 경로는 스킵되고 나머지는 정상 완료되는지 확인

## 제한 사항

- `http://`, `https://`, 내부 시트 링크(`#Sheet1!A1`)는 파일 복사 대상이 아닙니다.
- openpyxl 특성상 INPUT XLSX 전체를 메모리에 로드합니다.
- 일부 Excel 하이퍼링크 URI 형식은 경로 해석에 실패할 수 있으며, 이 경우 로그에 기록 후 스킵합니다.
