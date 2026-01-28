# CONVERT

업무 자동화 모듈 집합(문서 변환·이메일·AGI·CIPL·Stability 등). 루트 규칙은 **AGENTS.md**를 따른다.

## CONVERT 모듈 맵

| Module | Entry Points | Inputs | Outputs |
|--------|--------------|--------|---------|
| **mrconvert_v1** | `mrconvert` / `python -m mrconvert` | PDF, DOCX, XLSX | out/*.txt, *.md, *.json, tables/*.csv |
| **email_search** | `export_email_threads_cli.py`, `dashboard/app.py` (Streamlit) | Excel(Outlook export), sheet | threads.json, edges.csv, search_result.csv |
| **CIPL** | `CIPL_PATCH_PACKAGE/make_cipl_set.py` | voyage_input*.json | 4-sheet CIPL xlsx (out/) |
| **vessel_stability_python** | example_usage.py, tests/ | Stability Booklet.xls | 결과/검증 리포트 |
| **JPT71** | content_calendar_app.py, run_excel_engine.py | content-calendar.xlsx, voyage_input.json | calendar_output.*, xlsx |
| **AGI DOCS** | run_all_builders.py (Ask first) | 템플릿/소스 | xlsm/xlsx 템플릿 |
| **AGI TR 1-6 Gantt** | run_local.bat, requirements | xlsx/tsv, 시나리오 | TEST_OUT/*.xlsx |
| **scripts** | 유틸(날씨/검증/변환 등) | 다양 | out/, output/ |

## 스모크(검증)

- **공통**: `python -m compileall -q .`
- **pytest(조건부)**: `pytest -q` (pytest 설정 있는 모듈)
- **고정 커맨드 전체**: 루트 **AGENTS.md** §5.4 참고.

## 문서

- **AGENTS.md** — 개발 규칙, 권한, 검증 루틴, 고정 커맨드(스모크 후보 1~8).
