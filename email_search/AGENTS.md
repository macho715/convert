# email_search — AGENTS.md

Outlook Excel 기반 검색 + 스레드 추적 + Streamlit 대시보드. 루트 **AGENTS.md** 우선.

## 엔트리포인트

| 유형 | 경로 | 용도 |
|------|------|------|
| CLI | `email_search/scripts/export_email_threads_cli.py` | 스레드/검색 export |
| Streamlit | `email_search/dashboard/app.py` | 대시보드 UI |

실행(README 기준):

- `python email_search/scripts/run_full_export.py --excel <Excel> --sheet "전체_데이터" --out <out_dir> --query "<쿼리>"`
- `streamlit run email_search/dashboard/app.py`

## 입출력

- **입력**: Outlook Excel export, 시트명(예: `"전체_데이터"`).
- **출력**: `threads.json`, `edges.csv`, `search_result.csv` (지정한 out 디렉터리); 대시보드는 `../outputs/threads_full` 등 데이터 루트 사용.

## 고정 스모크(AGENTS.md §5.4 #4)

- `streamlit run email_search/dashboard/app.py` (실행 가능 여부 확인)
- `python email_search/scripts/run_full_export.py --excel <익명 Excel> --sheet "전체_데이터" --out email_search/outputs/threads_smoke --query "LPO-1599"` (익명 샘플 1회)

## 주의

- PII 포함 Excel은 익명 샘플로만 스모크. 운영 데이터 실행 시 루트 규칙 "Ask first" 적용.
