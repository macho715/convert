# AGI_TR6 READY PACK (Excel LTSC 2021 + VBA + Python)

## 빠른 시작 (3단계)

1. `ONE_CLICK_INSTALL.bat` 실행 -> `AGI_TR6_VBA_Enhanced_AUTOMATION_READY.xlsm` 생성
2. 생성된 `.xlsm` 파일 열기 -> "콘텐츠 사용(Enable Content)" 클릭
3. `Control_Panel` 시트 버튼으로 실행

---

## 패키지 구성

| 파일 | 설명 |
| --- | --- |
| `AGI_TR6_VBA_Enhanced_AUTOMATION.xlsx` | 베이스 워크북 (매크로 없음) |
| `AGI_TR6_VBA_Enhanced_AUTOMATION_READY.xlsm` | 설치 후 생성되는 매크로 포함 워크북 |
| `AGI_TR_AutomationPack.bas` | VBA 전체 모듈 (버튼/로그/검증/내보내기/파이썬 실행) |
| `ThisWorkbook_EventCode.txt` | Workbook_Open/BeforeClose 이벤트 코드 |
| `ONE_CLICK_INSTALL.bat` / `.vbs` | VBA 자동 주입 + XLSM 생성 |
| `agi_tr_runner.py` | Python 업데이트/검증 엔진 |
| `tr6_pipeline.py` | Python 리포트 엔진 |
| `requirements.txt` | Python 라이브러리 목록 |
| `run_local.bat` | Python 로컬 실행 예시 |

---

## 1) Excel 설치 (원클릭)

### 방법 A: 자동 설치 (권장)
```bat
ONE_CLICK_INSTALL.bat
```

### 방법 B: 수동 설치 (자동 설치 실패 시)
1. `AGI_TR6_VBA_Enhanced_AUTOMATION.xlsx` 열기
2. 다른 이름으로 저장 -> `.xlsm` 형식으로 저장
3. `Alt+F11` -> VBA Editor 열기
4. `Insert > Module` -> `AGI_TR_AutomationPack.bas` Import
5. `ThisWorkbook` 더블클릭 -> `ThisWorkbook_EventCode.txt` 내용 붙여넣기
6. 저장

### 설치 실패 시 해결책
Excel > 파일 > 옵션 > 보안 센터 > 보안 센터 설정 > 매크로 설정 >
`VBA 프로젝트 개체 모델에 대한 신뢰할 수 있는 액세스` 체크

---

## 2) Python 설치 (선택)

### 가상환경 설정
```bat
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
```

### 실행
```bat
REM 방법 1: 배치 파일 사용 (권장)
run_local.bat

REM 방법 2: 직접 실행
python agi_tr_runner.py --in "AGI_TR6_VBA_Enhanced_AUTOMATION_READY.xlsm" --out ".\out" --mode update
python tr6_pipeline.py --in "AGI_TR6_VBA_Enhanced_AUTOMATION_READY.xlsm" --out ".\out" --log ".\out\tr6_ops.log"
```

---

## 3) 테스트

### Python 테스트
```bat
run_local.bat
```
`out` 폴더에 `*_OUT_*.xlsx` / `*_PY_REPORT.xlsx` 생성되면 OK

### VBA 테스트
1. `AGI_TR6_VBA_Enhanced_AUTOMATION_READY.xlsm` 열기
2. `Alt+F8` -> `TR6_SelfTest` 실행
3. `LOG` 시트에 결과 기록되면 OK

---

## 주요 기능

### VBA 기능
- UpdateAllDates: D0 기반 일정 자동 업데이트 (Ctrl+Shift+U)
- RefreshGanttChart: 간트 차트 자동 갱신
- GenerateDailyBriefing: 일일 브리핑 생성
- ExportToCSV: CSV 내보내기
- BackupWorkbook: 워크북 백업
- RunPythonUpdate: Python 스크립트 실행 (VBA에서 호출)

### Python 기능
- agi_tr_runner.py: 스케줄 업데이트/검증/간트 재생성
- tr6_pipeline.py: 리포트 생성 (Phase/Status/Risk 분석)

---

## 로그 파일

- `INSTALL_LOG.txt`: 원클릭 설치 로그
- `out\tr6_ops.log`: Python 실행 로그 (JSONL 형식)
- `LOG` 시트: Excel 내부 로그

---

## 문제 해결

### Q: "VBA 프로젝트 개체 모델에 대한 액세스" 오류
A: Excel 옵션에서 위 설정을 활성화하세요 (위 "설치 실패 시 해결책" 참조)

### Q: Python 스크립트 실행 실패
A: `requirements.txt`의 패키지가 설치되었는지 확인:
```bat
pip list | findstr "pandas openpyxl"
```

### Q: READY.xlsm이 생성되지 않음
A: `INSTALL_LOG.txt`를 확인하고, Excel이 실행 중이 아닌지 확인하세요.

---

## 버전 정보

- Excel: LTSC 2021 호환
- Python: 3.11.8+ 권장
- pandas: >=2.2.3
- openpyxl: >=3.1.5

---

생성일: 2026-01-07
프로젝트: Samsung C&T HVDC Logistics - AGI TR Transportation
