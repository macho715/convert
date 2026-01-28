# AGI TR1..TR7 Dynamic Gantt (v13) – 운영 가이드

## 1) 파일 구성
- **AGI_TR7_Dynamic_Gantt_Template_v13.xlsx**
  - **Plan_A_Realistic**: 설치(Install)를 항차 크리티컬패스에 삽입(보수적/리스크 반영)
  - **Plan_B_Fast**: 설치를 항차와 분리(설치팀 병행 가정으로 3월 전 완료 목표)
  - **Tide_Peaks_MZP**: Mammoet tide PDF에서 추출한 일별 최고조(Planning)
  - **Weather_Forecast**: VBA로 자동 채움(없어도 계획은 동작)

- **generate_agi_tr7_dynamic_gantt_v13.py**: 엑셀 재생성 파이썬 스크립트
- **AGI_TR7_Automation_v4.bas**: VBA 모듈(선택)

## 2) 엑셀 사용(매크로 없이)
1. **Inputs** 탭에서 다음 셀만 수정하면 전체 일정이 자동 이동합니다.
   - **B5**: TR1 LO 시작일(Commencement @MZP)
   - **B4**: 설치 병행 팀 수(Install parallel teams) – Plan_B 기간 단축 레버
   - **B20:B28**: 단계별 Duration/Buffer(일)
   - **B10:B13**: 항차별 적재 대수(기본: 1-2-2-2 = 총 7기)
   - **E10:E12**: 설치 트리거(누적 도착 대수, 기본: 3/5/7)

2. 확인 방법
   - **Plan_A_Realistic / Plan_B_Fast** 탭에서 Gantt grid가 자동 재색칠됩니다.
   - 타겟: **Inputs!B6 (기본 2026-03-01)** 이전 완료 여부는 VBA 없이도 육안 확인 가능합니다.

## 3) VBA 사용(선택 – 업무 편의 기능)
### 3.1 모듈 Import
1) 엑셀 **Developer** 탭 활성화
2) **Alt+F11** → VBA Editor
3) **File → Import File…** → `AGI_TR7_Automation_v4.bas` 선택

### 3.2 추천 실행
- `UpdateAll(True)` : 전체 갱신 + 날씨 가져오기 + NO-GO 플래그 + 14D 룩어헤드 생성
- `UpdateAll(False)` : 재계산 + QC만

### 3.3 시작일 입력 시 자동 실행(옵션)
**Inputs 시트 코드(Worksheet_Change)에 아래를 추가**하면 B5 변경 시 자동 갱신됩니다.
```vb
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("B5")) Is Nothing Then
        Call AGI_TR7_Automation_v4.UpdateAll(False)
    End If
End Sub
```

## 4) 날씨 연동 아이디어(구현 포함)
- VBA `RefreshWeather_OpenMeteo`가 Inputs(D20:E27)의 좌표/임계치로 **Open‑Meteo** 일별 예보(CSV)를 받아 Weather_Forecast에 채웁니다.
- `FlagNoGoTasks`가 Plan 탭에서 **Start date가 NO-GO이면 Notes에 플래그**를 넣습니다.
- 실운영에서는 Port Control/Marine Forecast를 SSOT로 두고, VBA는 **1차 스크리닝(알림)** 용으로 사용 권장.

## 5) 파이썬으로 재생성(템플릿 구조 유지)
```bash
pip install openpyxl pandas pymupdf
python generate_agi_tr7_dynamic_gantt_v13.py --tide_pdf "MAMMOET_AGI TR.pdf" --out "AGI_TR7_Dynamic_Gantt_Template_v13.xlsx"
```

## 6) Fail-safe / QC
- Tide_Peaks_MZP는 **planning**입니다(최종은 AD Ports/Port tide table 우선).
- Weather_Forecast도 **planning gate**입니다(최종은 Marine forecast + Master/Port Control).
- 모든 Duration은 Inputs에서 즉시 조정 가능합니다(버퍼/리스크 반영).
