---
name: agi-schedule-pipeline-check
description: 통합 파이프라인 3단계. 모든 AGI Schedule 파이프라인 작업 완료 후 전체 점검. 데이터·일정·공지·날씨·KPI·히트맵·이미지·Go/No-Go 안내 등 모든 단계가 체크리스트에 포함되도록 검사 및 수정.
---

# agi-schedule-pipeline-check

## 파이프라인 위치

- **통합 파이프라인(에이전트)** 에서 **항상 3번째**로 적용되는 스킬이다.
- 순서: 1) agi-schedule-shift → 2) agi-schedule-daily-update → **3) agi-schedule-pipeline-check** → 4) weather-go-nogo.
- 모든 요청에 대해 1→2→**3**→4 가 누락 없이 수행되므로, 본 스킬은 **항상** 호출된다.

## 언제 사용

- **모든 파이프라인 작업이 완료된 후** 전체 파이프라인을 한 번 점검할 때 (통합 파이프라인 3단계).
- "파이프라인 점검", "전체 점검", "post-check", "마무리 점검" 요청 시.
- 공지·날씨·KPI·일정 시프트·히트맵 갱신 후 **최종 검증** 시.

## 작업 범위

- **대상**: `AGI TR 1-6 Transportation Master Gantt Chart/files/` 내 모든 파이프라인 산출물
- **소스 HTML**: 파일명 날짜(YYYYMMDD)가 **가장 최근인** `files/AGI TR SCHEDULE_*.html`

---

## 전체 파이프라인 작업 목록 (점검 시 모두 확인)

다음 단계가 **모두** 점검 대상임. 누락 없이 순서대로 확인한다.

| 단계 | 구분 | 작업 | 점검 항목 |
|------|------|------|-----------|
| **A** | 데이터 | `files/agi tr final schedule.json` | 일정 시프트 적용 여부, planned_start/finish 일관성, summary.date_range |
| **B** | 일정 시프트 | `files/schedule_shift.py` 적용 | JSON·HTML 동기화, pivot 이후만 시프트 |
| **C** | HTML 소스 | `files/AGI TR SCHEDULE_*.html` | 최신 날짜 파일 사용, projectStart/projectEnd, ganttData |
| **D** | 공지란 | Operational Notice 블록 | **날짜가 갱신일(YYYY-MM-DD)과 일치** |
| **E** | Weather & Marine Risk | Weather Alert 블록 | Last Updated 갱신, 4일치 예보(D~D+3), Mina Zayed·해상 문단 |
| **F** | KPI Grid | 6개 카드 | **KPI 갱신 수행 주체(유일):** 본 스킬(3단계)에서 **Total Days** 재계산 반영, **SPMT Set = 1** 확인·수정. 2단계(agi-schedule-daily-update)는 KPI를 수정하지 않음. |
| **G** | Voyage Cards | data-start/end, 표시 텍스트 | TR1~V7 일정과 JSON·ganttData 일치, Load-out/Sail/Load-in/Jack-down 표기 |
| **H** | ganttData·테이블 | JS activities, Schedule table | start/end YYYY-MM-DD, V1~V7 행 날짜가 JSON과 일치 |
| **I** | 날씨 파싱 | `files/weather/YYYYMMDD/` → `files/out/weather_parsed/` | (해당일 자료 있으면) JSON 생성 여부 |
| **J** | WEATHER_DASHBOARD | `files/WEATHER_DASHBOARD.py` | TARGET_DATE·VOYAGES, **하단 날짜 가로(rotation=0)**, **레이아웃(height_ratios·bottom)** |
| **K** | 히트맵 PNG | `files/out/weather_4day_heatmap.png` | 파일 존재, 필요 시 dashboard 복사 |
| **L** | 이미지 참조 | `embed_heatmap_base64.py` / `replace_img_ref.py` | HTML 내 히트맵 img src(파일 또는 Base64) 정상 반영 |
| **M** | weather-go-nogo 연계 | 스킬 `weather-go-nogo` | 파싱 JSON(`files/out/weather_parsed/...`) 존재 시 Go/No-Go 평가 가능 안내; 입력(파고·풍속·한도) 있으면 4단계에서 평가. |
| **N** | **물때 테이블** | 스킬 `water-tide-voyage` / `tide_to_voyage_overview.py` | **WATER TIDE.csv** 기반 6:00~17:00 상위 3시간대가 각 Voyage Overview `table.tide-table`에 반영되었는지 확인; 미반영 시 실행. |

---

## 점검·수정 체크리스트 (상세)

### 1) 공지 사항(Operational Notice) 날짜 — **D**

- **문제**: 공지란의 날짜가 갱신일로 변경되지 않음.
- **위치**: HTML 내 `<!-- Operational Notice -->` 다음 `div.weather-alert` 블록. `<strong ...>YYYY-MM-DD</strong>` 또는 동일 형식.
- **점검**: 갱신 수행일(또는 사용자 지정 공지일)과 일치하는지 확인.
- **수정**: 불일치 시 해당 블록의 날짜를 **갱신일(YYYY-MM-DD)** 로 교체. 공지 본문은 사용자 제공 시에만 교체, 미제공 시 날짜만 갱신하고 본문은 비움(규칙 유지).

### 2) SPMT Set — **F**

- **규칙**: **항상 1** (1 set). 고정값.
- **위치**: HTML 내 KPI Grid 블록. 6개 카드 중 "SPMT Set" 또는 "🛠️" 라벨 카드.
- **점검**: 표시값이 `1` 인지 확인.
- **수정**: `1` 이 아니면 `1` 로 변경.

### 3) Total Days — **F**

- **규칙**: **프로젝트 종료일 − (파일 수정 요청일 또는 오늘)** 의 일수. 갱신 시마다 요청일/오늘 기준 잔여 일수로 재계산.
- **위치**: HTML 내 KPI Grid 블록. "Total Days" 또는 "📅" 라벨 카드. 또한 JS 변수로 `totalDays` 등이 있으면 동일 값으로 맞춤.
- **점검**: 프로젝트 종료일(예: End Date Mar 23 → 2026-03-23)에서 요청일(또는 오늘)을 뺀 일수와 일치하는지 확인.
  - 예: 종료일 2026-03-23, 요청일 2026-01-28 → Total Days = 55 (또는 54, 0-based/1-based 규칙에 따라 일관 유지).
- **수정**: 불일치 시 KPI 카드 표시값과 필요 시 JS 내 `totalDays`(또는 동일 역할 변수)를 위 규칙으로 재계산해 반영.

### 4) 히트맵(WEATHER_DASHBOARD.py) 하단 날짜·레이아웃 — **J**

- **문제 1**: 히트맵 하단(Daily Operation Status)의 **날짜가 가로(rotation=0)** 로 표시되어야 함.
- **위치**: `files/WEATHER_DASHBOARD.py` 내 `ax3.set_xticklabels(..., rotation=..., ha=...)` 및 관련 서브플롯 설정.
- **점검**: `rotation=0`, `ha="center"` 로 설정되어 있는지 확인.
- **수정**: `rotation=0`, `ha="center"` 로 되어 있지 않으면 수정.

- **문제 2**: 날짜를 가로로 하면 아래 공간이 남으므로 **히트맵 레이아웃 조정**이 필요함.
- **위치**: `files/WEATHER_DASHBOARD.py` 내 `height_ratios`, `plt.subplots_adjust(left=..., right=..., top=..., bottom=...)`, 및 `gs = fig.add_gridspec(...)`.
- **점검**: `height_ratios`(예: `[2.0, 1.0, 0.6]`)로 하단 Daily Operation 구간 비율이 적절한지, `bottom` 값으로 하단 여백이 과하지 않은지 확인. 가로 날짜 시 여백을 줄이면 히트맵 영역이 더 넓어짐.
- **수정**: 하단 날짜가 가로일 때 여백이 과하면 `bottom` 값을 줄이거나(예: 0.10 → 0.08), `height_ratios`의 세 번째 값(0.6)을 조정해 Daily Operation 영역을 줄이고 히트맵 비중을 키움. (레이아웃이 깨지지 않도록 저장 후 PNG 확인 권장.)

### 5) A~C·G~H — JSON·HTML 일정 일관성

- **점검**: `files/agi tr final schedule.json` 의 TR1 Load-out 등 핵심 일정과, HTML 내 `projectStart`/`projectEnd`, `ganttData` 각 row의 `start`/`end`, voyage-card `data-start`/`data-end`, Schedule 테이블 V1~V7 행 날짜가 **동일한 일정**을 가리키는지 확인.
- **수정**: 불일치 시 시프트 규칙(pivot 이후만 +delta)에 맞춰 JSON 또는 HTML 중 한쪽을 기준으로 맞춘 뒤, `files/schedule_shift.py` 또는 수동으로 나머지 동기화.

### 6) E — Weather & Marine Risk 블록

- **점검**: "Last Updated: DD Mon YYYY" 가 갱신일과 맞는지, 갱신일 기준 **4일치**(D, D+1, D+2, D+3) 예보 문단이 있는지, Mina Zayed 인근·해상(Marine) 요약이 있는지 확인.
- **수정**: 누락/오류 시 스킬 `agi-schedule-daily-update` 절차에 따라 갱신.

### 7) I~L — 날씨 파싱·히트맵·이미지

- **I**: `files/weather/YYYYMMDD/` 에 자료가 있으면 `files/out/weather_parsed/YYYYMMDD/weather_for_weather_py.json` 생성 여부 확인.
- **K**: `files/out/weather_4day_heatmap.png` 존재 여부, 필요 시 `files/weather_4day_heatmap_dashboard.png` 복사 여부 확인.
- **L**: Schedule HTML 내 날씨 히트맵 이미지가 `src="weather_4day_heatmap_dashboard.png"` 또는 `src="data:image/png;base64,..."` 로 정상 표시되는지 확인. 필요 시 `embed_heatmap_base64.py` 또는 `replace_img_ref.py` 실행.

### 8) N — 물때 테이블 (Voyage Overview)

- **규칙**: `files/WATER TIDE.csv` 에서 **6:00~17:00** 구간만 사용해 voyage별 기간 내 **최고 물때 상위 3시간대**를 계산하고, 각 Voyage Overview의 `table.tide-table` tbody에 TIME/HEIGHT 3행으로 반영.
- **위치**: HTML 내 각 `.voyage-card` 내 `table.tide-table` tbody.
- **점검**: 각 voyage-card의 data-start~data-end 구간에 대해 WATER TIDE.csv 기반 상위 3시간대가 tide-table에 반영되었는지 확인.
- **수정**: 미반영 시 `files/` 폴더에서 `python tide_to_voyage_overview.py` 실행. 스킬 `water-tide-voyage` 참조.

---

## 실행 순서 (권장) — 모든 파이프라인 확인

1. **소스 결정**: `files/` 내 `AGI TR SCHEDULE_*.html` 중 **파일명 날짜가 가장 최근인** 파일 선택.
2. **A·B·C**: JSON 일정·시프트 적용·HTML projectStart/End·ganttData 일관성 점검.
3. **D**: 공지란 날짜 → 갱신일로 수정.
4. **E**: Weather & Marine Risk 블록(Last Updated, 4일치, Mina Zayed·해상) 점검·수정.
5. **F**: KPI Grid — Total Days 재계산 반영, SPMT Set = 1 확인·수정.
6. **G·H**: Voyage Cards·Schedule 테이블·ganttData가 JSON과 일치하는지 점검·수정.
7. **I**: (해당일 자료 있으면) weather 파싱 결과 JSON 존재 여부 확인.
8. **J**: WEATHER_DASHBOARD.py — 하단 날짜 `rotation=0`, `ha="center"` 및 `height_ratios`/`subplots_adjust(bottom)` 확인·수정 후 필요 시 `python WEATHER_DASHBOARD.py` 재실행.
9. **K**: `files/out/weather_4day_heatmap.png` 존재, 필요 시 dashboard 복사.
10. **L**: HTML 내 히트맵 img 반영 여부 확인, 필요 시 `embed_heatmap_base64.py` 또는 `replace_img_ref.py` 실행.
11. **M**: `files/out/weather_parsed/YYYYMMDD/weather_for_weather_py.json` 존재 시, 4단계 스킬 `weather-go-nogo` 로 해상 Go/No-Go 평가 가능함을 안내. (파고·풍속·한도값 입력 시 평가.)
12. **N**: WATER TIDE.csv 기반 Voyage Overview 물때 테이블 — 각 voyage의 tide-table에 6~17시 상위 3시간대 반영 여부 확인; 필요 시 `python tide_to_voyage_overview.py` 실행.

---

## 안전 규칙

- **`files/` 폴더 밖** 파일은 수정하지 않는다.
- 공지·날씨·KPI 블록 **밖**의 HTML(간트 데이터, 스크립트, Voyage Cards 등)은 점검 시 참고만 하고, 변경은 규칙에 맞는 범위에서만 한다.
- 히트맵 레이아웃 변경 후에는 생성 PNG를 한 번 확인하는 것이 좋다.

---

## 통합

- **통합 파이프라인**: 에이전트는 **항상** 1) agi-schedule-shift → 2) agi-schedule-daily-update → **3) agi-schedule-pipeline-check** → 4) weather-go-nogo 순으로 4개 스킬을 적용한다. 본 스킬은 3단계로 **항상** 호출된다.
- Subagent `/agi-schedule-updater`: 모든 요청에 대해 위 파이프라인을 수행한 뒤, 본 스킬로 **전체 파이프라인 작업(A~N)** 점검·수정 및 4단계(weather-go-nogo) 연계 안내를 한다.
