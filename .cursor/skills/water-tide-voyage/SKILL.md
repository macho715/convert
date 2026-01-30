---
name: water-tide-voyage
description: WATER TIDE.csv 기반 주간(오전 6시~오후 5시) 최고 물때 상위 3시간대를 Voyage Overview의 tide-table에 voyage별로 연동. files/ 전용.
---

# water-tide-voyage

## 목적

- **WATER TIDE.csv** (`files/WATER TIDE.csv`)에서 **주간 기준** · **오전 6시~오후 5시(6:00~17:00)** 구간만 사용해, **가장 물때가 높은 3시간대**를 산출한다.
- 산출한 TIME / HEIGHT를 **Voyages Overview**의 각 voyage 카드 내 **Voyage Overview** 블록의 `table.tide-table`에 **동일 형식**으로 연동한다.

## 작업 범위

- **모든 작업은 `AGI TR 1-6 Transportation Master Gantt Chart/files/` 폴더 안에서만 수행한다.**
- 입력: `files/WATER TIDE.csv`
- 출력 갱신: `files/AGI TR SCHEDULE_*.html` 내 각 voyage-card의 `table.tide-table` tbody

## 언제 사용

- "WATER TIDE Voyage Overview 연동", "물때 상위 3시간 Voyage에 반영", "tide table 갱신", "Voyages Overview 물때" 요청 시
- 파이프라인 점검(agi-schedule-pipeline-check) 시 **N) 물때 테이블** 항목으로 호출 가능

## 입력·출력

| 구분 | 내용 |
|------|------|
| **입력** | `files/WATER TIDE.csv` — 컬럼: 날짜, 0:00~23:00 (단위 m) |
| **시간 필터** | **6:00, 7:00, …, 17:00** 만 사용 (오전 6시~오후 5시) |
| **voyage 구간** | 각 voyage-card의 `data-start`, `data-end` (YYYY-MM-DD) |
| **산출** | voyage별로 해당 기간 내 일자에 대해 6~17시 각 시간대의 (평균 또는 최대) 높이 계산 → **상위 3개 시간대** (내림차순) |
| **출력 형식** | HTML 내 `<table class="tide-table">` tbody에 3행: `<td>HH:00</td><td>X.XXm</td>` |

예시 (한 voyage의 해당 주/기간 대표):

```
TIME    HEIGHT
9:00    2.03m
10:00   1.99m
8:00    1.97m
```

## 절차

### 1) CSV 파싱

- `files/WATER TIDE.csv` 읽기. 인코딩 UTF-8.
- 첫 행: `날짜,0:00,1:00,...,23:00` — 시간 컬럼 중 **6:00~17:00**만 사용.
- 날짜 컬럼 파싱 (YYYY-MM-DD).

### 2) voyage별 기간 확보

- 대상 HTML: `files/AGI TR SCHEDULE_*.html` (최신 날짜 파일 또는 모두).
- 각 `.voyage-card[data-start][data-end]`에서 `data-start`, `data-end` 추출.

### 3) voyage별 상위 3시간대 계산

- 해당 voyage의 `[data-start, data-end]` 구간에 포함되는 CSV 행만 사용.
- 6:00~17:00 컬럼만 사용해, **시간대별**로 해당 기간 내 **최대값**(또는 평균) 계산.
- 높이 내림차순 정렬 후 **상위 3개** 시간대 선택 → `[(time, height), ...]` (예: `[("9:00", 2.03), ("10:00", 1.99), ("8:00", 1.97)]`).

### 4) HTML 갱신

- 각 voyage-card 내 `table.tide-table tbody`를 찾아, 기존 3개 `<tr>`를 **산출한 TIME/HEIGHT 3행**으로 교체.
- 형식 유지: `<tr><td>9:00</td><td>2.03m</td></tr>`.

### 5) 스크립트

- **경로**: `AGI TR 1-6 Transportation Master Gantt Chart/files/tide_to_voyage_overview.py`
- **실행**: `files/` 폴더에서 `python tide_to_voyage_overview.py`
- **옵션(권장)**: `--dry-run` — HTML 쓰지 않고 voyage별 상위 3시간대만 출력.

## 검증

- 시프트/일정 변경 후에도 `WATER TIDE.csv` 날짜 범위가 voyage 구간과 겹치는지 확인.
- 6~17시 외 시간대가 테이블에 들어가지 않았는지 확인.

## 안전 규칙

- **`files/` 폴더 밖**의 CSV/HTML을 읽거나 쓰지 않는다.
- 공지·날씨·ganttData 등 **tide-table 외** HTML 블록은 수정하지 않는다.

## 통합

- **agi-schedule-pipeline-check**: 점검 목록 **N) 물때 테이블** — **항상** `tide_to_voyage_overview.py` 실행. 일정 시프트 후 대시보드(7 Voyages Overview) 일관성 유지를 위해 파이프라인 3단계에서 조건 없이 수행.
