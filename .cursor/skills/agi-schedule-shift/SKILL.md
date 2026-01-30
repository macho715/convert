---
name: agi-schedule-shift
description: 통합 파이프라인 1단계. AGI TR 일정(JSON/HTML)에서 pivot date 이후 전체 일정을 delta일만큼 자동 시프트. 모든 작업은 files 폴더 안에서만 수행.
---

# agi-schedule-shift

## 파이프라인 위치

- **통합 파이프라인(에이전트)** 에서 **항상 1번째**로 적용되는 스킬이다.
- 순서: **1) agi-schedule-shift** → 2) agi-schedule-daily-update → 3) agi-schedule-pipeline-check → 4) weather-go-nogo.
- pivot_date·new_date 제공 시 시프트 실행; 미제공 시 "일정 시프트 없음" 확인만 한다.

## 작업 범위

- **모든 작업은 `AGI TR 1-6 Transportation Master Gantt Chart/files/` 폴더 안에서만 수행한다.** JSON·HTML 읽기/쓰기는 `files/` 내 파일만 대상으로 한다.

## 언제 사용

- 특정 일자(예: 2월 1일)가 다른 일자(예: 2월 3일)로 변경될 때, **그 일자 이후** 모든 일정을 자동으로 같은 일수만큼 밀고 싶을 때
- "일정 시프트", "schedule shift", "일정 2일 연기", "AGI schedule delay", "전체 일정 자동 수정" 요청 시

## 입력

| 항목 | 설명 | 예시 |
|------|------|------|
| pivot_date | 기준일 (이 날짜 이후만 시프트) | 2026-02-01 |
| new_date | 바꿀 목표일 (pivot_date가 이동할 날짜) | 2026-02-03 |
| delta_days | (자동 계산) new_date − pivot_date | +2 |

## 대상 파일 (모두 files/ 내)

- **JSON**: `AGI TR 1-6 Transportation Master Gantt Chart/files/agi tr final schedule.json` (또는 동일 스키마의 `files/` 내 JSON)
- **HTML**: `AGI TR 1-6 Transportation Master Gantt Chart/files/AGI TR SCHEDULE_*.html`

- **기본 파일**: `files/AGI TR SCHEDULE_20260128.html` (업데이트 시에는 파일명 날짜가 가장 최근인 파일 사용 규칙과 동일하게, 시프트 시에는 `files/` 내 모든 `AGI TR SCHEDULE_*.html` 에 동일 delta 적용하여 일정 일치 유지.)

## 절차

### 1) JSON 시프트

- `files/agi tr final schedule.json` 의 `activities[]` 내 각 항목의 `planned_start`, `planned_finish`를 확인.
- **pivot_date 이상**인 날짜만 `+ delta_days` 일 적용 (날짜 파싱 → timedelta 더함 → YYYY-MM-DD 문자열로 저장).
- `duration`은 변경하지 않음 (날짜만 이동).
- `summary.date_range` 의 start/finish 도 pivot 이후이면 동일 delta 적용.

### 2) HTML 시프트 (files/ 내 AGI TR SCHEDULE_*.html)

- `projectStart`, `projectEnd`: pivot 이후 구간이면 동일 delta 적용.
- `ganttData` 내 각 row의 `activities[]`에서 `start`, `end`가 pivot_date 이상이면 `+ delta_days` 적용.
- voyage-card 의 `data-start`, `data-end`: pivot_date 이상이면 동일 delta 적용.
- 날짜 형식 유지: `'YYYY-MM-DD'`.

### 3) 옵션 (권장)

- **--dry-run**: 파일은 쓰지 않고, 변경될 날짜 목록만 출력.
- **--backup**: 수정 전 JSON/HTML 복사본을 **`files/` 안에** 생성 (예: `files/agi tr final schedule_backup_YYYYMMDD.json`).

### 4) 스크립트

- **경로**: `AGI TR 1-6 Transportation Master Gantt Chart/files/schedule_shift.py`
- **실행**: `files/` 폴더에서 `python schedule_shift.py [--dry-run] [--backup]` 실행. 스크립트 내 `PIVOT`, `DELTA_DAYS` 값을 pivot_date·new_date에 맞게 수정한 뒤 실행.

## 검증

- 시프트 후 모든 `planned_start` ≤ `planned_finish` 유지.
- 필요 시 "이전 활동 finish ≤ 다음 활동 start" 등 순서 일관성 확인.

## 안전 규칙

- **`files/` 폴더 밖**의 JSON·HTML을 읽거나 쓰지 않는다.
- pivot_date **이전** 날짜는 변경하지 않는다.
- JSON과 HTML을 **동시에** 시프트하여 `files/` 내 두 파일이 같은 일정을 유지하도록 한다.

## 대시보드 동기화 (일정 변경 시)

시프트 후 **파이프라인 3단계(agi-schedule-pipeline-check)** 를 반드시 실행하여 대시보드 전체에 동일한 날짜가 적용되도록 한다:

| 대시보드 영역 | 갱신 항목 |
|---------------|-----------|
| **7 Voyages Overview** | voyage-card data-start/end, Load-out/Sail/Load-in/Jack-down, **tide-table** (water-tide-voyage) |
| **Detailed Voyage Schedule** | Schedule 테이블 V1~V7, ganttData activities |
| **Gantt Chart** (Jan 26 - Mar 25, 2026) | projectStart/projectEnd, 차트 제목 날짜 범위 |

- **물때**: `tide_to_voyage_overview.py`는 pipeline-check 단계 N에서 **항상** 실행되어 voyage별 data-start~data-end에 맞는 tide-table을 반영한다.

## 통합

- Subagent `/agi-schedule-updater`와 별개. 일정 시프트는 이 스킬만으로 수행 가능. 시프트 시에도 작업 범위는 `files/` 로 한정.
