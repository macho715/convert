---
name: agi-schedule-shift
description: AGI TR 일정(JSON/HTML)에서 pivot date 이후 전체 일정을 delta일만큼 자동 시프트. 예: 2월 1일 → 2월 3일이면 이후 일정 +2일.
---

# agi-schedule-shift

## 언제 사용

- 특정 일자(예: 2월 1일)가 다른 일자(예: 2월 3일)로 변경될 때, **그 일자 이후** 모든 일정을 자동으로 같은 일수만큼 밀고 싶을 때
- "일정 시프트", "schedule shift", "일정 2일 연기", "AGI schedule delay", "전체 일정 자동 수정" 요청 시

## 입력

| 항목 | 설명 | 예시 |
|------|------|------|
| pivot_date | 기준일 (이 날짜 이후만 시프트) | 2026-02-01 |
| new_date | 바꿀 목표일 (pivot_date가 이동할 날짜) | 2026-02-03 |
| delta_days | (자동 계산) new_date − pivot_date | +2 |

## 대상 파일

- **JSON**: `AGI TR 1-6 Transportation Master Gantt Chart/agi tr final schedule.json` (또는 동일 스키마의 `option_c.json` 등)
- **HTML**: `AGI TR 1-6 Transportation Master Gantt Chart/AGI TR Unit 1 Schedule_*.html`

## 절차

### 1) JSON 시프트

- `activities[]` 내 각 항목의 `planned_start`, `planned_finish`를 확인.
- **pivot_date 이상**인 날짜만 `+ delta_days` 일 적용 (날짜 파싱 → timedelta 더함 → YYYY-MM-DD 문자열로 저장).
- `duration`은 변경하지 않음 (날짜만 이동).

### 2) HTML 시프트

- `projectStart`, `projectEnd`: pivot 이후 구간이면 동일 delta 적용.
- `ganttData` 내 각 row의 `activities[]`에서 `start`, `end`가 pivot_date 이상이면 `+ delta_days` 적용.
- 날짜 형식 유지: `'YYYY-MM-DD'`.

### 3) 옵션 (권장)

- **--dry-run**: 파일은 쓰지 않고, 변경될 날짜 목록만 출력.
- **--backup**: 수정 전 JSON/HTML 복사본 생성 (예: `*_backup_YYYYMMDD.json`).

## 검증

- 시프트 후 모든 `planned_start` ≤ `planned_finish` 유지.
- 필요 시 "이전 활동 finish ≤ 다음 활동 start" 등 순서 일관성 확인.

## 안전 규칙

- pivot_date **이전** 날짜는 변경하지 않는다.
- JSON과 HTML을 **동시에** 시프트하여 두 파일이 같은 일정을 유지하도록 한다.

## 통합

- Subagent `/agi-schedule-updater`와 별개. 일정 시프트는 이 스킬만으로 수행 가능.
