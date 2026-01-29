---
name: weather-go-nogo
description: 통합 파이프라인 4단계. SEA TRANSIT(해상 운행) 전용 Go/No-Go 의사결정 로직. Wave(ft)·Wind(kt) 입력, 3단 Gate(임계값·Squall/피크파 버퍼·연속 Weather window) 적용.
---

# Weather Go/No-Go (SEA TRANSIT)

## 파이프라인 위치

- **통합 파이프라인(에이전트)** 에서 **항상 4번째**로 적용되는 스킬이다.
- 순서: 1) agi-schedule-shift → 2) agi-schedule-daily-update → 3) agi-schedule-pipeline-check → **4) weather-go-nogo**.
- 파고·풍속·한도값(또는 파싱 JSON) 제공 시 Go/No-Go 평가 실행; 미제공 시 "입력 시 평가 가능" 안내.

## 언제 사용

- **해상 운행(SEA TRANSIT)** Go/No-Go 판단 요청 시
- "weather window", "Hs/Hmax", "squall buffer", "연속 GO", "marine weather decision" 언급 시
- 파이프라인(parsed_to_weather_json, WEATHER.PY)에서 추출한 wave_ft·wind_kt를 **운항 의사결정**으로 평가할 때
- 입력 데이터는 `AGI TR 1-6 Transportation Master Gantt Chart/files/` 내 파싱 결과(예: `files/out/weather_parsed/YYYYMMDD/weather_for_weather_py.json`) 사용 가능.

## 입력·출력

| 구분 | 내용 |
|------|------|
| **입력** | Wave Height(ft, Combined Sea+Swell), Wind Speed(kt). (선택) Wave Period(s). |
| **가정** | 차트 "combined sea+swell(ft)" → **Hs(유의파고)** 근사. Hs = 상위 1/3 평균 파고. |
| **출력** | `Decision: GO | NO-GO | CONDITIONAL`, `ReasonCodes[]`: WX_WAVE, WX_WIND, WX_SQUALL_BUFFER, WX_PEAK_WAVE, WX_WINDOW_GAP |

## 파라미터(사용자/SOP 입력)

- `Hs_limit_m`: 운항 허용 Hs 상한 (m)
- `Wind_limit_kt`: 운항 허용 풍속 상한 (kt)
- `SailingTime_hr`, `Reserve_hr`: 항해 시간 + 예비 시간 (연속 window 계산)
- (보수 운용 시) `ΔHs_squall_m`, `ΔGust_kt`: squall 미반영 시 보수 버퍼
- (피크파 운용 시) `Hmax_allow_m` 또는 `Hs_limit_m = Hmax_allow_m / 1.86`

**값 미제공 시**: 사용자에게 위 3개(Hs_limit_m 또는 Hmax_allow_m, Wind_limit_kt, SailingTime_hr+Reserve_hr) 입력 요청 후 버킷별 GO/NO-GO 출력 가능하다고 안내.

## 3단 Gate (의사결정 규칙)

### 1) Gate-A: 기본 임계값

- `Hs_m(t) = wave_ft(t) × 0.3048`
- `Wind_kt(t) = wind_kt(t)`
- **GO(t)** iff `Hs_m(t) ≤ Hs_limit_m` AND `Wind_kt(t) ≤ Wind_limit_kt`
  아니면 **NO-GO(t)**

### 2) Gate-B: Squall/피크파 보수화 (옵션 B·C)

- 차트에 "squall 미반영"이면:
  `Hs_eff = Hs_m + ΔHs_squall_m`, `Wind_eff = Wind_kt + ΔGust_kt`
  → 위 조건을 Hs_eff, Wind_eff로 평가.
- 피크파 기준 사용 시:
  `Hmax_est = 1.86 × Hs_eff` ([infoplaza.com])
  **NO-GO(t)** if `Hmax_est > Hmax_allow_m`

### 3) Gate-C: 연속 Weather window

- `NeedWindow_hr = SailingTime_hr + Reserve_hr`
- 출항 시각 `t0` 기준 **[t0, t0+NeedWindow]** 구간 **모든** 버킷이 GO → **최종 GO**
- 중간에 1개라도 NO-GO → **최종 NO-GO**, 해당 ReasonCode 기록

## 적용 옵션

| 옵션 | 내용 |
|------|------|
| **A** | Gate-A + Gate-C (Hs 기준 단순) |
| **B (권장)** | Gate-A + Gate-B(Δ 버퍼) + Gate-C |
| **C** | B + Hmax_est(=1.86×Hs) Gate |

## SSOT

- 2-days(시간대) 자료 있으면 **그걸 SSOT**
- 없으면 7-days(일 단위)로 후보 window만 뽑고, **최종은 최신 단기 예보로 재확정**

## QA

- **Hs**(유의파고, 상위 1/3 평균)와 **Hmax**(피크파, ≈1.86×Hs) 혼동 금지. 운영 기준이 Hs인지 Hmax인지 먼저 고정.

## 상세 로직·출처

- 전체 규칙·로직 테이블·출처 링크: [reference.md](reference.md)
- 원문: `AGI TR 1-6 Transportation Master Gantt Chart/weathergonnologic.md`
