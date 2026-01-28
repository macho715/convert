# Weather Go/No-Go — 상세 로직 (SEA TRANSIT)

원문: `AGI TR 1-6 Transportation Master Gantt Chart/weathergonnologic.md`

---

## Exec (해상 운행=SEA TRANSIT 전용 Go/No-Go 로직)

* 입력은 **Wave Height(=Combined Sea+Swell, ft)**, **Wind Speed(kt)**, (가능 시) **Wave Period(s)**로 통일하고, Wave는 기본적으로 **Hs(Significant Wave Height)**로 취급합니다. Hs 정의는 "상위 1/3 평균 파고"입니다. ([ndbc.noaa.gov][1])
* Go/No-Go는 **(1) 임계값 비교 Gate + (2) Squall/피크파고 버퍼 Gate + (3) 항해시간만큼 연속 Weather window Gate** 3단으로 결정합니다.
* 단일 최고파고(Hmax)는 통계적으로 **Hmax ≈ 1.86×Hs**로 다뤄지므로, "피크파" 안전여유를 로직에 반영합니다. ([infoplaza.com][2])

---

## EN Sources (≤3)

1. NOAA/NDBC — "WVHT(Significant Wave Height) 정의" ([ndbc.noaa.gov][1])
2. NDBC — "Hs(스펙트럼 기반) 산정 개요" ([ndbc.noaa.gov][3])
3. Hmax ≈ 1.86×Hs (피크파 관계) ([infoplaza.com][2])

---

## 로직 테이블 (SEA TRANSIT 전용)

| No | Item      | Value                                                          | Risk | Evidence/가정                                                              |
| -: | --------- | -------------------------------------------------------------- | ---- | ------------------------------------------------------------------------ |
|  1 | Wave 표준화  | `Hs_m = wave_ft × 0.3048`                                      | 중    | **가정:** 업로드 차트의 "combined sea+swell(ft)"를 Hs 근사로 사용 ([ndbc.noaa.gov][1]) |
|  2 | Wind 표준화  | `Wind_kt = chart_wind_kt`                                      | 하    | 단위 일치                                                                    |
|  3 | 피크파 보수화   | `Hmax_m ≈ 1.86 × Hs_m`                                         | 중    | ([infoplaza.com][2])                                                     |
|  4 | Squall 버퍼 | 차트에 "squall 미반영"이면 `Hs_eff=Hs_m+ΔHs`, `Wind_eff=Wind_kt+ΔGust` | 상    | Δ는 SOP 파라미터(현장 보수치)                                                      |
|  5 | 연속 window | `연속 GO 시간 ≥ (SailingTime + Reserve)`                           | 상    | 운항은 "한 시점 GO"가 아니라 "지속 GO" 필요 ([MBM Consultancy][4])                     |

---

## Go/No-Go 의사결정 규칙 (Pseudo-Logic)

### 0) 파라미터(현장/선장/SOP로 고정)

* `Hs_limit_m` : 운항 허용 Hs 상한
* `Wind_limit_kt` : 운항 허용 풍속 상한
* `SailingTime_hr` : 항해 예상 시간
* `Reserve_hr` : 예비시간(회항/대기/감속)
* `ΔHs_squall_m`, `ΔGust_kt` : squall/돌풍 보수 버퍼
* (선택) `Period_max_s` 또는 `Steepness_gate` : 주기/급파 리스크 게이트(텍스트 예보에 period가 있으면 사용)

### 1) SSOT 선택(해상 운항용)

* 2-days(시간대) 자료가 있으면 **그걸 SSOT**
* 없으면 7-days(일 단위)로 후보 window만 뽑고, **최종은 최신 단기 예보로 재확정**

### 2) Gate-A: 기본 임계값 비교

각 시간 버킷 `t`에 대해

* `Hs_m(t) = wave_ft(t)*0.3048`
* `Wind_kt(t) = wind_kt(t)`
* **GO(t)** if `Hs_m(t) ≤ Hs_limit_m` AND `Wind_kt(t) ≤ Wind_limit_kt`
* 아니면 **NO-GO(t)**

### 3) Gate-B: Squall/피크파 보수화

차트에 "squall 미반영"이 있으면(당신 업로드 자료에 명시)

* `Hs_eff(t) = Hs_m(t) + ΔHs_squall_m`
* `Wind_eff(t) = Wind_kt(t) + ΔGust_kt`
* **GO(t)** if `Hs_eff(t) ≤ Hs_limit_m` AND `Wind_eff(t) ≤ Wind_limit_kt`

추가로 **피크파 기준**을 쓰는 운영이면(예: "Hmax가 X m 넘으면 NO-GO")

* `Hmax_est(t) = 1.86 × Hs_eff(t)` ([infoplaza.com][2])
* **NO-GO(t)** if `Hmax_est(t) > Hmax_allow_m`
* (동일식 변형) `Hs_limit_m = Hmax_allow_m / 1.86`

### 4) Gate-C: 연속 Weather window 확보

* `NeedWindow_hr = SailingTime_hr + Reserve_hr`
* 출항시각 `t0`를 기준으로 **[t0, t0+NeedWindow]** 구간의 모든 버킷이 GO이면 **최종 GO**
* 중간에 1개라도 NO-GO가 끼면 **최종 NO-GO** (Reason Code 기록)

> 시간 버킷 지속시간은 관측/예보 특성상 "풍속 1h, 파고 3h"처럼 가정하는 방법이 문헌에 존재합니다(운영 계산용). ([MBM Consultancy][4])

### 5) 출력(로직 결과)

* `Decision: GO | NO-GO | CONDITIONAL`
* `ReasonCodes[]`: `WX_WAVE`, `WX_WIND`, `WX_SQUALL_BUFFER`, `WX_PEAK_WAVE`, `WX_WINDOW_GAP`

---

## Options (A/B/C) — 로직 적용 방식만

* **A (Hs 기준 단순 운용)**: Gate-A + Gate-C
* **B (보수 운용, 권장)**: Gate-A + Gate-B(Δ 버퍼) + Gate-C
* **C (피크파 관리 운용)**: B + `Hmax_est(=1.86×Hs)` Gate 추가 ([infoplaza.com][2])

---

## (Auto/QA)

* Auto: 2-days 그래프를 **시간대별 wave_ft/wind_kt**로 구조화 → 위 Gate를 자동 평가 → "연속 window" 탐색
* QA: **Hs 정의(상위 1/3 평균)**와 **피크파(≈1.86×Hs)**를 혼동하지 않도록, 운영 기준이 Hs인지 Hmax인지 먼저 고정 ([ndbc.noaa.gov][1])

---

## Acc (가정:)

* 업로드된 "combined sea & swell height(ft)"를 Hs 근사로 사용(차트 원문에 지표 정의가 없어서 **가정 처리**). ([ndbc.noaa.gov][1])
* ΔHs/ΔGust 및 limit 값은 **선장/Owner SOP 입력값**으로 둠(여기서 임의 수치 확정하지 않음).

---

## ZERO log (로직 확정은 가능, 값 확정은 입력 필요)

| 단계           | 이유                                                                               | 위험       | 요청데이터                                                                                          | 다음조치                         |
| ------------ | -------------------------------------------------------------------------------- | -------- | ---------------------------------------------------------------------------------------------- | ---------------------------- |
| Threshold 고정 | SEA TRANSIT의 `Hs_limit_m`, `Wind_limit_kt`, `ΔHs`, `ΔGust`, `SailingTime_hr` 미제공 | 과소/과대 제한 | **(1) Hs_limit_m (또는 Hmax_allow_m)**, **(2) Wind_limit_kt**, **(3) SailingTime_hr+Reserve_hr** | 입력 즉시 "시간대별 GO/NO-GO"로 출력 가능 |

원하면, 위 로직 그대로 **당신 업로드 2-days 시간대(08:00/14:00/20:00/02:00 …)**에 적용해 **버킷별 GO/NO-GO만** 뽑아드릴 수 있습니다(단, limit/버퍼 값 3개만 주면 됨).

---

[1]: https://www.ndbc.noaa.gov/faq/measdes.shtml "Measurement Descriptions and Units - NDBC"
[2]: https://www.infoplaza.com/en/blog/labeling-waves-the-nuances-of-hs-and-hmax "Labeling waves: the nuances of Hs and Hmax"
[3]: https://www.ndbc.noaa.gov/faq/wavecalc.shtml "How are significant wave height, dominant period, average ..."
[4]: https://www.mbm-consultancy.com/wp-content/uploads/2020/11/Guidelines-for-Marine-Transportations.pdf "technical policy board guidelines for marine transportations"
