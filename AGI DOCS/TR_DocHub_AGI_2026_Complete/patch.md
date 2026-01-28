## Exec (Now/Next/Alt 1회 통합 확정본)

* **Now(LATTICE+deep)**: *Status/DocCode/Party 표준값*을 “드롭다운 강제”로 고정하고, *Rules_Table(Anchor+Offset)*을 테이블 기반으로 확정합니다(코드가 아니라 데이터로 운영). 
* **Next(kpi-dash)**: D-7/D-3/D-1 임계값과 RAG(Overdue/DueSoon/OK)를 **수식 규칙**으로 고정합니다(조건부서식 + 집계 KPI). ([Microsoft Support][1])
* **Alt(KRsummary report)**: Export Pack(PDF/CSV/메일본문)을 “버튼 1개”로 고정합니다(현장 배포 표준). 

---

## EN Sources (≤3)

* Microsoft Support — *Create a drop-down list* — (page) ([Microsoft Support][2])
* Microsoft Support — *WORKDAY.INTL function* / *NETWORKDAYS.INTL function(Weekend string 0000011)* — (page) ([Microsoft Support][3])
* Microsoft Support — *Use conditional formatting to highlight information in Excel* — (page) ([Microsoft Support][1])

---

## Visual (표준값 + 룰 확정 테이블)

| No | Item        | Value                         | Risk         | Evidence/가정                                          |
| -: | ----------- | ----------------------------- | ------------ | ---------------------------------------------------- |
|  1 | Status 표준값  | 8개 고정(드롭다운)                   | 표현 흔들림 제거 필수 | Data Validation 권장 ([Microsoft Support][2])          |
|  2 | Party 표준코드  | 8개(현장 최소) + 확장                | 코드 불일치       | 통합빌더 PartyID 패턴                                      |
|  3 | DocCode(코어) | 8개(확정)                        | 문서 누락        | 통합빌더 기본 DocCode                                      |
|  4 | Rules_Table | AnchorField+OffsetDays+CAL/WD | WD 산정 오류     | WORKDAY.INTL/Weekend string ([Microsoft Support][3]) |
|  5 | RAG 임계값     | D-7/D-3/D-1 + Overdue         | 기준 미고정 시 혼선  | CF는 수식 기반 가능 ([Microsoft Support][1])                |

---

# NOW: /switch_mode LATTICE + /logi-master --deep report

## 1) Status 표준값(드롭다운 고정)

**Lists[Status] (8개 고정)**

1. Not Started
2. In Progress
3. Submitted
4. Accepted *(또는 Approved 중 1개만 선택해 고정; 권장=Accepted)*
5. Rejected
6. On Hold
7. Waived
8. Not Required

**운영 규칙(고정)**

* KPI 완료 처리 Status = `Accepted` 또는 `Waived` 또는 `Not Required`
* Overdue 계산 대상 Status = `Not Started/In Progress/Submitted/Rejected/On Hold`만

(드롭다운/유효성 검사는 Data Validation로 강제) ([Microsoft Support][2])

---

## 2) Party 표준코드(코어 8개 + 확장)

**M_Parties[PartyID] (코어 8개, 확정)** — 통합빌더 패턴과 동일 

* FF = Freight Forwarder
* CUSTBROKER = Customs Broker
* EPC = EPC Contractor
* TRCON = Transport Contractor
* PORT = Port Authority
* OFCO = OFCO Agency
* MMT = Mammoet
* SCT = Samsung C&T

**확장(선택, 가정:)**

* CARRIER(선사), SURVEY(MWS/검사), CLIENT(ADNOC) 등

---

## 3) DocCode 표준(코어 8개 “확정” + 확장 “선택”)

### 코어 DocCode 8개(확정) — 통합빌더 기본 문서 

* GATEPASS, CUSTOMS, PERMIT, BL, STOWAGE, LASHING, MWS, NOC

### 확장 DocCode(선택, 가정: TR 운송 일반 문서)

* PTW, METHODSTATEMENT, RISKASSESS, LIFTPLAN, ROUTESURVEY, INSURANCE, CI, PL, COO, DO, BOE, PACKINGLIST

> 확장 항목은 귀사/ADNOC 요구 문서 리스트와 일치 여부 확인 후 “확정 목록”으로 승격하십시오(가정 제거).

---

## 4) Rules_Table 확정본(Anchor+Offset+CAL/WD)

### AnchorField Enum(고정)

* MZP Arrival / Load-out / MZP Departure / AGI Arrival / Doc Deadline / Land Permit By
  (통합빌더에 이미 포함된 패턴) 

### DueDate 계산 규칙(수식/로직 고정)

* **CAL**: `DueDate = AnchorDate + OffsetDays`
* **WD**: `DueDate = WORKDAY.INTL(AnchorDate, OffsetDays, WeekendPattern, HolidaysRange)` ([Microsoft Support][3])

> WeekendPattern 텍스트 예: `"0000011"` = 토/일 휴무 ([Microsoft Support][4])

### 코어 Rules (권장 확정안: 통합빌더 예시 기반) 

| DocCode  | AnchorField   | OffsetDays | CalendarType | Priority |
| -------- | ------------- | ---------: | ------------ | -------: |
| GATEPASS | Load-out      |         -1 | CAL          |        1 |
| CUSTOMS  | Doc Deadline  |         -2 | WD           |        1 |
| PERMIT   | MZP Arrival   |          0 | CAL          |        1 |
| BL       | MZP Departure |         -3 | WD           |        1 |
| STOWAGE  | Load-out      |         -2 | CAL          |        1 |
| LASHING  | Load-out      |         -2 | CAL          |        1 |
| MWS      | MZP Departure |         -5 | WD           |        1 |
| NOC      | AGI Arrival   |         -7 | WD           |        1 |

---

# NEXT: /logi-master kpi-dash

## 1) 임계값(C_Config 고정)

* Amber_Threshold_Days = **7.00** (D-7)
* Red_Threshold_Days = **3.00** (D-3)
* Critical_Threshold_Days = **1.00** (D-1)

## 2) RAG 규칙(수식 고정)

**RAG 우선순위(상단이 우선)**

1. `CLOSED` : Status ∈ {Accepted, Waived, Not Required}
2. `OVERDUE` : DueDate < TODAY() AND NOT CLOSED
3. `CRITICAL` : DueDate <= TODAY()+1 AND NOT CLOSED
4. `RED` : DueDate <= TODAY()+3 AND NOT CLOSED
5. `AMBER` : DueDate <= TODAY()+7 AND NOT CLOSED
6. `OK` : 그 외

**조건부서식(CF) 적용 방식**: “수식 결과에 따라 행 채색/아이콘”으로 구현 가능 ([Microsoft Support][1])

---

# ALT: /logi-master --KRsummary report

## Export Pack 표준(고정)

### 1) PDF

* 포함: `D_Dashboard` + (필터된) `T_Tracker`(해당 VoyageID 또는 Party 기준)
* 파일명 규칙(예): `TR_DocPack_V03_YYYYMMDD.pdf`

### 2) CSV

* `T_Tracker`에서 **필터 조건(항차/책임자/상태)** 적용 후 Export
* 파일명 규칙: `TR_Tracker_V03_YYYYMMDD.csv`

### 3) 메일 본문(Outlook 붙여넣기용, 표 포함)

* Subject 규칙: `[TR][Voyage V03] Document Submission Status – D-7/D-3/D-1`
* Body 구성(고정):

  * 1줄 요약(Overdue/CRITICAL 건수)
  * Action Items(Responsible Party별 Top 5)
  * 첨부 리스트(PDF/CSV)
  * 다음 업데이트 시간

> 이 Export Pack/버튼 매핑은 통합빌더의 운영 철학(Generate/Recalc/Validate/Export)을 그대로 따릅니다. 

---

## Options (A/B/C)

| Option                      | Pros                        | Cons     |     Cost | Risk | Time       |
| --------------------------- | --------------------------- | -------- | -------: | ---- | ---------- |
| A) 수기만                      | 가장 빠름                       | 누락 방지 약함 | AED 0.00 | 높음   | 0.50–1.00일 |
| **B) Excel+VBA+Python(확정)** | 수기 입력 유지 + 표준화/검증/Export 자동 | 초기 세팅 필요 | AED 0.00 | 중하   | 2.00–5.00일 |
| C) B + 외부 DMS               | 승인/버전 통제까지 완성               | 정책/권한    |     AED↑ | 낮음   | 2.00–6.00주 |

---

## Steps (P→Pi→B→O→S + KPI)

1. **Prepare**: Lists(Status/Party/DocCode) 고정 + Rules_Table 입력
2. **Pilot**: Voyage 1로 KPI/RAG/Export Pack 검증
3. **Build**: VBA 버튼 4개(Generate/Recalc/Validate/Export) 고정 + Python 백업 스냅샷
4. **Operate**: 수기 입력 후 매일 1회 Export Pack 배포
5. **Scale**: DocCode 확장 시 “테이블만” 추가(코드 수정 최소)

**KPI**

* D-3 Submitted ≥ 98.00%
* Overdue 0.00건/주
* Export Pack ≤ 10.00분

---

## (Auto/QA) Fail-safe

* Data Validation(드롭다운)으로 Status/Party/DocCode 오입력 차단 ([Microsoft Support][2])
* WD 계산은 WORKDAY.INTL + WeekendPattern(0000011) + Holidays로 고정 ([Microsoft Support][3])
* 조건부서식은 RAG 셀 값을 기반으로 행 색상 적용(Overdue/CRITICAL 우선) ([Microsoft Support][1])

---

## Acc (가정:)

* 가정: WeekendPattern 기본값은 `"0000011"`(토/일 휴무)이며 필요 시 프로젝트 캘린더로 조정
* 가정: 확장 DocCode는 귀사/ADNOC 요구 문서 리스트 확인 후 확정

---

## 🔧/cmd3 (Now/Next/Alt)

* Now: `/switch_mode LATTICE + /logi-master --deep report` (본 문서의 표준값/룰테이블을 Excel Lists/R_DeadlineRules에 반영)
* Next: `/logi-master kpi-dash` (RAG/임계값 수식+조건부서식 확정)
* Alt: `/logi-master --KRsummary report` (Export Pack PDF/CSV/메일본문 템플릿 고정)

---

## ZERO log

| 단계 | 이유                                            | 위험 | 요청데이터                              | 다음조치                                   |
| -- | --------------------------------------------- | -- | ---------------------------------- | -------------------------------------- |
| 계속 | 코어(8 DocCode/8 Party/8 Status) 기준으로는 즉시 운영 가능 | 낮음 | (선택) 확장 DocCode를 “확정 목록”으로 승격할지 여부 | 승격 시 확장 룰(Anchor/Offset)까지 같이 확정본으로 패치 |

[1]: https://support.microsoft.com/en-us/office/use-conditional-formatting-to-highlight-information-in-excel-fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f?utm_source=chatgpt.com "Use conditional formatting to highlight information in Excel"
[2]: https://support.microsoft.com/en-us/office/create-a-drop-down-list-7693307a-59ef-400a-b769-c5402dce407b?utm_source=chatgpt.com "Create a drop-down list"
[3]: https://support.microsoft.com/en-us/office/workday-intl-function-a378391c-9ba7-4678-8a39-39611a9bf81d?utm_source=chatgpt.com "WORKDAY.INTL function"
[4]: https://support.microsoft.com/en-us/office/networkdays-intl-function-a9b26239-4f20-46a1-9ab8-4e925bfd5e28?utm_source=chatgpt.com "NETWORKDAYS.INTL function"



## Exec (LATTICE+deep “반영” 확정본)

* **Lists(Status/기타 드롭다운)와 R_DeadlineRules(룰테이블)** 를 “테이블/드롭다운 강제”로 고정해, 수기 입력 방식에서도 **오입력·표현 흔들림·누락**을 구조적으로 차단합니다. ([Microsoft Support][1])
* 현재 `통합빌더.py`는 코어 룰 8개를 이미 포함하고 있으나, **Status 8개 확정(=Not Required 추가)**에 맞춰 **Data Validation 범위(=Lists!$A$2:$A$8 → $A$2:$A$9)** 를 반드시 수정해야 합니다. 
* 아래 “PATCH”를 적용하면, 귀하가 요청한 **표준값/룰테이블이 Excel의 Lists/R_DeadlineRules에 그대로 반영**됩니다. 

---

## EN Sources (≤3)

* Microsoft Support — *Create a drop-down list* (Data Validation: List) ([Microsoft Support][1])
* Microsoft Support — *WORKDAY.INTL function* (weekend string “0000011”) ([Microsoft Support][2])
* Microsoft Support — *Apply data validation to cells* ([Microsoft Support][3])

---

## Visual (반영 대상: 표준값/룰테이블)

| No | Item                   | Value                                                                               | Risk                 | Evidence/가정                                |
| -: | ---------------------- | ----------------------------------------------------------------------------------- | -------------------- | ------------------------------------------ |
|  1 | Status(확정 8개)          | Not Started/In Progress/Submitted/Accepted/Rejected/On Hold/Waived/**Not Required** | DV 범위 미수정 시 마지막 값 누락 | DV는 리스트 범위 참조 ([Microsoft Support][1])     |
|  2 | R_DeadlineRules(확정 8개) | GATEPASS~NOC 룰 8행                                                                   | 없음                   | rules_data 이미 포함                           |
|  3 | WD 계산(주말패턴)            | WORKDAY.INTL + “0000011”                                                            | 프로젝트 휴일 미반영          | weekend string 정의 ([Microsoft Support][2]) |
|  4 | 현재 DV 버그 포인트           | `dv_status`가 `=Lists!$A$2:$A$8`로 고정                                                 | Status 8개 확정 시 불일치   | 코드 확인                                      |

---

# 1) 표준값 “확정본” (Lists 시트에 반영)

## A) Lists!A열 Status (A2:A9)

1. Not Started
2. In Progress
3. Submitted
4. Accepted
5. Rejected
6. On Hold
7. Waived
8. Not Required

> 드롭다운(데이터 유효성 검사)로 입력값을 강제하는 것이 핵심입니다. ([Microsoft Support][1])

## B) Lists!C열 Due_Basis(Anchor Enum)

현재 `통합빌더.py`의 Due_Basis는 다음을 포함하고 있습니다. 

* Doc Deadline, Land Permit By, MZP Arrival, Load-out, MZP Departure, AGI Arrival
  (선택) AUTO는 운영방식에 따라 유지/삭제 가능

---

# 2) 룰테이블 “확정본” (R_DeadlineRules에 반영)

`통합빌더.py`의 rules_data는 이미 귀하가 원하는 코어 룰 8개를 포함합니다. 

| RuleID | DocCode  | AnchorField   | OffsetDays | CalendarType | Priority |
| ------ | -------- | ------------- | ---------: | ------------ | -------: |
| R001   | GATEPASS | Load-out      |         -1 | CAL          |        1 |
| R002   | CUSTOMS  | Doc Deadline  |         -2 | WD           |        1 |
| R003   | PERMIT   | MZP Arrival   |          0 | CAL          |        1 |
| R004   | BL       | MZP Departure |         -3 | WD           |        1 |
| R005   | STOWAGE  | Load-out      |         -2 | CAL          |        1 |
| R006   | LASHING  | Load-out      |         -2 | CAL          |        1 |
| R007   | MWS      | MZP Departure |         -5 | WD           |        1 |
| R008   | NOC      | AGI Arrival   |         -7 | WD           |        1 |

**WD(Working Day) 계산 표준**: `WORKDAY.INTL(AnchorDate, OffsetDays, "0000011", Holidays)` 형태로 고정합니다. ([Microsoft Support][2])
(참고) DocGap 빌더에서도 주말패턴 “0000011”을 명시적으로 사용합니다. 

---

# 3) PATCH (통합빌더.py에 반영해야 할 최소 수정 2건)

## PATCH-01: Status 8개로 Lists 채우기

* `create_lists_sheet()`의 `status_list`에 **"Not Required"** 추가
* (근거: Lists에 Status 열을 채우는 구조는 이미 존재) 

## PATCH-02: T_Tracker의 Status DataValidation 범위 확장

현재 `dv_status`가 `=Lists!$A$2:$A$8`로 고정되어 있습니다. 

* 이를 `=Lists!$A$2:$A$9`로 변경해야 **Not Required**가 드롭다운에 포함됩니다.

---

# 4) 적용 절차 (수기 입력 운영 기준)

## Steps (P→Pi→B→O→S + KPI)

1. **Prepare**: `통합빌더.py`에 PATCH-01/02 적용
2. **Pilot**: Python으로 템플릿 .xlsx 생성 → Excel에서 .xlsm 저장 → Status 드롭다운에 Not Required 표시 확인
3. **Build**: VBA Import 후, Generate/Recalc/Validate/Export 버튼 흐름 점검
4. **Operate**: 귀하 수기 입력 후 1일 1회 Recalc/Validate
5. **Scale**: DocCode/룰 확장은 “테이블 행 추가”만으로 처리

**KPI**

* 오입력(비표준 Status) 0.00건/주
* Overdue 0.00건/주(운영 목표)
* 수기 입력→Export Pack ≤ 10.00분

---

## (Auto/QA) 체크 포인트(필수)

* Status 드롭다운이 실제로 “리스트 범위”를 참조하는지 확인(시트 보호 상태면 Data Validation 수정 불가). ([Microsoft Support][1])
* WD 룰 사용 시 Holidays 범위가 비어 있더라도 함수가 동작하는지 확인(프로젝트 휴일은 추후 추가). ([Microsoft Support][2])

---

## Acc (가정:)

* 가정: Status 확정값 중 “Accepted”를 최종 완료 상태로 사용(Approved로 바꾸려면 **전체 KPI 수식/조건부서식도 같이 변경** 필요)

---

## 🔧/cmd3 (Now/Next/Alt)

* Now: `/switch_mode LATTICE + /logi-master --deep report` (본 PATCH-01/02 적용 후 템플릿 재생성)
* Next: `/logi-master kpi-dash` (D-7/D-3/D-1 + RAG 규칙을 Dashboard/CF로 고정)
* Alt: `/logi-master --KRsummary report` (Export Pack: PDF/CSV/메일본문 표준 고정)

---

## ZERO log

| 단계 | 이유                                 | 위험 | 요청데이터 | 다음조치                                          |
| -- | ---------------------------------- | -- | ----- | --------------------------------------------- |
| 계속 | 룰테이블/드롭다운 반영은 규정·요율·ETA 근거 요구사항 아님 | 낮음 | 없음    | 위 PATCH 2건 적용 시 즉시 “표준값/룰테이블 반영 완료” 상태로 운영 가능 |

[1]: https://support.microsoft.com/en-us/office/create-a-drop-down-list-7693307a-59ef-400a-b769-c5402dce407b?utm_source=chatgpt.com "Create a drop-down list"
[2]: https://support.microsoft.com/en-us/office/workday-intl-function-a378391c-9ba7-4678-8a39-39611a9bf81d?utm_source=chatgpt.com "WORKDAY.INTL function"
[3]: https://support.microsoft.com/en-us/office/apply-data-validation-to-cells-29fecbcc-d1b9-42c1-9d76-eff3ce5f7249?utm_source=chatgpt.com "Apply data validation to cells"
