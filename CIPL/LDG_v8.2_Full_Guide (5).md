# 📗 ChatGPT 고정밀 OCR·구조화 **Full Guide v2.4 – LDG Ready (Linked·Agent Mode)**
*(2025‑08‑10 · Asia/Dubai · LDG v8.2 연동 · 본 문서는 Mini 지침의 보조 자료입니다)*

## 목차
- [1) 목적·적용 범위](#1)-목적적용-범위)
- [2) 원칙(증거·신뢰·투명)](#2)-원칙(증거신뢰투명))
- [3) KPI·품질 게이트](#3)-kpi품질-게이트)
- [4) 파이프라인 상세(6단계)](#4)-파이프라인-상세(6단계))
- [5) 자동 검증(5단계)](#5)-자동-검증(5단계))
- [6) LDG_PAYLOAD v2.4 스키마](#6)-ldg_payload-v24-스키마)
- [7) Compliance·RegTech 운영](#7)-complianceregtech-운영)
- [8) DEM/DET Forecast 2.0](#8)-demdet-forecast-20)
- [9) 비용·환율 — COST‑GUARD](#9)-비용환율-costguard)
- [10) HS‑RISK](#10)-hsrisk)
- [11) Fail‑Safe ZERO](#11)-failsafe-zero)
- [12) 보고서(7+2 섹션)](#12)-보고서(72-섹션))
- [13) 명령셋](#13)-명령셋)
- [14) 운영 체크리스트](#14)-운영-체크리스트)
- [15) 중단 로그 템플릿](#15)-중단-로그-템플릿)
- [부록 A) 스키마 종류 맵 & 필드 정의(현업형)](#부록-a)-스키마-종류-맵-필드-정의(현업형))
  - [A-1) LDG 스키마 전반(업무 요약)](#a-1)-ldg-스키마-전반(업무-요약))
- [1) 한눈에: 스키마 분류 맵](#1)-한눈에-스키마-분류-맵)
  - [1) 코어 인입·정규화](#1)-코어-인입정규화)
  - [2) 규정·컴플라이언스](#2)-규정컴플라이언스)
  - [3) 운영·예측](#3)-운영예측)
  - [4) 비용·단가(참조 테이블 계열)](#4)-비용단가(참조-테이블-계열))
  - [5) 조정·국내 레퍼런스](#5)-조정국내-레퍼런스)
  - [6) 레인·승인지도](#6)-레인승인지도)
  - [7) 감사·출처 추적(프루프·영수증)](#7)-감사출처-추적(프루프영수증))
  - [8) 코스트가드 판정](#8)-코스트가드-판정)
  - [9) 리포트·UI 응답(대화 출력 규격)](#9)-리포트ui-응답(대화-출력-규격))
  - [10) FAIL-SAFE](#10)-fail-safe)
- [1) LDG_PAYLOAD v2.4](#1)-ldg_payload-v24)
- [2) LDG_AUDIT v2.4](#2)-ldg_audit-v24)
- [3) evidence 레코드](#3)-evidence-레코드)
- [4) CERT_CHK 결과](#4)-cert_chk-결과)
- [5) HS_RISK](#5)-hs_risk)
- [6) DEMDET_FORECAST v2.0](#6)-demdet_forecast-v20)
- [7) COST_GUARD_RESULT (표 + 요약)](#7)-cost_guard_result-(표-요약))
- [8) container_cargo_rates (참조 JSON)](#8)-container_cargo_rates-(참조-json))
- [9) bulk_cargo_rates (참조 JSON)](#9)-bulk_cargo_rates-(참조-json))
- [10) air_cargo_rates (참조 JSON)](#10)-air_cargo_rates-(참조-json))
- [11) inland_trucking_reference_rates_clean (참조 JSON)](#11)-inland_trucking_reference_rates_clean-(참조-json))
- [12) domestic_reference.json (번들)](#12)-domestic_referencejson-(번들))
- [13) ref_adjusters.json](#13)-ref_adjustersjson)
- [14) ref_min_fare.json](#14)-ref_min_farejson)
- [15) ApprovedLaneMap_ENHANCED](#15)-approvedlanemap_enhanced)
- [16) Unified Response JSON v1.2 (대화/리포트 래퍼)](#16)-unified-response-json-v12-(대화리포트-래퍼))
- [17) ZERO_FAILSAFE_LOG](#17)-zero_failsafe_log)
  - [A-2) Full Guide 본문(동기화본)](#a-2)-full-guide-본문(동기화본))
- [1) 목적·적용 범위](#1)-목적적용-범위)
- [2) 원칙(증거·신뢰·투명)](#2)-원칙(증거신뢰투명))
- [3) KPI·품질 게이트](#3)-kpi품질-게이트)
- [4) 파이프라인 상세(6단계)](#4)-파이프라인-상세(6단계))
- [5) 자동 검증(5단계)](#5)-자동-검증(5단계))
- [6) LDG_PAYLOAD v2.4 스키마](#6)-ldg_payload-v24-스키마)
- [7) Compliance·RegTech 운영](#7)-complianceregtech-운영)
- [8) DEM/DET Forecast 2.0](#8)-demdet-forecast-20)
- [9) 비용·환율 — COST-GUARD](#9)-비용환율-cost-guard)
- [10) HS-RISK](#10)-hs-risk)
- [11) Fail-Safe ZERO](#11)-fail-safe-zero)
- [12) 보고서(7+2 섹션)](#12)-보고서(72-섹션))
- [13) 명령셋](#13)-명령셋)
- [14) 운영 체크리스트](#14)-운영-체크리스트)
- [15) 중단 로그 템플릿](#15)-중단-로그-템플릿)

## 1) 목적·적용 범위
- 본 가이드는 업로드 문서(CIPL, BL, PL, Invoice, COO, AWB/MAWB 등)를 **고정밀 OCR→정규화→검증→LDG 연동**하는 전 과정을 표준화한다.
- **내부 도구(web.run, Python)**만 사용하며, 외부 API·비인가 저장소 사용을 금지한다. PII는 마스킹하고 모든 산출물에 해시를 포함한다.

## 2) 원칙(증거·신뢰·투명)
- **증거 우선(HallucinationBan)**: 정부·항만·세관 공시를 최우선 인용, 모든 인용은 `evidence[]`에 *제목/기관/발행일/URL* 저장.
- **투명한 가정**: 불확실·부족 데이터는 `가정:`으로 명시 후 사람 검토 게이트.
- **재현성**: 수치는 소수점 2자리, 테이블은 CSV로 임베드, 파이프라인·파라미터는 로그에 기록.

## 3) KPI·품질 게이트
- MeanConf ≥ 0.92, TableAcc ≥ 98.00%, NumericIntegrity = 1.00, EntityMatch ≥ 0.98, HashConsistency = PASS, CrossDocConsistency = PASS.
- **Excel 생성 모드**: KPI 미달 시 **Soft Warning** (생성 차단 없음, 경고만 기록, Human Gate 권장)
- **OCR 파이프라인 모드**: 미달 시 **ZERO** 전환 (ingest 차단, 중단 로그 출력, 교정 루틴(`/ocr_lowres_fix`, `/ocr_retry`, `/ocr_align`) 제안)

## 4) 파이프라인 상세(6단계)
1. **Pre‑Prep**: 오토 회전·데스큐, 콘트라스트/샤프닝, 노이즈 억제, DPI 보정(권장 ≥300dpi).
2. **Vision OCR**: 레이아웃 블록·라인 인식, 언어 감지, confidence 수집, 페이지별 메트릭 추적.
3. **Smart Table Parser 2.1**: 병합셀 해제, 다중 헤더 정규화, 단위/통화 분리, 세로→가로 피벗, 라인아이템 복원.
4. **NLP Refine**: 단위 규격화(kg, m³, EA), 수량·단가·합계 일관성 검사, NULL/??/추정 표기.
5. **Field Tagger**: Shipper/Consignee/BL_No/Invoice_No/Incoterms/Origin/HS/UN_No/IMDG 등을 키‑값으로 태깅.
6. **Payload Builder**: LDG_PAYLOAD v2.4 직렬화 + `doc_hash`·`pages`·시간·교차 링크 + CSV 테이블 임베드.

## 5) 자동 검증(5단계)
A. 1차: KPI/스키마/타입·필수 필드/Null·합계 등 기본 무결성.  
B. 표현 로직: 소계↔총계, 통화·환율, 반올림, 음수/0 값 처리, 단위 불일치 탐지.  
C. 교차 검증: OCR Raw↔Refined, CIPL↔BL↔PL 키필드 매핑 및 Δ 계산(허용: Qty ±1, Wt ±15.00kg, CBM ±0.02m³).  
D. 적합도 산출: 항목 가중(테이블 일치·키필드·수치·규정 근거)로 0–100%.  
E. 최종 보고: LDG_PAYLOAD + LDG_AUDIT(Warnings, CrossDocCheck, ZERO_Failsafe, HashConsistency).

## 6) LDG_PAYLOAD v2.4 스키마
```json
{
  "meta": {
    "source_file": "…",
    "doc_hash": "sha256:…",
    "ocr_version": "v2.4",
    "mean_conf": 0.94,
    "table_acc": 0.99,
    "numeric_integrity": true,
    "pages": 7,
    "created_at": "2025-08-10T12:34:56+04:00"
  },
  "data": {
    "DocType": "B/L + CIPL",
    "BL_No": "…",
    "Invoice_No": "…",
    "Shipper": "…",
    "Consignee": "…",
    "Incoterms": "FOB",
    "Currency": "USD",
    "Packages": 14,
    "Gross_Weight_kg": 18250.00,
    "Net_Weight_kg": 17780.00,
    "CBM": 38.40,
    "HS_Candidate": ["85044090","85049010","85049090"],
    "Origin": "KR",
    "port_profile": {"discharge":"Khalifa","free_time":{"dem":5,"det":10}},
    "eta": {"value":"2025-08-18","source":"carrier notice","confidence":0.80},
    "demdet_profile": {"rule_id":"khalifa:5/10","valid_until":"2025-12-31"},
    "tables": [{"table_id":"bl_items","format":"csv","data":"Header1,Header2\nVal1,Val2"}]
  },
  "evidence": [
    {"type":"regulation","title":"Customs Notice …","url":"https://…","published":"2025-07-28"}
  ],
  "warnings": ["Page 3 설명 누락 → ?? 표시"]
}
```

## 7) Compliance·RegTech 운영
- **Mini‑RAG(24h)**: 정부·세관·항만 공시를 주기적 갱신. `/rag view [kw]`로 캐시 점검, 필요 시 `/system flush cache minirag`.
- **CERT‑CHK**: HS/키워드 기반으로 FANR/MOIAT/TRA/MOFAIC 요구 가능성 추정(결정 아님). 인용은 `evidence[]` 기록.
- **IMDG/UN**: 위험물 키워드 감지 시 클래스/UN_No 후보를 태깅하고, 확인 불가 시 `null`과 조치 제안.

## 8) DEM/DET Forecast 2.0
- **입력**: ETA 값·소스·신뢰도, Port Profile(Free Time), 공휴일/주말, 혼잡 가중치(가정 또는 출처).  
- **출력**: dem/det 만료일·잔여일·위험도·비용발생 예상·Gate Out 일자.  
- **명령**: `/demdet set profile {port}:{dem}d/{det}d [valid:YYYY MM DD]`, `/demdet set eta {cntr#} {YYYY MM DD} [source:"…"]`.

## 9) 비용·환율 — COST‑GUARD
- 기준단가 대비 초과율로 위험 등급. 원화폐 보존, 환율은 문서 기준 없으면 `fx_source:"none"`, 수동 입력 시 `fx_source:"manual"`과 로그.
- 출력: PASS/FAIL, 초과 항목 표, 요약 통계, 근거 링크.

## 10) HS‑RISK
- HS 후보 최대 3개 + confidence. 다의성·면제/허가 가능성을 태그하고 규정 근거 수집 요청.
- 결론 전 반드시 사람 검토. 분쟁 시 규정 상위 근거(정부 공시) 우선.

## 11) Fail‑Safe ZERO

### Excel 생성 모드 (Soft Warning)
- ZERO는 **Soft Warning**으로 처리 (Excel 생성 차단 없음)
- LDG_AUDIT.zero_flags에 기록하되 최선 노력으로 출력
- Fallback 전략:
  - 템플릿 로드 실패 → 기본 템플릿 사용
  - 입력 파일 파손 → 부분 데이터 추출, 누락 필드는 NULL
  - Excel 저장 실패 → JSON 출력 + 에러 메시지

### OCR 파이프라인 모드 (Hard Block)
- 트리거: KPI 미달, 핵심 식별자 결손, Hash 실패, 규정 근거 부재/상충
- **중단 로그**에는 사유·수치·근거·권고 조치(`/ocr_lowres_fix`, 재스캔, 정정본 요청, 캐시 초기화)를 포함

## 12) 보고서(7+2 섹션)
1) Auto Guard Summary  1.5) Risk Assessment(동인·신뢰도)  
2) Discrepancy Table(Δ·허용오차·상태)  3) Compliance Matrix(상태/근거/조치)  
4) Auto‑Fill: Freight & Insurance  5) Auto Action Hooks  
6) DEM/DET & Gate Out Forecast  7) Evidence & Citations  
8) Weak Spot & Improvements  9) Changelog

## 13) 명령셋
- `/ocr_basic {file} [mode:LDG|LDG+]`, `/ocr_table {file} [--colfix --unitnorm --as=csv]`, `/ocr_techdoc {file}`  
- `/ocr_retry {file}`, `/ocr_lowres_fix {file}`, `/ocr_align {cipl} {bl} [pl]`, `/ocr_compare {A} {B} [--tol=0.02]`  
- `/ocr_hs_seed {file} [--origin=AE --max=3]`, `/ocr_certchk {payload.json} [--moiat --fanr --imdg]`, `/ocr_costguard {file} {cost_table.csv} [--lock-fx]`  
- `/workflow scan email {file} {to}`, `/workflow full_checkup {file} [to]`

## 14) 운영 체크리스트
- OCR KPI 수치(2자리) 기록, Δ 계산·상태 이모지, evidence 수집, DEM/DET 프로필·ETA 소스/신뢰도, COST‑GUARD·HS‑RISK 결과, ZERO 로그.

## 15) 중단 로그 템플릿
```text
[중단|ZERO] LDG v8.2
사유: TableAcc 96.40% + CIPL↔BL 패키지 Δ +3
조치: 재스캔(≥300dpi), 정정본 업로드, /system flush cache all 후 재분석
근거: evidence[0] 발행일 경과(2023-…)
```

---

## 부록 A) 스키마 종류 맵 & 필드 정의(현업형)
> 통합 근거(업로드 파일): `LDG.MD`, `7257fb4b-64b5-4c27-b401-dfa7523e8403.md`  
> 병합시각(Asia/Dubai): 2025-11-03T19:53:07+04:00

### A-1) LDG 스키마 전반(업무 요약)
아래는 **LDG v8.2 × OCR v2.4 워크플로**에서 실제로 쓰는 **스키마 종류(전체)**를 용도별로 정리한 것입니다. (핵심 항목과 근거 라인 함께 표기)

---

# 1) 직렬화(Serialization) 스키마

* **LDG_PAYLOAD v2.4**

  * Top-level: `meta`, `data`, `evidence`, `warnings`. `meta`에 `doc_hash/ocr_version/pages/created_at` 등, 수치 KPI(`mean_conf`, `table_acc`, `numeric_integrity`) 포함.
  * `data` 요약: `DocType/Identifiers(BL_No, Invoice_No)/Parties(Shipper, Consignee)/Trade(Incoterms, Currency)/Logistics(Packages, Weights, CBM)/HS_Candidate/RegTechFlags/Financials(Total_Amount)/Linked Ops(port_profile, eta, demdet_profile)/Tables(csv 임베드)` .
  * `evidence[]` 표준 필드(기관·발행일·URL 등) 및 우선순위.

* **LDG_AUDIT (표준 감사 결과 JSON)**

  * 필드: `SelfCheck{critic_mode, fx_source, fx_locked}`, `TotalsCheck{sum_lines, doc_total, delta}`, `CrossDocCheck[]`, `HashConsistency`, `ZERO_Failsafe`, `Warnings[]` (합계 오차=0.00 원칙).

---

# 2) 도메인 온톨로지(RDF/OWL/JSON-LD) 스키마

* **핵심 클래스(요약)**:
  `ldg:Document, ldg:Page, ldg:Image, ldg:OCRBlock/OCRToken, ldg:Table, ldg:RefinedText, ldg:EntityTag, ldg:Payload, ldg:Validation, ldg:Metric, ldg:Audit, ldg:CrossLink, ldg:RegTechFlag, ldg:HSCandidate, ldg:CostGuardCheck`(rate 검증).

  * 처리단계와의 연결(입력→OCR→정제→검증→CostGuard→RegTech).

* **주요 객체·관계(OWL ObjectProperty) 스키마**:
  `ldg:hasPage/hasImage/partOf/contains/extractedFrom/parsedFrom/refines/tags/buildsFrom/validates/measures/audits/links/triggeredBy/proposedBy` 등 파이프라인 전단계 연결 정의.

* **주요 속성(OWL DatatypeProperty) 스키마**:
  `ldg:hasPageNumber/hasImageRef/hasImageHash/hasResolution/hasText/hasPosition/hasSchema/hasType …` 등 레이아웃·OCR 좌표·스키마명 보유.

* **JSON-LD 예시 스키마(컨텍스트·타입·속성 매핑)**:

  * Document Guardian 예시(문서→엔티티 타입·값·신뢰도 매핑).
  * OCR Pipeline 예시(Document/Page/Image/OCRBlock→Payload/Validation/Metric 연결).

* **Flow Code 관련 추출·검증 스키마(표/규칙·RDF 속성)**:
  문서유형별 추출필드 표/검증 규칙 + `ldg:extractedFlowCode`, `ldg:flowCodeConfidence` 등.

---

# 3) OFCO 인보이스 계열 스키마

* **OFCO 인보이스 시트 JSON Schema (2024-09~2025-03 배치)**
  Top-level: `type:"object"`, `required:["invoice_no","invoice_date","currency","lines"]`, `properties{invoice_no, invoice_date, supplier, buyer, currency, vat_rate_pct, total_excl_tax, total_tax, total_incl_tax, lines…}` 정의.

* **OFCO 정규화 라인(Line) 스키마(표준 컬럼)**
  열 정의: `invoice_no, line_no, tariff_id, description, unit1~unit3, rate, amount_excl_tax, tax_rate_pct, tax_amount, total_incl_tax, calc_check, evidence`(CSV/테이블로 고정).

* **OFCO 매핑·검증 로직(스키마 적용 위치)**
  대상 시트/입력열(예: `BJ=Total`, `BB:BI=EA/Rate`, `K:BA=코드/설명/VAT/통화`), KPI 및 ZERO 규칙.

---

# 4) 운영·예측 부가 스키마

* **DEM/DET Forecast 2.0 입력/링크 필드**:
  `eta{value, source, confidence}`, `port_profile.free_time{dem, det}` 및 LDG payload 내 Linked Ops로 직렬화.

* **COST-GUARD 판단 스키마(요율 감사 단위)**:
  온톨로지 엔티티 `ldg:CostGuardCheck{standardRate, draftRate, exceedPct, verdict}`.

---

# 5) 파이프라인·품질 게이트 스키마(참조)

* **KPI 게이트 & ZERO Fail-safe 트리거 정의**:
  MeanConf≥0.92, TableAcc≥0.98, NumericIntegrity=1.00, CrossDoc 일치율 기준 등(감사·리포트 연동).

---

## 1) 한눈에: 스키마 분류 맵

* 직렬화: **LDG_PAYLOAD v2.4 / LDG_AUDIT**
* 온톨로지: **클래스·관계(OWL)** + **JSON-LD 컨텍스트** + **Flow Code 규칙·속성**
* OFCO 전용: **인보이스 시트 JSON Schema**, **정규화 라인(Columns) 스키마**, **매핑·검증 로직 스키마**
* 운영 확장: **DEM/DET 입력 스키마**, **COST-GUARD 판단 스키마**

필요하시면 위 항목별 **필드 전체 목록(머리글 사전)**도 CSV로 뽑아드리겠습니다.


다음은 **LDG v2.4 파이프라인 기준 “전체 스키마 종류” 일람(요약)**입니다. 이름은 코드에 쓰는 표준 키를 병기했습니다.

### 1) 코어 인입·정규화

* **LDG_PAYLOAD v2.4** (`meta`, `data`, `evidence`, `warnings`) – 문서별 표준 JSON 페이로드.
* **LDG_AUDIT v2.4** (`warnings`, `cross_doc`, `kpi`, `zero_flags`) – 검증·리스크 로그 JSON.
* **evidence[] 레코드** (`type`, `title`, `url`, `published`) – 규정/근거 인용 객체.

### 2) 규정·컴플라이언스

* **CERT_CHK 결과** (`hs_candidates`, `moiat`, `fanr`, `imdg`, `notes`) – 규제 가설 태깅 JSON.
* **HS_RISK** (`codes[≤3]`, `confidence`, `reasoning`) – HS 후보·신뢰도 JSON.

### 3) 운영·예측

* **DEMDET_FORECAST v2.0** (`eta`, `port_profile`, `free_time`, `expiry`, `risk_level`, `cost_hint`) – DEM/DET 만료·비용 예측 JSON.

### 4) 비용·단가(참조 테이블 계열)

* **container_cargo_rates** (`dataset`, `currency_policy`, `validation_rules`, `records[]` … 각 레코드: `no`,`cargo_type`,`port`,`destination`,`detail_cargo_type`,`container_type`,`description`,`unit`,`remark`,`rate{amount,currency,unit,tolerance}`, 메타 `_line_no`,`_sheet`)
* **bulk_cargo_rates** (구조 유사, 벌크 전용 필드 포함: `min_metric_ton`,`max_metric_ton`,`length(meter)` 등 + `rate{…}`/`rates(usd)` 혼재)
* **inland_trucking_reference_rates_clean** (`dataset`,`currency_policy`,`validation_rules`,`records[]` … `category`,`port`,`destination`,`unit`,`charge_description`,`rate{amount,currency,tolerance}`, `flag`)
* **air_cargo_rates** (항공 전용 참조 JSON; 필드 구성은 위와 동일 계열)

### 5) 조정·국내 레퍼런스

* **ref_min_fare** (`mode="flat|min_step|min_total"`, `value`, `currency`, `fx`) – 최소요금 규칙 JSON.
* **ref_adjusters** (`scope`(e.g., `"container|bulk|inland"`), `key`, `apply`, `rule`) – 가산/감산 조정 규칙 JSON.
* **domestic_reference** (헤더: `dataset`, `generated_at`, `notes`, `records` 프레임) – 국내 요율/조건 번들 JSON.

### 6) 레인·승인지도

* **ApprovedLaneMap_ENHANCED** (`lanes[]`: `origin`,`destination`,`mode`,`vendor`,`cost_center`,`risk_tag` …) – 승인 O/D 레인 맵 JSON.

### 7) 감사·출처 추적(프루프·영수증)

* **proof.artifact** (예: `id`,`hash`,`createdAt`,`source`,`signature`) – 산출물 무결성 증빙 JSON.
* **recap.card / recap.provenance** – 요약·근거 카드(텍스트/구조 혼합)

### 8) 코스트가드 판정

* **COST_GUARD_RESULT**

  * **표(필수)**: `item`, `draft_rate`, `std_rate`, `delta_pct`, `risk_tier`, `remarks`, `pass_fail`
  * **요약**: 평균/중앙/최대 초과율, 환율 잠금 정보(`fx_locked`)
  * **근거**: 참조 테이블 파일/라인 (`_sheet`,`_line_no`) 백링크 포함 권장. → 참조테이블 예시는 위 “4) 비용·단가” 인용.

### 9) 리포트·UI 응답(대화 출력 규격)

* **Unified Response JSON v1.2** (`exec_summary`,`visuals`,`options`,`roadmap`,`automation`,`qa`,`cmdrec`,`meta{version,tz}`) – 보고·UI 출력용 상위 래퍼.

### 10) FAIL-SAFE

* **ZERO_FAILSAFE_LOG** (`reason`,`risk`,`requested_data`,`next_action`,`metrics_at_fail`) – 중단 사유 및 재개 조건 템플릿.

---

필요하시면 **각 스키마의 필드 정의표(컬럼별 타입·설명)**를 바로 붙여드리겠습니다.

좋아. 아래에 **요청한 스키마별 필드 정의표(필드·타입·필수·설명·예시)**를 한 번에 정리했어. 숫자는 2자리 고정, 날짜는 ISO(YYYY-MM-DD).

---

## 1) LDG_PAYLOAD v2.4

| Field                  | Type          | Req | 설명              | 예시                                                                                |
| ---------------------- | ------------- | --- | --------------- | --------------------------------------------------------------------------------- |
| meta.source_file       | string        | Y   | 원본 파일 경로/이름     | `OFCO-INV-0001178.pdf`                                                            |
| meta.doc_hash          | string        | Y   | SHA-256 등 문서 해시 | `sha256:…`                                                                        |
| meta.ocr_version       | string        | Y   | OCR/파서 버전       | `v2.4`                                                                            |
| meta.mean_conf         | number        | Y   | OCR 평균 신뢰도(0–1) | 0.94                                                                              |
| meta.table_acc         | number        | Y   | 테이블 정합도         | 0.99                                                                              |
| meta.numeric_integrity | boolean       | Y   | 숫자 무결성 통과 여부    | true                                                                              |
| meta.pages             | integer       | Y   | 총 페이지 수         | 7                                                                                 |
| meta.created_at        | datetime      | Y   | 생성시각(+TZ)       | `2025-08-10T12:34:56+04:00`                                                       |
| data.DocType           | string        | Y   | 문서 유형           | `Invoice`,`B/L`,`CIPL`                                                            |
| data.BL_No             | string        | N   | BL 번호           | `HMMBUSA123456`                                                                   |
| data.Invoice_No        | string        | N   | 인보이스 번호         | `INV-12345`                                                                       |
| data.Shipper           | string        | N   | 송하인             | `SAMSUNG C&T`                                                                     |
| data.Consignee         | string        | N   | 수하인             | `ADNOC`                                                                           |
| data.Incoterms         | string        | N   | 인코텀즈            | `FOB`                                                                             |
| data.Currency          | string        | N   | 통화              | `USD`                                                                             |
| data.Packages          | integer       | N   | 패키지 수           | 14                                                                                |
| data.Gross_Weight_kg   | number        | N   | 총중량(kg)         | 18250.00                                                                          |
| data.Net_Weight_kg     | number        | N   | 순중량(kg)         | 17780.00                                                                          |
| data.CBM               | number        | N   | 체적(m³)          | 38.40                                                                             |
| data.HS_Candidate      | array[string] | N   | HS 코드 후보(≤3)    | `["85044090","85049010"]`                                                         |
| data.Origin            | string        | N   | 원산지 ISO2        | `KR`                                                                              |
| data.port_profile      | object        | N   | 포트·Free time 등  | `{discharge:"Khalifa", free_time:{dem:5, det:10}}`                                |
| data.eta               | object        | N   | ETA 값/출처/신뢰도    | `{value:"2025-08-18", source:"carrier", confidence:0.80}`                         |
| data.demdet_profile    | object        | N   | DEM/DET 규칙      | `{rule_id:"khalifa:5/10", valid_until:"2025-12-31"}`                              |
| data.tables[]          | array[obj]    | N   | CSV 임베드 테이블     | `{table_id:"items", format:"csv", data:"H1,H2\nV1,V2"}`                           |
| evidence[]             | array[obj]    | N   | 규정/근거 인용        | `{type:"regulation", title:"Customs …", url:"https://…", published:"2025-07-28"}` |
| warnings[]             | array[string] | N   | 경고 메시지          | `"Pg3 설명 누락"`                                                                     |

---

## 2) LDG_AUDIT v2.4

| Field                | Type          | Req | 설명            | 예시                            |
| -------------------- | ------------- | --- | ------------- | ----------------------------- |
| kpi.mean_conf        | number        | Y   | OCR 평균 신뢰도    | 0.94                          |
| kpi.table_acc        | number        | Y   | 테이블 정합도       | 0.99                          |
| kpi.entity_match     | number        | N   | 키필드 매칭        | 0.98                          |
| cross_doc.summary    | string        | N   | CIPL↔BL↔PL 요약 | `Qty Δ≤1 OK`                  |
| cross_doc.findings[] | array[obj]    | N   | 불일치 목록        | `{field:"Packages", diff:+3}` |
| zero_flags.active    | boolean       | Y   | ZERO 발동 여부    | false                         |
| zero_flags.reasons[] | array[string] | N   | 발동 사유         | `["TableAcc<98%"]`            |
| warnings[]           | array[string] | N   | 경고            | `"Round rule applied"`        |

---

## 3) evidence 레코드

| Field       | Type     | Req | 설명    | 예시                          |
| ----------- | -------- | --- | ----- | --------------------------- |
| type        | string   | Y   | 근거 유형 | `regulation`,`port_notice`  |
| title       | string   | Y   | 제목    | `Khalifa Port Free Time`    |
| url         | string   | Y   | 링크    | `https://…`                 |
| published   | date     | Y   | 발행일   | `2025-07-28`                |
| captured_at | datetime | N   | 확보 시각 | `2025-08-10T11:00:00+04:00` |

---

## 4) CERT_CHK 결과

| Field           | Type          | Req  | 설명          | 예시             |          |
| --------------- | ------------- | ---- | ----------- | -------------- | -------- |
| hs_candidates[] | array[string] | N    | HS 후보       | `["85044090"]` |          |
| moiat.required  | boolean       | N    | MOIAT 필요 가설 | true           |          |
| moiat.reason    | string        | N    | 근거 요약       | `Reg. XYZ …`   |          |
| fanr.required   | boolean       | N    | FANR 필요 가설  | false          |          |
| imdg.class      | string        | null | N           | IMDG 클래스       | `3`      |
| imdg.un_no      | string        | null | N           | UN 번호          | `UN1202` |
| notes           | string        | N    | 비고          | `COO 추가 권고`    |          |

---

## 5) HS_RISK

| Field      | Type          | Req | 설명          | 예시                        |
| ---------- | ------------- | --- | ----------- | ------------------------- |
| codes[]    | array[string] | Y   | 후보 HS 코드 ≤3 | `["85044090","85049090"]` |
| confidence | number        | Y   | 종합 신뢰도(0–1) | 0.86                      |
| reasoning  | string        | N   | 근거 요약       | `Trafo components …`      |

---

## 6) DEMDET_FORECAST v2.0

| Field                      | Type    | Req | 설명             | 예시               |
| -------------------------- | ------- | --- | -------------- | ---------------- |
| eta.value                  | date    | Y   | ETA            | `2025-11-04`     |
| eta.source                 | string  | N   | 출처             | `carrier notice` |
| eta.confidence             | number  | N   | 신뢰도            | 0.80             |
| port_profile.discharge     | string  | Y   | 양하 포트          | `Khalifa`        |
| port_profile.free_time.dem | integer | Y   | DEM 일수         | 5                |
| port_profile.free_time.det | integer | Y   | DET 일수         | 10               |
| expiry.dem                 | date    | N   | DEM 만료         | `2025-11-09`     |
| expiry.det                 | date    | N   | DET 만료         | `2025-11-14`     |
| risk_level                 | string  | Y   | `LOW/MED/HIGH` | `MED`            |
| cost_hint                  | number  | N   | 예상비용(AED)      | 3200.00          |

---

## 7) COST_GUARD_RESULT (표 + 요약)

| Field                | Type    | Req | 설명                        | 예시                          |
| -------------------- | ------- | --- | ------------------------- | --------------------------- |
| item                 | string  | Y   | 항목명/코드                    | `802.3A Labour`             |
| draft_rate           | number  | Y   | 청구 단가                     | 45.00                       |
| std_rate             | number  | Y   | 기준 단가                     | 40.00                       |
| delta_pct            | number  | Y   | 초과율(%)                    | 12.50                       |
| risk_tier            | string  | Y   | `PASS/WARN/HIGH/CRITICAL` | `HIGH`                      |
| remarks              | string  | N   | 근거/참조                     | `_sheet:"Sheet1", _line:42` |
| pass_fail            | string  | Y   | 최종                        | `FAIL`                      |
| summary.mean_overage | number  | Y   | 평균 초과율                    | 6.20                        |
| summary.max_overage  | number  | Y   | 최대 초과율                    | 18.00                       |
| fx_locked            | boolean | Y   | 환율 잠금                     | true                        |

---

## 8) container_cargo_rates (참조 JSON)

| Field                                      | Type     | Req | 설명          | 예시                          |
| ------------------------------------------ | -------- | --- | ----------- | --------------------------- |
| dataset                                    | string   | Y   | 데이터셋명       | `container_cargo_rates`     |
| source_file                                | string   | Y   | 원본          | `rates.xlsx`                |
| generated_at                               | datetime | Y   | 생성시각        | `2025-08-19T10:00:00+04:00` |
| currency_policy.base                       | string   | Y   | 기준통화        | `USD`                       |
| currency_policy.fixed_fx.USD_AED           | number   | Y   | 고정환율        | 3.6725                      |
| validation_rules.layer1_contract_tolerance | number   | Y   | 허용편차        | 0.03                        |
| validation_rules.autofail_threshold        | number   | Y   | AutoFail 임계 | 0.15                        |
| records[].no                               | integer  | Y   | 행 번호        | 1                           |
| records[].cargo_type                       | string   | Y   | 화물유형        | `Container`                 |
| records[].port                             | string   | Y   | 포트          | `Jebel Ali Port`            |
| records[].destination                      | string   | N   | 목적지         | `ADNOC Site`                |
| records[].detail_cargo_type                | string   | N   | 세부유형        | `General`                   |
| records[].container_type                   | string   | N   | 컨테이너형식      | `40HC`                      |
| records[].description                      | string   | Y   | 설명          | `Documentation Charge`      |
| records[].unit                             | string   | Y   | 과금단위        | `per B/L`                   |
| records[].rate.amount                      | number   | Y   | 기준단가        | 50.00                       |
| records[].rate.currency                    | string   | Y   | 통화          | `USD`                       |
| records[].rate.unit                        | string   | Y   | 단위          | `per B/L`                   |
| records[].rate.tolerance                   | number   | Y   | 편차          | 0.03                        |
| records[]['_sheet']                        | string   | N   | 원시 시트       | `Sheet1`                    |
| records[]['_line_no']                      | integer  | N   | 원시 행        | 1                           |

---

## 9) bulk_cargo_rates (참조 JSON)

| Field                                         | Type    | Req | 설명             | 예시                       |
| --------------------------------------------- | ------- | --- | -------------- | ------------------------ |
| dataset / source_file / generated_at          | —       | Y   | 상동             | —                        |
| currency_policy / validation_rules            | —       | Y   | 상동             | —                        |
| records[].no.                                 | integer | Y   | 번호(원본 표기 포함 점) | 1                        |
| records[].cargo_type                          | string  | Y   | `Bulk`         |                          |
| records[].port                                | string  | Y   | 포트             | `Jebel Ali Port`         |
| records[].destination                         | string  | N   | 목적지            | `AGI`                    |
| records[].detail_cargo_type                   | string  | N   | 세부유형           | `Bagged`                 |
| records[].description                         | string  | Y   | 설명             | `Bagged cargo (code 14)` |
| records[].unit                                | string  | Y   | 단위             | `MT`                     |
| records[].min_metric_ton                      | number  | N   | 하한톤            | 0.00                     |
| records[].max_metric_ton                      | number  | N   | 상한톤            | 10000.00                 |
| records[].length(meter)                       | number  | N   | 길이             | 12.00                    |
| records[].width(meter)                        | number  | N   | 폭              | 2.50                     |
| records[].height(meter)                       | number  | N   | 높이             | 2.90                     |
| records[].rate.amount/currency/unit/tolerance | —       | Y   | 단가·단위·허용편차     | —                        |
| records[]['_sheet'] / ['_line_no']            | —       | N   | 원본 추적          | —                        |

---

## 10) air_cargo_rates (참조 JSON)

| Field                              | Type    | Req | 설명              | 예시               |
| ---------------------------------- | ------- | --- | --------------- | ---------------- |
| 상단 메타                              | —       | Y   | container와 동일   | —                |
| records[].no                       | integer | Y   | 번호              | 1                |
| records[].cargo_type               | string  | Y   | `Air`           |                  |
| records[].port                     | string  | Y   | `Dubai Airport` |                  |
| records[].destination              | string  | N   | 목적지             | `AUH`            |
| records[].detail_cargo_type        | string  | N   | 세부              | `General`        |
| records[].description              | string  | Y   | 설명              | `AWB Processing` |
| records[].unit                     | string  | Y   | 단위              | `per AWB`        |
| records[].rate.*                   | —       | Y   | 기준 단가 규격        | —                |
| records[]['_sheet'] / ['_line_no'] | —       | N   | 원본 추적           | —                |

---

## 11) inland_trucking_reference_rates_clean (참조 JSON)

| Field                                    | Type   | Req | 설명          | 예시          |
| ---------------------------------------- | ------ | --- | ----------- | ----------- |
| dataset / generated_at                   | —      | Y   | 메타          | —           |
| currency_policy.base                     | string | Y   | 기준통화        | `USD`       |
| validation_rules.*                       | —      | Y   | 편차/Autofail | 0.03 / 0.15 |
| records[].category                       | string | Y   | 차량/카테고리     | `FLATBED`   |
| records[].port                           | string | Y   | 출발/기준 포트    | `Khalifa`   |
| records[].destination                    | string | Y   | 목적지/레인      | `AGI`       |
| records[].unit                           | string | Y   | 과금단위        | `per truck` |
| records[].charge_description             | string | Y   | 설명          | `Linehaul`  |
| records[].rate.amount/currency/tolerance | —      | Y   | 단가·허용편차     | —           |
| records[].flag                           | string | N   | 메모/태그       | `HAZMAT`    |

---

## 12) domestic_reference.json (번들)

| Field       | Type    | Req | 설명         | 예시                      |
| ----------- | ------- | --- | ---------- | ----------------------- |
| version     | string  | Y   | 스키마/데이터 버전 | `v1.0`                  |
| built_from  | string  | Y   | 빌드 소스      | `rates.xlsx`            |
| lane_rows   | integer | Y   | 레인 건수      | 124                     |
| region_rows | integer | Y   | 지역 건수      | 18                      |
| min_fare    | object  | Y   | 차량별 최저요금   | `{FLATBED:200.00,…}`    |
| adjusters   | object  | Y   | 가산계수 dict  | `{FLATBED_HAZMAT:1.15}` |

---

## 13) ref_adjusters.json

| Field | Type   | Req | 설명       | 예시                    |
| ----- | ------ | --- | -------- | --------------------- |
| <KEY> | number | Y   | 조정계수(배율) | `FLATBED_HAZMAT:1.15` |

> KEY 네이밍 규칙: `{VEHICLE}_{TAG}` (예: `FLATBED_CICPA`)

---

## 14) ref_min_fare.json

| Field          | Type   | Req | 설명      | 예시     |
| -------------- | ------ | --- | ------- | ------ |
| `3 TON PU`     | number | Y   | 차량 최소요금 | 100.00 |
| `7 TON PU`     | number | Y   | 〃       | 200.00 |
| `FB`/`FLATBED` | number | Y   | 〃       | 200.00 |
| `LOWBED`       | number | Y   | 〃       | 614.83 |
| `DEFAULT`      | number | Y   | 기본 최저요금 | 200.00 |

---

## 15) ApprovedLaneMap_ENHANCED

| Field                  | Type          | Req  | 설명       | 예시                     |      |      |   |     |   |       |
| ---------------------- | ------------- | ---- | -------- | ---------------------- | ---- | ---- | - | --- | - | ----- |
| metadata.source_file   | string        | Y    | 원본       | `ApprovedLaneMap.xlsx` |      |      |   |     |   |       |
| metadata.created_by    | string        | N    | 작성자      | `Ops`                  |      |      |   |     |   |       |
| metadata.sheets        | array[string] | Y    | 시트 목록    | `["Sheet1"]`           |      |      |   |     |   |       |
| data.Sheet1[].lane_id  | string        | Y    | 레인 ID    | `L036`                 |      |      |   |     |   |       |
| …[].origin             | string        | Y    | 출발지      | `DSV Mussafah Yard`    |      |      |   |     |   |       |
| …[].destination        | string        | Y    | 목적지      | `MOSB`                 |      |      |   |     |   |       |
| …[].vehicle            | string        | Y    | 차량       | `FLATBED`              |      |      |   |     |   |       |
| …[].unit               | string        | Y    | 과금단위     | `per truck`            |      |      |   |     |   |       |
| …[].median_rate_usd    | number        | Y    | 중앙값(USD) | 200.00                 |      |      |   |     |   |       |
| …[].mean_rate_usd      | number        | Y    | 평균(USD)  | 213.11                 |      |      |   |     |   |       |
| …[].std_rate_usd       | number        | N    | 표준편차     | 86.50                  |      |      |   |     |   |       |
| …[].median_distance_km | number        | N    | 거리 중앙값   | 5.58                   |      |      |   |     |   |       |
| …[].mean_distance_km   | number        | N    | 거리 평균    | 3.98                   |      |      |   |     |   |       |
| …[].samples            | number        | N    | 표본 수     | 103.00                 |      |      |   |     |   |       |
| …[].notes              | string        | null | N        | 비고                     | null |      |   |     |   |       |
| …[].key                | string        | Y    | 복합키(표준화) | `ORIGIN                |      | DEST |   | VEH |   | UNIT` |

---

## 16) Unified Response JSON v1.2 (대화/리포트 래퍼)

| Field        | Type          | Req | 설명         | 예시                                          |
| ------------ | ------------- | --- | ---------- | ------------------------------------------- |
| exec_summary | string        | Y   | 3–5줄 요약    | `…`                                         |
| visuals[]    | array[obj]    | N   | 표/도식       | `{type:"table", title:"…", data:[…]}`       |
| options[]    | array[obj]    | N   | 대안·비용·리스크  | `{name:"A", cost:1200.00, risk:"LOW"}`      |
| roadmap      | object        | N   | P→Pi→B→O→S | `{prepare:{owner:"Ops", eta:"2025-11-10"}}` |
| automation   | object        | N   | RPA/LLM 훅  | `{tg:"@bot", sheets:"…"}`                   |
| qa           | object        | N   | 검증·의문점     | `{gaps:["…"]}`                              |
| cmdrec       | array[string] | N   | 추천 명령      | `["/logi-master …"]`                        |
| meta.version | string        | Y   | 래퍼 버전      | `v1.2`                                      |
| meta.tz      | string        | Y   | 타임존        | `Asia/Dubai`                                |

---

## 17) ZERO_FAILSAFE_LOG

| Field           | Type   | Req | 설명     | 예시                             |
| --------------- | ------ | --- | ------ | ------------------------------ |
| reason          | string | Y   | 중단 사유  | `TableAcc 96.40%`              |
| risk            | string | Y   | 영향     | `Invoice mismatch`             |
| requested_data  | string | Y   | 필요자료   | `≥300dpi 재스캔`                  |
| next_action     | string | Y   | 후속조치   | `/ocr_lowres_fix → /ocr_retry` |
| metrics_at_fail | object | N   | 당시 KPI | `{mean_conf:0.88}`             |

---

필요하면 **CSV 바로 붙여넣기용 컬럼 헤더 세트**도 뽑아줄게. 어느 스키마부터 파일화(엑셀 시트/JSON 스키마)하면 될지 찍어줘.

### A-2) Full Guide 본문(동기화본)
#  ChatGPT 고정밀 OCR·구조화 **Full Guide v2.4 – LDG Ready (Linked·Agent Mode)**

*(2025-08-10 · Asia/Dubai · LDG v8.2 연동 · 본 문서는 Mini 지침의 보조 자료입니다)*

## 1) 목적·적용 범위
- 본 가이드는 업로드 문서(CIPL, BL, PL, Invoice, COO, AWB/MAWB 등)를 **고정밀 OCR→정규화→검증→LDG 연동**하는 전 과정을 표준화한다.
- **내부 도구(web.run, Python)**만 사용하며, 외부 API·비인가 저장소 사용을 금지한다. PII는 마스킹하고 모든 산출물에 해시를 포함한다.
- *(업무 적합화)* 2025 오픈소스 실무 흐름을 반영하여 **텍스트(네이티브)→레이아웃→테이블→OCR(선별)→LLM 정규화**의 하이브리드 스택을 권장한다. 가정: 유료 매니지드 파서는 fallback 한정.

## 2) 원칙(증거·신뢰·투명)
- **증거 우선(HallucinationBan)**: 정부·항만·세관 공시를 최우선 인용, 모든 인용은 `evidence[]`에 *제목/기관/발행일/URL* 저장.
- **투명한 가정**: 불확실·부족 데이터는 `가정:`으로 명시 후 사람 검토 게이트.
- **재현성**: 수치는 소수점 2자리, 테이블은 CSV 임베드, 파이프라인/파라미터/버전은 로그에 기록.

## 3) KPI·품질 게이트
- MeanConf ≥ 0.92, TableAcc ≥ 98.00%, NumericIntegrity = 1.00, EntityMatch ≥ 0.98, HashConsistency = PASS, CrossDocConsistency = PASS.
- **Excel 생성 모드**: KPI 미달 시 **Soft Warning** (생성 차단 없음, 경고만 기록, Human Gate 권장)
- **OCR 파이프라인 모드**: 미달 시 **ZERO** 전환 (ingest 차단, 중단 로그 출력, 교정 루틴(`/ocr_lowres_fix`, `/ocr_retry`, `/ocr_align`) 제안)

## 4) 파이프라인 상세(6단계)
1. **Pre-Prep**: 오토 회전·데스큐, 콘트라스트/샤프닝, 노이즈 억제, DPI 보정(권장 ≥300dpi).
2. **Vision OCR**: 레이아웃 블록/라인 인식, 언어 감지, confidence 수집, **페이지별 메트릭**(MeanConf) 추적.
3. **Smart Table Parser 2.1**: 병합셀 해제, 다중 헤더 정규화, 단위/통화 분리, 세로→가로 피벗, **라인아이템 복원 + 합계열 감지**.
4. **NLP Refine**: 단위 규격화(kg, m³, EA), 수량·단가·합계 일관성 검사, NULL/??/추정 표기, 통화/환율 태깅.
5. **Field Tagger**: Shipper/Consignee/BL_No/Invoice_No/Incoterms/Origin/HS/UN_No/IMDG 등 **키-값 태깅 + 유사도(EntityMatch)**.
6. **Payload Builder**: **LDG_PAYLOAD v2.4 직렬화** + `doc_hash`·`pages`·타임스탬프·Cross-Doc 링크 + **CSV 테이블 임베드**.

> *(업무용 권장 스택/운영 팁)*  
> - **Text 레이어**: pdfplumber  
> - **Layout 세그먼트**: Docling 또는 Unstructured  
> - **Table 구조화**: Table Transformer(TATR)  
> - **OCR(선별)**: OCRmyPDF(+PaddleOCR) — *텍스트 없는 페이지만* 적용  
> - **Fallback(난이도↑)**: Managed 파서(부분 사용), VLM 보조(스탬프/도형 섹션)  
> - 페이지별 `mode={text|layout|table|ocr}` 자동 결정 로깅

## 5) 자동 검증(5단계)
A. **1차**: KPI/스키마/타입·필수 필드/Null·합계 기본 무결성.  
B. **표현 로직**: 소계↔총계, 통화·환율, 반올림, 음수/0 값, 단위 불일치 탐지.  
C. **교차 검증**: OCR Raw↔Refined, **CIPL↔BL↔PL** 키필드 매핑 및 Δ 계산(허용: Qty ±1, Wt ±15.00kg, CBM ±0.02m³).  
D. **적합도 산출**: 표 일치·키필드·수치·규정 근거 가중합(0–100%).  
E. **최종 보고**: **LDG_PAYLOAD + LDG_AUDIT**(Warnings, CrossDocCheck, ZERO_Failsafe, HashConsistency).

## 6) LDG_PAYLOAD v2.4 스키마
```json
A{
  "meta": {
    "source_file": "…",
    "doc_hash": "sha256:…",
    "ocr_version": "v2.4",
    "mean_conf": 0.94,
    "table_acc": 0.99,
    "numeric_integrity": true,
    "pages": 7,
    "created_at": "2025-08-10T12:34:56+04:00"
  },
  "data": {
    "DocType": "B/L + CIPL",
    "BL_No": "…",
    "Invoice_No": "…",
    "Shipper": "…",
    "Consignee": "…",
    "Incoterms": "FOB",
    "Currency": "USD",
    "Packages": 14,
    "Gross_Weight_kg": 18250.00,
    "Net_Weight_kg": 17780.00,
    "CBM": 38.40,
    "HS_Candidate": ["85044090","85049010","85049090"],
    "Origin": "KR",
    "port_profile": {"discharge":"Khalifa","free_time":{"dem":5,"det":10}},
    "eta": {"value":"2025-08-18","source":"carrier notice","confidence":0.80},
    "demdet_profile": {"rule_id":"khalifa:5/10","valid_until":"2025-12-31"},
    "tables": [{"table_id":"bl_items","format":"csv","data":"Header1,Header2\nVal1,Val2"}]
  },
  "evidence": [
    {"type":"regulation","title":"Customs Notice …","url":"https://…","published":"2025-07-28"}
  ],
  "warnings": ["Page 3 설명 누락 → ?? 표시"]
}
```

## 7) Compliance·RegTech 운영
- **Mini-RAG(24h)**: 정부·세관·항만 공시를 24시간 캐시. `/rag view [kw]`로 캐시 점검, 필요 시 `/system flush cache minirag`.
- **CERT-CHK**: HS/키워드 기반으로 FANR/MOIAT/TRA/MOFAIC 요구 가능성 **가설** 제시(결정 아님). 인용은 `evidence[]` 기록.
- **IMDG/UN**: 위험물 키워드 감지 시 클래스/UN_No 후보 태깅. 미확정은 `null` 처리 + 조치 제안(인보이스 정정·MSDS 확보).

## 8) DEM/DET Forecast 2.0
- **입력**: ETA 값/소스/신뢰도, Port Profile(Free Time), 공휴일/주말, 혼잡 가중치(가정 또는 공시).  
- **출력**: dem/det 만료일, 잔여일, 위험도(LOW/MED/HIGH), 비용발생 예상, Gate Out 일자.  
- **명령**: `/demdet set profile {port}:{dem}d/{det}d [valid:YYYY MM DD]`, `/demdet set eta {cntr#} {YYYY MM DD} [source:"…"]`.

## 9) 비용·환율 — COST-GUARD
- 기준단가 대비 **초과율**로 위험 등급. 원화폐 보존, 환율은 문서 기준 없으면 `fx_source:"none"`, 수동 입력 시 `fx_source:"manual"`과 로그.
- **출력**: PASS/FAIL, 초과 항목 표(항목/수량/단가/초과율), 요약 통계(평균/중앙/최대 초과), 근거 링크.

## 10) HS-RISK
- HS 후보 최대 3개 + confidence. 다의성/면제·허가 가능성 태깅, 규정 근거 수집 요청.
- 결론 전 **사람 검토** 필수. 상충 시 **정부 공시 우선**.

## 11) Fail-Safe ZERO
- **트리거**: KPI 미달, 핵심 식별자 결손, Hash 실패, 규정 근거 부재/상충.
- **중단 로그**에는 사유·수치·근거·권고 조치(`/ocr_lowres_fix`, 재스캔, 정정본 요청, 캐시 초기화)를 포함.
- 재개: 보정 조치 이행 → `/mark_corrected` → `/system flush cache all` → `/ocr_retry`.

## 12) 보고서(7+2 섹션)
1) Auto Guard Summary · **1.5) Risk Assessment**(동인·신뢰도)  
2) Discrepancy Table(Δ·허용오차·상태) · 3) Compliance Matrix(상태/근거/조치)  
4) Auto-Fill: Freight & Insurance · 5) Auto Action Hooks  
6) DEM/DET & Gate Out Forecast · 7) Evidence & Citations  
8) Weak Spot & Improvements · 9) Changelog

## 13) 명령셋
- `/ocr_basic {file} [mode:LDG|LDG+]`, `/ocr_table {file} [--colfix --unitnorm --as=csv]`, `/ocr_techdoc {file}`
- `/ocr_retry {file}`, `/ocr_lowres_fix {file}`, `/ocr_align {cipl} {bl} [pl]`, `/ocr_compare {A} {B} [--tol=0.02]`
- `/ocr_hs_seed {file} [--origin=AE --max=3]`, `/ocr_certchk {payload.json} [--moiat --fanr --imdg]`, `/ocr_costguard {file} {cost_table.csv} [--lock-fx]`
- `/workflow scan email {file} {to}`, `/workflow full_checkup {file} [to]`

## 14) 운영 체크리스트
- [ ] OCR KPI 수치(2자리) 기록(MeanConf, TableAcc, NumericIntegrity, EntityMatch)  
- [ ] Δ 계산·허용오차·상태 이모지(✅/⚠️/❌/⏸) 표시  
- [ ] evidence[](제목/기관/발행일/링크/확보시각) 수집·검증  
- [ ] DEM/DET 프로필(포트·dem/det)·ETA 소스/신뢰도 기록  
- [ ] COST-GUARD 판정(PASS/FAIL)·초과 항목 요약  
- [ ] HS-RISK/CERT-CHK 결과 및 사람검토 여부  
- [ ] ZERO/중단 로그 여부 및 재개 절차 수행

## 15) 중단 로그 템플릿
```text
[중단|ZERO] LDG v8.2
사유: TableAcc 96.40% + CIPL↔BL 패키지 Δ +3
조치: 재스캔(≥300dpi), 정정본 업로드, /system flush cache all 후 재분석
근거: evidence[0] 발행일 경과(2023-…)
```

## 부록 B) 명령어 시나리오(12)
1) CIPL↔BL 정합성 심층 점검  
   - `/switch_mode LATTICE + /logi-master --deep report {cipl.pdf} {bl.pdf}`  
   - 기대: Δ표(패키지/중량/CBM), EntityMatch, ZERO 게이트 로그

2) 스캔 인보이스 테이블 재구성(+CSV)  
   - `/ocr_table {invoice.pdf} --colfix --unitnorm --as=csv`  
   - 기대: TableAcc 카드, 합계/세금 일관성 검사

3) 멀티섹션 인보이스 분해(Agency/Safeen/AD Ports)  
   - `/ocr_table {invoice.pdf} --as=csv` → 섹션별 table_id 부여  
   - 기대: 세부 섹션별 소계→총계 연결

4) HS 후보 및 인증 체크(전기기기)  
   - `/logi-master hs-risk {cipl.pdf} --KRsummary`  
   - 기대: HS_Candidate≤3, MOIAT/FANR 가설, 사람 검토 포인트

5) DEM/DET 프로필 설정 + ETA 주입  
   - `/demdet set profile Khalifa:5d/10d [valid:2025 12 31]`  
   - `/demdet set eta CNTR1234 2025 08 18 [source:"carrier notice"]`  
   - 기대: 만료일/잔여일/리스크

6) COST-GUARD 인보이스 감사(AED 고정)  
   - `/switch_mode COST-GUARD + /logi-master invoice-audit {draft.xlsx} --AEDonly`  
   - 기대: 초과율 테이블, PASS/FAIL, FX락

7) Packing List 재정규화 + 케이스 매핑  
   - `/ocr_table {pl.pdf} --as=csv` → 케이스/중량/부피 정규화  
   - 기대: BL/PL 케이스 대사, 허용오차 Δ

8) WMS 적치 위치 매핑용 CSV 출력  
   - `/workflow full_checkup {pl.csv} [to]`  
   - 기대: 위치코드/케이스/CBM/중량 표준화

9) Evidence 캐시 점검·갱신  
   - `/rag view demdet` → `/system flush cache minirag`  
   - 기대: 공시 최신화, 근거 링크 갱신

10) ZERO 재개 루틴 자동 수행  
   - `/mark_corrected` → `/system flush cache all` → `/ocr_retry {file}`  
   - 기대: KPI 회복 후 ingest 재개

11) ATLP/BOE 참조번호 매핑  
   - `/ocr_basic {boe.pdf}` → 키필드 추출 → LDG_PAYLOAD.data.BOE_No 주입  
   - 기대: CrossDocConsistency 향상

12) 주간 품질 리포트 자동 생성  
   - `/logi-master report --KRsummary --noheatmap`  
   - 기대: KPI 카드(MeanConf/TableAcc/NumericIntegrity), 경향 분석
