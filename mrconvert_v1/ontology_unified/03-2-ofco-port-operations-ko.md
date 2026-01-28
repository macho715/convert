---
title: "OFCO System - Ontology Perspective (Korean)"
type: "ontology-design"
domain: "port-operations"
version: "4.1"
date: "2025-01-19"
language: "ko"
tags: ["ontology", "ofco", "port-ops", "invoice", "hvdc", "unified", "korean"]
status: "active"
---

# OFCO 시스템 — 온톨로지 관점 (HVDC Logistics Unified v4.1)

## Executive Summary

> 핵심 한 줄: **OFCO는 '항만 대행(Agency)·항만요금·장비/인력·수배(수자원/게이트패스) 서비스**를 묶는 **온톨로지 기반 Port Ops & Invoice 허브**이며, 문서(Invoice)↔운영(PortCall/ServiceEvent)↔요율(Tariff/PriceCenter)을 **Multi-Key Identity Graph**로 한 번에 해석합니다. (EN-KR: Any-key in → Resolve → PortCall & Services → Rate & Cost mapping.)

---

## 0) Executive Summary (3–5)
- **Multi‑Key Identity Graph**: 입력 키는 *OFCO/SAFEEN/ADP 인보이스번호, Rotation#, Samsung Ref, hvdc_code* 등 아무 키든 OK → **동일 실체(PortCall·Shipment·Invoice) 클러스터**로 귀결.
- **Ontology‑First**: *Invoice, InvoiceLine, ServiceEvent(채널크로싱/접안/예인/조종/PHC/장비/인력/수배/게이트패스/문서수수료), PortCall, Rotation, TariffRef, PriceCenter, CostCenter(A/B/C)* 클래스로 정규화.
- **검증 표준**: **LDG v8.2 ↔ OCR v2.4** 연동, KPI(MeanConf≥0.92, TableAcc≥0.98, NumericIntegrity=1.00), **ZERO failsafe**.
- **매핑 규칙**: Cost Center v2.5 **17‑Regex** + Subject 패턴(“SAFEEN”→Channel Transit, “ADP INV + Port Dues”→Port Dues, “Cargo Clearance/Arranging FW/BA/5000 IG FW” 등) → **Price Center 3‑Way(A/B/C)** 분개.
- **회계 일관성**: EA×Rate 합=Amount(±0.01), Σ라인=Invoice Total(±2.00%), 통화/VAT 100.00% 일치, **[EXT] 메타 금액 집계 제외**.

---

## 1) Ontology Core (RDF/OWL)
### 1.1 주요 클래스
- **org:Organization** ⟶ *ofco:OFCO, adports:ADPorts, safeen:SAFEEN, sct:SCT*
- **vsl:Vessel / vsl:Voyage / port:PortCall** *(RotationNo 포함)*
- **fin:Invoice** *(source=OFCO)*, **fin:InvoiceLine** *(최대 4 RatePair)*
- **ops:ServiceEvent** ⟶ *ops:ChannelCrossing, ops:Berthing, ops:Pilotage, ops:PilotLaunch, ops:PHC_BulkHandling, ops:PortDues, ops:Waste, ops:FW_Supply, ops:EquipmentHire, ops:Manpower, ops:GatePass, ops:DocProcessing*
- **rate:TariffRef / rate:RatePair(EA,Rate,Amount)**
- **cost:CostCenterA/B, cost:PriceCenter** *(A/B/C 3‑Way)*
- **id:Key** ⟶ *id:OFCOInvNo, id:SAFEENInvNo, id:ADPInvNo, id:RotationNo, id:SamsungRef, id:HVDCCode*

### 1.2 핵심 프로퍼티(요지)
- **ops:relatesToPortCall**(InvoiceLine→PortCall), **ops:hasRotationNo**, **fin:belongsToInvoice**, **fin:lineNo**(NO.), **fin:subject**(SUBJECT), **fin:currency**(AED), **fin:vat**(0.00/5.00), **rate:hasEA_i / hasRate_i / lineAmount**, **cost:hasCostCenterA/B / hasPriceCenter**, **prov:hasEvidence**(file,page,line or ref‑row), **id:hasSamsungRef / hasOFCOInvNo / hasRotationNo / hasHVDCCode**.

### 1.3 예시 IRI 정책(요지)
- `ofco:invoice/OFCO-INV-0000181`
- `ofco:line/OFCO-INV-0000181#2015` *(NO.=2015)*
- `ops:portcall/ROT-2504053298` *(RotationNo)*
- `id:samsung/HVDC-AGI-GRM-J71-50`

### 1.4 Mini‑TTL 예시
```ttl
ofco:invoice/OFCO-INV-0000181 a fin:Invoice ;
  fin:currency "AED" ; fin:total "2799.99"^^xsd:decimal .

ofco:line/OFCO-INV-0000181#2002 a fin:InvoiceLine ;
  fin:belongsToInvoice ofco:invoice/OFCO-INV-0000181 ;
  fin:lineNo 2002 ; fin:subject "SAFEEN … Channel Crossing – Rot# 2503123133" ;
  rate:hasEA_1 2.00 ; rate:hasRate_1 3091.25 ;
  rate:hasEA_2 2.00 ; rate:hasRate_2 100.00 ;
  rate:hasEA_3 1.00 ; rate:hasRate_3 239.00 ;
  fin:lineAmount 6621.52 ;
  ops:relatesToPortCall ops:portcall/ROT-2503123133 ;
  cost:hasCostCenterA cost:PORT_HANDLING_CHARGE ;
  cost:hasCostCenterB cost:CHANNEL_TRANSIT_CHARGES ;
  cost:hasPriceCenter  cost:CHANNEL_TRANSIT_CHARGES .
```

---

## 2) Multi‑Key Identity Graph
**문제**: 단일 키 의존 시 연쇄조인 실패·누락 위험.
**해법**: **id:Key** 슈퍼클래스 아래 모든 키를 동등 1급 엔터티로 수집하고, **Same‑As/LinkSet**으로 실체를 클러스터링.

**링크 소스(예)**
- *InvoiceNo(OFCO/SAFEEN/ADP)*, *RotationNo*, *SamsungRef(HVDC‑AGI‑…)*, *HVDCCode*, *Vessel+ETA*.

**클러스터링 규칙(요지)**
1) `RotationNo` 같고, 날짜 창(±7d)·항만 동일 → 같은 **PortCall** 후보.
2) `SamsungRef` 동일 + Subject 패턴 일치 → 같은 **Operation Batch**.
3) `InvoiceNo` 묶음 Σ(lineAmount) = Invoice Total(±2.00%) → 같은 **Invoice**.

---

## 3) SHACL 검증(요약)
- **InvoiceLineShape**
  - `rate:hasEA_* × rate:hasRate_*`의 합 = `fin:lineAmount` ±0.01
  - RatePair 최대 4, 결측 시 0.00 채움
  - `fin:currency = "AED"`, `fin:vat ∈ {0.00, 5.00}`
  - `prov:hasEvidence` **필수**
- **InvoiceShape**
  - Σ(InvoiceLine.fin:lineAmount) = `fin:total` ±2.00%
  - `[EXT]` 라벨 행 금액 **집계 제외**
- **PortCallLinkShape**
  - Subject에 `Rot#` 있으면 `ops:relatesToPortCall` **필수**

---

## 4) Cost/Price Center 매핑 규칙(OFCO 전용)
- **Regex v2.5 + Subject 패턴**(요지)
  - `"SAFEEN"` → `PORT HANDLING CHARGE / CHANNEL TRANSIT CHARGES`
  - `"ADP INV"`+`"Port Dues"` → `PORT HANDLING CHARGE / PORT DUES & SERVICES CHARGES`
  - `"Cargo Clearance"` → `CONTRACT / AF FOR CC`
  - `"Arranging FW Supply"|"FW Supply"` → `CONTRACT / AF FOR FW SA`
  - `"Berthing Arrangement"` → `CONTRACT / AF FOR BA`
  - `"5000 IG FW"` → `AT COST / AT COST(WATER SUPPLY)`

- **PRICE CENTER 3‑Way**
  - **A/B**: Tariff·키워드 기반(예: Channel Crossing/Port Dues/PHC/Equipment/Manpower)
  - **C**: 수수료/Pass/Document(예: Gate Pass, Doc Processing)
  - **규칙**: *C=0 의심 재검토*, *A>B 또는 B<0 시 일부 C로 이동*, *A+B+C=Original_TOTAL, Diff=0.00*

---

## 5) 파이프라인(운영·검증)
1) **Pre‑Prep**: 회전/데스큐/샤프닝(DPI<300 경고)
2) **OCR v2.4**: 레이아웃·토큰 conf 수집
3) **Smart Table Parser 2.1**: 병합셀 해제·세로표 가로화·단위/통화 분리
4) **NLP Refine**: NULL/단위 보정, 추정 금지
5) **Field Tagger**: Parties/IDs/Incoterms/Rotation/Subject
6) **LDG Payload Build**: 해시·CrossLinks·Evidence
7) **Mapping & QC**: EA×Rate 분해, Cost/Price Center 적용, VAT·통화·합계 검증
8) **COST‑GUARD**: 기준요율 대비 Δ% 밴드(PASS/WARN/HIGH/CRITICAL)
9) **Report(7+2)**: Discrepancy Table, Compliance Matrix, DEM/DET Forecast 등

**KPI 게이트**: MeanConf≥0.92, TableAcc≥0.98, NumericIntegrity=1.00 → 미달 시 **ZERO** 중단 로그.

---

## 6) 데이터 맵(Excel/JSON → Ontology)
| Source Field | Ontology Property | Note |
|---|---|---|
| `NO.` | `fin:lineNo` | Row key |
| `SUBJECT` | `fin:subject` | 패턴 매핑 트리거 |
| `SAMSUNG REF` | `id:hasSamsungRef` | 클러스터 anchor |
| `Channel Crossing Charges…` 등 금액열 | `fin:lineAmount` 또는 `rate:hasRate_i`/`rate:hasEA_i` | 금액→Line, EA/Rate 분해 |
| `EA_1..4` | `rate:hasEA_i` | 최대 4 쌍 |
| `Rate_1..4` | `rate:hasRate_i` | 금액=Σ(EA×Rate) |
| `Amount (AED)` | `fin:lineAmount` | 2 decimals |
| `INVOICE NUMBER (OFCO)` | `id:hasOFCOInvNo` | Invoice join |
| `Rotation#`(Subject 내) | `ops:hasRotationNo` | PortCall link |

---

## 7) Report 표준(7+2)
1) **Auto Guard Summary**
1.5) **Risk Assessment**(등급/동인/신뢰도)
2) **Discrepancy Table**(Δ·허용오차·상태)
3) **Compliance Matrix**(UAE·근거 링크)
4) **Auto‑Fill**(Freight/Insurance)
5) **Auto Action Hooks**(명령·가이드)
6) **DEM/DET & Gate‑Out Forecast**
7) **Evidence & Citations**
8) **Weak Spot & Improvements**
9) **Changelog**

---

## 8) 운영 명령 & 자동화 훅
- **인식/검증**: `/ocr_basic {file} mode:LDG+` → KPI Pass 확인 → `/ocr_table`/`/ocr_retry`
- **코스트가드**: `/switch_mode COST-GUARD + /logi-master invoice-audit --AEDonly`
- **매핑**: `/mapping run` → `/run pricecenter map` → `/mapping update pricecenter`
- **규제 체크**: `/logi-master cert-chk`(MOIAT/FANR/TRA)
- **배치**: `/workflow bulk …` → `/export excel`

---

## 9) 운영 규칙(정합성)
- `Σ(BB:BI)=BJ` ±2.00% / EA 결측 시 EA=1.00 & Rate=Amount 규칙 허용(내 ±2.00%)
- VAT=0.00% 또는 5.00% 외 **[MISMATCH]**
- `[EXT]` 메타는 **금액 집계 제외**, 근거(M열) 필수
- **증거(Evidence)**: 파일명/페이지/라인 또는 참조시트(Row) **필수 기록**

---

## 10) 로드맵 (P→Pi→B→O→S + KPI)
- **Prepare(2주)**: 스키마/네임스페이스/IRI 설계, SHACL 초안, 키‑링크 룰 정의
  KPI: 스키마 커버리지 ≥90.00%
- **Pilot(3주)**: 1개 인보이스 묶음(예: OFCO‑INV‑0000181) E2E, Δ오차≤2.00%
  KPI: ZERO 트리거=0, Evidence 100.00%
- **Build(4주)**: CostCenter v2.5·3‑Way 분개·COST‑GUARD 통합
  KPI: Pass율≥95.00%
- **Operate(지속)**: 배치 처리 및 리포트(7+2) 자동 발행
  KPI: TAT ≤ 0.50h/건
- **Scale(지속)**: SAFEEN/ADP 직조인, PortCall API, DEM/DET 2.0 연계
  KPI: 오탐율 ≤ 2.00%

---

## 11) 리스크 & 완화
- **키 불일치/누락**: Multi‑Key 흡수 + 휴리스틱 윈도우(±7d)
- **OCR 품질 저하**: KPI 게이트 + `/ocr_lowres_fix` + ZERO 중단
- **요율 변동**: TariffRef 버전드(유효일) + COST‑GUARD Δ% 밴드

---

## 12) 부록 — Subject→Cost/PriceCenter 예시(발췌)
| Subject 큐 | Cost A | Cost B | PriceCenter |
|---|---|---|---|
| SAFEEN … Channel Crossing | PORT HANDLING CHARGE | CHANNEL TRANSIT CHARGES | CHANNEL TRANSIT CHARGES |
| ADP INV … Port Dues | PORT HANDLING CHARGE | PORT DUES & SERVICES CHARGES | PORT DUES |
| Agency fee: Cargo Clearance | CONTRACT | AF FOR CC | AGENCY FEE FOR CARGO CLEARANCE |
| Arranging FW Supply | CONTRACT | AF FOR FW SA | SUPPLY WATER 5000IG |
| Berthing Arrangement | CONTRACT(AF FOR BA) | CONTRACT | AGENCY FEE FOR BERTHING ARRANGEMENT |
| 5000 IG FW | AT COST | AT COST(WATER SUPPLY) | SUPPLY WATER 5000IG |

---

## 13) 구현 노트
- 코드베이스: `logiontology/`(mapping/validation/reasoning/rdfio/report/pipeline)
- SHACL Runner 옵션, JSON‑LD 컨텍스트 제공, RDFLib + DuckDB로 라인‑레벨 집계 검증.
- 외부 연계: PortCall(AD Ports)·SAFEEN 청구 스냅샷 → `TariffRef` Evidence로 보관.

---

### 끝. (숫자 2 decimals, ISO 날짜 사용, NDA/PII 마스킹 준수)

