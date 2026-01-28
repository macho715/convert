---
title: "JPT 71 - Ontology Blueprint"
type: "ontology-design"
domain: "vessel-operations"
version: "1.0"
date: "2025-01-19"
tags: ["ontology", "jpt71", "vessel", "stability", "load", "hvdc", "operations"]
vessel: "JOPETWIL 71 (JPW71)"
status: "active"
---

# JPT 71 — Ontology Blueprint v1.0 (HVDC Logistics Multi-Key Identity Graph)

## 개요

> TL;DR — Any-key in → Resolve to the same JPW71 cluster (Vessel·Contract·Stability·Load·Ops) → Reason over constraints (Class vs. Field) → Output a load/ops envelope with compliance and KPIs. (EN-KR one-liner)

---

## 0) Purpose & Scope
- 목적: LCT **JOPETWIL 71 (JPW71)** 관련 ‘문서·운영·안전’ 정보를 **온톨로지(지식 그래프)** 로 통합.
- 범위: 선박(제원·검사) / 계약(SUPPLYTIME, Amend) / Deck Upgrade / 안정성(Rev.8) / 화물(골재·모래·프리캐스트·A‑Frame) / 실적(Loading Record) / 운영 제약(Sea state, Padeye SWL, CEP/Vetting) / 기상 윈도우 / KPI.
- 원칙: **Multi‑Key Identity Graph** — 입력 키는 BL/Case/ShipmentID/ContractID/hvdc_code/TripNo 중 아무거나.

---

## 1) Identity & URI Scheme
- Base IRI: `https://hvdc.example.org/` (로컬 배포 시 환경 변수화)
- 네임스페이스:  
  `hvdc:` 도메인 클래스·속성,  `hvdci:` 개별 인스턴스,  `doc:` 문서 레퍼런스,  `org:` 조직,  `geo:` 위치.
- UID 규칙: `hvdci:{Type}/{slug or hash}`  
  예) `hvdci:Vessel/JPW71`, `hvdci:Charter/OLCOF24-ALS086`, `hvdci:Stability/Rev8-BV`, `hvdci:LoadEvent/J71-067`.

---

## 2) Core Classes (요지)
- **hvdc:Vessel** — LCT 선박(제원, Class, Flag, IMO)
- **hvdc:ConditionSurvey** — 선박 상태조사 리포트(일자, 발행처, 주요 소견)
- **hvdc:DeckUpgrade** — 데크 콘크리트 보호층(시공 범위, 완료일, 두께/면적, 목적)
- **hvdc:CharterParty** — SUPPLYTIME 2017 계약(기간, Hire, 범위, Amend)
- **hvdc:Amendment** — 연장/조건 변경(효력일, 조항 변경)
- **hvdc:StabilityAddendum** — Trim & Stability Addendum Rev.x (허용 톤수, Deck Strength, VCG 조건, Weather Criteria)
- **hvdc:CargoType** — Aggregate(5/10/20mm), Sand, Precast(HCS, Wall), A‑Frame, Jumbo Bag, Soil 등
- **hvdc:LoadEvent** — 항차·일자·톤·품목 기록(실적)
- **hvdc:Operation** — Loading/Offloading/Sailing·Berth·GatePass·MWS·CEP 상태
- **hvdc:LashingHardware** — Padeye, Chain, Belt, SWL, 구속 방향/각
- **hvdc:WeatherWindow** — 해상 상태(Sea State, Hs, Wind), 운항 허용치
- **hvdc:VettingCEP** — 섬/항만 CEP·Vetting(유효기간, 상태)
- **hvdc:Stakeholder** — Owner/Operator/Charterer/Sub‑con/Class/Authority

---

## 3) Key Individuals (대표 인스턴스)
- `hvdci:Vessel/JPW71`  — 이름 JOPETWIL 71, IMO 9582829, BV Class, LOA 70.90 m, B 15.00 m, UAE Flag
- `hvdci:DeckUpgrade/2024-07` — 목적: Bulk(골재·모래) 취급용 보호층; 범위: Main Deck 콘크리트 레이어; 결과: 허용 Deck Strength 연계
- `hvdci:Charter/OLCOF24-ALS086` — Charter Party(SUPPLYTIME 2017), Day Rate, Area, Cargo 범위
- `hvdci:Amend/No2-2025-07-08` — Charter 기간 연장 ~ 2025‑10‑05, 연장 조건
- `hvdci:Stability/Rev8-BV` — Aggregate Deck Load 최대치, Deck Strength 10 MT/m², 유효 Footprint 274 m², Weather/VCG 조건
- `hvdci:ConditionSurvey/2024-07-23` — 상태: 작동/설비 점검, 비고 항목
- `hvdci:Constraint/AFrame` — Deck padeye SWL 2.5T, Sea State 4~5 ft 제한(라싱 계산 필요)
- `hvdci:Record/J71` — 2024–2025 전 항차 LoadEvent 테이블(tonnage 트렌드 포함)

---

## 4) Object/Data Properties (발췌)
- `hvdc:hasOwner (Vessel→org)`
- `hvdc:hasOperator (Vessel→org)`
- `hvdc:hasCharterer (Vessel→org)`
- `hvdc:governedBy (Vessel→CharterParty)`
- `hvdc:hasAmendment (CharterParty→Amendment)`
- `hvdc:hasDeckUpgrade (Vessel→DeckUpgrade)`
- `hvdc:hasStabilityAddendum (Vessel→StabilityAddendum)`
- `hvdc:permitsCargo (CharterParty→CargoType)`
- `hvdc:hasDeckStrength (StabilityAddendum→xsd:decimal)  # MT/m2`
- `hvdc:allowsMaxDeckAggregate (StabilityAddendum→xsd:decimal)  # MT`
- `hvdc:hasEffectiveFootprint (StabilityAddendum→xsd:decimal)  # m2`
- `hvdc:hasPadeyeSWL (LashingHardware→xsd:decimal)  # T`
- `hvdc:limitedBySeaState (Operation→xsd:string)`
- `hvdc:hasCEPStatus (VettingCEP→xsd:string)`
- `hvdc:records (Vessel→LoadEvent)`
- `hvdc:hasMaterial, hvdc:hasSize, hvdc:deliveredTon, hvdc:onDate, hvdc:onTripNo (LoadEvent→…)`

---

## 5) Inference — Load/Ops Envelope (규칙 요약)
**A. Class(이론) 경계**  
- `MaxAggregateByStability = hvdc:allowsMaxDeckAggregate (e.g., 800.00 MT)`
- `DeckStrengthCap = hvdc:hasDeckStrength × hvdc:hasEffectiveFootprint (e.g., 10 × 274 = 2,740 MT)`, 다만 실제 분포·안식각 반영

**B. Field(현장) 경계**  
- `A‑Frame`: Padeye SWL 2.5T → 라싱 체인/벨트 인장력 계산 필요 → 미계산 시 해상조건 `SeaState ≤ 4~5 ft` 제한
- `CEP/Vetting`: CEP 만료 시 특정 항만(예: DAS) Bulk/특수 화물 출입 제한 → Vetting close‑out 후 갱신

**C. 합성 Envelope**  
- `PermissibleLoad = min(MaxAggregateByStability, FieldConstraints, WeatherWindow)`  
- 운영 시뮬레이션: `Aggregate only` vs `Mix with Precast/A‑Frame` vs `Jumbo Bag`로 Case별 Envelope 산출

---

## 6) SHACL Shapes (스케치)
```ttl
@prefix sh: <http://www.w3.org/ns/shacl#> .
@prefix hvdc: <https://hvdc.example.org/ns#> .
@prefix xsd: <http://www.w3.org/2001/XMLSchema#> .

hvdc:LoadEventShape a sh:NodeShape ;
  sh:targetClass hvdc:LoadEvent ;
  sh:property [
    sh:path hvdc:deliveredTon ;
    sh:datatype xsd:decimal ;
    sh:minInclusive 0.0 ;
  ] ;
  sh:property [
    sh:path hvdc:hasMaterial ;
    sh:in ( hvdc:Aggregate5 hvdc:Aggregate10 hvdc:Aggregate20 hvdc:Sand hvdc:Precast hvdc:AFrame hvdc:JumboBag hvdc:Soil ) ;
  ] ;
  # 규칙형 제약: deliveredTon ≤ PermissibleLoad(날짜별 Envelope)
  # 구현: SHACL-SPARQL 또는 룰 엔진에서 계산 결과를 sh:maxInclusive 로 바인딩
.
```

---

## 7) Example Triples (TTL)
```ttl
hvdci:Vessel/JPW71 a hvdc:Vessel ;
  hvdc:imo "9582829" ;
  hvdc:name "JOPETWIL 71" ;
  hvdc:hasOperator org:ADNOC_LS ;
  hvdc:hasCharterer org:Samsung_CT ;
  hvdc:governedBy hvdci:Charter/OLCOF24-ALS086 ;
  hvdc:hasDeckUpgrade hvdci:DeckUpgrade/2024-07 ;
  hvdc:hasStabilityAddendum hvdci:Stability/Rev8-BV ;
  hvdc:records hvdci:LoadEvent/J71-067 .

hvdci:Stability/Rev8-BV a hvdc:StabilityAddendum ;
  hvdc:hasDeckStrength 10.00 ;       # MT/m2
  hvdc:hasEffectiveFootprint 274.00 ; # m2
  hvdc:allowsMaxDeckAggregate 800.00 .

hvdci:LoadEvent/J71-067 a hvdc:LoadEvent ;
  hvdc:onDate "2025-09-08"^^xsd:date ;
  hvdc:hasMaterial hvdc:Aggregate5 ;
  hvdc:deliveredTon 840.86 .
```

---

## 8) JSON‑LD Context (발췌)
```json
{
  "@context": {
    "hvdc": "https://hvdc.example.org/ns#",
    "name": "hvdc:name",
    "imo":  "hvdc:imo",
    "hasStabilityAddendum": {"@id":"hvdc:hasStabilityAddendum","@type":"@id"},
    "deliveredTon": {"@id":"hvdc:deliveredTon","@type":"xsd:decimal"}
  }
}
```

---

## 9) Data Sources → Graph Mapping (요지)
- Condition Survey → `hvdc:ConditionSurvey` (제원, 인증 유효기간, 설비 상태)
- Deck Upgrade 견적/완료 → `hvdc:DeckUpgrade`
- SUPPLYTIME + Amend → `hvdc:CharterParty`·`hvdc:Amendment` (기간, Hire, Cargo 범위)
- Stability Addendum Rev.8 → `hvdc:StabilityAddendum` (800MT, Deck Strength, Weather Criteria, VCG)
- 운영 서신(CEP/Vetting, Sea state, Padeye) → `hvdc:VettingCEP`, `hvdc:LashingHardware`, `hvdc:Operation`
- Loading Record(2024–2025) → `hvdc:LoadEvent` 시계열

---

## 10) Reasoning Recipes (의사코드)
1) **Aggregate‑only**  
`permit = min( stability.maxAggregate , weather.window , ops.ceiling )`

2) **With A‑Frames**  
`lash = solveLashingForces(m, cog, μ, roll, sea_state)`  
`permit = (lash ≤ padeyeSWL && sea_state ≤ limit) ? min(stability, weather) : reduced_cap`

3) **Jumbo Bag 시나리오**  
`permit = min(stability, palletization/stacking rules, forklift ops limits)`

---

## 11) KPIs & Dash Hooks
- **Load per Trip (t)**, **OTIF %**, **Envelope Utilization %**(actual / permissible), **CEP Validity Days**, **MWS Observations Close‑out**
- 알림: CEP 만료 14일 전, Vetting 재검 7일 전, Sea State>limit 시 ‘NO‑GO’ 플래그

---

## 12) Next Steps (실행)
- (P) 온톨로지 스키마/컨텍스트 `ontology.ttl`, `context.jsonld` 생성
- (Pi) 소스→RDF 매핑(Condition Survey, Stability Rev.8, Charter, Record) 1차 적재
- (B) **SHACL** 검증 + **룰엔진**(PermissibleLoad) 적용
- (O) 운영보드: Trip 계획 입력→Envelope 자동 산출, 위험(Sea/Padeye/CEP) 경고
- (S) A‑Frame 라싱계산 모듈 연동(ECGM, chain grade, α/β angles) 및 해상조건 룰 고도화

---

## 13) Glossary
- **Envelope**: 시점별 허용 적재·운항 범위(이론×현장×기상)
- **SWL**: Safe Working Load
- **VCG**: Vertical Center of Gravity
- **CEP/Vetting**: 통제 구역 운영허가/검사

