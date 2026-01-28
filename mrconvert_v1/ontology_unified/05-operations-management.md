---
title: "Operations Management Ontology"
type: "ontology-design"
domain: "operations-management"
sub-domains: ["warehouse", "bulk-cargo", "vessel-operations"]
version: "unified-1.0"
date: "2025-01-19"
tags: ["ontology", "warehouse", "bulk-cargo", "vessel", "jpt71", "operations", "hvdc"]
status: "active"
---

# Part 1: Warehouse Management

__ExecSummary__

창고 pjt를 __온톨로지\(지식그래프\) 관점__으로 보면, 엑셀/ERP의 각 행은 TransportEvent\(이동\), StockSnapshot\(재고 스냅샷\), Invoice\(청구\), Case\(개별 케이스\) 같은 __클래스__로 귀속되고, 열들은 hasDate/hasLocation/hasQuantity/hasLogisticsFlowCode 같은 __속성__으로 정규화됩니다\. 이 구조가 “창고 트랙\(WH\)”·“현장 트랙\(Site\)”·“Flow Code\(0–4\)”를 한 장의 그래프로 __동일 실체__에 묶어 줍니다\. \(Any\-key in → Resolve→Cluster→Downstream\)   
매핑된 데이터는 __RDF/OWL__로 변환되어 SPARQL로 검증/집계가 가능하고, 비용 분류\(OFCO\)나 월별 입출고·재고·SQM 과금까지 __한 체계__에서 굴러갑니다\.   
핵심은 “2\-트랙 날짜 컬럼\(창고 vs 현장\)”과 __시간순 출고 판정__·__이중계산 방지__·__Flow 0–4 일관성__을 코드 레벨로 보증하는 것입니다\.

__Visual — Ontology Map \(요약표\)__

__Layer__

__Ontology 객체/속성__

__소스 열\(예\)__

__역할/효과__

장소모델

Warehouse\(Indoor/Outdoor/AAA/MZP/MOSB\), Site\(AGI/DAS/MIR/SHU\)

DSV Indoor/Outdoor, AAA Storage, MOSB, AGI…

창고/현장 계층 표현\(Indoor/Outdoor/Offshore\) → 의미론적 위치 집계

이벤트

TransportEvent \+ hasDate/hasLocation/hasQuantity

창고/현장 날짜, Pkg/CBM

“언제, 어디로, 몇 개/면적” 이동을 그래프에 기록

흐름

hasLogisticsFlowCode\(0~4\)

wh handling 또는 창고 방문 횟수

Port→WH→\(MOSB\)→Site 경로를 정규화\(0=Pre\-Arrival…4\)

재고

StockSnapshot

Status\_Location, Status\_Location\_Date

월말 스냅샷/누계 재고 산출의 기준 노드

비용

Invoice/InvoiceLineItem \+ OFCO 매핑

Description/Rate/Amount

AT\-COST/CONTRACT 등 비용센터 자동 분류

__파이프라인 to KG \(요약\)__  
Ingest\(Excel\) → 정규화\(헤더/날짜/공백\) → 매핑\(JSON rules\) → RDF 변환 → SPARQL 검증\(12 rules\) → Flow/WH·Site 집계 → 리포트/과금\(SQM\) 

__How it works \(핵심 동작 원리, EN\-KR one\-liners\)__

1. __2\-트랙 날짜 모델__: 창고 컬럼\(DSV Indoor/Al Markaz/AAA/MOSB…\)과 현장 컬럼\(AGI/DAS/MIR/SHU\)을 분리 인식 → 최신 위치/이동 추론 강화\. 
2. __Flow Code 계산\(0–4\)__: Pre\-Arrival\(0\)~WH/MOSB 경유~Site 도착까지 hop 수\+오프쇼어 경유로 표준화\.
3. __출고 판정\(시간순\)__: “창고에 찍힌 날짜 < 다음 위치\(다른 창고/현장\) 날짜”일 때만 출고로 인정\(동일일자 중복 방지\)\. 
4. __이중계산 방지 \+ 검증__: 창고간 이동 목적지는 입고에서 제외, 재고는 Status\_Location vs 물리위치 __교차검증__\(불일치 0건 목표\)\. 
5. __RDF/OWL & SPARQL__: DataFrame→RDF 자동 변환, 금액/패키지/위치/시간 일관성 규칙 12종으로 품질게이트\. 
6. __리포팅 아키텍처__: 5\-시트 요약\(Flow/WH·Site 월별/Pre\-Arrival/전체 트랜잭션\) \+ 27시트 스냅샷\(B5 날짜 기반 시계열\) \+ SQM 과금\.

__Options \(구현 옵션 ≥3 · pros/cons/$/risk/time\)__

1. __Option A — Lite KG\(매핑\+피벗 중심\)__

- Pros: 빠른 적용, 5\-시트 리포트 즉시화, 기존 엑셀 호환 우수\. 
- Cons: 실시간 추론/질의 한계, 규칙 변경 시 수작업 많음\.
- Cost/Time: $ · 1–2주\.
- Risk: 규칙 누락/헤더 변형에 민감\(중\)\.

1. __Option B — Full KG\(\+SPARQL 검증/자동 추론\)__

- Pros: RDF 변환\+12개 규칙 검증, 의미론 질의/벤더·월·창고 통합 시계열 안정\. 
- Cons: 온톨로지/삼중저장소 운영 필요\.
- Cost/Time: $$ · 3–5주\.
- Risk: 초기 스키마 설계 미스매치\(중\)\.

1. __Option C — Ops Twin\(\+Flow 추적·SQM 과금\)__

- Pros: 시간순 출고·이중계산 방지, SQM 누적/요율 기반 월별 과금 자동화\. 
- Cons: 데이터 품질\(SQM 실측률\)에 민감\.
- Cost/Time: $$ · 4–6주\.
- Risk: 일부 항목 SQM 추정치 사용 시 오차\(중\)\. 

__Roadmap \(Prepare→Pilot→Build→Operate→Scale \+ KPI\)__

__Prepare \(1주\)__

- 헤더/날짜 정규화, 전각공백\(‘\\u3000’\) 처리, 중복제거 파이프라인 정리\. *KPI: 정제 성공률 ≥ 94\.60%\.* 

__Pilot \(1–2주\)__

- 2\-트랙 매핑 \+ Flow 0–4 적용, 5\-시트 리포트 생성\. *KPI: Flow 계산 일치율 100\.00%\.* 

__Build \(2–3주\)__

- RDF 변환 \+ SPARQL 12규칙, OFCO 비용센터 매핑 연결\. *KPI: 검증 규칙 통과율 100\.00%\.* 

__Operate \(지속\)__

- 시간순 출고/재고 교차검증, 이중계산 0건 유지, SQM 월별 과금\. *KPI: PKG Accuracy ≥ 99\.00% / Inventory 불일치 0건\.* 

__Scale \(지속\)__

- 27시트 스냅샷 도입\(B5 기반 시계열\), 트렌드/변동 자동 감지\. *KPI: 스냅샷 커버리지 100\.00%\.* 

__Automation Hooks \(RPA\+LLM\)__

- __/logi\-master kpi\-dash__: Flow/WH·Site 월별 피벗 \+ KPI 리포트 생성\. 
- __/logi\-master report \-\-deep__: RDF 변환→SPARQL 검증→요약 리포트\. 
- __/logi\-master cert\-chk | invoice\-audit__: OFCO/비용센터 라벨링과 교차 검증\. 
- __/visualize\_data \-\-type=pkg\-flow__: Port→WH→\(MOSB\)→Site 흐름 시각화\(Flow 0–4\)\. 

__QA / Gap 체크리스트__

- 창고 vs 현장 컬럼 __완전 분리__ 적용 여부\(이중계산 방지\)\. 
- 출고 판정이 “다음 위치가 더 늦은 날짜” 규칙을 지키는가\. 
- Flow 0–4 경로 정의와 hop 계산 일치 여부\. 
- 전처리\(전각공백/날짜 정규화/중복제거\) 성공 여부\. 
- SPARQL 12 규칙 통과\(금액 음수/패키지 양수/시간 일관성 등\)\. 
- SQM 실측 vs 추정 비율 보고\(정책: 실측 비중을 단계적으로 상향\)\. 

__CmdRec \(바로 실행\)__

1. __/logi\-master kpi\-dash \-\-KRsummary__ → 월별 WH/Site·Flow 요약 5\-시트 생성\. 
2. __/logi\-master report \-\-deep__ → RDF 변환\+SPARQL 검증\+OFCO 라벨링\. 
3. __/visualize\_data \-\-type=pkg\-flow__ → Flow 0–4 동선 확인\(이상 경로 탐지\)\. 

__한 줄 정리__

__창고 pjt의 ‘한 몸체’는 온톨로지다\.__ 장소·시간·흐름·재고·비용을 __하나의 그래프__에 올려두면, 어떤 키로 들어와도\(케이스·BL·Site…\) 같은 실체로 모이고, 그다음은 계산이 아니라 __질의__가 된다\. \(그리고, 그게 가장 덜 고생한다\.\)



---

# Part 2: Bulk Cargo Operations

# BULK CARGO OPERATION - 온톨로지 관점

## 개요

현장 용어와 데이터를 한 언어로 묶어, **무엇(Thing)**—**어디(Location)**—**언제(Time)**—**어떻게(Operation)**—**무엇으로(Resource)**—**왜/규정(Compliance)**을 서로 연결하는 지식 그래프로 설계합니다. 아래는 바로 적용 가능한 최소 핵심 스키마와 운영 포인트입니다.

__1\) 최상위 개념\(Top\-level Classes\)__

- __Cargo__: CargoItem, Package/Bundle, Lot
	- 속성: weight, dimensions\(L/W/H\), COG\(x,y,z\), stackable, hazardousNote …
- __TransportMeans__: Vessel, Barge, Truck, Trailer
	- 속성: deckStrength, deckArea, coordOrigin, capacity…
- __Location__: Port, Terminal, Jetty, __DeckZone__\(구역/그리드\), StorageBay, Berth
- __Operation__:
	- __Loading/Discharging__, __Stowage__, __Lashing/Seafastening__, __Lifting__, Pre\-carriage, SeaPassage, Inspection
	- 상태: Planned → Ready → InProgress → Completed → Verified
- __Resource__:
	- __Equipment__\(Crane, Forklift, Spreader, RiggingGear: sling/shackle/beam\), __Workforce__\(Rigger, Banksman, Operator\)
- __Document__: StowagePlan, LashingPlan, StabilityReport, LiftingPlan, MS/JSA, P/L, B/L, Permit
- __Condition/Measurement__: Weather, SeaState, Wind, Motion\(accel g\), Clearance
- __Organization/Agent__: OFCO, DSV, SCT, Client, Surveyor, Class
- __Constraint/Rule__: DeckLoadLimit, SWL/ WLL, ClearanceRule, RegulatoryRule
- __Time__: Instant/Interval\(ETA/ETD, Shift\), Milestone

말 그대로, “화물—작업—장비—장소—시간—문서—규정”을 모두 1개의 그래프에서 ‘연결’해 질문이 통과되게 만듭니다\.

__2\) 핵심 관계\(Object Properties\)__

- cargoLocatedAt\(Cargo → DeckZone | StorageBay\)
- assignedTo\(Cargo → Operation\) / produces\(Operation → Document\)
- securedBy\(Cargo → RiggingGear\) / performedBy\(Operation → Workforce|Organization\)
- uses\(Operation → Equipment\) / occursAt\(Operation → Location\) / scheduledFor\(Operation → TimeInterval\)
- constrainedBy\(TransportMeans|Operation → Constraint\)
- hasMeasurement\(… → Measurement\) / hasStatus\(… → StatusConcept\)

__3\) 필수 데이터 속성\(Data Properties\) 예__

- weight\(kg|t\), length/width/height\(m\), cogX/Y/Z\(m\)
- deckStrength\(t/m2\), radius\(m\), swl/wll\(t\)
- windSpeed\(m/s\), roll/pitch\(deg\), accelLong/Trans/Vert\(g\)
- startAt/endAt\(ISO 8601\), docVersion, approvalState

__4\) 표준 연계\(Interoperability\)__

- __단위__: QUDT/UCUM \(kg, t, m, deg, m/s²\)
- __시간__: OWL\-Time \(Instant/Interval\)
- __측정/센서__: SOSA/SSN \(가속도, 풍속\)
- __위치/좌표__: GeoSPARQL \(DeckZone도 폴리곤/그리드로 모델링\)
- __어휘/상태표__: SKOS \(작업상태/허가상태 코드셋\)
- __근거성__: PROV\-O \(문서가 어떤 작업/데이터에서 파생됐는지\)

표준을 재사용하면 시스템 간 데이터 교환이 편해지고, 단위 오류를 줄입니다\.

__5\) 규칙/검증\(Constraints\) — SHACL로 예시__

- __Deck 접지압__: Σ\(cargo\.weight / contactArea\) ≤ deckStrength
- __Lashing 용량__: Σ\(WLL × cosθ\) ≥ designLoad × safetyFactor
- __Crane 반경 SWL__: SWL\(radius\) ≥ liftedWeight × factor
- __Clearance__: cargo\.height \+ grillage ≤ allowableHeight\(zone\)

규칙은 __SHACL__\(또는 규칙엔진\)로 선언해 “데이터가 들어오는 순간” 자동 검증하게 합니다\.

__6\) 컴피턴시 질문\(이 온톨로지가 반드시 답해야 할 질문\)__

1. 현재 선적안\(버전 X\)에서 __DeckZone A__의 총 하중과 접지압은 안전한가?
2. __LashingPlan \#123__에서 각 슬링의 예상 장력과 WLL 대비 사용률은?
3. 반경 R에서 __선정 크레인__의 SWL이 리프트에 충분한가?
4. __COG가 높은 화물__만 필터해 추가 시추/보강이 필요한 후보는?
5. 오늘 야간\(19:00–07:00\) __필요 인력/장비__와 공석은?
6. __SeaState ≥ 5__ 조건에서 가속도\(g\) 가정이 바뀌면 어떤 화물의 라싱이 불합격되는가?
7. 특정 __B/L__에 포함된 Cargo들의 __Stowage 위치/문서/승인 현황__은?
8. __OFCO/DSV/SCT__ 각각 담당 작업과 책임 경계는 어디까지인가?
9. 마지막 승인된 __StabilityReport__와 실제 탑재 데이터\(중량/VCG\)의 차이는?
10. 적재 순서 변경 시 __크리티컬 경로/대기시간__ 변화는?

__7\) 아주 작은 예시\(Turtle\)__

@prefix bco:   <https://example\.com/bco\#> \.

@prefix time:  <http://www\.w3\.org/2006/time\#> \.

@prefix qudt:  <http://qudt\.org/schema/qudt/> \.

@prefix unit:  <http://qudt\.org/vocab/unit/> \.

bco:CARGO\_001 a bco:Cargo ;

  bco:weight "18\.5"^^qudt:QuantityValue ;

  qudt:unit unit:T ;

  bco:length "12\.0" ; bco:width "2\.4" ; bco:height "2\.8" ;

  bco:cogX "6\.0" ; bco:cogY "1\.2" ; bco:cogZ "1\.4" ;

  bco:cargoLocatedAt bco:DeckZone\_A ;

  bco:assignedTo bco:OP\_Load\_20251019 \.

bco:OP\_Load\_20251019 a bco:LoadingOperation ;

  bco:occursAt bco:Vessel\_TRUE\_Deck ;

  bco:uses bco:Crane\_80T ;

  bco:scheduledFor bco:Shift\_Night\_20251019 \.

bco:Shift\_Night\_20251019 a time:Interval ;

  time:hasBeginning "2025\-10\-19T19:00:00\+04:00" ;

  time:hasEnd       "2025\-10\-20T07:00:00\+04:00" \.

__8\) SHACL 스케치\(간단 아이디어\)__

bco:DeckLoadShape a sh:NodeShape ;

  sh:targetClass bco:DeckZone ;

  sh:sparql \[

    sh:message "Deck load exceeds allowable pressure\." ;

    sh:select """

      SELECT ?this WHERE \{

        ?this a bco:DeckZone ; bco:deckStrength ?limit \.

        \{

          SELECT ?this \(SUM\(?w/?area\) AS ?pressure\)

          WHERE \{

            ?cargo bco:cargoLocatedAt ?this ; bco:weight ?w ; bco:contactArea ?area \.

          \} GROUP BY ?this

        \}

        FILTER \(?pressure > ?limit\)

      \}

    """ ;

  \] \.

__9\) 운영 설계 팁\(현장 맞춤\)__

- __ID 정책__: VSL\_TRUE/ZONE\-A/2025\-10\-19/LOT\-xxx처럼 사람과 시스템이 같이 읽히는 URI/ID\.
- __DeckZone 그리드화__: 2D 좌표계 기준\(Origin, X fwd, Y port\)과 격자 크기\(예: 1×1 m\)를 그래프에 저장\.
- __문서\-데이터 연결__: LashingPlan, StabilityReport를 __produces/validates__ 관계로 작업/데이터와 연결\.
- __버전/승인 추적__: PROV\-O로 “누가, 언제, 무엇을” 승인/수정했는지 이력 관리\.
- __상태어휘\(SKOS\)__: Planned/Ready/InProgress/OnHold/Completed/Rejected 같은 컨트롤 타워용 코드셋 고정\.
- __규정 계층화__: SOLAS/IMSBC/AD Ports 규정을 Rule 노드로 선언하고, 작업/장비에 constrainedBy로 링크\.

__10\) 시스템 아키텍처\(간단 청사진\)__

- __Triple Store/Graph DB__\(RDF/OWL\) \+ __SHACL Validator__
- __Ingest 파이프라인__: CSV/Excel\(화물, 장비, 인력, 스케줄\) → 매핑\(R2RML/ETL\) → RDF
- __Query API__: SPARQL endpoint \+ GraphQL façade\(현장 앱/대시보드 용\)
- __Rule/Calc__: SHACL\(SP\), 파이프라인 계산\(예: 라싱 각도/장력\), 결과를 Measurement로 귀속
- __문서화__: 그래프에서 최신 상태를 끌어와 Stowage/Lashing/Lifting/Logistics Plan 자동 채움

__11\) 지금 있는 데이터와의 핏__

당신이 이미 관리하는 __화물/선박/장비/인력/환경/스케줄 표__는 그대로 쓰되,

- 열\(Column\)마다 __어떤 클래스/속성__으로 들어갈지 맵핑 테이블만 정하면 됩니다\.
- 이후부터는 “질문”이 곧 “SPARQL 쿼리”가 되고, 검증은 SHACL이 담당합니다\.

__12\) 한 줄 요약__

온톨로지는 __현장 데이터를 하나의 지식 그래프__로 엮어, “안전·용량·일정·책임” 질문에 즉답하게 합니다\. 한 번 골격을 세워두면, 선적 변경·야간 교대·기상 변수 같은 __변동성__에도 빠르게 재검증·재생성할 수 있습니다\. 엔진은 단순합니다\. \*\*개념\(클래스\)\*\*를 작게, __관계__는 명확하게, __규칙__은 선언적으로\. 그러면 일은 훨씬 덜 복잡해집니다\.

---

# Part 3: Vessel Operations (JPT 71)

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
