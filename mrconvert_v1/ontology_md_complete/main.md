래는 __삼성 C&T 건설물류\(UA E 6현장, 400 TEU/100 BL·월\)__ 업무를 __온톨로지 관점__으로 재정의한 “작동 가능한 설계서”입니다\.  
핵심은 \*\*표준\(UN/CEFACT·WCO DM·DCSA·ICC Incoterms·HS·MOIAT·FANR\)\*\*을 상위 스키마로 삼아 __문서·화물·설비·프로세스·이벤트·계약·규정__을 하나의 그래프\(KG\)로 엮고, 여기서 __Heat‑Stow·WHF/Cap·HSRisk·CostGuard·CertChk·Pre‑Arrival Guard__ 같은 기능을 \*\*제약\(Constraints\)\*\*으로 돌리는 것입니다\. \(Incoterms 2020, HS 2022 최신 적용\)\. [Wcoomd\+4UNECE\+4Wcoomd\+4](https://unece.org/trade/uncefact/rdm?utm_source=chatgpt.com)

__1\) Visual — Ontology Stack \(요약표\)__

__Layer__

__표준/근거__

__범위__

__당신 업무 매핑\(예\)__

__Upper__

__IOF/BFO Supply Chain Ontology__, __ISO 15926__

상위 개념\(행위자/행위/자산/이벤트\)·플랜트 라이프사이클

자산\(크레인, 스키드, 모듈\)·작업\(리깅, 해상 보급\)·상태\(검사/격납\) 정합성 프레임

__Reference Data \(Process/Data\)__

__UN/CEFACT Buy‑Ship‑Pay RDM & CCL__

주문–선적–결제 전과정 공통 데이터·용어

*Party, Shipment, Consignment, Transport Means, Invoice/LineItem* 공통 정의

__Border/Customs__

__WCO Data Model v4\.2\.0__, __HS 2022__

신고/승인/통관 데이터·코드셋

BOE\(수입신고\), 원산지·보증·증명, HS 분류·위험도

__Ocean/Carrier__

__DCSA Booking 2\.0 & eBL 3\.0__

예약/BL 데이터 모델·API

BL 데이터 정규화, eBL 규칙·검증

__Trade Terms__

__ICC Incoterms® 2020__

비용/리스크 이전 지점

EXW/FOB/CIF/DAP별 의무·리스크 노드 매핑

__UAE Reg\.__

__MOIAT ECAS/EQM__, __FANR 수입허가__, __CICPA/ADNOC 출입__

규제/인증/출입 통제

CertChk\(MOIAT·FANR\), 게이트패스 제약, 위험물 통제

__Offshore 계약__

__BIMCO SUPPLYTIME 2017__

OSV 타임차터 KfK 책임체계

보트/바지선 운영 KPI·책임 분기 조건

Hint: Abu Dhabi는 역사적으로 __CICPA/구 CNIA 보안패스__ 체계가 근간이며, 항만 __e‑pass__ 디지털화가 병행되었습니다\(현장 Gate 규정은 매년 공지 확인 필요\)\. [HLB Abudhabi\+1](https://hlbabudhabi.com/a-comprehensive-guide-on-cicpa-passes-in-abu-dhabi/?utm_source=chatgpt.com)

__2\) Domain Ontology — 클래스/관계\(업무 단위 재정의\)__

__핵심 클래스 \(Classes\)__

- __Party__\(Shipper/Consignee/Carrier/3PL/Authority\)
- __Asset__\(Container ISO 6346, OOG 모듈, 장비/스프레더, OSV/바지선\)
- __Document__\(CIPL, Invoice, BL/eBL, BOE, DO, INS, MS\(Method Statement\), Port Permit, Cert\[ECAS/EQM/FANR\], SUPPLYTIME17\)
- __Process__\(Booking, Pre‑alert, Export/Import Clearance, Berth/Port Call, Stowage, Gate Pass, Last‑mile, WH In/Out, Returns\)
- __Event__\(ETA/ATA, CY In/Out, Berth Start/End, DG Inspection, Weather Alert, FANR Permit Granted, MOIAT CoC Issued\)
- __Contract__\(IncotermTerm, SUPPLYTIME17\)
- __Regulation__\(HS Rule, MOIAT TR, FANR Reg\.\)
- __Location__\(UN/LOCODE, Berth, Laydown Yard, Site Gate\)
- __KPI__\(DEM/DET Clock, Port Dwell, WH Util, Delivery OTIF, Damage Rate, Cert SLA\)

__대표 관계 \(Object Properties\)__

- Shipment → hasIncoterm → IncotermTerm \(리스크/비용 이전 노드\) [ICC \- International Chamber of Commerce](https://iccwbo.org/business-solutions/incoterms-rules/?utm_source=chatgpt.com)
- InvoiceLineItem → classifiedBy → HSCode \(HS 2022\) [Wcoomd](https://www.wcoomd.org/en/topics/nomenclature/instrument-and-tools/hs-nomenclature-2022-edition/hs-nomenclature-2022-edition.aspx?utm_source=chatgpt.com)
- BL → conformsTo → DCSA\_eBL\_3\_0 \(데이터 검증 규칙\) [dcsa\.org](https://dcsa.org/newsroom/final-versions-of-booking-bill-of-lading-standards-released?utm_source=chatgpt.com)
- CustomsDeclaration\(BOE\) → usesDataModel → WCO\_DM\_4\_2\_0 \(전자신고 필드 정합\) [Wcoomd](https://www.wcoomd.org/en/media/newsroom/2025/july/world-customs-organization-releases-data-mode.aspx?utm_source=chatgpt.com)
- Equipment/OOG → requiresCertificate → MOIAT\_ECAS|EQM \(규제 제품\) [Ministry of Industry\+1](https://moiat.gov.ae/en/services/issue-conformity-certificates-for-regulated-products/?utm_source=chatgpt.com)
- Radioactive\_Source|Gauge → requiresPermit → FANR\_ImportPermit \(60일 유효\) [Fanr](https://www.fanr.gov.ae/en/services/import-and-export-permit/issue-import-permit-for-radiation-sources-and-nuclear-materials?utm_source=chatgpt.com)
- PortAccess → governedBy → CICPA\_Policy \(게이트패스\) [HLB Abudhabi](https://hlbabudhabi.com/a-comprehensive-guide-on-cicpa-passes-in-abu-dhabi/?utm_source=chatgpt.com)
- OSV\_Charter → governedBy → SUPPLYTIME2017 \(KfK 책임\) [BIMCO](https://www.bimco.org/contractual-affairs/bimco-contracts/contracts/supplytime-2017/?utm_source=chatgpt.com)

__데이터 속성 \(Data Properties\)__

- grossMass, dims\(L×W×H\), isOOG\(boolean\), dgClass, UNNumber, tempTolerance, stowHeatIndex, demClockStartAt, detClockStartAt, gatePassExpiryAt, permitId, costCenter, tariffRef\.

__3\) Use‑case별 제약\(Constraints\) = 운영 가드레일__

__3\.1 CIPL·BL Pre‑Arrival Guard \(eBL‑first\)__

- __Rule‑1__: BL 존재 → BL\.conformsTo = DCSA\_eBL\_3\_0 AND Party·Consignment·PlaceOfReceipt/Delivery 필수\. 미충족 시 *Berth Slot* 확정 금지\. [dcsa\.org](https://dcsa.org/newsroom/final-versions-of-booking-bill-of-lading-standards-released?utm_source=chatgpt.com)
- __Rule‑2__: 모든 InvoiceLineItem는 HSCode 필수 \+ OriginCountry·Qty/UM·FOB/CI 금액\. __WCO DM 필드__ 매핑 누락 시 __BOE 초안 생성 차단__\. [Wcoomd](https://www.wcoomd.org/en/media/newsroom/2025/july/world-customs-organization-releases-data-mode.aspx?utm_source=chatgpt.com)
- __Rule‑3__: IncotermTerm별 책임/비용 그래프 확인\(예: __DAP__면 현지 내륙운송·통관 리스크=Buyer\)\. [ICC \- International Chamber of Commerce](https://iccwbo.org/business-solutions/incoterms-rules/?utm_source=chatgpt.com)

__3\.2 Heat‑Stow \(고온 노출 최소화\)__

- stowHeatIndex = f\(DeckPos, ContainerTier, WeatherForecast\) → 임계치 초과 시 __Under‑deck/센터 베이__ 유도, __berth 시간대 조정__\. \(기상 이벤트는 Event로 연결\)
- dgClass ∈ \{1,2\.1,3,4\.1,5\.1,8\} → Heat‑Stow 규칙 엄격 적용\(위치·분리거리\)\.

__3\.3 WHF/Cap \(Warehouse Forecast/Capacity\)__

- InboundPlan\(TEU/주\)·Outplan → WHUtil\(%\) 예측, 임계치\(85\.00%\) 초과 시 *overflow yard* 예약, __DET 발생 예측__과 연결\.

__3\.4 HSRisk__

- RiskScore = g\(HS, Origin, DG, Cert 요구, 과거검사빈도\) → __검사·추징·지연 확률__ 추정\. \(HS·규제요건: HS 2022·MOIAT·FANR 근거\) [Wcoomd\+2Ministry of Industry\+2](https://www.wcoomd.org/en/topics/nomenclature/instrument-and-tools/hs-nomenclature-2022-edition/hs-nomenclature-2022-edition.aspx?utm_source=chatgpt.com)

__3\.5 CertChk \(MOIAT·FANR\)__

- 규제제품 → ECAS/EQM 승인서 필수 없으면 __DO·GatePass 발행 금지__, __선하증권 인도 보류__\. [Ministry of Industry\+1](https://moiat.gov.ae/en/services/issue-conformity-certificates-for-regulated-products/?utm_source=chatgpt.com)
- 방사선 관련 기자재 → FANR Import Permit\(유효 60일\) 없으면 __BOE 제출 중단__\. [Fanr](https://www.fanr.gov.ae/en/services/import-and-export-permit/issue-import-permit-for-radiation-sources-and-nuclear-materials?utm_source=chatgpt.com)

__4\) 최소 예시\(표현\) — JSON‑LD \(요지\)__

\{

  "@context": \{"incoterm":"https://iccwbo\.org/incoterms/2020\#","dcsa":"https://dcsa\.org/bl/3\.0\#","wco":"https://www\.wcoomd\.org/datamodel/4\.2\#"\},

  "@type":"Shipment",

  "id":"SHP\-ADNOC\-2025\-10\-001",

  "hasIncoterm":\{"@type":"incoterm:DAP","deliveryPlace":"Ruwais Site Gate"\},

  "hasDocument":\[

    \{"@type":"dcsa:BillOfLading","number":"DCSA123\.\.\.", "status":"original\-validated"\},

    \{"@type":"wco:CustomsDeclarationDraft","items":\[\{"hsCode":"850440", "qty":2, "value":120000\.00\}\]\}

  \],

  "consistsOf":\[\{"@type":"Container","isoCode":"45G1","isOOG":true,"dims":\{"l":12\.2,"w":2\.44,"h":2\.90\}\}\]

\}

__5\) 선택지\(3\) — 구축 옵션 \(pro/con/$·risk·time\)__

1. __Reference‑first \(표준 우선, 얇은 구현\)__

- __Pro__: 대외 연계 쉬움\(UN/CEFACT·WCO·DCSA\)\. __Con__: 현장 특성 반영 속도↓\.
- __$__: 초기 낮음\(₩·$$\)\. __Risk__: 커스터마이즈 지연\. __Time__: 6–8주 MVP\. [UNECE\+2Wcoomd\+2](https://unece.org/trade/uncefact/rdm?utm_source=chatgpt.com)

1. __Hybrid \(표준\+현장제약 동시\)__ ← *추천*

- __Pro__: 표준 적합 \+ GatePass/Heat‑Stow/WH 바로 적용\. __Con__: 설계 복잡\.
- __$__: 중간\. __Risk__: 스키마 복잡성\. __Time__: 10–12주 POC→Rollout\.

1. __Ops‑first \(현장 규칙 우선\)__

- __Pro__: 즉효\(DEM/DET·GatePass\)\. __Con__: 표준 정합 나중 기술부채\.
- __$__: 낮음→중간\. __Risk__: 대외 API 통합 시 재작업\. __Time__: 4–6주\.

__6\) Roadmap \(P→Pi→B→O→S \+ KPI\)__

- __P\(Plan\)__: 스코프 확정\(문서: CIPL/BL/BOE/DO/INS/Permit, 프로세스: Berth/Gate Pass/WH\)\. __KPI__: 데이터 필드 완전성 ≥ 98\.00%\.
- __Pi\(Pilot\)__: __eBL‑Pre‑Arrival Guard__ \+ __WHF/Cap__ 1현장 적용\. __KPI__: Port dwell ↓ 12\.50%, DET 비용 ↓ 18\.00% *\(가정\)*\.
- __B\(Build\)__: __HSRisk__·__CertChk__·__CostGuard__ 추가, __SUPPLYTIME17__ 운영지표 연계\. __KPI__: 검사로 인한 Leadtime 분산 ↓ 15\.00%\. [BIMCO](https://www.bimco.org/contractual-affairs/bimco-contracts/contracts/supplytime-2017/?utm_source=chatgpt.com)
- __O\(Operate\)__: 규칙/SHACL 자동검증, Slack/Telegram 알림\. __KPI__: 규칙 위반 건당 처리시간 ≤ 0\.50h\.
- __S\(Scale\)__: 6현장→글로벌 재사용, __UN/CEFACT Web Vocabulary__로 공개 스키마 매핑\. __KPI__: 시스템 간 매핑 공수 ↓ 30\.00%\. [Vocabulary UNCEFACT](https://vocabulary.uncefact.org/about?utm_source=chatgpt.com)

__7\) Data·Sim·BI \(운영 숫자 관점\)__

- __DEM/DET 시계__: ClockStart = \(CY In or FreeTime Start by Carrier\) → 컨테이너별 __DEM/DET Clock__ 노드 운영\.
- __WH Capacity Forecast__: Util\_t\+1 = Util\_t \+ Inbound \- Outbound \(ARIMA/Prophet 가능\)\.
- __Heat‑Stow 점수__: HI = α\*DeckExposure \+ β\*Tier \+ γ\*ForecastTemp\(°C\) → 임계 0\.70 이상 __스택 변경__\.
- __Risk@HS__: 로지스틱 회귀/GBT로 검사확률·추징금 기대값\.

__8\) Automation \(RPA·LLM·Sheets·TG\) — Slash Cmd 예시__

- __/logi\-master \-\-fast invoice\-audit__ → CIPL/Invoice 라인 __HS·Origin·Qty·Value 누락__ 탐지→BOE 초안 블록\. \(WCO DM/HS 2022\) [Wcoomd\+1](https://www.wcoomd.org/en/media/newsroom/2025/july/world-customs-organization-releases-data-mode.aspx?utm_source=chatgpt.com)
- __/logi\-master predict \-\-AEDonly weather\-tie__ → 기상경보 Event→Berth 스케줄 재배치\(Heat‑Stow 임계\)\.
- __/switch\_mode COST\-GUARD LATTICE__ → DET/DEM 예측비용 알림 \+ eBL 상태/도착지연 교차검증\(DCSA eBL 3\.0\)\. [dcsa\.org](https://dcsa.org/newsroom/final-versions-of-booking-bill-of-lading-standards-released?utm_source=chatgpt.com)
- __/visualize\_data \-\-type=heatmap <stow\.csv>__ → HI>0\.70 구간 강조\.

__9\) QA — Gap/Recheck 리스트__

- __eBL 상태 신뢰도__: Carrier별 DCSA 3\.0 호환 여부 점검\. [dcsa\.org](https://dcsa.org/newsroom/final-versions-of-booking-bill-of-lading-standards-released?utm_source=chatgpt.com)
- __HS·CCL 정합성__: UN/CEFACT CCL 릴리스\(예: __24A__\)와 로컬 속성 매핑 재검\. [UNECE](https://unece.org/trade/uncefact/unccl?utm_source=chatgpt.com)
- __UAE 인증__: MOIAT ECAS/EQM 최신 규제 범위/코드 확인, FANR 퍼밋 유효일\(60일\) 자동 만료 체크\. [Ministry of Industry\+2SGSCorp\+2](https://moiat.gov.ae/en/services/issue-conformity-certificates-for-regulated-products/?utm_source=chatgpt.com)
- __GatePass 체계__: 현장 보안 주체\(CICPA/ADNOC\) 최신 공지 확인\(사내 SOP 연결\)\. [HLB Abudhabi](https://hlbabudhabi.com/a-comprehensive-guide-on-cicpa-passes-in-abu-dhabi/?utm_source=chatgpt.com)

__10\) Fail‑safe “중단” 테이블 \(ZERO 전략\)__

__트리거\(중단\)__

__ZERO 액션__

__재개 조건__

eBL 비정합\(DCSA 3\.0 스키마 오류\)

Berth 확정 보류, 선적명세 수동검증

eBL 재검증 Pass

BOE 필수필드 미충족\(WCO DM\)

신고중단, Shipper 보완요청

필드 완전성 ≥ 98\.00%

규제제품 Cert 부재\(MOIAT/FANR\)

DO/게이트패스 발급 금지

유효 CoC/Permit 업로드

DET 임계 초과 예측\(>AED x/TEU\)

내륙 인도·반납 계획 재조정

비용 추정 < 임계값

__11\) 운영에 바로 쓰는 SHACL\(요지\)__

@prefix sh: <http://www\.w3\.org/ns/shacl\#> \.

:InvoiceLineItemShape a sh:NodeShape ;

  sh:targetClass :InvoiceLineItem ;

  sh:property \[

    sh:path :hsCode ; sh:minCount 1 ; sh:pattern "^\[0\-9\]\{6,10\}$"

  \] ;

  sh:property \[

    sh:path :originCountry ; sh:minCount 1

  \] ;

  sh:property \[

    sh:path :quantity ; sh:minInclusive 0\.01

  \] \.

__12\) GitHub·재사용__

- 리포지토리 __macho715/ontology\-insight__에 __/models \(TTL/JSON‑LD\)__, __/rules \(SHACL\)__, __/mappings \(CCL/HS/WCO/DCSA\)__ 디렉토리 구조 권장\.
- OFCO CostCenter 매핑은 __Invoice/LineItem → CostCenter__ 룰로 두고, __/mappings/ofco\.csv__로 관리\. *\(가정: OFCO는 내규 분류표\)*

__13\) Assumptions & Sources__

- __가정:__ OFCO/CostCenter·Heat‑Stow·HSRisk·CostGuard는 사내용 명칭/로직\. UAE 게이트패스\(e‑pass, CICPA/ADNOC\)는 현장별 차이 존재\(사는 공지 우선\)\.
- __표준/근거:__ UN/CEFACT BSP‑RDM & CCL, WCO DM v4\.2\.0, HS 2022, DCSA eBL 3\.0/Booking 2\.0, ICC Incoterms 2020, MOIAT ECAS/EQM, FANR Import Permit, SUPPLYTIME 2017\. [BIMCO\+10UNECE\+10UNECE\+10](https://unece.org/trade/uncefact/rdm?utm_source=chatgpt.com)

__14\) 다음 액션\(짧게\)__

- __/logi\-master \-\-fast invoice\-audit__ 로 현재 월간 100 BL 대상 __Pre‑Arrival Guard__ 돌리고,
- __/switch\_mode COST\-GUARD__ 로 __DEM/DET__ 시계 추정부터 막읍시다\(목표 ↓ 15\.00%\)\.
- __/wh\-capacity oracle \-\-horizon 12m__ 로 __WH Util__ 예측 배치\.

원하시면, 위 스택으로 __CIPL·BL Pre‑Arrival Guard__와 \*\*CertChk\(MOIAT·FANR\)\*\*부터 SHACL/룰팩을 묶어 드리겠습니다\.

