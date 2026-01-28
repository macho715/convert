---
title: "HVDC Node Infrastructure Ontology - Unified Framework"
type: "ontology-design"
domain: "hvdc-node-infrastructure"
sub-domains: ["port-operations", "node-logistics", "transport-network", "cargo-management"]
version: "unified-3.0"
date: "2025-10-25"
tags: ["ontology", "hvdc", "node-network", "samsung-ct", "adnoc", "mosb", "transport", "container", "bulk", "heavy-cargo"]
standards: ["UN/LOCODE", "CICPA", "DOT-UAE", "ADNOC-L&S", "BIMCO-SUPPLYTIME", "ISO-6346", "Hitachi-Preservation", "ADOPT", "Free-Zone"]
status: "active"
source: "HVDC Material Handling Workshop 2024-11-13"
verified_facts: "All cargo types (Container/Bulk/Heavy) - 8 nodes network"
---

# hvdc-infra-nodes · 1_CORE-02

아래는 __HVDC 프로젝트 물류 노드 네트워크(UAE 8거점)__를 __온톨로지 관점__으로 정의한 "작동 가능한 설계서"입니다.
핵심은 __Port(입항)·Hub(집하)·Site(수령/설치)__ 를 하나의 그래프(KG)로 엮고, __컨테이너·벌크·중량화물 전반__을 포함한 __DOT 허가·LCT 운항·MOSB 중심 체계·보존조건__ 같은 제약을 **Constraints**로 운영하는 것입니다.

__1) Visual — Ontology Stack (요약표)__

| __Layer__                         | __표준/근거__                                    | __범위__                                       | __HVDC 업무 매핑(예)__                                        |
| --------------------------------- | ------------------------------------------------ | ---------------------------------------------- | ------------------------------------------------------------- |
| __Upper__                         | __IOF/BFO Supply Chain Ontology__, __ISO 15926__ | 상위 개념(행위자/행위/자산/이벤트)·플랜트 라이프사이클 | 노드(Port/Hub/Site)·행위(Transport/Storage)·상태(MRR/OSDR) 프레임 |
| __Reference Data (Location)__     | __UN/LOCODE__, __ISO 3166__                      | 항만·지역 코드 표준화                          | Zayed(AEZYD), Mugharaq, MOSB(Mussafah), Site 좌표             |
| __Transport/Marine__              | __BIMCO SUPPLYTIME 2017__, __ISO 6346__          | OSV/LCT 운항, Container 코드                   | LCT 운항(MOSB→DAS 20h, →AGI 10h), Roll-on/off                |
| __Heavy Transport__               | __DOT UAE Permit System__                        | 중량물(>90톤) 육상 운송 허가                   | MIR/SHU 트랜스포머 SPMT 이송, DOT 승인 필수                   |
| __Port Access Control__           | __CICPA/ADNOC Gate Pass__                        | 항만·현장 출입 통제                            | MOSB/Port 게이트패스, ALS 운영 규정                           |
| __Preservation Standards__        | __Hitachi Specification__, __IEC__               | 보존 환경 조건                                 | Dry air/N₂ 충전, +5~40°C, RH ≤85%, 습도 모니터링            |
| __Quality Control__               | __MRR/OSDR/MIS Standards__                       | 자재 검수·상태 리포팅                          | 수령 검수(MRR), 해상 상태(OSDR), 설치 전 검증(MIS)            |
| __Offshore Operations__           | __ADNOC L&S (ALS) Regulations__                  | 해상 작업·리프팅·안전                          | DAS/AGI 하역, Sea fastening, 기상 제약                        |

Hint: MOSB는 **ADNOC Logistics & Services (ALS)** 관할 Yard(20,000㎡)이며, **삼성물산(SCT) 물류본부**가 상주하는 실질적 중앙 노드입니다.

__2) Domain Ontology — 클래스/관계(노드 단위 재정의)__

__핵심 클래스 (Classes)__

- __Node__(Port/Hub/OnshoreS ite/OffshoreSite)
- __Party__(SCT/JDN/ALS/ADNOC/Vendor/Subcon)
- __Asset__(Transformer/Cable/CCU/Module/Container/Bulk_Cargo/Heavy_Cargo/General_Materials)
- __TransportEvent__(노드 간 이동 및 상태 변경 이벤트)
- __Warehouse__(IndoorWarehouse/OutdoorWarehouse/DangerousCargoWarehouse)
- __Transport__(InlandTruck/SPMT/LCT/Vessel)
- __Document__(CI/PL/BL/COO/eDAS/MRR/OSDR/MIS/DOT_Permit/FRA/PTW)
- __Process__(Import_Clearance/Yard_Storage/Preservation/Inland_Transport/Marine_Transport/Site_Receiving/Installation)
- __Event__(ETA/ATA/Berth_Start/Berth_End/CY_In/CY_Out/LCT_Departure/LCT_Arrival/MRR_Issued/OSDR_Updated)
- __Permit__(DOT_Heavy_Transport/FANR_Import/MOIAT_CoC/CICPA_GatePass/FRA/PTW)
- __Location__(UN/LOCODE: AEZYD/AEMFA, Berth, Laydown_Yard, Site_Gate)
- __Regulation__(Customs_Code/DOT_Rule/ADNOC_Policy/Hitachi_Preservation_Spec)
- __KPI__(Port_Dwell/Transit_Time/Storage_Duration/MRR_SLA/OSDR_Timeliness/Delivery_OTIF)

__대표 관계 (Object Properties)__

- Node → connectedTo → Node (물류 연결성)
- MOSB → centralHubFor → (SHU, MIR, DAS, AGI) (중앙 허브 역할)
- Port → importsFrom → Origin_Country (수입 출발지)
- Transformer → transportedBy → LCT/SPMT (운송 수단)
- Cargo → storedAt → Node (보관 위치)
- Transport → requiresPermit → DOT_Permit/FRA (허가 요구)
- Site → receivesFrom → MOSB (수령 관계)
- Asset → hasDocument → MRR/OSDR (검수 문서)
- LCT_Operation → operatedBy → ALS (운영 주체)
- Node → governedBy → ADNOC_Policy/CICPA_Rule (규정 적용)
- Asset → preservedBy → Hitachi_Spec (보존 기준)

__데이터 속성 (Data Properties)__

- grossMass, dims(L×W×H), laydownArea_sqm, transitTime_hours, storageCapacity_teu, gatePassExpiryAt, permitId, preservationTemp_min, preservationTemp_max, relativeHumidity_max, dryAirPressure_bar, n2ChargePressure_bar, lctVoyageDuration_hours, distanceFromMOSB_nm, dotPermitRequired(boolean), customsCode, operatingOrg, sctTeamLocation, hasLogisticsFlowCode, hasWHHandling.

__3) Use-case별 제약(Constraints) = 운영 가드레일__

__3.1 Port Import & Clearance Guard__

- __Rule-1__: Port(Zayed/Mugharaq) → hasDocument(CI, PL, BL, COO) 필수. 미충족 시 *Customs Clearance 차단*.
- __Rule-2__: 통관 코드 검증: ADNOC(47150) for Abu Dhabi, ADOPT(1485718/89901) for Dubai/Free Zone. 미일치 시 *BOE 제출 거부*.
- __Rule-3__: 방사선 기자재 → FANR Import Permit(유효 60일) 필수. 없으면 *입항 승인 보류*.

__3.2 MOSB Central Hub Operations__

- __Rule-4__: 모든 자재는 MOSB를 경유. MOSB → consolidates → Cargo_from_Ports AND MOSB → dispatches → (SHU/MIR/DAS/AGI).
- __Rule-5__: Yard 용량 체크: MOSB.storageCapacity(20,000㎡) > CurrentUtilization. 초과 시 *overflow yard* 확보 또는 *출하 스케줄 조정*.
- __Rule-6__: 보존 조건: Indoor storage, Temp(+5~40°C), RH(≤85%). 미준수 시 *자재 손상 리스크 알림* + *재검수(MRR) 필수*.

__3.3 Heavy Inland Transport (DOT Permit)__

- __Rule-7__: Cargo.grossMass > 90_ton → DOT_Permit 필수. 없으면 *MIR/SHU 이송 금지*.
- __Rule-8__: SPMT 이송 시 routeApproval + escortVehicle 필수. 미확보 시 *이송 연기*.
- __Rule-9__: Laydown area capacity: SHU(10,556㎡), MIR(35,006㎡). 용량 초과 시 *site receiving schedule 재조정*.

__3.4 Marine Transport (LCT Operations)__

- __Rule-10__: LCT_Operation → operatedBy → ALS (ADNOC L&S 전담). 비승인 선박 *출항 금지*.
- __Rule-11__: 항로 및 소요시간: MOSB→DAS(≈20h), MOSB→AGI(≈10h). 기상 경보 시 *출항 연기* (Weather-Tie 규칙).
- __Rule-12__: Roll-on/off, Sea fastening 필수. 검증 미완료 시 *선적 중단*.
- __Rule-13__: 보존 조건 유지: Dry air/N₂ 충전 상태 체크. 압력 이탈 시 *즉시 재충전* + *OSDR 업데이트*.

__3.5 Site Receiving & Quality Control__

- __Rule-14__: 자재 수령 시 MRR(Material Receiving Report) 즉시 발행. 미발행 시 *납품 미완료 처리*.
- __Rule-15__: 해상 현장(DAS/AGI) → OSDR(Offshore Storage & Delivery Report) 주기적 업데이트. 지연 시 *상태 불명확 경고*.
- __Rule-16__: 설치 전 MIS(Material Installation Sheet) 최종 검증. 미통과 시 *설치 작업 보류*.

__3.6 Logistics Flow Code System__

- __Rule-17__: 모든 화물은 Flow Code(0~4) 부여 필수.
  - **0**: Pre Arrival (Planning → Port)
  - **1**: Direct Port→Site
  - **2**: Port→WH→Site
  - **3**: Port→WH→MOSB→Site
  - **4**: Port→WH→WH→MOSB→Site
- __Rule-18__: WH Handling Count = 경유 창고 횟수(0~3). Flow Code와 일치 필수.
- __Rule-19__: 비표준 Flow Code(예: 6) 감지 시 *자동 정규화* 또는 *데이터 검증 실패*.

__4) 최소 예시(표현) — JSON-LD (요지)__

```json
{
  "@context": {
    "hvdc": "https://hvdc-project.ae/ontology#",
    "geo": "http://www.w3.org/2003/01/geo/wgs84_pos#",
    "time": "http://www.w3.org/2006/time#"
  },
  "@type": "hvdc:LogisticsFlow",
  "id": "HVDC-FLOW-2025-10-001",
  "origin": {
    "@type": "hvdc:Port",
    "name": "Zayed Port",
    "locode": "AEZYD",
    "customsCode": "47150",
    "location": "Abu Dhabi"
  },
  "centralHub": {
    "@type": "hvdc:Hub",
    "name": "MOSB",
    "operatedBy": "ADNOC L&S",
    "sctTeamLocation": true,
    "storageCapacity_sqm": 20000,
    "role": "Central consolidation and dispatch hub"
  },
  "destinations": [
    {
      "@type": "hvdc:OnshoreS ite",
      "name": "SHUWEIHAT (SHU)",
      "laydownArea_sqm": 10556,
      "receivesFrom": "Sweden",
      "transportMode": "Inland_SPMT",
      "requiresDOT": true
    },
    {
      "@type": "hvdc:OnshoreS ite",
      "name": "MIRFA (MIR)",
      "laydownArea_sqm": 35006,
      "receivesFrom": "Brazil",
      "transportMode": "Inland_SPMT",
      "requiresDOT": true
    },
    {
      "@type": "hvdc:OffshoreSite",
      "name": "DAS Island",
      "cluster": "Zakum",
      "transportMode": "LCT",
      "voyageDuration_hours": 20,
      "preservationMethod": "Dry_air_N2"
    },
    {
      "@type": "hvdc:OffshoreSite",
      "name": "Al Ghallan Island (AGI)",
      "cluster": "Zakum",
      "transportMode": "LCT",
      "voyageDuration_hours": 10,
      "parallelTo": "DAS"
    }
  ],
  "hasDocument": [
    {"@type": "hvdc:CI", "status": "validated"},
    {"@type": "hvdc:PL", "status": "validated"},
    {"@type": "hvdc:BL", "status": "original"},
    {"@type": "hvdc:COO", "origin": "Brazil/Sweden"}
  ],
  "consistsOf": [
    {
      "@type": "hvdc:Transformer",
      "origin": "Brazil",
      "grossMass_ton": 120,
      "dims": {"l": 12.5, "w": 3.2, "h": 4.8},
      "requiresDOT": true,
      "preservationTemp": {"min": 5, "max": 40},
      "preservationRH_max": 85,
      "hasLogisticsFlowCode": 3,
      "hasWHHandling": 1
    }
  ],
  "hasTransportEvent": [
    {
      "@type": "hvdc:TransportEvent",
      "hasCase": "HE-208221",
      "hasDate": "2025-05-13T08:00:00",
      "hasLocation": "DSV Indoor",
      "hasLogisticsFlowCode": 3,
      "hasWHHandling": 1
    }
  ]
}
```

__5) 선택지(3) — 구축 옵션 (pro/con/$·risk·time)__

1. __Reference-first (표준 우선, 글로벌 호환)__

- __Pro__: UN/LOCODE·BIMCO·ISO 표준 즉시 적용, 대외 연계 용이.
- __Con__: HVDC 특화 제약(DOT/CICPA/ALS 규정) 반영 속도↓.
- __$__: 초기 낮음(₩·$$). __Risk__: 현장 커스터마이즈 지연. __Time__: 8–10주 MVP.

2. __Hybrid (표준+현장제약 동시)__ ← *추천*

- __Pro__: UN/LOCODE + MOSB 중심 체계 + DOT/LCT/보존 규칙 즉시 적용.
- __Con__: 스키마 복잡성↑.
- __$__: 중간. __Risk__: 초기 설계 공수. __Time__: 12–14주 POC→Rollout.

3. __Ops-first (현장 규칙 우선)__

- __Pro__: MOSB 운영·DOT 허가·LCT 스케줄 즉효.
- __Con__: 표준 정합 나중 기술부채.
- __$__: 낮음→중간. __Risk__: 글로벌 확장 시 재작업. __Time__: 6–8주.

__6) Roadmap (P→Pi→B→O→S + KPI)__

- __P(Plan)__: 스코프 확정(노드: 7개, 문서: CI/PL/BL/MRR/OSDR, 프로세스: Import/Storage/Transport/Receiving). __KPI__: 노드 정의 완전성 ≥ 100%.
- __Pi(Pilot)__: __MOSB Central Hub__ + __DOT Permit Guard__ 1현장 적용. __KPI__: Transit time ↓ 15%, DOT 지연 건수 ↓ 25%.
- __B(Build)__: __LCT Operations__ + __Preservation Monitoring__ + __MRR/OSDR 자동화__ 추가. __KPI__: 보존 이탈 건수 ↓ 30%, MRR SLA ≥ 95%.
- __O(Operate)__: 규칙/SHACL 자동검증, Slack/Telegram 알림, KPI 대시보드. __KPI__: 규칙 위반 건당 처리시간 ≤ 0.5h.
- __S(Scale)__: 7거점→글로벌 재사용, __UN/LOCODE Web Vocabulary__로 공개 스키마 매핑. __KPI__: 타 프로젝트 적용 공수 ↓ 40%.

__7) Data·Sim·BI (운영 숫자 관점)__

- __Transit Time Clock__: TransitStart = (Port CY Out or MOSB Dispatch) → 노드별 __Transit Clock__ 운영.
- __MOSB Capacity Forecast__: Util_t+1 = Util_t + Inbound - Outbound (ARIMA/Prophet 가능).
- __DOT Permit Lead Time__: 평균 승인 기간 추적, 지연 시 *대안 경로* 제시.
- __LCT Voyage Risk__: Weather score + Cargo weight + Voyage distance → 출항 적합성 판정.
- __Preservation Compliance__: Temp/RH 센서 데이터 실시간 수집 → 이탈 시 *자동 알림*.

__8) Automation (RPA·LLM·Sheets·TG) — Slash Cmd 예시__

- __/logi-master --fast node-audit__ → 7개 노드별 __CI/PL/BL/MRR 누락__ 탐지→import 차단.
- __/logi-master predict --AEDonly transit-time__ → MOSB→Site 경로별 예상 소요시간 + DOT 지연 반영.
- __/switch_mode LATTICE RHYTHM__ → MOSB 용량 알림 + LCT 스케줄 교차검증.
- __/visualize_data --type=network <nodes.csv>__ → 7-노드 관계망 시각화(방사형).
- __/weather-tie check --port=MOSB__ → 기상 경보→LCT 출항 연기 여부 판단.
- __/compliance-check DOT-permit__ → 중량물(>90톤) 대상 DOT 승인 상태 일괄 체크.

__9) QA — Gap/Recheck 리스트__

- __UN/LOCODE 정합성__: Zayed(AEZYD), Mugharaq 코드 재확인.
- __DOT 규정__: 90톤 임계값, 승인 절차, escortVehicle 요구사항 최신화.
- __ALS 운영 규정__: MOSB Yard 규칙, LCT 출항 승인 프로세스 변경 추적.
- __CICPA/GatePass__: 최신 출입 통제 정책, e-pass 디지털화 상태 확인.
- __Hitachi Preservation Spec__: 온습도 기준, Dry air/N₂ 충전 압력, 모니터링 주기 재검.
- __MRR/OSDR/MIS 양식__: 최신 템플릿 및 필수 필드 매핑 점검.

__10) Fail-safe "중단" 테이블 (ZERO 전략)__

| __트리거(중단)__                           | __ZERO 액션__                              | __재개 조건__                         |
| ------------------------------------------ | ------------------------------------------ | ------------------------------------- |
| CI/PL/BL/COO 미충족                        | Customs clearance 보류, Shipper 보완요청   | 필수 문서 완전성 ≥ 100%               |
| 통관코드 불일치(ADNOC/ADOPT)               | BOE 제출 중단, 코드 재확인                 | 올바른 코드 적용 확인                 |
| FANR Permit 부재(방사선 기자재)            | 입항 승인 보류, Vendor 퍼밋 요청           | 유효 FANR Permit 업로드(60일 이내)    |
| MOSB 용량 초과(>20,000㎡)                  | 추가 입고 중단, overflow yard 확보         | 용량 < 임계값 or 출하 완료            |
| 보존 조건 이탈(Temp/RH)                    | 자재 격리, 재검수(MRR) 필수                | 환경 조건 복구 + MRR Pass             |
| DOT Permit 부재(>90톤)                     | 내륙 이송 금지, DOT 승인 대기              | 유효 DOT Permit 발급                  |
| 기상 경보(LCT 출항 부적합)                 | LCT 출항 연기, 기상 재평가                 | Weather score < 임계값                |
| Sea fastening 검증 미완료                  | 선적 중단, 고박 재작업                     | Sea fastening 검증 Pass               |
| Dry air/N₂ 압력 이탈                       | 해상 운송 중단, 즉시 재충전 + OSDR 업데이트 | 보존 압력 정상 범위 복구              |
| MRR 미발행(자재 수령 후 24h 초과)          | 납품 미완료 처리, Site 검수팀 긴급 투입    | MRR 발행 + 승인                       |
| OSDR 업데이트 지연(해상 현장 >7일)         | 상태 불명확 경고, 현장 긴급 점검           | OSDR 최신화 + 보존 상태 확인          |
| MIS 최종 검증 미통과                       | 설치 작업 보류, QAQC 재검증                | MIS Pass + OE(Owner's Engineer) 승인 |

__11) 운영에 바로 쓰는 SHACL(요지)__

```turtle
@prefix sh: <http://www.w3.org/ns/shacl#> .
@prefix hvdc: <https://hvdc-project.ae/ontology#> .

hvdc:PortNodeShape a sh:NodeShape ;
  sh:targetClass hvdc:Port ;
  sh:property [
    sh:path hvdc:hasDocument ;
    sh:minCount 4 ;  # CI, PL, BL, COO 필수
    sh:message "Port must have CI, PL, BL, COO documents"
  ] ;
  sh:property [
    sh:path hvdc:customsCode ;
    sh:minCount 1 ;
    sh:pattern "^(47150|1485718|89901)$" ;
    sh:message "Invalid customs code for UAE"
  ] .

hvdc:HeavyCargoShape a sh:NodeShape ;
  sh:targetClass hvdc:Transformer ;
  sh:property [
    sh:path hvdc:grossMass_ton ;
    sh:minInclusive 0.01
  ] ;
  sh:sparql [
    sh:message "Cargo >90 ton requires DOT Permit" ;
    sh:select """
      SELECT $this
      WHERE {
        $this hvdc:grossMass_ton ?mass .
        FILTER (?mass > 90)
        FILTER NOT EXISTS { $this hvdc:requiresPermit ?permit .
                           ?permit a hvdc:DOT_Permit }
      }
    """
  ] .

hvdc:MOSBCapacityShape a sh:NodeShape ;
  sh:targetClass hvdc:MOSB ;
  sh:property [
    sh:path hvdc:storageCapacity_sqm ;
    sh:hasValue 20000
  ] ;
  sh:sparql [
    sh:message "MOSB storage capacity exceeded" ;
    sh:select """
      SELECT $this
      WHERE {
        $this hvdc:currentUtilization_sqm ?util .
        $this hvdc:storageCapacity_sqm ?cap .
        FILTER (?util > ?cap)
      }
    """
  ] .

hvdc:PreservationShape a sh:NodeShape ;
  sh:targetClass hvdc:Asset ;
  sh:property [
    sh:path hvdc:preservationTemp_min ;
    sh:hasValue 5
  ] ;
  sh:property [
    sh:path hvdc:preservationTemp_max ;
    sh:hasValue 40
  ] ;
  sh:property [
    sh:path hvdc:preservationRH_max ;
    sh:maxInclusive 85
  ] .

# Flow Code 검증 규칙
hvdc:FlowCodeShape a sh:NodeShape ;
  sh:targetClass hvdc:Asset ;
  sh:property [
    sh:path hvdc:hasLogisticsFlowCode ;
    sh:datatype xsd:integer ;
    sh:minInclusive 0 ;
    sh:maxInclusive 4 ;
    sh:message "Flow Code must be 0-4"
  ] ;
  sh:property [
    sh:path hvdc:hasWHHandling ;
    sh:datatype xsd:integer ;
    sh:minInclusive 0 ;
    sh:maxInclusive 3 ;
    sh:message "WH Handling must be 0-3"
  ] .

# Flow Code와 WH Handling 일치성 검증
hvdc:FlowCodeConsistencyShape a sh:NodeShape ;
  sh:targetClass hvdc:Asset ;
  sh:sparql [
    sh:message "WH Handling count must match Flow Code" ;
    sh:select """
      SELECT $this
      WHERE {
        $this hvdc:hasLogisticsFlowCode ?flow .
        $this hvdc:hasWHHandling ?wh .
        FILTER (
          (?flow = 0 && ?wh != 0) ||
          (?flow = 1 && ?wh != 0) ||
          (?flow = 2 && ?wh != 1) ||
          (?flow = 3 && (?wh < 1 || ?wh > 2)) ||
          (?flow = 4 && (?wh < 2 || ?wh > 3))
        )
      }
    """
  ] .
```

__12) GitHub·재사용__

- 리포지토리 __macho715/hvdc-node-ontology__에 __/models (TTL/JSON-LD)__, __/rules (SHACL)__, __/mappings (UN-LOCODE/CICPA/DOT)__ 디렉토리 구조 권장.
- MOSB 중심 흐름은 __Node → centralHubFor → Site__ 룰로 두고, __/mappings/mosb-dispatch.csv__로 관리.
- LCT 운항 스케줄은 __/data/lct-operations.json__으로 버전 관리.

__13) Assumptions & Sources__

- __가정:__ MOSB는 모든 자재의 필수 경유지. DOT 90톤 임계값은 UAE 법규 기준. ALS 운영 규정은 ADNOC L&S 내부 정책 따름. CICPA/e-pass는 현장별 차이 존재(현장 공지 우선).
- __표준/근거:__ UN/LOCODE, BIMCO SUPPLYTIME 2017, ISO 6346(Container), DOT UAE Heavy Transport Regulation, CICPA/ADNOC Gate Pass Policy, Hitachi Preservation Specification, IEC Standards, HVDC Material Handling Workshop 2024-11-13.

__14) 다음 액션(짧게)__

- __/logi-master --fast node-audit__ 로 7개 노드 대상 __필수 문서·허가__ 일괄 점검,
- __/switch_mode LATTICE__ 로 __MOSB 용량__ 및 __DOT 지연__ 모니터링 시작,
- __/visualize_data --type=network <hvdc-nodes.csv>__ 로 __노드 관계망__ 시각화.

원하시면, 위 스택으로 __Port Import Guard__와 __MOSB Central Hub Operations__부터 SHACL/룰팩을 묶어 드리겠습니다.

---

# Part 2: HVDC Node Lifecycle Framework

## 개요

HVDC 프로젝트의 7개 물류 노드를 **온톨로지 관점**에서 정리하면, '물류 생명주기'를 하나의 **지식그래프(Ontology)**로 모델링할 수 있습니다.

핵심은 **"노드 간 행위(Activity)"가 아닌 "관계(Relation)"** 중심으로 보는 것입니다 — Port, Hub, Site, Actor, Document, Permit 간의 연결망.

__🔶 1. Ontology Root Class__

**hvdc-node-ontology**

| __Layer__ | __Ontology Domain__ | __대표 엔티티__                        | __관계 키(Relation)__                                |
| --------- | ------------------- | -------------------------------------- | ---------------------------------------------------- |
| __L1__    | Physical Flow       | Cargo, Port, MOSB, Site, LCT, SPMT    | movesFrom, movesTo, storedAt, consolidatedAt         |
| __L2__    | Document Flow       | CI, PL, BL, COO, eDAS, MRR, OSDR, MIS | certifies, refersTo, attachedTo, validates           |
| __L3__    | Actor Flow          | SCT, JDN, ALS, ADNOC, Vendor, Subcon  | responsibleFor, operates, approves, reportsTo        |
| __L4__    | Regulatory Flow     | DOT, FANR, MOIAT, CICPA, Customs      | requiresPermit, compliesWith, auditedBy, governedBy  |
| __L5__    | System Flow         | eDAS, SAP, NCM, LDG, KPI Dashboard    | feedsDataTo, validates, monitoredBy, alertsOn        |

__🔶 2. Core Classes (from Workshop + Verified Facts)__

| __Class__               | __Subclass of__ | __Description__                                              | __Onto-ID__       |
| ----------------------- | --------------- | ------------------------------------------------------------ | ----------------- |
| __Node__                | Location        | 물류 거점(Port/Hub/OnshoreS ite/OffshoreSite)                | hvdc-loc-node     |
| __Cargo__               | Asset           | 자재 및 기자재(Transformer, Cable, CCU, Module)              | hvdc-asset-cargo  |
| __TransportEvent__      | Activity        | Inland(SPMT), Marine(LCT), Offloading, Receiving             | hvdc-act-trans    |
| __Storage__             | Process         | Yard Storage, Preservation(Dry air/N₂), Laydown              | hvdc-proc-stor    |
| __Inspection__          | Process         | MRR(Material Receiving), OSDR(Offshore Status), MIS(Install) | hvdc-proc-insp    |
| __Permit__              | Document        | DOT Heavy Transport, FANR Import, CICPA GatePass, FRA, PTW   | hvdc-doc-perm     |
| __Actor__               | Agent           | SCT Logistics Team, ADNOC L&S, Vendor, Subcon                | hvdc-agent-role   |
| __PortOperation__       | Activity        | Import Clearance, CY In/Out, Customs BOE                     | hvdc-act-port     |
| __PreservationStandard__ | Specification   | Hitachi Spec(Temp/RH), Dry air/N₂ Charging                   | hvdc-spec-presrv  |

__🔶 3. Relation Model (Partial)__

```turtle
Cargo --hasDocument--> MRR
Cargo --transportedBy--> TransportEvent
TransportEvent --departsFrom--> MOSB
TransportEvent --arrivesAt--> Site
TransportEvent --requires--> DOT_Permit
DOT_Permit --approvedBy--> DOT_Authority
Storage --locatedAt--> MOSB
Storage --monitoredBy--> SCT_Team
Inspection --reportedAs--> MRR/OSDR/MIS
Actor(SCT) --usesSystem--> eDAS
LCT_Operation --operatedBy--> ALS
Site --receivesFrom--> MOSB
MOSB --consolidates--> Cargo_from_Ports
Port(Zayed) --importsFrom--> Brazil
Port(Mugharaq) --importsFrom--> Sweden
```

이 관계망은 `hvdc-node-ontology.ttl`로 구현 가능:

```turtle
:MOSB rdf:type :Hub ;
      :hosts :SCT_Logistics_Team ;
      :operatedBy :ALS ;
      :storageCapacity_sqm 20000 ;
      :consolidates :Cargo_from_Zayed, :Cargo_from_Mugharaq ;
      :dispatches :SHU, :MIR, :DAS, :AGI .

:TR_001 rdf:type :Transformer ;
        :origin "Brazil" ;
        :grossMass_ton 120 ;
        :hasDocument :MRR_20241113 ;
        :storedAt :MOSB ;
        :transportedBy :SPMT_Operation_20241120 ;
        :requiresPermit :DOT_Permit_20241115 ;
        :preservedBy :Hitachi_Spec .

:SPMT_Operation_20241120 rdf:type :InlandTransport ;
                          :departsFrom :MOSB ;
                          :arrivesAt :MIR ;
                          :requiresPermit :DOT_Permit_20241115 ;
                          :operatedBy :Mammoet .

:LCT_Operation_20241125 rdf:type :MarineTransport ;
                         :departsFrom :MOSB ;
                         :arrivesAt :DAS ;
                         :voyageDuration_hours 20 ;
                         :operatedBy :ALS ;
                         :cargo :TR_002 ;
                         :preservationMethod "Dry_air_N2" .
```

__🔶 4. Lifecycle Ontology (Node-based Material Flow)__

__Stage 1 – Import & Clearance__
→ arrivesAt(Port: Zayed/Mugharaq) → hasDocument(CI, PL, BL, COO) → customsClearedBy(ADNOC/ADOPT) → storedAt(Port Yard)

__Stage 2 – Consolidation at MOSB__
→ transportedBy(Inland Truck) → consolidatedAt(MOSB) → storedAt(MOSB Yard 20,000㎡) → preservedBy(Hitachi Spec: +5~40°C, RH≤85%)

__Stage 3 – Inland Transport (Onshore Sites)__
→ requiresPermit(DOT >90ton) → transportedBy(SPMT) → arrivesAt(SHU/MIR) → inspectedBy(QAQC) → resultsIn(MRR)

__Stage 4 – Marine Transport (Offshore Sites)__
→ requiresPermit(FRA) → transportedBy(LCT) → operatedBy(ALS) → arrivesAt(DAS/AGI ≈10~20h) → resultsIn(OSDR) → preservationMonitored(Dry air/N₂)

__Stage 5 – Installation Preparation__
→ finalInspection(MIS) → approvedBy(OE) → installedAt(Site) → commissionedBy(Hitachi/Vendor)

__🔶 5. Alignment with AI-Logi-Guide__

| __Ontology Node__      | __대응 모듈__     | __기능적 의미__                 |
| ---------------------- | ----------------- | ------------------------------- |
| Node                   | mapping           | 7-거점 좌표·연결성              |
| Activity               | pipeline          | Import→Storage→Transport→Install |
| Document               | rdfio, validation | CI/PL/BL/MRR/OSDR triple 구조   |
| Agent                  | core              | SCT/ALS/ADNOC 역할/권한 모델    |
| Permit                 | compliance        | DOT/FANR/CICPA 규제 검증        |
| RiskEvent              | reasoning         | Weather-Tie·Delay 추론          |
| Report                 | report            | KPI/MRR/OSDR 리포트 생성        |

__🔶 6. Semantic KPI Layer (Onto-KPI)__

| __KPI Class__              | __Onto Property__ | __계산식__                         | __Source__      |
| -------------------------- | ----------------- | ---------------------------------- | --------------- |
| __Port Dwell Time__        | portDwellDays     | (CY Out - CY In) days              | Port Event Log  |
| __MOSB Storage Duration__  | storageDays       | (Dispatch - Arrival) days          | MOSB Yard Data  |
| __Transit Time Accuracy__  | meetsETA          | ETA vs Actual ≤12%                 | Transport Event |
| __MRR SLA Compliance__     | mrrIssuedWithin   | MRR Issued ≤ 24h after Receiving   | QC Gate         |
| __OSDR Timeliness__        | osdrUpdatedWithin | OSDR Updated ≤ 7 days              | Offshore Report |
| __DOT Permit Lead Time__   | permitApprovalDays | (Issued - Requested) days          | DOT System      |
| __Preservation Compliance__ | tempRHWithinSpec  | Temp(5~40°C) AND RH(≤85%) %        | Sensor Data     |
| __Flow Code Distribution__ | flowCodeCoverage | Count per Flow Code (0-4) | Transport Events |

__🔶 7. Ontological Integration View__

```
[Origin: Sweden/Brazil]
     │
     ▼
[Port: Zayed/Mugharaq]
  ⟶ [Document: CI/PL/BL/COO]
  ⟶ [Customs: BOE·Duty]
     │
     ▼
[Hub: MOSB (Central Node)]
  ⟶ [Storage: 20,000㎡ Yard]
  ⟶ [Preservation: Hitachi Spec]
  ⟶ [Actor: SCT Team + ALS]
     │
     ├──→ [Onshore: SHU/MIR]
     │     ⟶ [Transport: SPMT + DOT Permit]
     │     ⟶ [Inspection: MRR]
     │     ⟶ [Installation: MIS + OE Approval]
     │
     └──→ [Offshore: DAS/AGI]
           ⟶ [Transport: LCT + FRA + ALS]
           ⟶ [Inspection: OSDR]
           ⟶ [Preservation: Dry air/N₂]
           ⟶ [Installation: MIS + Hitachi]
```

이 전체를 `hvdc-node-ontology.ttl`로 export하면,
GitHub macho715/hvdc-node-ontology에서 RDF 시각화 및 reasoning 연결 가능.

__🔶 8. 요약 메타 구조__

```json
{
 "Ontology": "hvdc-node-ontology",
 "CoreNodes": [
   {"name": "Zayed Port", "type": "Port", "locode": "AEZYD"},
   {"name": "Mugharaq Port", "type": "Port", "locode": null},
   {"name": "MOSB", "type": "Hub", "role": "Central consolidation", "capacity_sqm": 20000},
   {"name": "SHUWEIHAT (SHU)", "type": "OnshoreS ite", "laydown_sqm": 10556},
   {"name": "MIRFA (MIR)", "type": "OnshoreS ite", "laydown_sqm": 35006},
   {"name": "DAS Island", "type": "OffshoreSite", "voyageTime_h": 20},
   {"name": "Al Ghallan (AGI)", "type": "OffshoreSite", "voyageTime_h": 10}
 ],
 "PrimaryRelations": [
   "Port → consolidatedAt → MOSB",
   "MOSB → dispatches → (SHU, MIR, DAS, AGI)",
   "Cargo → transportedBy → (SPMT, LCT)",
   "Transport → requiresPermit → (DOT, FANR, CICPA)",
   "Site → receivesFrom → MOSB",
   "Asset → hasDocument → (MRR, OSDR, MIS)",
   "Operation → operatedBy → (SCT, ALS, ADNOC)"
 ],
 "AlignmentModule": "AI-Logi-Guide v2.1+",
 "ExportFormat": ["RDF/XML", "TTL", "JSON-LD"],
 "VerifiedSource": "HVDC Material Handling Workshop 2024-11-13"
}
```

이 프레임이면, HVDC 프로젝트 전체가 __"Port-Hub-Site의 지식망"__으로 정규화됩니다.
다음 단계는 `hvdc-node-ontology.reasoning` 모듈에서 __Rule-based inference__ 정의 — 예컨대 "DOT Permit가 누락된 중량물(>90톤)은 Site 이송 불가" 같은 정책을 OWL constraint로 명세하면 완성됩니다.

---

## 🔶 9. 핵심 노드 상세 정보 (검증된 사실 기반 - v3.0)

### 9.1 Core Node Set (8개 노드)

| 구분                                       | 유형                | 위치                       | 주요 기능                                                                                          | 연계 관계                                  |
| ------------------------------------------ | ------------------- | -------------------------- | -------------------------------------------------------------------------------------------------- | ------------------------------------------ |
| **자이드항 (Zayed Port)**                  | 해상입항노드         | 아부다비                   | **중량 및 일반 벌크 화물 처리항.** 변압기, 케이블드럼, 구조물 등 비컨테이너 자재 중심. SCT·JDN 확보 야드(1,100㎡) 존재. ADNOC 코드(47150)로 통관. | → MOSB / MIR                               |
| **칼리파항 (Khalifa Port)**                | 해상입항노드         | 아부다비                   | **컨테이너 전용항.** 해외(한국, 일본 등) 공급 자재 대부분 도착. ADNOC L&S 또는 DSV 관리하 적출. 자재는 트럭으로 MOSB 또는 현장 직송. | → MOSB / MIR / SHU                         |
| **제벨알리항 (Jebel Ali Port)**             | 해상입항노드 (특수케이스) | 두바이               | Free Zone 및 비ADNOC 공급사 사용. 일부 파이어파이팅, 전기부품 등 통관 후 ADOPT 코드로 재이송. SCT가 관세 납부 후 ADNOC에 비용 환급 요청. | → MOSB (재통관 경유)                       |
| **MOSB (Mussafah Offshore Supply Base)**  | **중앙 물류 허브**  | 아부다비 무사파            | ADNOC L&S 운영 Yard (20,000㎡). **SCT 물류본부 상주.** 해상화물(LCT/RoRo/Barge) 집하 및 적재. 컨테이너·CCU(약 80EA) 임시보관. 운송계획·FRA·Permit·Gate Pass 관리. | ← Zayed/Khalifa/Jebel Ali → MIR/SHU/DAS/AGI |
| **MIRFA SITE (MIR)**                       | 육상 현장           | 아부다비 서부              | 내륙 시공현장. 컨테이너·일반자재·중량화물 도착 후 설치. 35,000㎡ Laydown. 저장컨테이너(방화, 온도조절) 비치. 자재관리절차(SJT-19LT-QLT-PL-023) 적용. | ← MOSB / Zayed / Khalifa                  |
| **SHUWEIHAT SITE (SHU)**                   | 육상 현장           | 아부다비 서부              | 내륙 시공현장. Laydown 약 10,500㎡. 공간 제약으로 **운송순서·HSE 통제** 중요. 전기/기계류, 포설장비 등 일반자재 도착지. | ← MOSB / Khalifa                           |
| **DAS ISLAND (DAS)**                       | 해상 현장           | ADNOC 해역 (Zakum Cluster) | ADNOC 운영 해상기지. MOSB→LCT 약 20시간 항해. 컨테이너·벌크 혼재 화물 하역 및 적재장 운영. ADNOC HSE 표준, Lifting inspection, Gate control 준수. | ← MOSB                                     |
| **AL GHALLAN ISLAND (AGI)**                | 해상 현장           | ADNOC 해역 (DAS 병렬)     | MOSB→LCT 약 10시간 항해. 일반자재, 설치기구, 전기부품 운송. Laydown 47,000㎡ (3구역), 보안 강화. ADNOC L&S 동일 절차로 하역·보존 수행. | ← MOSB / ↔ DAS                             |

### 9.2 물류 흐름 구조 (v3.0 - All Cargo Types)

```
[해외 공급사 (Asia/EU 등)]
         ↓ (선적)
┌───────────────────────────┐
│   ZAYED PORT   KHALIFA PORT   JEBEL ALI PORT   │
└───────────────────────────┘
         ↓ (통관·운송)
             MOSB
    ┌────────┼────────┐
    ↓        ↓        ↓
  MIR      SHU     DAS / AGI
```

* **컨테이너 화물:** 주로 Khalifa Port → MOSB → 육상/해상 현장.
* **일반 벌크 화물:** Zayed Port → MOSB 또는 직접 MIR/SHU.
* **특수자재(Free Zone):** Jebel Ali → 재통관 → MOSB 경유.

### 9.3 기능 계층 구조 (v3.0)

| 계층                       | 설명                                     | 대표 노드                     |
| -------------------------- | ---------------------------------------- | ----------------------------- |
| **① 입항·통관 계층**       | 선적서류 검토(CI/PL/COO/eDAS), BL Endorsement, 통관코드 관리 | Zayed, Khalifa, Jebel Ali    |
| **② 집하·분류 계층**       | Port cargo 집하, 임시보관, Crane/Forklift 배차, Gate Pass, FRA 관리 | **MOSB**                      |
| **③ 육상 운송·시공 계층**  | 컨테이너·벌크 화물의 도로 운송 및 현장 인수, MRR/MRI 관리 | MIR, SHU                      |
| **④ 해상 운송·설치 계층**  | LCT/Barge 출항, ADNOC 해상안전기준(HSE), 하역·보존 | DAS, AGI                      |

### 9.4 운영·관리 사실 (v3.0)

* **SCT 물류본부:** MOSB 상주. 현장·항만·해상 노드 통합 관리.
* **운항 주체:** ADNOC Logistics & Services (ALS).
* **통관 관리:** ADOPT/ADNOC 코드 사용.
* **저장 관리:** MOSB + 인근 실내창고(6,000~8,000㎡) + 각 Site Laydown.
* **운송수단:** 트럭 / SPMT / CCU / LCT / Barge.
* **HSE 절차:** FRA, Method Statement, PTW, Lifting Certificate.
* **문서 체계:** MRR, MRI, OSDR, Gate Pass, Delivery Note.
* **중량물 운송 허가:** DOT 승인 필수(90톤 초과).
* **보존조건:** 실내 +5~40 °C, RH ≤ 85 % (Hitachi 권장).
* **항로거리:** MOSB→DAS 약 20 h, MOSB→AGI 약 10 h.

### 9.5 온톨로지 관계 (3중 구조 요약 - v3.0)

```
(MOSB, hosts, SCT_Logistics_Team)
(MOSB, consolidates, Container_and_Bulk_Cargo)
(MOSB, dispatches, MIR)
(MOSB, dispatches, SHU)
(MOSB, dispatches, DAS)
(MOSB, dispatches, AGI)
(Zayed_Port, handles, Heavy_and_Bulk_Cargo)
(Khalifa_Port, handles, Container_Cargo)
(Jebel_Ali_Port, handles, Freezone_Shipments)
(DAS, connected_to, AGI)
(MIR, and, SHU are Onshore_Receiving_Sites)
```

### 9.6 검증된 사실 요약 (v3.0)

1. **입항 및 통관:**
   * 중량·벌크 화물 → 자이드항,
   * 컨테이너 화물 → 칼리파항,
   * 일부 특수품 → 제벨알리항(Free Zone).

2. **중앙 허브(MOSB):**
   * 모든 화물의 **집하·검수·보존·해상출하** 기능 수행.
   * SCT 물류팀 본사 및 ADNOC L&S 현장운영팀 상주.

3. **육상 현장(MIR·SHU):**
   * 설치 및 시공 자재 수령지.
   * Laydown 내 임시보관, MRR/MRI·HSE 통제 중심.

4. **해상 현장(DAS·AGI):**
   * LCT 운항으로 자재 운송 및 하역.
   * ADNOC 해상안전 절차에 따라 작업.

5. **전체 구조:**
   > "**Zayed/Khalifa/Jebel Ali → MOSB → (MIR·SHU·DAS·AGI)**"
   > 형태의 다계층 물류 체계이며, **MOSB가 중앙 온톨로지 노드**로 작동한다.

---

**결론:**

HVDC 물류 시스템은 트랜스포머뿐 아니라 **컨테이너·벌크·일반자재 전반을 포함하는 복합 네트워크**이다.
모든 자재는 항만(자이드·칼리파·제벨알리)에서 통관 후 **MOSB를 중심으로 집하·분류·출하**되며,
최종 목적지는 육상(MIR·SHU) 또는 해상(DAS·AGI)으로 구분된다.
MOSB는 이 전체 체계의 **운영·정보·의사결정의 중심 노드**다.

---

🔧 **추천 명령어:**
`/logi-master node-audit` [8개 노드 필수 문서·허가 일괄 점검 - MOSB 중심 검증]
`/visualize_data --type=network hvdc-nodes` [노드 관계망 시각화 - 다계층 구조 확인]
`/compliance-check DOT-permit` [중량물(>90톤) DOT 승인 상태 검증 - MIR/SHU 대상]
`/cargo-flow analyze --type=all` [컨테이너·벌크·중량화물 전체 흐름 분석]
`/flow-code validate --strict` [Flow Code + WH Handling 일치성 검증 - 데이터 품질 보장]

