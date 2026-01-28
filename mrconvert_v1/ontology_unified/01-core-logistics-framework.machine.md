```json
{
  "@context": {
    "@vocab": "https://hvdc.example/ont#",
    "schema": "https://schema.org/",
    "incoterm": "https://iccwbo.org/incoterms/2020#",
    "dcsa": "https://dcsa.org/bl/3.0#",
    "wco": "https://www.wcoomd.org/datamodel/4.2#"
  },
  "@type": "UnifiedLogisticsOntology",
  "name": "Core Logistics Framework - HVDC Unified Ontology",
  "standards": ["UN/CEFACT", "WCO-DM", "DCSA", "ICC-Incoterms-2020", "HS-2022", "MOIAT", "FANR"],
  "layers": ["Upper", "ReferenceData", "Border/Customs", "Ocean/Carrier", "TradeTerms", "UAEReg", "OffshoreContract"],
  "coreClasses": ["Party", "Asset", "Document", "Process", "Event", "Contract", "Regulation", "Location", "KPI"],
  "objectProperties": ["hasIncoterm", "classifiedBy", "conformsTo", "usesDataModel", "requiresCertificate", "requiresPermit", "governedBy"],
  "rules": [
    {"appliesTo": "BL", "rule": "conformsTo", "object": "DCSA_eBL_3_0"},
    {"appliesTo": "InvoiceLineItem", "rule": "classifiedBy", "object": "HSCode"},
    {"appliesTo": "CustomsDeclaration(BOE)", "rule": "usesDataModel", "object": "WCO_DM_4_2_0"}
  ]
}
```

# Core Logistics Framework (Machine-readable edition)

---

### Part 1: Samsung C&T Construction Logistics Ontology {#top}

래는 __삼성 C&T 건설물류(UA E 6현장, 400 TEU/100 BL·월)__ 업무를 __온톨로지 관점__으로 재정의한 "작동 가능한 설계서"입니다.
핵심은 **표준(UN/CEFACT·WCO DM·DCSA·ICC Incoterms·HS·MOIAT·FANR)**을 상위 스키마로 삼아 __문서·화물·설비·프로세스·이벤트·계약·규정__을 하나의 그래프(KG)로 엮고, 여기서 __Heat‑Stow·WHF/Cap·HSRisk·CostGuard·CertChk·Pre‑Arrival Guard__ 같은 기능을 **제약(Constraints)**으로 돌리는 것입니다. (Incoterms 2020, HS 2022 최신 적용). 

---

### Domain Ontology — 클래스/관계 {#domain-ontology}

**핵심 클래스 (Classes)**

- **Party**(Shipper/Consignee/Carrier/3PL/Authority)
- **Asset**(Container ISO 6346, OOG 모듈, 장비/스프레더, OSV/바지선)
- **Document**(CIPL, Invoice, BL/eBL, BOE, DO, INS, MS(Method Statement), Port Permit, Cert[ECAS/EQM/FANR], SUPPLYTIME17)
- **Process**(Booking, Pre‑alert, Export/Import Clearance, Berth/Port Call, Stowage, Gate Pass, Last‑mile, WH In/Out, Returns)
- **Event**(ETA/ATA, CY In/Out, Berth Start/End, DG Inspection, Weather Alert, FANR Permit Granted, MOIAT CoC Issued)
- **Contract**(IncotermTerm, SUPPLYTIME17)
- **Regulation**(HS Rule, MOIAT TR, FANR Reg.)
- **Location**(UN/LOCODE, Berth, Laydown Yard, Site Gate)
- **KPI**(DEM/DET Clock, Port Dwell, WH Util, Delivery OTIF, Damage Rate, Cert SLA)

**대표 관계 (Object Properties)**

- Shipment → hasIncoterm → IncotermTerm (리스크/비용 이전 노드)
- InvoiceLineItem → classifiedBy → HSCode (HS 2022)
- BL → conformsTo → DCSA_eBL_3_0 (데이터 검증 규칙)
- CustomsDeclaration(BOE) → usesDataModel → WCO_DM_4_2_0 (전자신고 필드 정합)
- Equipment/OOG → requiresCertificate → MOIAT_ECAS|EQM (규제 제품)
- Radioactive_Source|Gauge → requiresPermit → FANR_ImportPermit (60일 유효)
- PortAccess → governedBy → CICPA_Policy (게이트패스)
- OSV_Charter → governedBy → SUPPLYTIME2017 (KfK 책임)

<!-- data:ref=#domain-ontology -->

---

### Use‑case 제약(Constraints) = 운영 가드레일 {#constraints}

**CIPL·BL Pre‑Arrival Guard (eBL‑first)**

- Rule‑1: BL 존재 → BL.conformsTo = DCSA_eBL_3_0 AND Party·Consignment·PlaceOfReceipt/Delivery 필수. 미충족 시 Berth Slot 확정 금지.
- Rule‑2: 모든 InvoiceLineItem는 HSCode 필수 + OriginCountry·Qty/UM·FOB/CI 금액. WCO DM 필드 매핑 누락 시 BOE 초안 생성 차단.
- Rule‑3: IncotermTerm별 책임/비용 그래프 확인(예: DAP면 현지 내륙운송·통관 리스크=Buyer).

**Heat‑Stow (고온 노출 최소화)**

- stowHeatIndex = f(DeckPos, ContainerTier, WeatherForecast) → 임계치 초과 시 Under‑deck/센터 베이 유도, berth 시간대 조정.
- dgClass ∈ {1,2.1,3,4.1,5.1,8} → Heat‑Stow 규칙 엄격 적용(위치·분리거리).

**WHF/Cap (Warehouse Forecast/Capacity)**

- InboundPlan(TEU/주)·Outplan → WHUtil(%) 예측, 임계치(85.00%) 초과 시 overflow yard 예약, DET 발생 예측과 연결.

**HSRisk**

- RiskScore = g(HS, Origin, DG, Cert 요구, 과거검사빈도) → 검사·추징·지연 확률 추정.

**CertChk (MOIAT·FANR)**

- 규제제품 → ECAS/EQM 승인서 필수 없으면 DO·GatePass 발행 금지, 선하증권 인도 보류.
- 방사선 관련 기자재 → FANR Import Permit(유효 60일) 없으면 BOE 제출 중단.

---

### 최소 예시 — JSON‑LD (요지) {#jsonld-sample}

```json
{
  "@context": {"incoterm":"https://iccwbo.org/incoterms/2020#","dcsa":"https://dcsa.org/bl/3.0#","wco":"https://www.wcoomd.org/datamodel/4.2#"},
  "@type":"Shipment",
  "id":"SHP-ADNOC-2025-10-001",
  "hasIncoterm":{"@type":"incoterm:DAP","deliveryPlace":"Ruwais Site Gate"},
  "hasDocument":[
    {"@type":"dcsa:BillOfLading","number":"DCSA123...", "status":"original-validated"},
    {"@type":"wco:CustomsDeclarationDraft","items":[{"hsCode":"850440", "qty":2, "value":120000.00}]}
  ],
  "consistsOf":[{"@type":"Container","isoCode":"45G1","isOOG":true,"dims":{"l":12.2,"w":2.44,"h":2.90}}]
}
```

---

### Roadmap & QA & SHACL 하이라이트 {#roadmap-shacl}

상세 내용은 원본 문서의 파트 구조(옵션·로드맵·QA·Fail-safe·SHACL·재사용 가이드)를 유지하며, 키 포인트를 앵커로 교차참조 합니다.

See also: `../TR.machine.md`, `../TR2.machine.md` (process/timeline/gates)



