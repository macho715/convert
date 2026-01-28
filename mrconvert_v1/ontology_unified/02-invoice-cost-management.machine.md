```json
{
  "@context": {
    "@vocab": "https://hvdc.example/ont#",
    "schema": "https://schema.org/"
  },
  "@type": "InvoiceCostOntology",
  "name": "Invoice & Cost Management Ontology",
  "currencyPolicy": {"base": "USD", "fx": 3.6725},
  "classes": ["Invoice", "InvoiceLine", "ODLane", "RateRef", "CurrencyPolicy", "RiskResult"],
  "relations": [
    {"subject": "Invoice", "predicate": "hasLine", "object": "InvoiceLine"},
    {"subject": "InvoiceLine", "predicate": "hasLane", "object": "ODLane"},
    {"subject": "InvoiceLine", "predicate": "uses", "object": "RateRef"}
  ],
  "rules": [
    {"appliesTo": "InvoiceLine", "rule": "toleranceBySource", "contract": "±3%", "market|quotation": "±5%", "special": "±10%"},
    {"appliesTo": "Invoice", "rule": "singleCurrencyOrFlag", "flags": ["MULTI_CURRENCY", "CURRENCY_MISMATCH"]},
    {"appliesTo": "Invoice", "rule": "recalculateTotal", "precision": 0.01}
  ],
  "bands": ["PASS", "WARN", "HIGH", "CRITICAL"],
  "fxLocked": true
}
```

# Invoice & Cost Management (Machine-readable edition)

---

### Part 1: Invoice Ontology System {#invoice-ontology}

**온톨로지-퍼스트 청구서 시스템**은 **멀티-키 아이덴티티 그래프** 위에서 Invoice → Line → OD Lane → RateRef → Δ% → Risk로 캐스케이드합니다. 표준요율은 Air/Container/Bulk 계약 레퍼런스와 Inland Trucking(OD×Unit) 테이블을 온톨로지 클래스로 들고, 모든 계산은 **USD 기준·고정환율 1.00 USD=3.6725 AED** 규칙을 따릅니다.

---

### 핵심 클래스/관계 요약 {#classes-relations}

| Class | 핵심 속성 | 관계 | 근거/조인 소스 | 결과 |
|---|---|---|---|---|
| hvdc:Invoice | docId, vendor, issueDate, currency | hasLine → InvoiceLine | — | 상태, 총액, proof |
| hvdc:InvoiceLine | chargeDesc, qty, unit, draftRateUSD | hasLane → ODLane / uses → RateRef | Trucking, Air/Container/Bulk rate | Δ%, cg_band |
| hvdc:ODLane | origin_norm, destination_norm, vehicle, unit | joinedBy → ApprovedLaneMap | RefDestinationMap, Lane stats | median_rate_usd |
| hvdc:RateRef | rate_usd, tolerance(±3%), source(contract/market/special) | per Category/Port/Dest/Unit | Rate tables | ref_rate_usd |
| hvdc:CurrencyPolicy | base=USD, fx=3.6725 | validates Invoice/Line | currency_mismatch rule | 환산/락 |
| hvdc:RiskResult | delta_pct, cg_band, verdict | from Line vs Ref | COST-GUARD bands | PASS/FAIL |

<!-- data:ref=#classes-relations -->

---

### Flow (동작 개요) {#flow}

1. 키 해석(Identity): 다중 키 입력 → 동일 실체 클러스터 식별.
2. Lane 정규화: RefDestinationMap → ApprovedLaneMap.
3. Rate 조인: Category+Port+Destination+Unit 기준.
4. Δ% 및 밴드 산정: PASS/WARN/HIGH/CRITICAL.
5. 감사 아티팩트: PRISM.KERNEL 요약 + JSON proof.

---

### 규칙·SHACL 요점 {#rules-shacl}

- 필수 필드/단위/포맷 강제 (Invoice/DO/Stowage/WH/Status shape).
- 출처별 허용오차: Contract ±3%, Market/Quotation ±5%, Special ±10%.
- 통화 규칙: 원문 문서 통화 유지, 다중 통화는 MULTI_CURRENCY, 참조와 다르면 CURRENCY_MISMATCH.
- 합계 재계산: rate × quantity 오차 ≤ 0.01.

---

### 그래프 예시 (축약 Turtle) {#ttl-example}

```ttl
@prefix hvdc: <https://example.com/hvdc#> .
@prefix prov: <http://www.w3.org/ns/prov#> .
@prefix xsd:  <http://www.w3.org/2001/XMLSchema#> .

hvdc:Invoice123 a hvdc:Invoice ;
  hvdc:invoiceKey "INV-123" ;
  hvdc:currency "AED" ;
  hvdc:rateSource hvdc:Contract ;
  hvdc:hasRate "150.00"^^xsd:decimal ;
  hvdc:hasQuantity "10"^^xsd:decimal ;
  hvdc:hasTotal "1500.00"^^xsd:decimal ;
  hvdc:references hvdc:Contract789 .

hvdc:Invoice123_Validation a hvdc:ValidationActivity ;
  hvdc:verificationStatus hvdc:PENDING_REVIEW ;
  hvdc:hasDiscrepancy [ a hvdc:Discrepancy ; hvdc:discrepancyType hvdc:RateTolerance ; hvdc:deltaPercent "0.045"^^xsd:decimal ] ;
  prov:used hvdc:Contract789 ; prov:generated hvdc:Invoice123 .
```

---

See also: `../../TR.machine.md` (actors/process/timeline), `../../TR2.machine.md` (relations/gates)



