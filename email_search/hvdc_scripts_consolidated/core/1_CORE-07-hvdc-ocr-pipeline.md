---
title: "LDG High-Precision OCR Pipeline Ontology"
type: "ontology-design"
domain: "document-ocr-processing"
sub-domains: ["ocr-extraction", "data-refinement", "validation-framework", "cost-guard"]
version: "unified-2.4"
date: "2025-01-23"
tags: ["ontology", "ldg", "ocr", "document-processing", "validation", "cost-guard", "regtech"]
status: "active"
---

# hvdc-ocr-pipeline · 1_CORE-07

## Executive Summary

**고정밀 OCR·구조화 지침 v2.4 – LDG Ready**를 온톨로지 관점으로 보면, 단순 파이프라인이 아니라 "문서→추출→정제→검증→감사"로 이어지는 **의미 그래프**다. 핵심은 각 단계가 **명시적 클래스와 관계**로 연결되고, KPI와 Fail-safe가 **제약(Constraint)**으로 모델에 박혀 있다는 점이다.

**상위 개념 계층(Top Taxonomy)**:
```
Document Processing Pipeline
└── LDG OCR Pipeline
    ├── Document Input (CI/PL/BL/Invoice 등)
    ├── OCR Processing (Vision OCR, Smart Table Parser)
    ├── Data Refinement (NLP Refine, Field Tagger)
    ├── Validation Framework (Auto-Validation 5단계)
    ├── Cost Guard (표준요율 대비, FX 잠금)
    ├── RegTech Integration (MOIAT/FANR/IMDG/Dual-Use)
    └── Audit & Reporting (LDG_AUDIT, Cross-Doc Links)
```

**Visual — 핵심 클래스/관계(요약)**

| Class | 핵심 속성 | 관계 | 근거/조인 소스 | 결과 |
|-------|-----------|------|----------------|------|
| ldg:Document | docType, docId, fileHash | hasPage→Page, hasImage→Image | Document Registry | 처리 상태 |
| ldg:Page | pageNumber, imageRef | partOf→Document | OCR Engine | 추출 결과 |
| ldg:Image | imageHash, resolution | contains→OCRBlock | Vision OCR | 신뢰도 점수 |
| ldg:OCRBlock/OCRToken | text, confidence, position | extractedFrom→Image | OCR Processing | 정제 텍스트 |
| ldg:Table | schema, type, footnote | parsedFrom→OCRBlock | Smart Table Parser | 구조화 데이터 |
| ldg:RefinedText | formatted, unit, currency | refines→OCRToken | NLP Refine | 정규화 텍스트 |
| ldg:EntityTag | entityType, value, confidence | tags→RefinedText | Field Tagger | 엔티티 매핑 |
| ldg:Payload | version, trade, logistics | buildsFrom→EntityTag | Payload Builder | LDG_PAYLOAD |
| ldg:Validation | stage, result, percentage | validates→Payload | Auto-Validation | 검증 상태 |
| ldg:Metric | meanConf, tableAcc, numericIntegrity, entityMatch | measures→Validation | KPI Calculation | 성능 지표 |
| ldg:Audit | selfCheck, totalsCheck, crossDocCheck, hashConsistency | audits→Payload | LDG_AUDIT | 감사 결과 |
| ldg:CrossLink | sourceDoc, targetDoc, relation | links→Document | Cross-Doc Analysis | 문서 연관 |
| ldg:RegTechFlag | flagType, severity, jurisdiction | triggeredBy→EntityTag | RegTech Analysis | 규제 플래그 |
| ldg:HSCandidate | hsCode, confidence, source | proposedBy→EntityTag | HS Classification | HS 코드 후보 |
| ldg:CostGuardCheck | standardRate, draftRate, exceedPct, verdict | evaluates→Payload | Cost Guard | 비용 검증 |

자료: LDG Pipeline 단계별 처리 결과, KPI 임계값, 제약 조건.

**How it works (flow)**

1. **Document Input**: CI/PL/BL/Invoice 등 문서 업로드 → Document 객체 생성 → Page/Image 분할
2. **OCR Processing**: Vision OCR → OCRBlock/OCRToken 추출 (confidence 포함) → Smart Table Parser → Table 구조화
3. **Data Refinement**: NLP Refine → RefinedText 생성 (형식·단위 보정) → Field Tagger → EntityTag 자동 태깅
4. **Validation Framework**: Payload Builder → LDG_PAYLOAD 생성 → Auto-Validation 5단계 → Validation 결과
5. **Cost Guard**: 표준요율 대비 초과율 계산 → FX 잠금 정책 적용 → CostGuardCheck 판정
6. **RegTech Integration**: HS 후보/키워드 분석 → MOIAT/FANR/IMDG/Dual-Use 플래그 설정 → RegTechFlag 생성
7. **Audit & Reporting**: Cross-Doc Links 분석 → LDG_AUDIT 생성 → HITL 승인 → Report Lock

**Options (설계 선택지)**

1. **엄격형**: 모든 단계를 OWL/SHACL로 엄격하게 모델링. *Pros* 의미적 추론↑ / *Cons* 초기 모델링 복잡도↑
2. **하이브리드형(권장)**: OWL + JSON-LD + SHACL 제약, 부족 구간은 유사 패턴 추천. *Pros* 실용성↑ / *Cons* 온톨로지 일관성 유지 필요
3. **실무형**: 핵심 클래스만 모델링하고 나머지는 확장 가능한 구조. *Pros* 빠른 적용↑ / *Cons* 확장성 제한

**Roadmap (P→Pi→B→O→S + KPI)**

- **Prepare**: 네임스페이스/컨텍스트 확정, 클래스 스키마 정의, SHACL 제약 규칙 작성
- **Pilot**: /switch_mode LATTICE + /logi-master document-guardian --deep --ocr-precision으로 샘플 문서 1회전. KPI: OCR 정확도 ≥97%, 검증 성공률 ≥95%
- **Build**: KPI 게이트, Fail-safe 시스템, HITL 승인 프로세스 구축
- **Operate**: 실시간 모니터링, 이상 상황 즉시 ZERO 모드 전환 + 중단 로그
- **Scale**: 다중 문서 타입 지원, RegTech 규정 업데이트 자동화, Cost Guard 임계값 동적 조정

**Automation notes**

- **입력 감지 →** /switch_mode LATTICE + /logi-master document-guardian (OCR→정제→검증→감사→보고서)
- **표준 근거**: LDG Pipeline 단계별 KPI 임계값, HallucinationBan/Deterministic 규칙
- **감사 포맷**: LDG_AUDIT JSON + 해시/서명/타임스탬프 + Changelog

**QA / Gap 체크**

- OCR 신뢰도가 **임계값 이상**인가?
- NumericIntegrity가 **100%**인가?
- EntityMatch가 **기준 이상**인가?
- HashConsistency가 **PASS**인가?
- KPI 게이트를 **모두 통과**했는가?

가정: (i) 모든 문서는 표준 형식을 따름, (ii) OCR 엔진이 최신 버전으로 유지됨, (iii) KPI 임계값이 사전에 정의됨.

---

# Part 2: Data Schema & Mapping

## LDG_PAYLOAD/AUDIT 매핑 (JSON → 그래프)

**예시 매핑(개념)**

* `payload.DocType` → `ldg:Payload ldg:docType`
* `Parties.Shipper/Consignee` → `ldg:EntityTag ldg:partyRole`
* `Ids.BL_No/Invoice_No` → `ldg:Identifier ldg:idValue`
* `Logistics.(Packages,GW,NW,CBM)` → `ldg:CargoStat`
* `Trade.(Incoterms,Origin,Destination)` → `ldg:TradeTerm`
* `Currency, Total_Amount` → `ldg:MonetaryTerm`
* `HS_Candidates[]` → `ldg:HSCandidate(ldg:code, ldg:confidence)`
* `RegTechFlags` → `ldg:RegTechFlag` (MOIAT/FANR/IMDG/DualUse)
* `Metrics` → `ldg:Metric` (검증 노드에 연결)
* `Hashes.image/text` → `ldg:HashDigest`
* `CrossLinks.PL_No` → `ldg:CrossLink`(문서 간 참조)
* `LDG_AUDIT.SelfCheck/TotalsCheck/CrossDocCheck/HashConsistency/Warnings` → `ldg:Audit` 하위 속성

문서에 정의된 JSON 구조를 위와 같이 온톨로지로 승격하면, **문서·값·검증·감사**가 한 그래프 위에서 **질의/추적/잠금** 가능.

## JSON-LD 컨텍스트(요약 예시)

```json
{
  "@context": {
    "ldg": "https://example.com/ldg#",
    "docType": "ldg:docType",
    "partyRole": "ldg:partyRole",
    "idValue": "ldg:idValue",
    "numericIntegrity": "ldg:numericIntegrity",
    "meanConf": "ldg:meanConf",
    "tableAcc": "ldg:tableAcc",
    "entityMatch": "ldg:entityMatch",
    "hashImage": "ldg:hashImage",
    "hashText": "ldg:hashText",
    "fxSource": "ldg:fxSource",
    "fxLocked": "ldg:fxLocked",
    "exceedPct": "ldg:exceedPct",
    "verdict": "ldg:verdict"
  }
}
```

이 컨텍스트를 쓰면 LDG_PAYLOAD/LDG_AUDIT JSON이 곧바로 **지식 그래프 문서화**된다(질의·감사·재현성 ↑).

---

# Part 3: Constraints & Validation

## 제약(Constraints)과 Fail-safe의 온톨로지 표현

**SHACL 제약**(개념 초안)

```turtle
# NumericIntegrity = 100% (라인합=수량×단가 ∧ 합계 일치 ∧ 통화/FX 일관)
ldg:NumericIntegrityShape a sh:NodeShape ;
  sh:targetClass ldg:Validation ;
  sh:property [
    sh:path ldg:numericIntegrity ;
    sh:datatype xsd:decimal ;
    sh:minInclusive 1.00 ; sh:maxInclusive 1.00
  ] .

# MeanConf ≥ 임계값
ldg:MeanConfShape a sh:NodeShape ;
  sh:targetClass ldg:Metric ;
  sh:property [
    sh:path ldg:meanConf ;
    sh:datatype xsd:decimal ;
    sh:minInclusive 0.95
  ] .

# TableAcc ≥ 임계값
ldg:TableAccShape a sh:NodeShape ;
  sh:targetClass ldg:Metric ;
  sh:property [
    sh:path ldg:tableAcc ;
    sh:datatype xsd:decimal ;
    sh:minInclusive 0.90
  ] .

# EntityMatch ≥ 임계값
ldg:EntityMatchShape a sh:NodeShape ;
  sh:targetClass ldg:Metric ;
  sh:property [
    sh:path ldg:entityMatch ;
    sh:datatype xsd:Decimal ;
    sh:minInclusive 0.85
  ] .
```

**ZERO Fail-safe**: 아래 중 하나라도 위반 시 `ldg:Validation ldg:state ldg:ZERO`
(MeanConf/TableAcc/EntityMatch 기준 미달, NumericIntegrity≠100%, HashMismatch 등).

## 질의·감사 시나리오(예)

* **합계 무결성 점검**: `SELECT` 그래프 질의로 `(qty×unitPrice)=lineTotal` 전 행 검증 → `doc_total`과 일치 확인(±0.00). 미일치 시 해당 `ldg:Validation`을 `ldg:ZERO`로 전이.
* **Cross-Doc 일관성**: `PL→CI weights`, `CIPL→BL pkgs`를 `ldg:CrossLink`로 추적, `ldg:Audit.CrossDocCheck` PASS/FAIL 로그 축적.
* **RegTech 트리거**: HS 후보/키워드로 MOIAT/FANR/IMDG/Dual-Use 플래그 세팅, 최종 결론은 사람 승인(HITL)로 잠금.
* **Cost-Guard**: `std rate` vs `draft rate` 초과율 계산 + FX 잠금 정책(`fx_source`, `fx_locked`)을 그래프에 기록, 항목별 리스크와 `verdict` 집계.

---

# Part 4: Operational Principles

## 운영 원칙의 온톨로지 내재화(요점)

* **HallucinationBan**: 값 부재는 `null`, 불명확은 `"??"`, 추정은 `"(추정·Med)"`로 **상태값**을 분리—모델이 *모른다*는 사실을 표현.
* **Deterministic 수리 규칙**: 소수점·합계·환율·단위 고정 → 모든 계산은 **재현 가능 함수**(Audit에 로그).
* **KPI 게이트**: KPI는 단순 숫자가 아니라 **통과 조건**. 온톨로지에서 SHACL/규칙으로 강제.
* **HITL & Report Lock**: 승인 전 체크리스트와 해시/타임스탬프로 **불변 스냅샷**을 남김(Changelog 누적).

## 바로 쓰는 최소 구현 순서(현장 적용용)

1. **네임스페이스/컨텍스트** 확정(`ldg:`), 문서키(`doc_hash`+파일 메타) 생성.
2. **스키마 실체화**: `ldg:Document/Validation/Audit/Metric/RegTech/CostGuard` 클래스와 속성 발행.
3. **파이프라인 이벤트→트리플**: 각 스텝 산출물(표, 태그, KPI)을 즉시 그래프에 `Append`.
4. **KPI/Fail-safe SHACL** 배치: 미달 즉시 `ldg:ZERO` 상태 전이 + 중단 로그.
5. **HITL 승인 시 Report Lock**: 해시/서명/타임스탬프 고정, 재생성 시 Changelog 추가.

## 프로세스 → 온톨로지 상태머신

```
ldg:Document
  ├─(ldg:hasImage)→ ldg:Image
  ├─(ldg:hasOCR)→ ldg:OCRBlock/Token(conf)
  ├─(ldg:hasTable)→ ldg:Table(type, unit/currency split, footnote)
  ├─(ldg:refines)→ ldg:RefinedText(??, (추정·Med) 명시)
  ├─(ldg:tagsEntity)→ ldg:EntityTag(Shipper/BL_No/Incoterms…)
  ├─(ldg:buildsPayload)→ ldg:Payload(LDG_Version, Trade, Logistics…)
  ├─(ldg:validatedBy)→ ldg:Validation(5단계, 적합도%)
  │    └─(ldg:hasMetric)→ ldg:Metric(MeanConf/TableAcc/NumericIntegrity/EntityMatch)
  ├─(ldg:hasAudit)→ ldg:Audit(SelfCheck/Totals/CrossDoc/Hash)
  ├─(ldg:crossLinksTo)→ ldg:CrossLink(CIPL↔BL↔PL)
  ├─(ldg:triggersRegTech)→ ldg:RegTechFlag(MOIAT/FANR/IMDG/DualUse)
  └─(ldg:evaluatedByCostGuard)→ ldg:CostGuardCheck(FX locked, exceed_pct, verdict)
```

* **KPI 게이트**(MeanConf≥…, TableAcc≥…, NumericIntegrity=100%, EntityMatch≥…, HashConsistency=PASS)는 `ldg:Validation`의 **필수 제약조건**으로 모델링 → 미달 시 `ldg:FailState(ZERO)`로 전이.
* **HallucinationBan/Deterministic** 규칙은 `ldg:RefinedText`와 `ldg:Payload`에 **표현 제약**으로 귀속(불명확="??", 추정="(추정·Med)", 합계 오차 ±0.00 고정).

원하는 만큼 깊게 들어갈 수 있는데, 여기까지면 뼈대는 잡혔다.
다음 스텝으로, 네가 쓰는 **RDF/JSON-LD 템플릿**이나 **SHACL 세트**를 바로 만들어줄 수도 있어. 어떤 문서타입(CI/PL/BL/Invoice)부터 묶을지만 콕 집어줘.
