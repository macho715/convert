---
title: "Logistics Document Guardian (LDG) Ontology"
type: "ontology-design"
domain: "document-verification"
sub-domains: ["document-guardian", "trust-ontology", "semantic-verification"]
version: "unified-1.0"
date: "2025-01-23"
tags: ["ontology", "ldg", "trust-layer", "semantic-reasoning", "knowledge-graph"]
status: "active"
---

# hvdc-doc-guardian · 1_CORE-06

## Executive Summary

**온톨로지 관점에서 Logistics Document Guardian (LDG)**은 "문서 인식·검증 자동화 시스템"이 아니라, **지식 그래프 기반의 신뢰 체계(Trust Ontology System)**로 보는 게 정확하다.

LDG는 각 문서(CIPL, BL, PL, Invoice 등)를 **객체(Entity)**로 보고, 그 속성(Shipper, BL_No, HS_Code, Weight 등)을 **관계(Relation)**로 연결한다. 즉 "한 송장의 무게 필드가 B/L과 일치한다"는 것은 **데이터 일치가 아니라 관계의 정합성**을 의미한다.

이런 삼중 구조는 단순 데이터베이스가 아닌 **지식 기반(knowledge base)**이 되며, 문서 간 의미적 추론(Semantic Reasoning)이 가능하다.

**Visual — 핵심 클래스/관계(요약)**

| Class | 핵심 속성 | 관계 | 근거/조인 소스 | 결과 |
|-------|-----------|------|----------------|------|
| hvdc:Document | docId, docType, docHash | hasEntity → DocumentEntity | OCR/Table Parser | 상태, 정합성 |
| hvdc:DocumentEntity | entityType, value, confidence | linkedTo → CrossDocEntity | Field Tagger | URI 연결 |
| hvdc:TrustLayer | evidence, provenance, kpi | validates → DocumentGraph | SHACL Validation | PASS/FAIL |
| hvdc:LDGPayload | cascadedData, auditTrail | contains → VerificationResult | Knowledge Serialization | JSON/RDF |
| hvdc:CrossDocLink | sourceDoc, targetDoc, relation | crossReferences → Document | Entity Linking | 그래프 관계 |
| hvdc:VerificationResult | status, confidence, discrepancy | validates → Document | Auto-Verification | 검증 상태 |

자료: RDF 삼중 구조, SHACL 제약, 지식 그래프 기반 신뢰 체계.

**How it works (flow)**

1. **Data Acquisition**: 문서 이미지 → OCR → 디지털 트리플화 시작점 (관찰 노드 생성)
2. **Schema Alignment**: 문서별 속성을 온톨로지 클래스 구조에 맞춰 정규화
3. **Semantic Normalization**: 단위, 통화, 수량 등 의미 정규화 — "동일 의미 다른 표현"을 하나의 속성으로 매핑
4. **Entity Linking**: BL_No, Invoice_No 등을 URI로 연결 — 문서 간 그래프 관계 생성
5. **Knowledge Serialization**: LDG_PAYLOAD(JSON) = RDF 그래프의 직렬화 표현 (doc_hash는 Identity Anchor)
6. **SHACL Validation**: LDG_AUDIT은 그래프 제약 검증 결과 — 불일치 시 ZERO Fail-safe 트리거

**Options (설계 선택지)**

1. **RDF 삼중 기반 엄격형**: 모든 문서 관계를 RDF 삼중으로 모델링. *Pros* 의미적 추론↑ / *Cons* 초기 모델링 복잡도↑
2. **하이브리드형(권장)**: RDF + JSON 직렬화 + SHACL 제약, 부족 구간은 유사 문서 추천. *Pros* 실용성↑ / *Cons* 온톨로지 일관성 유지 필요
3. **지식 그래프 확장형**: FANR, MOIAT, Customs API 등 외부 규정도 동일한 URI 체계로 연결. *Pros* 확장성↑ / *Cons* 외부 데이터 동기화 필요

**Roadmap (P→Pi→B→O→S + KPI)**

- **Prepare**: 문서 타입별 RDF 스키마 정의, SHACL 제약 규칙 작성
- **Pilot**: /switch_mode LATTICE + /logi-master document-guardian --deep --trust-validation으로 샘플 문서 1회전. KPI: 검증정확도 ≥97%, 신뢰도 ≥95%
- **Build**: CrossDoc 관계 매핑, Trust Layer 증빙 시스템 구축, KPI 실시간 추적
- **Operate**: 불일치 감지 시 즉시 ZERO 모드 전환 + 감사 로그 생성
- **Scale**: 문서 그래프 스냅샷/변동 추적, 분기별 신뢰도 임계치 튜닝

**Automation notes**

- **입력 감지 →** /switch_mode LATTICE + /logi-master document-guardian (OCR→정규화→링킹→검증→신뢰도 측정)
- **신뢰 근거**: evidence[]와 doc_hash는 데이터의 provenance(출처·무결성)를 RDF 형태로 기록
- **감사 포맷**: SHACL Validation 결과 + Trust Layer KPI + CrossDoc 관계 맵

**QA / Gap 체크**

- 문서 간 관계 매핑이 **RDF 삼중 형태**로 올바르게 모델링되었는가?
- **SHACL 제약** 규칙이 모든 문서 타입에 대해 정의되었는가?
- Trust Layer의 **provenance 추적**이 완전한가?
- CrossDoc 링크의 **URI 연결**이 일관성 있게 유지되는가?

가정: (i) 모든 문서는 RDF 스키마에 따라 정규화됨, (ii) SHACL 제약은 내부 표준에 따라 배포됨, (iii) Trust Layer KPI는 실시간으로 업데이트됨.

---

# Part 2: LDG 파이프라인의 온톨지적 구조 대응

## LDG 파이프라인 온톨로지 매핑

| LDG 단계 | 온톨지 개념 | 설명 |
|----------|-------------|------|
| Pre-Prep / Vision OCR | **Data Acquisition Layer** | 관찰(observation) 노드 생성 — 물리적 문서의 디지털 트리플화 시작점 |
| Smart Table Parser | **Schema Alignment** | 문서별 속성(Property)을 Ontology의 클래스 구조에 맞춰 정규화 |
| NLP Refine | **Semantic Normalization** | 단위, 통화, 수량 등 의미 정규화 — "동일 의미 다른 표현"을 하나의 속성으로 매핑 |
| Field Tagger | **Entity Linking** | BL_No, Invoice_No 등을 URI로 연결 — 문서 간 그래프 관계 생성 |
| Payload Builder | **Knowledge Serialization** | LDG_PAYLOAD(JSON) = RDF 그래프의 직렬화 표현 (doc_hash는 Identity Anchor) |
| Auto-Verification | **SHACL Validation** | LDG_AUDIT은 그래프 제약 검증 결과 — 불일치 시 ZERO Fail-safe 트리거 |

## 온톨지적 가치

* **추론(Reasoning)**: "CIPL의 GrossWeight와 BL의 GrossWeight 차이가 ±15kg 이하"라는 규칙은 SHACL/OWL 제약으로 표현 가능
* **신뢰(Trust Layer)**: `evidence[]`와 `doc_hash`는 데이터의 provenance(출처·무결성)를 RDF 형태로 기록
* **확장성(Linked Data)**: FANR, MOIAT, Customs API 등 외부 규정도 동일한 URI 체계로 연결 가능 — LDG는 '물류 온톨지 허브'가 된다

## 요약 도식 (Mermaid)

```mermaid
graph TD
A[Document Image] -->|OCR| B[LDG Entity Graph]
B --> C[Field Tagging / URI Mapping]
C --> D[CrossDoc Links (CIPL↔BL↔PL)]
D --> E[LDG_PAYLOAD(JSON/RDF)]
E --> F[LDG_AUDIT(SHACL Validation)]
F --> G[Trust Layer (Evidence + Hash + KPI)]
```

## 핵심 포인트

1. LDG는 **문서의 진실성을 검증하는 온톨지적 "관계망"**이다
2. 각 문서는 개별 데이터셋이 아니라 **상호 연결된 의미 노드**다
3. KPI (MeanConf, TableAcc 등)는 **관계 신뢰도의 수량화된 지표**다
4. ZERO Fail-safe는 **온톨지 무결성 파손 시 그래프를 중단시키는 제어 규칙**이다

요약하자면, LDG는 OCR·AI·RPA의 조합으로 문서를 읽는 시스템이 아니라, **문서 간 관계를 '의미적으로 이해하고 검증'하는 온톨지 기반 신뢰 엔진**이다 — 즉, 물류 세계의 "Semantic Customs Officer."
