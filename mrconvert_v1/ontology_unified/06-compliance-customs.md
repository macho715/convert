---
title: "Compliance & Customs Clearance Ontology"
type: "ontology-design"
domain: "customs-clearance"
version: "1.0"
date: "2025-01-19"
tags: ["ontology", "cipl", "bl", "pre-arrival", "customs", "hvdc", "compliance"]
status: "active"
---

# CIPL·BL Pre-Arrival Guard - 온톨로지 청사진

## 개요

"엔티티→관계→규칙→증거"로 뼈대를 세워서 온톨로지 청사진을 정의합니다. 클래스/프라퍼티, 15-스텝 상태머신, SHACL 제약, PRISM.KERNEL 연동, hvdc 네이밍 규칙까지 포함합니다.

티저로 핵심만 짚자면:

- 상위 개념: Document/Shipment/GoodItem/Event\(Regulatory·ETA/DEM\)/Regulation·Permit/RiskClass\(S\-Class\)/StorageCondition/Rate/Alert/Evidence\.
- 허용오차·게이트\(±1\.00%·±2\.00%, OCR 0\.60, TableAcc 98\.00, EntityMatch 0\.98\)는 __데이터 속성__으로 고정값 선언\.
- Gov≥Carrier 원칙과 14종 Alert는 __택소노미__로 분리해서 추론·필터링이 쉬움\.
- Incoterm\-aware 운임/보험 로직·ETA→DEM\_END 계산은 __Computation 노드__ \+ __Evidence 링크__로 투명화\.
- PRISM\.KERNEL 방식의 recap\.card \+ proof\.artifact\(JSON\)은 온톨리지 질의결과에서 바로 뽑도록 바인딩\.

근거는 너랑 쓰던 내부 레퍼런스들에 맞춰 깔았고\(세부 수치는 캔버스에만, 여기선 경량화\):

다음 선택지는 셋 중 하나야—네가 고르면 바로 이어서 반영할게\.

1. __/switch\_mode ORACLE__: HS↔UN↔S\-Class 룰을 실제 SHACL/SPARQL 규칙으로 자세화\(프로덕션 규칙 세트로 내림\)\.
2. __/switch\_mode LATTICE__: CIPL/BL 필드 매핑을 RDF/JSON\-LD 스키마로 세분화\(필드별 정규화·코드리스트 포함\)\.
3. __/switch\_mode COST\-GUARD__: Ocean/Air Rate를 Rate 인스턴스 \+ 유효기간으로 로드하고, CIF/CFR 계산 파이프 연결\.

원하는 모드랑, 필요하면 __hvdc\-adopt\-xxx\-xxxx__ 코드값도 알려줘\. 그러면 그 코드로 네임스페이스 고정해서 이어갈게\.
