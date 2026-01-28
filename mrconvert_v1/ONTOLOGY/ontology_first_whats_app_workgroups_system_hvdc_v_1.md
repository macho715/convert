---
title: "WhatsApp Workgroups System - Ontology First"
type: "ontology-design"
domain: "communication-workflow"
version: "1.0"
date: "2025-10-19"
tags: ["ontology", "whatsapp", "workgroups", "hvdc", "communication"]
timezone: "Asia/Dubai"
scope: "UAE HVDC Project Lightning"
status: "active"
---

# Ontology-First Spec — WhatsApp Workgroups System (HVDC)

## Executive Summary

**Version:** v1.0
**Date:** 2025-10-19 (GST)
**Scope:** UAE HVDC Project Lightning — WhatsApp 업무 그룹 전반(Abu Dhabi Logistics, Jopetwil 71 Group, UPC – Precast Transportation, AGI – Wall Panel – GCC Storage, [HVDC] Project Lightning)

---

## 1) Executive Summary (3–5 lines)
- 본 시스템은 **Multi‑Key Identity Graph** 위에서 *그룹↔스레드↔메시지↔태그↔자산/사이트/화물/승인*을 연결한다. (*Any‑key in → Resolve → Cluster → Tasks*).
- **Master Policy**의 태그·SLA·파일명·보안 규칙을 **온톨로지 계층**으로 승격하여, **일관된 자동 분류·SLA 타이머·PII 마스킹·자동 리포트**를 실행한다.
- 각 그룹의 **고유 패턴(항만/타이드/장비/증빙)**은 *도메인 보카블러리(SKOS)*로 관리, **키워드→태스크**를 표준화한다.
- 결과물: RDF/OWL 온톨로지 + SHACL 검증 + JSON‑LD 컨텍스트 + SPARQL 질의 + 자동화 훅(08:30/17:30, 목 16:00).

---

## 2) Conceptual Model (개념)
- **hvdc:Workgroup** ⟶ hasMember **hvdc:Participant** (역할/RACI)
- **hvdc:Workgroup** ⟶ hasThread **hvdc:Thread** ⟶ hasMessage **hvdc:Message**
- **hvdc:Message** ⟶ hasTag **hvdc:Tag** (고정 9종 + 확장)
- **hvdc:Message** ⟶ about **hvdc:Asset | hvdc:Site | hvdc:Cargo | hvdc:Approval**
- **hvdc:Message** ⟶ evokes **hvdc:Action** (예: book_crane_100t, submit_gate_pass_list)
- **hvdc:Message** ⟶ hasAttachment **hvdc:Document** (CIPL/BL/DO/Permit 등)
- **hvdc:SLAClock** attachesTo **hvdc:Message/Action** (업무시간·오프타임 규칙 내장)
- **hvdc:Policy** governs **Workgroup/Message/Document** (보안·PII·파일명·언어)

> 네임스페이스 권고: `hvdc:` (core), `prov:`(활동/책임), `schema:`(일반 메타), `skos:`(용어), `time:`(시간), `sh:`(SHACL)

---

## 3) Core Classes & Key Properties
| Class | 필수 속성 | 선택 속성 | 설명 |
|---|---|---|---|
| hvdc:Workgroup | id, name, purpose | kpiProfile, alias, policy | 그룹 단위 엔티티 |
| hvdc:Participant | id, displayName, role | org, phone*, email* | 역할/책임(RACI); *PII 마스킹 규칙 적용 |
| hvdc:Thread | id, startedAt | subject | WhatsApp 대화 스레드/하위 토픽 |
| hvdc:Message | id, sentAt, sender, textHash | header, replyTo, attachments | 첫 줄 태그·제목 규칙 파싱(`header`) |
| hvdc:Tag (skos:Concept) | prefLabel | altLabel, inScheme | 고정 9종 + 도메인 태그(예: [SITREP]) |
| hvdc:Site (skos:Concept) | code, prefLabel | altLabel | /sites/ 키워드 동의어 포함 |
| hvdc:Asset (skos:Concept) | type, prefLabel | altLabel | Crane-100T, A‑frame 등 |
| hvdc:Cargo (skos:Concept) | prefLabel | size/mm, unit | Aggregate‑5/10/20mm, HCS, JB |
| hvdc:Approval (skos:Concept) | prefLabel | issuer, docRef | TPI, MWS, Gate pass |
| hvdc:Action | verb, target | dueBy, owner | 키워드→태스크 맵핑 결과 |
| hvdc:Document | docType, fileName | shptNo, version | 파일명 규칙 검증 대상 |
| hvdc:SLAClock | classOfService, startedAt | breachedAt, owner | URGENT/ACTION/FYI SLA 집행 |
| hvdc:Policy | name, version | fileRef | Master Policy 규칙 묶음 |

---

## 4) Tagging & Controlled Vocabulary
### 4.1 고정 태그(메타)
`[URGENT][ACTION][FYI][ETA][COST][GATE][CRANE][MANIFEST][RISK]`

### 4.2 도메인 태그/보카블러리(발췌)
- **Site:** `/sites/AGI-West-Harbor`, `/sites/MW4`, `/sites/MOSB`, `/sites/DAS-Island`, `/sites/GCC-yard` …
- **Asset:** `Crane-100T`, `A-frame trailer`, `Forklift-10T`, `Wheel loader`, `Head engine (tractor)` …
- **Cargo:** `Aggregate-5mm/10mm/20mm`, `Jumbo bag`, `HCS`, `Wall panel`, `Pin rack` …
- **Approval:** `TPI`, `MWS`, `Lifting Plan`, `Gate pass`

### 4.3 동의어(예시)
- `ALS` = `ADNOC L&S`; `AGI`=Al Ghallan Island; `JPTW71`=`Jopetwil 71`
- `A‑frame` ↔ `aframe`, `af frames`; `Crane-100T` ↔ `100t`, `100 ton`

### 4.4 키워드→태스크(Action) 맵(예)
- `"100t crane" → hvdc:Action(book_crane_100t)`
- `"gate pass" → hvdc:Action(submit_gate_pass_list)`
- `"A‑frame rotation" → hvdc:Action(allocate_aframe_rotation)`

> 구현: SKOS `altLabel` + 정규화 사전(동의어) + 룰베이스(정규표현식) + ML 보조(옵션)

---

## 5) Event/Message Model & SLA Semantics
- **업무시간**: 08:00–20:00 (GST). 오프타임: *URGENT만 에스컬레이션*.
- **SLA**: URGENT 10분, ACTION 2시간, FYI 당일.
- **헤더 규칙**: 메시지 첫줄 `[TAG][TAG] [SHPTNO]/[SITE]/[ITEM]/[ETA]/[ACTION]` 권장.
- **파일명 규칙**: `YYYYMMDD_[SHPTNO]_[DOC]_v##` (예: `20250808_HVDC-AGI-J71-047_CIPL_v02`).
- **SLAClock 동작**: 메시지 수신→태그·업무시간 판단→타이머 시작→정답/진척 메시지로 멈춤→초과 시 `hvdc:SLAEvent` 기록+에스컬.

---

## 6) Policy Alignment (보안·언어·PII)
- **언어**: KR 기본 + EN 병기(수량·TON·시간·장비·위험은 EN 병기 권장).
- **PII**: 전화·이메일 마스킹(****). *이름은 원문 유지(조직 정책에 따름).*
- **중복/스팸**: 동일 내용 ≥3회 → `CONSOLIDATE` 스레드로 묶기.
- **미디어/이모지**: 활동 신호로만 사용(토픽 카운트 제외).

---

## 7) Group Instantiation (도메인 인스턴스)
- **Abu Dhabi Logistics**: 중앙 오케스트레이션(게이트패스·크레인·컨테이너·ETA/ETD) — `highCadence=true`, `policy.port=Al Jaber A‑Frame only`, `rule.FLIFT=10·20·30`.
- **Jopetwil 71 Group**: AGI Jetty#3, RORO 타이드 윈도우(~10:30/12:00), LOLO(점보백) 시 Telehandler 금지, ALS 크레인 예약.
- **UPC – Precast Transportation**: A‑frame 3대 이상 상시, Dunnage 높이 규정(웹슬링 삽입), 16:00 플랜, 주말 MOSB 크레인 N/A.
- **AGI – Wall Panel – GCC Storage**: 100T 크레인, GCC yard close(14:00/17:30 가변), A‑frame offloading/restoring 목표 5대/일.
- **[HVDC] Project Lightning**: Program‑level SITREP 07:30/16:00, CCU/FR/OT 순환, Empty 48h 회수.

> 각 그룹은 `hvdc:Workgroup`의 개별 인스턴스로, 고유 `kpiProfile`, `cadenceRule`, `riskPattern`을 붙인다.

---

## 8) SHACL Shapes (발췌)
```ttl
@prefix sh: <http://www.w3.org/ns/shacl#> .
@prefix hvdc: <https://example.org/hvdc#> .

hvdc:MessageShape a sh:NodeShape ;
  sh:targetClass hvdc:Message ;
  sh:property [ sh:path hvdc:sentAt ; sh:datatype xsd:dateTime ; sh:minCount 1 ] ;
  sh:property [ sh:path hvdc:sender ; sh:class hvdc:Participant ; sh:minCount 1 ] ;
  sh:property [ sh:path hvdc:hasTag ; sh:class hvdc:Tag ; sh:minCount 1 ] ;
  sh:property [ sh:path hvdc:attachments ; sh:class hvdc:Document ; sh:minCount 0 ] .

hvdc:DocumentShape a sh:NodeShape ;
  sh:targetClass hvdc:Document ;
  sh:property [ sh:path hvdc:fileName ; sh:pattern "^[0-9]{8}_[A-Z0-9\-]+_[A-Z]+_v[0-9]{2}$" ] ;
  sh:property [ sh:path hvdc:docType ; sh:in ("CIPL" "BL" "DO" "Permit") ] .
```

---

## 9) JSON‑LD Context (발췌)
```json
{
  "@context": {
    "hvdc": "https://example.org/hvdc#",
    "schema": "http://schema.org/",
    "skos": "http://www.w3.org/2004/02/skos/core#",
    "prov": "http://www.w3.org/ns/prov#",
    "Workgroup": "hvdc:Workgroup",
    "Message": "hvdc:Message",
    "Tag": "hvdc:Tag",
    "Site": "hvdc:Site",
    "Asset": "hvdc:Asset",
    "Cargo": "hvdc:Cargo",
    "Approval": "hvdc:Approval",
    "Action": "hvdc:Action",
    "fileName": "hvdc:fileName",
    "hasTag": {"@id": "hvdc:hasTag", "@type": "@id"},
    "about": {"@id": "hvdc:about", "@type": "@id"},
    "evokes": {"@id": "hvdc:evokes", "@type": "@id"}
  }
}
```

---

## 10) KPIs & SPARQL (예시)
**KPI 정의**
- `kpi:sla_on_time_rate`: (SLA 내 응답 건수 / 총 SLA 대상)
- `kpi:tag_usage_rate`: (고정 9태그 사용 메시지 / 총 메시지)
- `kpi:attachment_completeness`: (필수 문서 첨부 메시지 / 필수문서 요구 메시지)
- `kpi:daily_summary_on_time`: (08:30/17:30 요약 정시/총 요약)

**SLA 위반 메시지 조회**
```sparql
PREFIX hvdc: <https://example.org/hvdc#>
SELECT ?msg ?sentAt ?owner ?class
WHERE {
  ?msg a hvdc:Message ; hvdc:hasTag hvdc:URGENT ; hvdc:sentAt ?sentAt .
  ?clock a hvdc:SLAClock ; hvdc:attachesTo ?msg ; hvdc:classOfService ?class ; hvdc:breachedAt ?b .
  OPTIONAL { ?clock hvdc:owner ?owner }
}
ORDER BY DESC(?b)
```

**그룹별 상위 키워드**
```sparql
SELECT ?group (SAMPLE(?kw) AS ?topKw) (COUNT(*) AS ?cnt)
WHERE {
  ?m a hvdc:Message ; hvdc:inWorkgroup ?group ; hvdc:hasKeyword ?kw .
}
GROUP BY ?group
ORDER BY DESC(?cnt)
```

---

## 11) Automation Hooks (운영)
- **일일 요약**: 08:30 / 17:30 — `Workgroup → Thread(scan 24h) → Action Board` 자동 생성·배포.
- **주간 리포트**: 목 16:00 — KPI 카드(태그 사용률, SLA 준수율, 첨부 완전성, SITREP 정시율) 생성.
- **에스컬레이션 체인**: 담당→리드→PM→관리 (연락처는 PII 정책에 따라 마스킹 공유).
- **Keyword→Task**: `[CRANE]`→ 장비 예약 시트, `[GATE]`→ 게이트패스 제출 폼, `[MANIFEST]`→ PL/Manifest 체크.

---

## 12) Implementation Checklist
- [ ] **Ontology**: core.ttl(SKOS/도메인 보카블러리 포함), roles.ttl(RACI), policy.ttl(SLA/보안/파일명)
- [ ] **Validation**: shacl shapes(메시지/문서/SLAClock), PII 마스킹 유닛테스트
- [ ] **Ingest**: WhatsApp txt → `Message(header/body/attachments)` 파서 + 동의어 정규화
- [ ] **Action Mapper**: 키워드→태스크 룰셋(JSON) + 회귀테스트
- [ ] **KPI Jobs**: 일일/주간 배치 + 시각화(대시보드)
- [ ] **Export**: JSON‑LD & Parquet; Triple Store 적재(RDFLib/HTTP)

---

## 13) Change Log
- **v1.0 (2025-10-19)**: 초기 발행 — Master Policy/SLA/태그/파일명 규칙을 온톨로지 계층으로 재구성, 그룹별 도메인 보카블러리·Action 맵 반영.

