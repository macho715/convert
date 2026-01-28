---
title: "Chat Communication System - WhatsApp/Telegram Ontology"
type: "ontology-design"
domain: "multi-channel-communication"
sub-domains: ["whatsapp-workgroups", "multi-channel-fusion"]
version: "unified-1.0"
date: "2025-01-19"
tags: ["ontology", "whatsapp", "telegram", "email", "chat", "hvdc", "unified"]
timezone: "Asia/Dubai"
pii_policy: "담당자 이름 공개, 전화번호·이메일 마스킹"
status: "active"
---

# Part 1: WhatsApp Workgroups System

## Ontology-First Spec — WhatsApp Workgroups System (HVDC)

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


# Ontology ??Work Group Chat System (WhatsApp/TG/Email Fusion) v1.0

## 媛쒖슂

*EN-KR one-liner: Any chat ??normalize ??link to real-world work ??auto-SLA & SITREP.*

**Scope**: HVDC Logistics (Samsung C&T 쨌 ADNOC 쨌 DSV) ??WhatsApp 洹몃９諛?以묒떖, Telegram/Email ?곕룞 ?ы븿
**Timezone**: Asia/Dubai (UTC+04)
**PII Policy**: ?대떦??**?대쫫 怨듦컻**, ?꾪솕踰덊샇쨌?대찓?쇱? 留덉뒪??***1234)
**FX/Units**: ?좎쭨 ISO, ?섏튂 2 decimals, ?쒓컙/湲곌컙? h ?⑥쐞(?? 24.00h)

---

## 1) 紐⑹쟻쨌臾몄젣?뺤쓽
- **臾몄젣**: 洹몃９諛?硫붿떆吏??鍮좊Ⅴ吏留??⑹뼱吏怨??섎컻?쒕떎 ??*?낅Т ?ㅼ껜*? ?곌껐?섏? ?딆쑝硫?KPI, 梨낆엫, 鍮꾩슜, 由ъ뒪?ш? ?먮┸?댁쭊??
- **紐⑺몴**: 硫붿떆吏?믪쓽誘??쒓렇/?뷀떚???믪뾽臾?洹몃옒??Shipment/Invoice/WH/Marine Ops)濡?**?⑦넧濡쒖?-?쇱뒪??* ?꾪솚.
- **?듭떖 媛??*: *Multi?멚ey Identity Graph*?붿뼱???ㅻ뱺(BL/DO/Container/Case/Invoice/hvdc_code) 硫붿떆吏?먯꽌 諛쒓껄?섎㈃ **?숈씪 ?ㅼ껜 ?대윭?ㅽ꽣**???곌껐?섍퀬, ?꾨옒 ?먮쫫(ETA/Action/Cost/SITREP)源뚯? ?먮룞 罹먯뒪耳?대뱶.

---

## 2) 肄붿뼱 ?대옒??OWL/RDF)
### 2.1 ?ㅼ엫?ㅽ럹?댁뒪 ?덉떆
- `hvdc:` http://samsung.com/hvdc#
- `org:` http://www.w3.org/ns/org#
- `schema:` http://schema.org/

### 2.2 ?듭떖 ?대옒??- **hvdc:ChatGroup** (WA/TG/Email Thread)
- **hvdc:Channel** (WhatsApp|Telegram|Email)
- **hvdc:Message** (?띿뒪??誘몃뵒??紐낅졊)
- **hvdc:Person** (?대떦?? Review/Approver ?ы븿)
- **hvdc:Role** (PIC/Owner/Reviewer/Approver)
- **hvdc:Tag** ([URGENT][ACTION][ETA][COST][GATE][CRANE][MANIFEST][RISK]???붿씠?몃━?ㅽ듃)
- **hvdc:ActionItem** (????議곗튂, due/owner/SLA)
- **hvdc:Decision** (寃곗젙/?뱀씤, 洹쇨굅 硫붿떆吏 留곹겕)
- **hvdc:Attachment** (?뚯씪/?대?吏/?꾨㈃)
- **hvdc:EntityRef** (Shipment/Voyage/Invoice/Case/DO/BL/Container/Warehouse/WorkOrder ???꾨찓??媛앹껜???꾨줉??
- **hvdc:Event** (ETA/ETD/Delay/OSDR/COA ?쒖텧 ???곹깭 ?대깽??

### 2.3 ?듭떖 ?띿꽦
- `hvdc:hasMessage (Group?묺essage)`
- `hvdc:sentBy (Message?뭁erson)` / `hvdc:sentAt (xsd:dateTime)`
- `hvdc:hasTag (Message?뭈ag)` / `hvdc:mentions (Message?묮ntityRef)`
- `hvdc:referencesFile (Message?묨ttachment)`
- `hvdc:assertsAction (Message?묨ctionItem)` / `hvdc:assertsDecision (Message?묭ecision)`
- `hvdc:hasOwner (ActionItem?뭁erson)` / `hvdc:dueAt` / `hvdc:slaHours (xsd:decimal)` / `hvdc:status`
- `hvdc:linksTo (EntityRef??Shipment/Invoice/??)`
- `hvdc:channelType (ChatGroup?묬hannel)`

---

## 3) ?꾩씠?댄떚??& Multi?멚ey 留곹겕
- **??異붿텧湲?*: ?뺢퇋?쒗쁽??猷?NLP(?듭뀡)濡?硫붿떆吏?먯꽌 BL/DO/Container/Case/Invoice/`hvdc-adopt-xxxx-xxxx` ?⑦꽩 異붿텧.
- **?숈씪???댁꽍**: 異붿텧???ㅻ뒗 `hvdc:EntityRef`濡??쒖??????숈씪 Shipment/Item/Doc ?대윭?ㅽ꽣??**UUIDv5** 寃곗젙???앸퀎?먮줈 臾띠쓬.
- **罹먯뒪耳?대뱶**: EntityRef媛 臾띠씠硫??대떦 硫붿떆吏???먮룞?쇰줈 **?꾨찓???ㅽ듃由?*(Invoice ?%, WH ?대룞, Marine ETA, Gate Pass)濡??섎윭?ㅼ뼱媛?

> EN-KR: *Any-key in ??Resolve ??Cluster ??Downstream views (SITREP, KPI, Cost).*

---

## 4) ?쒓퉭 洹쒖튃 (?붿씠?몃━?ㅽ듃)
| Tag | ?섎?/異붾줎 | ?섎Т ?띿꽦(Shape) |
|---|---|---|
| **[ACTION]** | ?????앹꽦 | owner, dueAt, slaHours(湲곕낯 24.00h), scope(EntityRef) |
| **[ETA]** | ?꾩갑/以鍮??덉긽 | time(UTC+4), location, scope |
| **[COST]** | 鍮꾩슜/?붿쑉/?% ?쇱쓽 | currency(AED/USD), amount, src(怨꾩빟/鍮꾧퀎??, scope |
| **[GATE]** | Gate/Gate Pass/Port Pass | port/site, date/time, docRef |
| **[CRANE]** | ?щ젅??由ш굅/?λ퉬 ?섎같 | tonnage/boom/reach/slots, date/time |
| **[MANIFEST]** | CIPL/PL/BL/DO/Manifest 李몄“ | docRef, version/date |
| **[RISK]** | ?꾪뿕/李⑥쭏/李⑤떒?붿씤 | severity(1??), mitigation, owner |
| **[URGENT]** | SLA 媛???먯뒪而щ젅?댁뀡 | priority=High, notify=Lead |

- **?먮룞 蹂댁젙**: `[ACTION]`?몃뜲 owner/due媛 ?놁쑝硫?Shape ?꾨컲?쇰줈 **WARN**. `[ETA]`媛 怨쇨굅?쒓컙?대㈃ **ANOMALY**.
- **?대쫫 ?쒓린**: ?ㅻ챸 ?덉슜, ?꾪솕踰덊샇쨌?대찓?쇱? ***留덉뒪??**.

---

## 5) SHACL Shapes (諛쒖톸)
```turtle
@prefix sh: <http://www.w3.org/ns/shacl#> .
@prefix hvdc: <http://samsung.com/hvdc#> .
@prefix xsd: <http://www.w3.org/2001/XMLSchema#> .

hvdc:ActionShape a sh:NodeShape ;
  sh:targetClass hvdc:ActionItem ;
  sh:property [ sh:path hvdc:hasOwner ; sh:minCount 1 ] ;
  sh:property [ sh:path hvdc:dueAt ; sh:datatype xsd:dateTime ; sh:minCount 1 ] ;
  sh:property [ sh:path hvdc:slaHours ; sh:datatype xsd:decimal ; sh:minInclusive 0.00 ] .

hvdc:MessageShape a sh:NodeShape ;
  sh:targetClass hvdc:Message ;
  sh:property [ sh:path hvdc:sentBy ; sh:minCount 1 ] ;
  sh:property [ sh:path hvdc:sentAt ; sh:datatype xsd:dateTime ; sh:minCount 1 ] ;
  sh:property [ sh:path hvdc:hasTag ; sh:minCount 0 ] .
```

---

## 6) ?곗씠???뚯씠?꾨씪??(Ingest?묿ormalize?뭃eason?뭆erve)
1) **Ingest**: WhatsApp Export(.txt/.zip) / WA API 寃뚯씠?몄썾??/ Email(IMAP) / TG Bot.
2) **Normalize**: ?ㅻ뜑 ?뺢퇋?? ?щ엺 ?대쫫 留ㅽ븨, ?꾪솕쨌?대찓??留덉뒪?? 硫붿떆吏 UUIDv5(`group_id + timestamp + text`).
3) **Link**: ??異붿텧??EntityRef` ?앹꽦??linksTo`(Shipment/Invoice/Case/Container/BL/DO).
4) **Reason**: ?쒓렇 湲곕컲 ?≪뀡쨌寃곗젙쨌?대깽??異붾줎, ETA ?뚯떛, Gate Pass/COA/OSDR 洹쒖튃, **SLA ??대㉧**(24.00h 湲곕낯, 洹몃９蹂??ъ젙??.
5) **Serve**: **SITREP ?앹꽦**(AM/PM), **KPI ??쒕낫??*(?묐떟TAT, 誘명빐寃?ACTION, RISK ??, **/status** ?묐떟.

---

## 7) KPI & ?뚮┝
- **?묐떟 TAT(hrs)** = 泥??쒓렇 硫붿떆吏?믪껀 ?좎쓽誘??묐떟源뚯? ?쒓컙(紐⑺몴 ??2.00h)
- **Action Closure Rate(%)** = 二쇨컙 ?꾨즺/?좉퇋 (紐⑺몴 ??85.00%)
- **Overdue Actions(#)** = 湲고븳 珥덇낵 議곗튂 ??(紐⑺몴 0.00)
- **Risk Open Days(??** = RISK ?쒓렇 ???닿껐源뚯? ?쇱닔 (紐⑺몴 ??2.00d)
- **?먯뒪而щ젅?댁뀡**: SLA 珥덇낵 ??洹몃９ 由щ뱶쨌PMT ?먮룞 硫섏뀡.

---

## 8) SPARQL 吏덉쓽(?섑뵆)
**8.1 誘명빐寃?ACTION (24h 珥덇낵)**
```sparql
PREFIX hvdc: <http://samsung.com/hvdc#>
SELECT ?action ?owner ?due ?msg WHERE {
  ?action a hvdc:ActionItem ; hvdc:status "OPEN" ; hvdc:dueAt ?due ; hvdc:hasOwner ?owner ; hvdc:sourceMessage ?msg .
  FILTER(NOW() > (?due + "P0D"^^xsd:dayTimeDuration))
}
```
**8.2 ?뱀젙 hvdc_code 愿??????ㅻ젅??*
```sparql
SELECT ?msg ?at ?by WHERE {
  ?ref a hvdc:EntityRef ; hvdc:hasHVDCCode "hvdc-adopt-0123-4567" .
  ?msg a hvdc:Message ; hvdc:mentions ?ref ; hvdc:sentAt ?at ; hvdc:sentBy ?by .
} ORDER BY ?at
```
**8.3 ETA媛 12h ???꾨옒?섎뒗 ?대깽??*
```sparql
SELECT ?entity ?eta WHERE {
  ?evt a hvdc:Event ; hvdc:eventType "ETA" ; hvdc:etaAt ?eta ; hvdc:forEntity ?entity .
  FILTER(?eta < NOW() + "PT12H"^^xsd:dayTimeDuration)
}
```

---

## 9) SITREP & 而ㅻ㎤???곕룞
- **SITREP(AM/PM)**: Group?묺essage 吏묎퀎 ??Action/Risk/ETA/Cost/Manifest ?듭떖 1?멠ager ?앹꽦.
- **紐낅졊**: `/status`(?ㅻ뒛??誘멸껐 ACTION/ETA/CRANE/GATE ?붿빟), `/bot sync`(TG/SharePoint ?숆린??, `/approval list`(?꾧퀎 ?뱀씤 D?묒긽??, `/logi-master report --KRsummary`(二쇱슂 ?댁뒋 由ы룷??.

---

## 10) 蹂댁븞쨌而댄뵆?쇱씠?몄뒪
- **PDPL/GDPR**: 理쒖냼?섏쭛쨌紐⑹쟻?쒗븳쨌蹂댁〈湲곌컙 ?뺤콉.
- **PII 泥섎━**: ?대쫫 怨듦컻, ?꾪솕/?대찓?쇱? 留덉뒪??諛?SHA??56 ?댁떆濡?留곹겕.
- **媛먯궗 ?몃젅??*: 硫붿떆吏 ?댁떆, 蹂??濡쒓렇, 洹쒖튃 踰꾩쟾, ?⑦넧濡쒖? 踰꾩쟾.

---

## 11) ?듭뀡(紐⑤뜽留??꾨왂)
| ?듭뀡 | ?ㅻ챸 | ?μ젏 | ?⑥젏 | T(二? |
|---|---|---|---|---|
| **A. WA?멑irst** | WhatsApp 以묒떖, Email/TG??李몄“ | 援ы쁽 鍮좊쫫, ?꾩옣 ?곹빀 | 梨꾨꼸 ?뺤옣???쒗븳 | 2.00 |
| **B. Multi?멌hannel** | WA/TG/Email ?숇벑 痍④툒 | ?좎뿰/?뺤옣, 梨꾨꼸 ?낅┰ | ?뚯씠?꾨씪??蹂듭옟 | 4.00 |
| **C. Domain?멑irst** | ?꾨찓???대깽??以묒떖, 梨꾨꼸? ?낅젰 | 遺꾩꽍 理쒖쟻, KPI 吏곸젒異붿쟻 | 珥덇린 紐⑤뜽留?鍮꾩슜 | 5.00 |

---

## 12) ?④퀎蹂?濡쒕뱶留?(Prepare?뭁ilot?묪uild?뭀perate?뭆cale)
- **Prepare(1.00二?**: 洹몃９ 由ъ뒪?맞룻솕?댄듃由ъ뒪???쒓렇 ?뺤젙, 湲곕낯 SHA?묐쭏?ㅽ궧 猷? ??異붿텧湲??뺢퇋???명듃. *Owner: Ops Lead.*
- **Pilot(2.00二?**: 3媛?洹몃９(A/B/C) ?ㅼ쟻?? `/status`쨌SITREP ?먮룞諛고룷, KPI 寃뚯씠???ㅼ젙. *Owner: PMO.*
- **Build(4.00二?**: SHACL ?듯빀, SLA ?붿쭊, AI 蹂댁“ ?뚯꽌(NER), TG/Email 而ㅻ꽖?? *Owner: Data.*
- **Operate(吏??**: 二쇨컙 KPI 由щ럭, 洹쒖튃 Diff?멬atch, 鍮꾩슜/由ъ뒪???곌퀎. *Owner: Leads.*
- **Scale(吏??**: 洹몃９ ?꾨㈃ ?뺣?, 怨꾩빟/?몃낫?댁뒪/李쎄퀬/?댁긽 ?곌퀎 ?ы솕. *Owner: All.*

---

## 13) TTL/JSON?멛D ?ㅻ땲???붿빟)
**TTL**
```turtle
@prefix hvdc: <http://samsung.com/hvdc#> .
@prefix xsd: <http://www.w3.org/2001/XMLSchema#> .

hvdc:Group_HVDC a hvdc:ChatGroup ; hvdc:channelType "WhatsApp" .

hvdc:Msg_20251019_210200 a hvdc:Message ;
  hvdc:sentBy hvdc:Person_Minhaj ;
  hvdc:sentAt "2025-10-19T21:02:00+04:00"^^xsd:dateTime ;
  hvdc:hasTag hvdc:ACTION ;
  hvdc:mentions hvdc:Ref_Container_ABC123 ;
  hvdc:assertsAction hvdc:Act_001 .

hvdc:Act_001 a hvdc:ActionItem ;
  hvdc:hasOwner hvdc:Person_Jay ;
  hvdc:dueAt "2025-10-20T10:30:00+04:00"^^xsd:dateTime ;
  hvdc:slaHours 24.00 ;
  hvdc:status "OPEN" .
```
**JSON?멛D Context (諛쒖톸)**
```json
{
  "@context": {
    "hvdc": "http://samsung.com/hvdc#",
    "ChatGroup": "hvdc:ChatGroup",
    "Message": "hvdc:Message",
    "sentBy": {"@id": "hvdc:sentBy", "@type": "@id"},
    "sentAt": {"@id": "hvdc:sentAt", "@type": "xsd:dateTime"}
  }
}
```

---

## 14) 援ы쁽 ?ъ씤???붿??덉뼱留?
- **Parser**: PyParsing/regex + dateparser(tz=Asia/Dubai) + phone/email mask + UUIDv5.
- **Storage**: RDFLib + DuckDB/opensearch ?몃뜳??硫붿떆吏 ?댁떆쨌?쒓렇쨌??.
- **Rules**: SHACL + SPARQL Update + 媛꾨떒 猷곗뿏吏?ETA/COA/Gate/CRANE).
- **Dash**: Streamlit/Metabase ??SLA 寃뚯씠吏, 誘멸껐 ACTION, RISK ???
- **Ops**: `/bot sync`濡?SharePoint쨌TG ?묐갑?? `/status`濡??곗씪由??붿빟.

---

## 15) QA쨌由ъ뒪??泥댄겕由ъ뒪??- [ ] ??異붿텧 ?뺢퇋??理쒖떊??BL/DO/Container/Case/Invoice/hvdc_code)
- [ ] ?쒓렇 ?붿씠?몃━?ㅽ듃 ?쇱튂????95.00%
- [ ] SHACL ?꾨컲????2.00% (珥덇낵 ??**以묐떒**)
- [ ] PII 留덉뒪???꾨씫 0.00嫄?- [ ] SLA 珥덇낵 ?먮룞 ?먯뒪而щ젅?댁뀡 ?뺤긽 ?숈옉

---

## 16) 遺濡???洹몃９ ???& 湲곕낯 PIC(?덉떆)
| 洹몃９ | ???| PIC(Owner) | ?명듃 |
|---|---|---|---|
| HVDC Project Lightning | Ops/Marine | **Sajid H. Khan** | Marine ETA/CRANE/GATE 以묒떖 |
| DSV Jay Manaloto | 3PL Ops | **Jay Manaloto** | DO/Trucking/WH 諛곗감 |
| DSV Minhaj | 3PL Ops | **Minhaj** | Gate Pass/Port Admin |
| CJ ICM Logistics | Forwarding | **CJ ICM Lead** | CIPL/Manifest/Customs |

> *Example only ???ㅼ젣 PIC? ?댁쁺??湲곗? ?낅뜲?댄듃.*

---

### ?? (踰꾩쟾: v1.0, 2025-10-19, Asia/Dubai)

