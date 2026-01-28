---
title: "Work Group Chat System - WhatsApp/TG/Email Fusion"
type: "ontology-design"
domain: "multi-channel-communication"
version: "1.0"
date: "2025-01-19"
tags: ["ontology", "whatsapp", "telegram", "email", "chat", "hvdc", "fusion"]
timezone: "Asia/Dubai"
pii_policy: "담당자 이름 공개, 전화번호·이메일 마스킹"
status: "active"
---

# Ontology — Work Group Chat System (WhatsApp/TG/Email Fusion) v1.0

## 개요

*EN-KR one-liner: Any chat → normalize → link to real-world work → auto-SLA & SITREP.*

**Scope**: HVDC Logistics (Samsung C&T · ADNOC · DSV) — WhatsApp 그룹방 중심, Telegram/Email 연동 포함  
**Timezone**: Asia/Dubai (UTC+04)  
**PII Policy**: 담당자 **이름 공개**, 전화번호·이메일은 마스킹(***1234)  
**FX/Units**: 날짜 ISO, 수치 2 decimals, 시간/기간은 h 단위(예: 24.00h)

---

## 1) 목적·문제정의
- **문제**: 그룹방 메시지는 빠르지만 흩어지고 휘발된다 → *업무 실체*와 연결되지 않으면 KPI, 책임, 비용, 리스크가 흐릿해진다.  
- **목표**: 메시지→의미(태그/엔티티)→업무 그래프(Shipment/Invoice/WH/Marine Ops)로 **온톨로지-퍼스트** 전환.  
- **핵심 가설**: *Multi‑Key Identity Graph*—어떤 키든(BL/DO/Container/Case/Invoice/hvdc_code) 메시지에서 발견되면 **동일 실체 클러스터**에 연결되고, 아래 흐름(ETA/Action/Cost/SITREP)까지 자동 캐스케이드.

---

## 2) 코어 클래스(OWL/RDF)
### 2.1 네임스페이스 예시
- `hvdc:` http://samsung.com/hvdc#  
- `org:` http://www.w3.org/ns/org#  
- `schema:` http://schema.org/  

### 2.2 핵심 클래스
- **hvdc:ChatGroup** (WA/TG/Email Thread)  
- **hvdc:Channel** (WhatsApp|Telegram|Email)  
- **hvdc:Message** (텍스트/미디어/명령)  
- **hvdc:Person** (담당자, Review/Approver 포함)  
- **hvdc:Role** (PIC/Owner/Reviewer/Approver)  
- **hvdc:Tag** ([URGENT][ACTION][ETA][COST][GATE][CRANE][MANIFEST][RISK]… 화이트리스트)  
- **hvdc:ActionItem** (할 일/조치, due/owner/SLA)  
- **hvdc:Decision** (결정/승인, 근거 메시지 링크)  
- **hvdc:Attachment** (파일/이미지/도면)  
- **hvdc:EntityRef** (Shipment/Voyage/Invoice/Case/DO/BL/Container/Warehouse/WorkOrder 등 도메인 객체의 프록시)  
- **hvdc:Event** (ETA/ETD/Delay/OSDR/COA 제출 등 상태 이벤트)  

### 2.3 핵심 속성
- `hvdc:hasMessage (Group→Message)`  
- `hvdc:sentBy (Message→Person)` / `hvdc:sentAt (xsd:dateTime)`  
- `hvdc:hasTag (Message→Tag)` / `hvdc:mentions (Message→EntityRef)`  
- `hvdc:referencesFile (Message→Attachment)`  
- `hvdc:assertsAction (Message→ActionItem)` / `hvdc:assertsDecision (Message→Decision)`  
- `hvdc:hasOwner (ActionItem→Person)` / `hvdc:dueAt` / `hvdc:slaHours (xsd:decimal)` / `hvdc:status`  
- `hvdc:linksTo (EntityRef→{Shipment/Invoice/…})`  
- `hvdc:channelType (ChatGroup→Channel)`  

---

## 3) 아이덴티티 & Multi‑Key 링크
- **키 추출기**: 정규표현식/룰+NLP(옵션)로 메시지에서 BL/DO/Container/Case/Invoice/`hvdc-adopt-xxxx-xxxx` 패턴 추출.  
- **동일성 해석**: 추출된 키는 `hvdc:EntityRef`로 표준화 → 동일 Shipment/Item/Doc 클러스터에 **UUIDv5** 결정적 식별자로 묶음.  
- **캐스케이드**: EntityRef가 묶이면 해당 메시지는 자동으로 **도메인 스트림**(Invoice Δ%, WH 이동, Marine ETA, Gate Pass)로 흘러들어감.

> EN-KR: *Any-key in → Resolve → Cluster → Downstream views (SITREP, KPI, Cost).*

---

## 4) 태깅 규칙 (화이트리스트)
| Tag | 의미/추론 | 의무 속성(Shape) |
|---|---|---|
| **[ACTION]** | 할 일 생성 | owner, dueAt, slaHours(기본 24.00h), scope(EntityRef) |
| **[ETA]** | 도착/준비 예상 | time(UTC+4), location, scope |
| **[COST]** | 비용/요율/Δ% 논의 | currency(AED/USD), amount, src(계약/비계약), scope |
| **[GATE]** | Gate/Gate Pass/Port Pass | port/site, date/time, docRef |
| **[CRANE]** | 크레인/리거/장비 수배 | tonnage/boom/reach/slots, date/time |
| **[MANIFEST]** | CIPL/PL/BL/DO/Manifest 참조 | docRef, version/date |
| **[RISK]** | 위험/차질/차단요인 | severity(1–5), mitigation, owner |
| **[URGENT]** | SLA 가속/에스컬레이션 | priority=High, notify=Lead |

- **자동 보정**: `[ACTION]`인데 owner/due가 없으면 Shape 위반으로 **WARN**. `[ETA]`가 과거시간이면 **ANOMALY**.  
- **이름 표기**: 실명 허용, 전화번호·이메일은 ***마스킹***.

---

## 5) SHACL Shapes (발췌)
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

## 6) 데이터 파이프라인 (Ingest→Normalize→Reason→Serve)
1) **Ingest**: WhatsApp Export(.txt/.zip) / WA API 게이트웨이 / Email(IMAP) / TG Bot.  
2) **Normalize**: 헤더 정규화, 사람 이름 매핑, 전화·이메일 마스킹, 메시지 UUIDv5(`group_id + timestamp + text`).  
3) **Link**: 키 추출→`EntityRef` 생성→`linksTo`(Shipment/Invoice/Case/Container/BL/DO).  
4) **Reason**: 태그 기반 액션·결정·이벤트 추론, ETA 파싱, Gate Pass/COA/OSDR 규칙, **SLA 타이머**(24.00h 기본, 그룹별 재정의).  
5) **Serve**: **SITREP 생성**(AM/PM), **KPI 대시보드**(응답TAT, 미해결 ACTION, RISK 큐), **/status** 응답.  

---

## 7) KPI & 알림
- **응답 TAT(hrs)** = 첫 태그 메시지→첫 유의미 응답까지 시간(목표 ≤ 2.00h)  
- **Action Closure Rate(%)** = 주간 완료/신규 (목표 ≥ 85.00%)  
- **Overdue Actions(#)** = 기한 초과 조치 수 (목표 0.00)  
- **Risk Open Days(일)** = RISK 태그 후 해결까지 일수 (목표 ≤ 2.00d)  
- **에스컬레이션**: SLA 초과 시 그룹 리드·PMT 자동 멘션.

---

## 8) SPARQL 질의(샘플)
**8.1 미해결 ACTION (24h 초과)**
```sparql
PREFIX hvdc: <http://samsung.com/hvdc#>
SELECT ?action ?owner ?due ?msg WHERE {
  ?action a hvdc:ActionItem ; hvdc:status "OPEN" ; hvdc:dueAt ?due ; hvdc:hasOwner ?owner ; hvdc:sourceMessage ?msg .
  FILTER(NOW() > (?due + "P0D"^^xsd:dayTimeDuration))
}
```
**8.2 특정 hvdc_code 관련 대화 스레드**
```sparql
SELECT ?msg ?at ?by WHERE {
  ?ref a hvdc:EntityRef ; hvdc:hasHVDCCode "hvdc-adopt-0123-4567" .
  ?msg a hvdc:Message ; hvdc:mentions ?ref ; hvdc:sentAt ?at ; hvdc:sentBy ?by .
} ORDER BY ?at
```
**8.3 ETA가 12h 내 도래하는 이벤트**
```sparql
SELECT ?entity ?eta WHERE {
  ?evt a hvdc:Event ; hvdc:eventType "ETA" ; hvdc:etaAt ?eta ; hvdc:forEntity ?entity .
  FILTER(?eta < NOW() + "PT12H"^^xsd:dayTimeDuration)
}
```

---

## 9) SITREP & 커맨드 연동
- **SITREP(AM/PM)**: Group→Message 집계 → Action/Risk/ETA/Cost/Manifest 핵심 1‑Pager 생성.  
- **명령**: `/status`(오늘의 미결 ACTION/ETA/CRANE/GATE 요약), `/bot sync`(TG/SharePoint 동기화), `/approval list`(임계 승인 D‑상태), `/logi-master report --KRsummary`(주요 이슈 리포트).  

---

## 10) 보안·컴플라이언스
- **PDPL/GDPR**: 최소수집·목적제한·보존기간 정책.  
- **PII 처리**: 이름 공개, 전화/이메일은 마스킹 및 SHA‑256 해시로 링크.  
- **감사 트레일**: 메시지 해시, 변환 로그, 규칙 버전, 온톨로지 버전.  

---

## 11) 옵션(모델링 전략)
| 옵션 | 설명 | 장점 | 단점 | T(주) |
|---|---|---|---|---|
| **A. WA‑First** | WhatsApp 중심, Email/TG는 참조 | 구현 빠름, 현장 적합 | 채널 확장성 제한 | 2.00 |
| **B. Multi‑Channel** | WA/TG/Email 동등 취급 | 유연/확장, 채널 독립 | 파이프라인 복잡 | 4.00 |
| **C. Domain‑First** | 도메인 이벤트 중심, 채널은 입력 | 분석 최적, KPI 직접추적 | 초기 모델링 비용 | 5.00 |

---

## 12) 단계별 로드맵 (Prepare→Pilot→Build→Operate→Scale)
- **Prepare(1.00주)**: 그룹 리스트·화이트리스트 태그 확정, 기본 SHA‑마스킹 룰, 키 추출기 정규식 세트. *Owner: Ops Lead.*  
- **Pilot(2.00주)**: 3개 그룹(A/B/C) 실적용, `/status`·SITREP 자동배포, KPI 게이트 설정. *Owner: PMO.*  
- **Build(4.00주)**: SHACL 통합, SLA 엔진, AI 보조 파서(NER), TG/Email 커넥터. *Owner: Data.*  
- **Operate(지속)**: 주간 KPI 리뷰, 규칙 Diff‑Watch, 비용/리스크 연계. *Owner: Leads.*  
- **Scale(지속)**: 그룹 전면 확대, 계약/인보이스/창고/해상 연계 심화. *Owner: All.*

---

## 13) TTL/JSON‑LD 스니펫(요약)
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
**JSON‑LD Context (발췌)**
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

## 14) 구현 포인트(엔지니어링)
- **Parser**: PyParsing/regex + dateparser(tz=Asia/Dubai) + phone/email mask + UUIDv5.  
- **Storage**: RDFLib + DuckDB/opensearch 인덱스(메시지 해시·태그·키).  
- **Rules**: SHACL + SPARQL Update + 간단 룰엔진(ETA/COA/Gate/CRANE).  
- **Dash**: Streamlit/Metabase — SLA 게이지, 미결 ACTION, RISK 타일.  
- **Ops**: `/bot sync`로 SharePoint·TG 양방향, `/status`로 데일리 요약.

---

## 15) QA·리스크 체크리스트
- [ ] 키 추출 정규식 최신화(BL/DO/Container/Case/Invoice/hvdc_code)  
- [ ] 태그 화이트리스트 일치율 ≥ 95.00%  
- [ ] SHACL 위반율 ≤ 2.00% (초과 시 **중단**)  
- [ ] PII 마스킹 누락 0.00건  
- [ ] SLA 초과 자동 에스컬레이션 정상 동작  

---

## 16) 부록 — 그룹 타입 & 기본 PIC(예시)
| 그룹 | 타입 | PIC(Owner) | 노트 |
|---|---|---|---|
| HVDC Project Lightning | Ops/Marine | **Sajid H. Khan** | Marine ETA/CRANE/GATE 중심 |
| DSV Jay Manaloto | 3PL Ops | **Jay Manaloto** | DO/Trucking/WH 배차 |
| DSV Minhaj | 3PL Ops | **Minhaj** | Gate Pass/Port Admin |
| CJ ICM Logistics | Forwarding | **CJ ICM Lead** | CIPL/Manifest/Customs |

> *Example only — 실제 PIC은 운영표 기준 업데이트.*

---

### 끝. (버전: v1.0, 2025-10-19, Asia/Dubai)

