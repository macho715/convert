---
title: "LOGISTICS AI - HVDC Ontology Framework"
type: "ontology-design"
domain: "logistics-lifecycle"
version: "2.1"
date: "2025-01-19"
tags: ["ontology", "logistics", "ai", "hvdc", "knowledge-graph"]
status: "active"
---

# LOGISTICS AI - HVDC 물류 온톨로지

## 개요

프로젝트 전체를 **온톨로지 관점**에서 정리하면, 'HVDC 물류 생명주기'를 하나의 **지식그래프(Ontology)**로 모델링할 수 있습니다.

핵심은 **"물류행위(Activity)"가 아닌 "관계(Relation)"** 중심으로 보는 것입니다 — 사람, 문서, 물품, 절차, 시스템 간의 연결망.

__🔶 1\. Ontology Root Class__

__hvdc\-adopt\-logiontology__

__Layer__

__Ontology Domain__

__대표 엔티티__

__관계 키\(Relation\)__

__L1__

Physical Flow

Material, Cargo, Port, Site, Vessel

movesFrom, movesTo, storedAt, handledBy

__L2__

Document Flow

BL, CI, PL, COO, eDAS, MRR, OSDR

certifies, refersTo, attachedTo

__L3__

Actor Flow

SCT, JDN, ALS, ADNOC, Subcon

responsibleFor, approves, reportsTo

__L4__

Regulatory Flow

MOIAT, FANR, Customs, DOT

requiresPermit, compliesWith, auditedBy

__L5__

System Flow

eDAS, SAP, NCM, LDG

feedsDataTo, validates, monitoredBy

__🔶 2\. Core Classes \(from Workshop\)__

__Class__

__Subclass of__

__Description__

__Onto\-ID__

__Material__

Asset

자재 및 기자재\(Transformer, Cable, CCU 등\)

hvdc\-asset\-mat

__TransportEvent__

Activity

Inland, Marine, Offloading, SiteReceiving

hvdc\-act\-trans

__Storage__

Location

Yard, Warehouse, Laydown

hvdc\-loc\-stor

__Inspection__

Process

MRR, MRI, OSDR

hvdc\-proc\-insp

__Permit__

Document

PTW, Hot Work, FRA

hvdc\-doc\-perm

__Actor__

Agent

SCT, ADNOC L&S, Vendor

hvdc\-agent\-role

__PortOperation__

Activity

RORO/LOLO, Sea Fastening

hvdc\-act\-port

__🔶 3\. Relation Model \(Partial\)__

Material \-\-hasDocument\-\-> MRR

Material \-\-transportedBy\-\-> TransportEvent

TransportEvent \-\-operatedAt\-\-> Port

TransportEvent \-\-requires\-\-> Permit

Permit \-\-approvedBy\-\-> ADNOC

Storage \-\-monitoredBy\-\-> SCT

Inspection \-\-reportedAs\-\-> OSDR

Actor\(SCT\) \-\-usesSystem\-\-> eDAS

이 관계망은 logiontology\.mapping 모듈에서 RDF triple로 구현 가능:

:TR001 rdf:type :Transformer ;

       :hasDocument :MRR\_20240611 ;

       :storedAt :Mussafah\_Yard ;

       :handledBy :SCT ;

       :requiresPermit :FRA\_202405 ;

       :transportedBy :LCT\_Operation\_202405 \.

__🔶 4\. Lifecycle Ontology \(Material Handling Flow\)__

__Stage 1 – Importation__
→ hasDocument\(BL, CI, COO\) → customsClearedBy\(ADOPT\) → storedAt\(PortYard\)

__Stage 2 – Inland/Marine Transport__
→ transportedBy\(LCT/SPMT\) → requiresPermit\(DOT/FRA\) → monitoredBy\(ALS\)

__Stage 3 – Site Receiving__
→ inspectedBy\(QAQC\) → resultsIn\(MRR/OSDR\) → issuedAs\(MIS\)

__Stage 4 – Preservation & Foundation__
→ preservedBy\(HitachiStd\) → foundationBy\(Mammoet\) → approvedBy\(OE\)

__🔶 5\. Alignment with AI\-Logi\-Guide__

__Ontology Node__

__대응 모듈__

__기능적 의미__

Activity

pipeline

단계별 절차 정의

Document

rdfio, validation

eDAS·MRR 등 문서형 triple

Agent

core

역할/권한 모델

Location

mapping

Port/Site 좌표·거점

RiskEvent

reasoning

Weather\-Tie·Delay inference

Report

report

KPI/Inspection 리포트

__🔶 6\. Semantic KPI Layer \(Onto\-KPI\)__

__KPI Class__

__Onto Property__

__계산식__

__Source__

__On\-Time Delivery__

meetsETA

ETA vs Actual ≤12%

ETA MAPE Rule

__Inspection Compliance__

hasMRR

MRR Count / Total Deliveries

QC Gate

__Storage Efficiency__

occupies

Used m² / Available m²

WH Forecast

__Safety Conformance__

requiresPermit

Valid PTW/FRA %

HSE Docs

__🔶 7\. Ontological Integration View__

\[Material\]

   ⟶ \[Document: CI/PL/COO/eDAS\]

   ⟶ \[TransportEvent: LCT/SPMT\]

   ⟶ \[Location: Port → Yard → Site\]

   ⟶ \[Inspection: MRR/OSDR\]

   ⟶ \[Report: KPI/Dashboard\]

   ⟶ \[Governance: AI\-Logi\-Guide Rules\]

이 전체를 hvdc\-adopt\-ontology\.ttl로 export하면,
GitHub macho715/ontology\-insight에서 RDF 시각화 및 reasoning 연결 가능\.

__🔶 8\. 요약 메타 구조__

\{

 "Ontology":"hvdc\-adopt\-logiontology",

 "CoreClasses":\["Material","TransportEvent","Storage","Inspection","Permit","Actor","PortOperation"\],

 "PrimaryRelations":\["hasDocument","transportedBy","storedAt","requiresPermit","inspectedBy","approvedBy"\],

 "AlignmentModule":"AI\-Logi\-Guide v2\.1\+",

 "ExportFormat":\["RDF/XML","TTL","JSON\-LD"\]

\}

이 프레임이면, HVDC 프로젝트 전체가 __“문서\-행위\-공간\-주체\-규정”의 지식망__으로 정규화됩니다\.
다음 단계는 logiontology\.reasoning 모듈에서 __Rule\-based inference__ 정의 — 예컨대 “운송허가가 누락된 자재는 SiteReceiving 단계로 진행 불가” 같은 정책을 OWL constraint로 명세하면 완성됩니다\.

