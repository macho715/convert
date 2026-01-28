---
title: "INVOICE Ontology System"
type: "ontology-design"
domain: "invoice-verification"
version: "1.0"
date: "2025-01-19"
tags: ["ontology", "invoice", "cost-guard", "verification", "hvdc"]
status: "active"
---

# INVOICE - 온톨로지 퍼스트 청구서 시스템

## Executive Summary

**온톨로지-퍼스트 청구서 시스템**은 "**멀티-키 아이덴티티 그래프**(BL/Container/DO/Invoice/Case/Booking/ShipmentID/.../hvdc_code 아무 키든 OK)" 위에서 **Invoice→Line→OD Lane→RateRef→Δ%→Risk**로 한 번에 캐스케이드합니다. \(EN\-KR: Any\-key in → Resolve → Lane&Rate join → Δ% risk band\.\)
표준요율은 __Air/Container/Bulk 계약 레퍼런스__와 __Inland Trucking\(OD×Unit\) 테이블__을 온톨로지 클래스로 들고, 모든 계산은 __USD 기준·고정환율 1\.00 USD=3\.6725 AED__ 규칙을 따릅니다\.
OD 정규화·조인은 __ApprovedLaneMap/RefDestinationMap__을 통해 수행되고, 결과는 \*\*COST\-GUARD Δ% 밴드\(PASS/WARN/HIGH/CRITICAL\)\*\*로 귀결됩니다\.
감사 트레이스는 __PRISM\.KERNEL__ 포맷\(5\-line recap \+ proof\.artifact JSON\)으로 고정 형식으로 남깁니다\.

__Visual — 핵심 클래스/관계\(요약\)__

__Class__

__핵심 속성__

__관계__

__근거/조인 소스__

__결과__

hvdc:Invoice

docId, vendor, issueDate, currency

hasLine → InvoiceLine

—

상태, 총액, proof

hvdc:InvoiceLine

chargeDesc, qty, unit, draftRateUSD

hasLane → ODLane / uses → RateRef

Inland Trucking/Table, Air/Container/Bulk Rate

Δ%, cg\_band

hvdc:ODLane

origin\_norm, destination\_norm, vehicle, unit

joinedBy → ApprovedLaneMap

RefDestinationMap, Lane stats

median\_rate\_usd

hvdc:RateRef

rate\_usd, tolerance\(±3%\), source\(contract/market/special\)

per Category/Port/Dest/Unit

Air/Container/Bulk/Trucking tables

ref\_rate\_usd

hvdc:CurrencyPolicy

base=USD, fx=3\.6725

validates Invoice/Line

currency\_mismatch rule

환산/락

hvdc:RiskResult

delta\_pct, cg\_band, verdict

from Line vs Ref

COST\-GUARD bands

PASS/FAIL

자료: 표준요율 테이블\(계약\)·고정 FX 규정·Lane 정규화 지도\.

__How it works \(flow\)__

1. __키 해석\(Identity\)__: BL/Container/DO/Invoice/… 입력 → 동일 실체\(Shipment/Doc\) 클러스터 식별\. \(멀티\-키 그래프\)
2. __Lane 정규화__: 원지/착지 명칭을 __RefDestinationMap__으로 정규화 → __ApprovedLaneMap__에서 lane 통계/표준요율 후보 추출\.
3. __Rate 조인__: 라인별 __Category\+Port\+Destination\+Unit__로 계약 요율 테이블 매칭\(±3% 톨러런스\)\.
4. __Δ% & 밴드 산정__: Δ%=\(draft−ref\)/ref×100 → __PASS/WARN/HIGH/CRITICAL__ \(COST\-GUARD\)\. FX는 USD 고정\(3\.6725\)로 비교\.
5. __감사 아티팩트__: __PRISM\.KERNEL__로 5\-라인 요약 \+ JSON 증빙\(입력/계산/판정 해시\)\.

__Options \(설계 선택지\)__

1. __OWL/SHACL 엄격형__: 스키마·제약\(단위/Currency/OD 필수\)로 하드 밸리데이션\. *Pros* 규정준수↑ / *Cons* 초기 모델링 비용↑\.
2. __하이브리드형\(권장\)__: OWL\+Lane Map\+계약요율\+Δ% 밴드, 부족 구간은 유사 레인 추천\. *Pros* 커버리지↑ / *Cons* Ref 미보유 구간 튜닝 필요\.
3. __마켓레이트 보강형__: Market API\(At\-cost 항목\)에 한정 보조\. *Pros* 현실성↑ / *Cons* 출처 관리·증빙 필요\.

__Roadmap \(P→Pi→B→O→S \+ KPI\)__

- __Prepare__: RefDestinationMap 최신화, Lane 조인율≥80% 달성\.
- __Pilot__: /switch\_mode COST\-GUARD \+ /logi\-master invoice\-audit \-\-deep \-\-highlight\-mismatch로 월간 샘플 1회전\. KPI: 검증정확도 ≥97%, 자동화 ≥94%\.
- __Build__: 라인별 Δ%·밴드·증빙\(표준요율 근거 링크\) 자동 표기, 통화정책 락\.
- __Operate__: High/CrITICAL 즉시 TG 알림 \+ 반려 사유 템플릿\.
- __Scale__: Lane 그래프 스냅샷/변동 추적, 분기별 임계치 튜닝\.

__Automation notes__

- __입력 감지 →__ /switch\_mode COST\-GUARD \+ /logi\-master invoice\-audit \(OD 정규화→Rate 조인→Δ% 밴드→PASS/FAIL 표\)\.
- __표준 근거__: Air/Container/Bulk/Trucking 계약표 \+ Inland Ref\(정리본\)\.
- __감사 포맷__: PRISM 5\-라인 \+ JSON proof\(해시\)\.

__QA / Gap 체크__

- Line 단위가 __per truck vs per RT__ 혼재 시 단위환산 룰 적용? \(컨/Bulk 교차 구간\)
- __CURRENCY\_MISMATCH__/고정 FX\(3\.6725\) 락 확인?
- Lane 조인 실패\(REF\_MISSING\) 건은 유사도≥0\.60 Top\-3 제안 출력?
- 증빙 링크\(표준요율 원천, Lane 통계\)와 PRISM proof 동시 첨부?

가정: \(i\) 계약표는 최신판으로 동기화됨, \(ii\) 환율 고정 정책 유지, \(iii\) SHACL Shapes는 내부 표준에 따라 배포됨\.

__원한다면, 너한테 맞춘 “클래스·프로퍼티 TTL 스켈레톤 \+ SHACL” 바로 뽑아줄게\.__

