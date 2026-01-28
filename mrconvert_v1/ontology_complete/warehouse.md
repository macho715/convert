__ExecSummary__

창고 pjt를 __온톨로지\(지식그래프\) 관점__으로 보면, 엑셀/ERP의 각 행은 TransportEvent\(이동\), StockSnapshot\(재고 스냅샷\), Invoice\(청구\), Case\(개별 케이스\) 같은 __클래스__로 귀속되고, 열들은 hasDate/hasLocation/hasQuantity/hasLogisticsFlowCode 같은 __속성__으로 정규화됩니다\. 이 구조가 “창고 트랙\(WH\)”·“현장 트랙\(Site\)”·“Flow Code\(0–4\)”를 한 장의 그래프로 __동일 실체__에 묶어 줍니다\. \(Any\-key in → Resolve→Cluster→Downstream\)   
매핑된 데이터는 __RDF/OWL__로 변환되어 SPARQL로 검증/집계가 가능하고, 비용 분류\(OFCO\)나 월별 입출고·재고·SQM 과금까지 __한 체계__에서 굴러갑니다\.   
핵심은 “2\-트랙 날짜 컬럼\(창고 vs 현장\)”과 __시간순 출고 판정__·__이중계산 방지__·__Flow 0–4 일관성__을 코드 레벨로 보증하는 것입니다\.

__Visual — Ontology Map \(요약표\)__

__Layer__

__Ontology 객체/속성__

__소스 열\(예\)__

__역할/효과__

장소모델

Warehouse\(Indoor/Outdoor/AAA/MZP/MOSB\), Site\(AGI/DAS/MIR/SHU\)

DSV Indoor/Outdoor, AAA Storage, MOSB, AGI…

창고/현장 계층 표현\(Indoor/Outdoor/Offshore\) → 의미론적 위치 집계

이벤트

TransportEvent \+ hasDate/hasLocation/hasQuantity

창고/현장 날짜, Pkg/CBM

“언제, 어디로, 몇 개/면적” 이동을 그래프에 기록

흐름

hasLogisticsFlowCode\(0~4\)

wh handling 또는 창고 방문 횟수

Port→WH→\(MOSB\)→Site 경로를 정규화\(0=Pre\-Arrival…4\)

재고

StockSnapshot

Status\_Location, Status\_Location\_Date

월말 스냅샷/누계 재고 산출의 기준 노드

비용

Invoice/InvoiceLineItem \+ OFCO 매핑

Description/Rate/Amount

AT\-COST/CONTRACT 등 비용센터 자동 분류

__파이프라인 to KG \(요약\)__  
Ingest\(Excel\) → 정규화\(헤더/날짜/공백\) → 매핑\(JSON rules\) → RDF 변환 → SPARQL 검증\(12 rules\) → Flow/WH·Site 집계 → 리포트/과금\(SQM\) 

__How it works \(핵심 동작 원리, EN\-KR one\-liners\)__

1. __2\-트랙 날짜 모델__: 창고 컬럼\(DSV Indoor/Al Markaz/AAA/MOSB…\)과 현장 컬럼\(AGI/DAS/MIR/SHU\)을 분리 인식 → 최신 위치/이동 추론 강화\. 
2. __Flow Code 계산\(0–4\)__: Pre\-Arrival\(0\)~WH/MOSB 경유~Site 도착까지 hop 수\+오프쇼어 경유로 표준화\.
3. __출고 판정\(시간순\)__: “창고에 찍힌 날짜 < 다음 위치\(다른 창고/현장\) 날짜”일 때만 출고로 인정\(동일일자 중복 방지\)\. 
4. __이중계산 방지 \+ 검증__: 창고간 이동 목적지는 입고에서 제외, 재고는 Status\_Location vs 물리위치 __교차검증__\(불일치 0건 목표\)\. 
5. __RDF/OWL & SPARQL__: DataFrame→RDF 자동 변환, 금액/패키지/위치/시간 일관성 규칙 12종으로 품질게이트\. 
6. __리포팅 아키텍처__: 5\-시트 요약\(Flow/WH·Site 월별/Pre\-Arrival/전체 트랜잭션\) \+ 27시트 스냅샷\(B5 날짜 기반 시계열\) \+ SQM 과금\.

__Options \(구현 옵션 ≥3 · pros/cons/$/risk/time\)__

1. __Option A — Lite KG\(매핑\+피벗 중심\)__

- Pros: 빠른 적용, 5\-시트 리포트 즉시화, 기존 엑셀 호환 우수\. 
- Cons: 실시간 추론/질의 한계, 규칙 변경 시 수작업 많음\.
- Cost/Time: $ · 1–2주\.
- Risk: 규칙 누락/헤더 변형에 민감\(중\)\.

1. __Option B — Full KG\(\+SPARQL 검증/자동 추론\)__

- Pros: RDF 변환\+12개 규칙 검증, 의미론 질의/벤더·월·창고 통합 시계열 안정\. 
- Cons: 온톨로지/삼중저장소 운영 필요\.
- Cost/Time: $$ · 3–5주\.
- Risk: 초기 스키마 설계 미스매치\(중\)\.

1. __Option C — Ops Twin\(\+Flow 추적·SQM 과금\)__

- Pros: 시간순 출고·이중계산 방지, SQM 누적/요율 기반 월별 과금 자동화\. 
- Cons: 데이터 품질\(SQM 실측률\)에 민감\.
- Cost/Time: $$ · 4–6주\.
- Risk: 일부 항목 SQM 추정치 사용 시 오차\(중\)\. 

__Roadmap \(Prepare→Pilot→Build→Operate→Scale \+ KPI\)__

__Prepare \(1주\)__

- 헤더/날짜 정규화, 전각공백\(‘\\u3000’\) 처리, 중복제거 파이프라인 정리\. *KPI: 정제 성공률 ≥ 94\.60%\.* 

__Pilot \(1–2주\)__

- 2\-트랙 매핑 \+ Flow 0–4 적용, 5\-시트 리포트 생성\. *KPI: Flow 계산 일치율 100\.00%\.* 

__Build \(2–3주\)__

- RDF 변환 \+ SPARQL 12규칙, OFCO 비용센터 매핑 연결\. *KPI: 검증 규칙 통과율 100\.00%\.* 

__Operate \(지속\)__

- 시간순 출고/재고 교차검증, 이중계산 0건 유지, SQM 월별 과금\. *KPI: PKG Accuracy ≥ 99\.00% / Inventory 불일치 0건\.* 

__Scale \(지속\)__

- 27시트 스냅샷 도입\(B5 기반 시계열\), 트렌드/변동 자동 감지\. *KPI: 스냅샷 커버리지 100\.00%\.* 

__Automation Hooks \(RPA\+LLM\)__

- __/logi\-master kpi\-dash__: Flow/WH·Site 월별 피벗 \+ KPI 리포트 생성\. 
- __/logi\-master report \-\-deep__: RDF 변환→SPARQL 검증→요약 리포트\. 
- __/logi\-master cert\-chk | invoice\-audit__: OFCO/비용센터 라벨링과 교차 검증\. 
- __/visualize\_data \-\-type=pkg\-flow__: Port→WH→\(MOSB\)→Site 흐름 시각화\(Flow 0–4\)\. 

__QA / Gap 체크리스트__

- 창고 vs 현장 컬럼 __완전 분리__ 적용 여부\(이중계산 방지\)\. 
- 출고 판정이 “다음 위치가 더 늦은 날짜” 규칙을 지키는가\. 
- Flow 0–4 경로 정의와 hop 계산 일치 여부\. 
- 전처리\(전각공백/날짜 정규화/중복제거\) 성공 여부\. 
- SPARQL 12 규칙 통과\(금액 음수/패키지 양수/시간 일관성 등\)\. 
- SQM 실측 vs 추정 비율 보고\(정책: 실측 비중을 단계적으로 상향\)\. 

__CmdRec \(바로 실행\)__

1. __/logi\-master kpi\-dash \-\-KRsummary__ → 월별 WH/Site·Flow 요약 5\-시트 생성\. 
2. __/logi\-master report \-\-deep__ → RDF 변환\+SPARQL 검증\+OFCO 라벨링\. 
3. __/visualize\_data \-\-type=pkg\-flow__ → Flow 0–4 동선 확인\(이상 경로 탐지\)\. 

__한 줄 정리__

__창고 pjt의 ‘한 몸체’는 온톨로지다\.__ 장소·시간·흐름·재고·비용을 __하나의 그래프__에 올려두면, 어떤 키로 들어와도\(케이스·BL·Site…\) 같은 실체로 모이고, 그다음은 계산이 아니라 __질의__가 된다\. \(그리고, 그게 가장 덜 고생한다\.\)

