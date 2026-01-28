---
title: "CODE Writing - Ontology Perspective"
type: "ontology-design"
domain: "software-development"
version: "1.0"
date: "2025-01-19"
tags: ["ontology", "code", "tdd", "development", "process"]
status: "active"
---

# CODE 작성 - 온톨로지 관점

## 요약

"코드 작성"을 **온톨로지**(무엇이 있고 어떻게 얽히는가)로 정의합니다.

이 프로젝트는 **원칙(Constitution)**이 **작업 흐름(Commands)**을 **산출물(Artifacts)**로 이어 붙이고, 각 흐름은 **게이트(Quality Gates)**로 검증되며, 모든 조각이 **명세–계획–과업–분석–구현**이라는 **유형화된 과정 타입들**로 상호 제약되는 체계입니다. 헌법 규칙은 상위 규범이며 분석 단계에서 "불가침"으로 집행됩니다.

__1\) TBox\(스키마\): 핵심 클래스와 관계__

__클래스\(개념\)__

- __Project__: 코드 작성 전체를 아우르는 최상위 개체\.
- __Principle__: 변경 불가한 규범\. *Constitution* 문서로 구체화된다\(예: 의사결정·품질 기준\)\.
- __Command\(ProcessType\)__: /specify·/clarify·/plan·/tasks·/analyze·/implement 같은 실행 가능한 과정 유형\. 각 커맨드는 입력/출력 산출물이 정형화되어 있다\.
- __Artifact__: spec\.md, plan\.md, tasks\.md, research\.md, data\-model\.md, contracts/, quickstart\.md 같은 파일 산출물\.
- __QualityGate__: “진행해도 되는가?”를 판단하는 규범/검증 모음\. 예: 헌법 준수 체크, TDD/Tidy 루프, 선행 명세/정리 유무\.
- __State__: Artifact/Feature의 생애주기 상태\(초안 Draft → Clarified → Planned → Tasked → Analyzed → Implemented\)\.
- __Actor__: 참여 주체\(개발자/리뷰어/자동 에이전트\)\.
- __Tool__: 테스트/린트·스크립트 등\(예: 계획/검증 스크립트, 테스트 러너\)\. 커맨드가 이를 호출한다\.

__핵심 관계__

- governs\(Principle, Command/Artifact\): 헌법 원칙이 프로세스와 산출물을 규율\.
- produces\(Command, Artifact\): 각 커맨드가 특정 산출물을 생성/갱신\.
- requires\(Command, Artifact|Principle\): 실행 전제조건\(예: /plan은 __Clarifications__가 있어야 함\)\.
- validates\(QualityGate, Artifact|Process\): 분석/게이트가 요구사항 충족 여부를 판정\(헌법 충돌은 자동 *CRITICAL*\)\.
- orders\(ProcessType, ProcessType\): 수행 순서 제약\(예: 테스트→코드, 구조→행위\)\.

__2\) 프로세스 타입\(Commands\)로 본 수명주기__

1. __/specify__ → *특징*: 브랜치와 스펙 파일을 생성·체크아웃하고 템플릿에 맞춰 __spec\.md__를 초기화한다\. 산출: SPEC\_FILE, 분기: BRANCH\_NAME\.
2. __/clarify__ → *역할*: 최대 5개의 정밀 질문으로 __모호성 해소__ 후, 답을 __Clarifications__ 섹션과 관련 섹션\(요구사항·데이터모델·품질 속성 등\)에 __즉시 반영__한다\. 상태 전이: Draft → Clarified\.
3. __/plan__ → *전제*: Clarifications가 존재해야 진행\. *산출*: __research\.md__, __data\-model\.md__, __contracts/__, __quickstart\.md__, 그리고 최종적으로 __tasks\.md__\(Phase2\)\.
4. __/tasks__ → *규칙*: 의존성 순서와 병렬성\(\[P\]\)을 명시\. “__테스트 우선\(TDD\)__, 모델→서비스→엔드포인트, 코어→통합→폴리시” 등 정형 규칙을 반영해 __실행가능한__ 작업 목록을 생성\.
5. __/analyze__ → *성격*: __읽기 전용__ 교차검증\. Spec/Plan/Tasks 간 __중복·모호·미커버리지·헌법 충돌__을 표 형태로 보고하고 심각도\(CRITICAL/HIGH/…\)를 부여\. 헌법 위반은 자동 CRITICAL\.
6. __/implement__ → *집행*: __tasks\.md__의 단계/의존/병렬 규칙에 따라 __단계별__ 구현\. 실패 시\(비병렬\) 중단·보고, 병렬은 성공분 계속\. 완료 시 체크박스를 \[X\]로 갱신\.

__3\) 규범\(Principles\)과 게이트\(Quality Gates\)__

- __TDD 루프\(RED→GREEN→REFACTOR\)__: 기능 한 조각의 실패 테스트\(RED\) → 통과에 필요한 __최소 구현__\(GREEN\) → __중복 제거/의도 드러내기__\(REFACTOR\)\. 단일 테스트는 ms~수백 ms SLA, 느린 I/O는 더블/주입\.
- __Tidy First\(구조→행위 분리\)__: 리네임/추출/이동 등 __구조 변경__은 행위 불변으로 먼저, 그 다음 __행위 변경__\(feat/fix\)\. 커밋도 structural: ↔ behavioral:로 분리\.
- __진입 게이트__: /plan은 명시적으로 __Clarifications__ 유무를 검사\(없으면 중단·사전 정리 권고\)\. /analyze는 헌법 위반을 __CRITICAL__로 지정\.

이 세 가지는 온톨로지 상에서 QualityGate 인스턴스들로 모델링되어 각 Process 인스턴스에 연결된다\. 실패 시 __State__는 다음 단계로 전이되지 않는다\(불합격 전이 차단\)\.

__4\) ABox\(예시 인스턴스\) — 한 기능이 지나가는 이야기__

- Process\(/specify\) __produces__ Artifact\(spec\.md\) & __creates__ Branch\(feature/x\)\.
- Process\(/clarify\) __updates__ Artifact\(spec\.md\#Clarifications\); 상태 Draft→Clarified\.
- Process\(/plan\) __produces__ research\.md, data\-model\.md, contracts/, quickstart\.md, tasks\.md\.
- Process\(/tasks\) __orders__ 작업: __Setup→Tests→Core→Integration→Polish__, 파일 충돌 시 __순차__, 파일 분리 시 __\[P\] 병렬__\.
- Process\(/analyze\) __validates__ \(dup/ambiguity/coverage/constitution\) ⇒ __OK__면 다음\.
- Process\(/implement\) __executes__ tasks; 실패 시 중단·리포트, 성공은 \[X\] 체크\.

__5\) 소형 온톨로지 스케치\(의미 보존용 의사‑Turtle\)__

\#\#\# Classes

Project, Principle, Command, Artifact, QualityGate, State \.

\#\#\# Object properties

governs, produces, requires, validates, orders, updates, creates, blocksTransition \.

\#\#\# Instances \(축약 예\)

:Constitution        a Principle \.

:TDD\_Gate            a QualityGate \.

:TidyFirst\_Gate      a QualityGate \.

:CodeWriting         a Project \.

:Specify  a Command ; produces :SpecMD ; creates :FeatureBranch \.

:Clarify  a Command ; requires :SpecMD ; updates :SpecMD \.

:Plan     a Command ; requires :SpecMD ; produces :Research, :DataModel, :Contracts, :Quickstart, :TasksMD \.

:Tasks    a Command ; requires :Plan ; produces :TasksMD \.

:Analyze  a Command ; validates :SpecMD, :PlanMD, :TasksMD ; governedBy :Constitution \.

:Implement a Command ; requires :TasksMD ; validates :TDD\_Gate, :TidyFirst\_Gate \.

\# 규칙 예시

:Analyze \-\-governedBy\-\-> :Constitution \.   \# 헌법 불가침

:TDD\_Gate \-\-blocksTransition\-\-> :ImplementWhenFailed \.

- 헌법 규칙의 “불가침”과 분석의 읽기 전용·심각도 체계는 /analyze 문서의 규정으로 근거화된다\.
- Clarify의 “최대 5문항”·분류 택소노미·즉시 반영 규칙은 /clarify 흐름 정의로부터 온다\.
- Plan의 단계 산출물 세트와 Tasks의 의존·병렬 규칙은 각 명세에 기재되어 있다\.
- TDD/Tidy 루프와 구조↔행위 분리는 Tidy 지침에 명시되어 있다\.

__6\) 이 온톨로지가 답해야 하는 Competency Questions__

- 현재 Feature X는 __어떤 상태__이며\(Clarified? Planned?\) 다음 단계로 __전이 가능한가?__ 전이를 막는 __게이트 실패__는 무엇인가? \(예: Clarifications 없음, Analyze CRITICAL\)
- spec/plan/tasks 사이에 __중복·충돌·미커버리지__가 있는가? 있다면 __심각도__와 __수정 후보__는?
- 특정 과업 T\#\#\#는 __어떤 요구사항__을 충족하며, __병렬 실행__ 가능한가\(파일 경합 여부\)?
- 이번 커밋은 __구조 변경__인가 __행위 변경__인가? \(분리 커밋/리뷰 기준\)

__7\) 함께 볼 메타‑원칙\(현실과의 접점\)__

- __plan\.md 최우선 & RED→GREEN→REFACTOR & 구조→행위__: 설계→검증→집행 순서가 지식 구조로서 먼저 서야, 실제 코드 변화가 안정적으로 수렴한다\.
- __/plan 전 Clarifications 필수__: 모호성은 가장 비싼 결함\. 전방에서 줄인다\.
- __헌법 우선권__: 원칙을 바꾸려면 별도 절차\(버전/개정 내역/동기\)가 필요하다\.

__다음 선택지\(가벼운 옵션\)__

1. 위 스키마를 __TTL/OWL 초안__으로 뽑아드릴 수 있다\(프로세스·산출물·게이트·상태 포함\)\.
2. 현재 진행 중인 기능 하나를 골라 __ABox 인스턴스__\(실제 삼중\)로 채워 시각화해보자\.
3. /analyze 리포트 항목을 __온톨로지 질의__\(competency questions\)로 바로 매핑하는 규칙표를 만들어 드릴 수도 있다\.

원하면 바로 골라줘\. 난 틀을 잡아주고, 너는 도메인 언어를 더해 주면 된다\.
이 구조는 네가 원하는 “온톨로지 관점의 설계—운영—검증”을 같은 축 위에 올려준다\.

