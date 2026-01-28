다음은 **Transformer transportation process(RORO)2025**를 **온톨리지(ontology) 관점**으로 정리한 전체 그림이야. 말 그대로 “무엇(entities)–어떻게(relations)–언제(process)–왜(constraints)”를 명확히 연결해서, 문서 42종·승인·시뮬레이션·현장 조건이 한눈에 이어지도록 모델링한다.

---

# 1) 상위 개념 스키마 (Top-level classes)

* **Actor(주체)**: Samsung, Mammoet, ADNOC L&S, Harbour Master(HM), OFCO, Khalid Faraj Shipping(KFS), Sterling Technical(MWS). 역할/책임 충돌 이슈와 승인권한을 명시.  
* **Asset(자산)**: Transformer(169 t), Vessel(LCT *BUSHRA*), Ramp, Jetty, SPMT 등. 
* **Document(문서)**: **Master 42종**(행정/허가·승인/기술·HSE 등)과 서브타입(PTW, Mooring Plan, Stowage Plan, Ramp Plate Cert, DAS 등), 발행자·검증자·버전·제출채널(Maqta/ATLP·이메일) 속성 포함.    
* **Process(프로세스/페이즈)**: Pre-Arrival → Approval → Load-Out(RoRo) → Voyage → Load-In → Post-Ops. 날짜·의존관계·Gate(승인) 모델로 연결.   
* **Simulation/Calculation(시뮬/해석)**: Ramp FEA, Ballast/Trim, Mooring Tension, Stability(책임소재 쟁점 포함).  
* **Approval/Permit(승인/허가)**: HM Approval, PTW(Hot Work/Over-Water), MWS(DAS). 프로세스 Gate로 동작.  
* **Constraint/Condition(제약)**: **주간작업만, 풍속 ≤ 15 kt, Ramp angle≈0°, 특정 조위/조석 창**, Tug & Pilot 상시 대기 등. 규정 준수 실패 시 작업 중지/지연.   
* **Risk(위험)**: 문서 지연, 장비고장(SPMT/램프), 규정 불이행, 기상. 완화책과 보고주기 포함. 
* **Platform/Channel(제출 채널)**: **Maqta Gateway/ATLP**, 이메일(대용량 분할/수신확인). 검증흐름·타임스탬프 추적.   

---

# 2) 핵심 관계(객체속성) 예시

* `requiresDocument(Process, Document)` : 예) **Load-Out** → PTW, Mooring Plan, SPMT Cert, Ramp Cert. 
* `approvedBy(Document, Actor)` : PTW→AD Ports HSE/HM, DAS→MWS.  
* `submittedVia(Document, Platform)` : Maqta/ATLP 업로드 또는 이메일 분할 송부.  
* `constrains(Constraint, Operation)` : **Daylight only / Wind≤15 kt / Tug&Pilot 필수**가 Cargo Ops를 제약. 
* `dependsOn(Process, Simulation)` : Load-Out은 Ramp-FEA/ Ballast/ Mooring 해석 가정에 의존. 
* `hasResponsibility(Actor, Task)` : 책임분장(예: Stability/Mooring/FEA 범위 논쟁) 표준화.  

---

# 3) 프로세스 온톨로지 (Gate 기반)

* **Pre-Arrival**: 사전 신고·서류 번들 준비(항목·버전·영문·SI단위). → **Maqta/ATLP 제출 & 수신확인**.  
* **Approval**: HM 조건 충족(시뮬·Mooring·PTW) 없으면 다음 단계 불가. **DAS**가 실질적 “Go/No-Go” Gate.  
* **Load-Out(RoRo)**: Ramp angle≈0°, 조석창 준수, **풍속 15 kt 초과 시 즉시 중단**, Tug/Pilot 필수.  
* **Voyage/Load-In/Post**: Stability 재계산·보고(예: Capt. Joey 지시사항), 문서 아카이빙·교훈정리.  

> 일정 모델은 2024.10 롤온 데이(현장 승인·회의·작업 개시)와 2025.03~04 **TR 3–6** 배치까지 포괄한다(의존관계/마일스톤 연결).   

---

# 4) 문서 42종 온톨로지

* **Document** ⊃ {Administrative, Permits&Approvals, Technical/HSE, RoRo/Marine, Post-Ops}
  각각 **발행자–검증자–제출채널–유효기간–버전** 속성 포함.
* 예)

  * **PTW(Hot Work/Over-Water)**: *issuedBy* Samsung/OFCO ↔ *verifiedBy* AD Ports HSE/HM ↔ *submittedVia* Maqta ↔ *appliesTo* Load-Out. 
  * **Mooring Plan**: *basedOn* MMT 시뮬 ↔ *validatedBy* MMT/HM ↔ *gateFor* DAS.  
  * **Ramp Plate Certificate/FEA**: *derivedFrom* SPMT Axle Load ↔ *verifies* Ramp capacity. 
  * **DAS(MWS Approval)**: *issuedBy* Sterling Technical ↔ *unlocks* Roll-On. 
* 42개 항목의 **역할·발행주체·검증 적합성** 매핑 표는 별도 시트로 검증되어 있음(필수 제출/오류 가능 항목 포함).  

---

# 5) 규정·조건의 형식화 (Rule layer)

* **HM 규정**

  * `DaylightOnly(Operation)`; `MaxWind(Operation, 15kt)`; `Require(Tug & Pilot)`; `Require(PTW, MooringPlan)` → 미충족 시 `Operation=Halt`. 
* **시뮬/해석 가정**

  * `RampAngle(Operation) ≈ 0°` & `TideWindow(LAT 0.7 m)` & `FWDdraft=2 m(150 m³ FW 주입 후)` → 로딩 타이밍·트림 제약. 
* **게이트 조건**

  * `PermitApproved(PTW) ∧ MooringPlan(Validated) ∧ SPMT Cert ∧ Ramp Cert ∧ DAS` ⇒ `CanStart(Load-Out)`.  

---

# 6) 책임·분쟁 모델링

* `hasScope(Mammoet, RoRo Simulation & Stowage)` ∧ `scopeDispute(ballast/mooring/ramp)` → 대체 경로: `externalNavalArchitect` 또는 `KFS Documents`(Stability/GA/Structural) 참조.  
* `issues(DAS, MWS)`이 `dependsOn(ballast/mooring/ramp verification)`에 연결되어 지연 리스크를 생성. 

---

# 7) 데이터 플로우·추적성

* **제출 채널**: Maqta/ATLP 우선, 대용량은 이메일 분할+수신확인(파일명 규칙/보안 라벨/버전 필수). **검증 흐름**: Initial(HSE) → Secondary(MMT) → Final(HM). 
* **체크리스트·타임라인**: 접안–모어링–HSE 브리핑–승인회의–작업개시–출항까지의 시퀀스와 책임자·증빙 링크를 추적.  

---

# 8) 리스크 레이어

* 위험요인(기상 15 kt 초과·문서오류·SPMT/램프·규정미준수) ↔ **완화조치/보고주기(일일 RA 18:00)**를 클래스/속성으로 고정. 

---

# 9) 2025년(3~4월) 다기(3–6호기) 확장

* 동일 온톨로지로 3–6호기 일정(Load-Out/항해/Load-In/육상이동)과 HM 조건(주간·15 kt·Double Mooring·Ramp 0°)을 재사용. 마스터 문서·승인 게이트 그대로 확장 적용. 

---

## (참고) 초간단 TTL 스니펫

```ttl
:LoadOut_2024_10_24 a :Process ;
  :requiresDocument :PTW_HotWork, :PTW_OverWater, :MooringPlan, :SPMT_Cert, :Ramp_Cert, :DAS ;
  :constrainedBy :DaylightOnly, :WindMax15kt, :RampAngleZero, :TugAndPilot ;
  :dependsOn :RampFEA, :BallastTrimCalc, :MooringTensionCalc ;
  :performedBy :Samsung, :Mammoet, :KFS, :OFCO .

:HM_Guideline a :ConstraintSet ;
  :hasConstraint :DaylightOnly, :WindMax15kt, :TugAndPilot .
```

---

## 이 온톨로지로 답할 수 있는 질문(Competency Qs)

* “지금 **Load-Out** 시작 가능?” → 필요한 문서/승인 Gate 충족 여부·현장 조건(풍속/주간) 평가.  
* “**Ramp** 보강 필요?” → SPMT 축하중→Ramp FEA→Cert 상호추적. 
* “**DAS** 왜 지연?” → Mooring/Trim/FEA 검증 누락 경로 트레이스. 

---

딱 이 구조면, 사람·문서·승인·해석·현장조건이 하나의 지식그래프에서 깨끗하게 이어져. 현장선 ‘지금 할 수 있나’를, 본사선 ‘무엇이 막고 있나’를 즉시 묻고 답하게 되지. 필요한 만큼 더 세분화(예: SHACL 제약, 버전 규칙)도 붙일 수 있어.
