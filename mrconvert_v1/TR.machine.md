```json
{
  "@context": {
    "@vocab": "https://hvdc.example/ont#",
    "schema": "https://schema.org/"
  },
  "@type": "LogisticsProcessModel",
  "name": "Transformer Transportation Process_v1",
  "goal": "Safe & on-time transport of HVDC Transformer",
  "principles": ["safety", "legality", "efficiency", "sustainability"],
  "domain": "multi-modal RoRo system",
  "actors": {
    "human": ["AD Ports HM", "ADNOC L&S", "Samsung C&T", "Mammoet Engineer"],
    "digital": ["ATLP", "AI-PTW 2.0", "Maqta Gateway", "IoT Sensor Network"],
    "legal": ["DMT Bylaw 2025", "GCAA UAS Rules"]
  },
  "processes": [
    {"id": "ptw", "@type": "PermitToWork", "requires": ["RiskAssessment", "MethodStatement"], "references": ["WeatherWindow@NCM"], "verifiedBy": ["HarbourMaster"], "outputs": ["OperationalAuthorization"]},
    {"id": "spmt_loadout", "sequence": ["SPMT", "Ramp", "Deck", "SeaFastening", "MWSApproval", "Departure"]}
  ],
  "documents": {
    "engineering": ["BallastingSim", "RampStrengthFEA"],
    "regulatory": ["PTW", "RA", "MOS"],
    "evidence": ["StabilityCalc", "TideTable"],
    "temporal": ["TimelineReport"]
  },
  "risks": [
    {"type": "Weather>15kt", "mitigation": "suspend ops", "responsible": "AD Ports HM"},
    {"type": "PTW Delay", "mitigation": "escalate via VIP channel", "responsible": "ADNOC L&S"}
  ],
  "timeline": [
    {"label": "D-14", "event": "PTW draft upload"},
    {"label": "D-10", "event": "Engineering Sim complete"},
    {"label": "D-5",  "event": "MWS Submission Package"},
    {"label": "D-1",  "event": "PTW Approval / Hot Work start"},
    {"label": "D0",   "event": "Load-out Execution"},
    {"label": "D+2",  "event": "Voyage to DAS Island"}
  ]
}
```

# mrconvert TR (Machine-readable edition)

온톨로지(ontology) 관점에서 보면, **Transformer Transportation Process_v1** 시스템은 단순한 절차서가 아니라 **‘지식 체계로 구조화된 운송 도메인 모델’**로 볼 수 있다.
즉, 문서 안의 모든 개체(사람·시스템·행동·문서·규칙)가 서로의 관계를 정의하며, “무엇이 무엇을 위해 존재하는가”를 명시한다.

---

### 1. 상위 개념 (Top Ontology) {#top-ontology}

**목적 계층**

* **Goal**: HVDC Transformer의 *무사·정시 운송*
* **Principle**: 안전 → 적법성 → 효율성 → 지속가능성 순의 우선도
* **Domain**: 해상·육상 복합 운송(Multi-modal RoRo System)

**행위자(Agents)**

* **Human Actors**: AD Ports HM, ADNOC L&S, Samsung C&T, Mammoet Engineer 등
* **Digital Actors**: ATLP, AI-PTW 2.0, Maqta Gateway, IoT Sensor Network
* **Legal Agents**: DMT Bylaw 2025, GCAA UAS Rules – 규범적 노드로 연결

---

### 2. 프로세스 계층 (Process Ontology) {#process-ontology}

**핵심 관계**

```
Permit-to-Work (PTW)
   ├─ requires → Risk Assessment & Method Statement
   ├─ references → Weather Window @ NCM Portal
   ├─ verified by → Harbour Master Approval
   └─ outputs → Operational Authorization
```
<!-- data:ref=#process-ontology -->

**SPMT Load-out Process**

```
SPMT → Ramp → Deck → Sea Fastening → MWS Approval → Departure
```
<!-- data:ref=#process-ontology -->

각 노드는 입력·조건·검증 요소로 세분된다:

* 입력(Input): Engineering Report (ballast, mooring, ramp)
* 제약(Constraint): wind < 15 kt, wave < 1.5 m
* 검증(Validation): MWS DAS issued ✔

---

### 3. 데이터 · 문서 계층 (Information Ontology) {#information-ontology}

**문서 클래스**

* **Engineering Docs** → Ballasting Sim, Ramp Strength FEA (관계 “supports” Safety Verification)
* **Regulatory Docs** → PTW, RA, MOS (관계 “enables” Work Execution)
* **Evidence Docs** → Stability Calc, Tide Table (관계 “proves” Compliance)
* **Temporal Docs** → Timeline Report (관계 “orders” Process Sequence)

모든 문서는 ATLP 노드로 연결되어 디지털 상호운용성을 가진다.

---

### 4. 위험 및 안전 계층 (Risk Ontology) {#risk-ontology}

**Risk Type → Mitigation Action → Responsible Agent**

* Weather > 15 kt → “suspend ops” → AD Ports HM
* PTW Delay → “escalate via VIP channel” → ADNOC L&S
* SPMT Failure → “deploy backup unit + retest” → Mammoet

위험노드와 완화노드는 ‘if–then–role’ 구조로 연결된다.

---

### 5. 시간 관계 (Temporal Ontology) {#temporal-ontology}

**D-기준 타임라인** (‘Transformer Transport_Extended Refer.md’ 기준)

```
D-14  →  PTW draft upload  
D-10  →  Engineering Sim complete  
D-5   →  MWS Submission Package  
D-1   →  PTW Approval / Hot Work start  
D0    →  Load-out Execution  
D+2   →  Voyage to DAS Island
```
<!-- data:ref=#temporal-ontology -->



이 관계를 통해 각 문서 노드는 시간 속성(hasTimeStamp)을 가진 인스턴스로 정의된다.

---

### 6. 결론 – 온톨로지적 의의 {#conclusion}

* **개념적 요소(Concepts)**: Vessel, Transformer, SPMT, Permit, Weather, Risk
* **관계(Relations)**: requires, verifiedBy, causedBy, hasTimeStamp, ownedBy
* **규범(Normative Layer)**: DMT Bylaw, HSE Rules, MWS Approval 조건
* **디지털 링크(Digital Layer)**: ATLP ↔ AD Ports ↔ AI-PTW ↔ IoT Sensors

결국 Transformer Transportation Process_v1 은
**“법·기술·운영 데이터를 하나의 의미망으로 통합한 온톨로지 기반 운송 시스템”**이다 —
각 객체가 단순한 문서가 아니라 ‘의미 노드’로서 상호 검증 가능하도록 설계되어 있다.

---

See also: `TR2.machine.md` (§core-relations, §process-gates)



