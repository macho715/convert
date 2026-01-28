아래와 같이 **요청하신 운송 시퀀스**로 일정을 재구성했습니다. 본 변경은 **AGI Site 기준**이며, **DAS Island에는 적용하지 않습니다**(DAS는 Pilotage·PTW·Ramp 조건이 상이하여 별도 스케줄 필요).

---

## 1) 변경 해석 (요청 사항의 명확화)

요청하신 일정은 다음 **4항차 + 3회 현장 JACK‑UP(배치)** 구조입니다.

1. **1항차:** TR **1기** 운송
2. **2항차:** TR **2기**(TR2+TR3) 운송 → **현장 JACK‑UP 3기**(TR1+TR2+TR3)
3. **3항차:** TR **2기**(TR4+TR5) 운송 → **현장 JACK‑UP 2기**
4. **4항차:** TR **2기**(TR6+TR7) 운송 → **현장 JACK‑UP 2기**

> 총 변압기 **7기**, 총 **4 Voyages / 3 Jack‑Up Events**

---

## 2) 핵심 변경 요약 (기존 7항차 → 4항차)

* **Voyages:** 7 → **4**
* **Jack‑Down/Up Events:** 4 → **3** (배치 기준)
* **항차 구성:** 다기 적재(2기) 항차 도입
* **AGI Site 유지:** Mina Zayed ↔ AGI, LCT BUSHRA, Shamal(2/5–2/14) 유지
* **리스크 관리:** 항차 운항(SAIL)·하역(AGI_UNLOAD)은 Shamal 회피 유지

---

## 3) 코드 변경 (핵심부만 교체)

아래는 **`Schedule_Data`의 `tasks` 리스트를 그대로 교체**하면 동작하도록 구성한 **최종 실행용 스케줄**입니다.
(기간 파라미터 `DUR_*`는 기존 Control_Panel 값을 그대로 사용)

```python
# === TASK DEFINITIONS (4 Voyages + 3 Jack-Up Batches) ===
# Project Start = 2026-01-18 (Control_Panel!B4)

tasks = [
    # Mobilization (one-time)
    ("MOB-001","1.0","MOBILIZATION","MOBILIZATION","Mammoet",0,"DUR_MOB",""),
    ("PREP-001","1.1","Deck Preparations","DECK_PREP","Mammoet",1,"DUR_DECK","One-time setup"),

    # === V1: TR1 (1 unit) ===
    ("V1","2.0","VOYAGE 1: TR1 (1 unit)","MILESTONE","All",0,0,""),
    ("LO-101","2.1","TR1 Load-out","LOADOUT","Mammoet",0,"DUR_LO",""),
    ("SF-102","2.2","TR1 Sea Fastening","SEAFAST","Mammoet",0,"DUR_SF",""),
    ("MWS-103","2.3","MWS + MPI","BUFFER","Aries/Captain",0,"DUR_MWS",""),
    ("SAIL-104","2.4","V1 Sail","SAIL","LCT Bushra",2,"DUR_SAIL",""),
    ("ARR-105","2.5","AGI Unload TR1","AGI_UNLOAD","Mammoet",4,"DUR_UL",""),
    ("STORE-106","2.6","TR1 Laydown (Hold)","BUFFER","Mammoet",4,"DUR_BUF","Hold for batch JU-1"),
    ("RET-107","2.7","V1 Return","RETURN","LCT Bushra",4,"DUR_RET",""),

    # === V2: TR2+TR3 (2 units) ===
    ("V2","3.0","VOYAGE 2: TR2+TR3 (2 units)","MILESTONE","All",8,0,""),
    ("LO-201","3.1","TR2 Load-out","LOADOUT","Mammoet",8,"DUR_LO",""),
    ("SF-202","3.2","TR2 Sea Fastening","SEAFAST","Mammoet",8,"DUR_SF",""),
    ("LO-203","3.3","TR3 Load-out","LOADOUT","Mammoet",13,"DUR_LO",""),
    ("SF-204","3.4","TR3 Sea Fastening","SEAFAST","Mammoet",13,"DUR_SF",""),
    ("MWS-205","3.5","MWS + MPI (2 units)","BUFFER","Aries/Captain",14,"DUR_MWS",""),
    ("SAIL-206","3.6","V2 Sail","SAIL","LCT Bushra",15,"DUR_SAIL",""),
    ("ARR-207","3.7","AGI Unload TR2+TR3","AGI_UNLOAD","Mammoet",16,"DUR_UL",""),
    ("RET-209","3.8","V2 Return","RETURN","LCT Bushra",16,"DUR_RET",""),

    # JU-1: 3 units (TR1+TR2+TR3)
    ("TRN-211","3.10","TR1 Transport/Turn","TURNING","Mammoet",16,"DUR_TURN",""),
    ("TRN-212","3.11","TR2 Transport/Turn","TURNING","Mammoet",16,"DUR_TURN",""),
    ("TRN-213","3.12","TR3 Transport/Turn","TURNING","Mammoet",16,"DUR_TURN",""),
    ("JU-217","3.95","★ JU-1 JACK UP (TR1+TR2+TR3)","JACKDOWN","Mammoet",17,"DUR_JD","Batch"),

    # === V3: TR4+TR5 (2 units) ===
    ("V3","4.0","VOYAGE 3: TR4+TR5 (2 units)","MILESTONE","All",28,0,"Post-Shamal"),
    ("LO-301","4.1","TR4 Load-out","LOADOUT","Mammoet",28,"DUR_LO",""),
    ("SF-302","4.2","TR4 Sea Fastening","SEAFAST","Mammoet",28,"DUR_SF",""),
    ("LO-303","4.3","TR5 Load-out","LOADOUT","Mammoet",36,"DUR_LO",""),
    ("SF-304","4.4","TR5 Sea Fastening","SEAFAST","Mammoet",36,"DUR_SF",""),
    ("MWS-305","4.5","MWS + MPI (2 units)","BUFFER","Aries/Captain",36,"DUR_MWS",""),
    ("SAIL-306","4.6","V3 Sail","SAIL","LCT Bushra",36,"DUR_SAIL",""),
    ("ARR-307","4.7","AGI Unload TR4+TR5","AGI_UNLOAD","Mammoet",37,"DUR_UL",""),
    ("RET-309","4.8","V3 Return","RETURN","LCT Bushra",37,"DUR_RET",""),

    # JU-2: 2 units (TR4+TR5)
    ("TRN-311","4.10","TR4 Transport/Turn","TURNING","Mammoet",37,"DUR_TURN",""),
    ("TRN-312","4.11","TR5 Transport/Turn","TURNING","Mammoet",37,"DUR_TURN",""),
    ("JU-315","4.95","★ JU-2 JACK UP (TR4+TR5)","JACKDOWN","Mammoet",40,"DUR_JD","Batch"),

    # === V4: TR6+TR7 (2 units) ===
    ("V4","5.0","VOYAGE 4: TR6+TR7 (2 units)","MILESTONE","All",38,0,""),
    ("LO-401","5.1","TR6 Load-out","LOADOUT","Mammoet",38,"DUR_LO",""),
    ("SF-402","5.2","TR6 Sea Fastening","SEAFAST","Mammoet",38,"DUR_SF",""),
    ("LO-403","5.3","TR7 Load-out","LOADOUT","Mammoet",40,"DUR_LO",""),
    ("SF-404","5.4","TR7 Sea Fastening","SEAFAST","Mammoet",40,"DUR_SF",""),
    ("MWS-405","5.5","MWS + MPI (2 units)","BUFFER","Aries/Captain",40,"DUR_MWS",""),
    ("SAIL-406","5.6","V4 Sail","SAIL","LCT Bushra",40,"DUR_SAIL",""),
    ("ARR-407","5.7","AGI Unload TR6+TR7","AGI_UNLOAD","Mammoet",41,"DUR_UL",""),
    ("RET-409","5.8","V4 Return","RETURN","LCT Bushra",41,"DUR_RET",""),

    # JU-3: 2 units (TR6+TR7)
    ("TRN-411","5.10","TR6 Transport/Turn","TURNING","Mammoet",41,"DUR_TURN",""),
    ("TRN-412","5.11","TR7 Transport/Turn","TURNING","Mammoet",41,"DUR_TURN",""),
    ("JU-415","5.95","★ JU-3 JACK UP (TR6+TR7)","JACKDOWN","Mammoet",41,"DUR_JD","Final batch"),

    # Demobilization
    ("DEMOB","9.0","DEMOBILIZATION","MOBILIZATION","Mammoet",42,"DUR_MOB",""),
    ("END","99.0","★★★ PROJECT COMPLETE ★★★","MILESTONE","All",42,0,"7 TRs complete"),
]
```

### 함께 바꿔야 할 문자열(요약)

* `Schedule_Data!A2`
  **기존:** `7 Voyages | 4 Jack-down Events`
  **변경:** `4 Voyages | 3 Jack-up Events`
* `Summary` 시트:

  * **Total Voyages:** 4
  * **Jack-down/Up Events:** 3
  * **Voyage Summary:** V1 / V2(+JU‑1) / V3(+JU‑2) / V4(+JU‑3)

---

## 4) SITE 구분 적용

### AGI Site

* 본 코드 **적용 대상**
* Pilotage Exemption 가능, PTW ≥24h
* Shamal(2/5–2/14) 회피 유지

### DAS Island

* **미적용**
* LCT Barge 전용, Pilotage 의무, PTW ≥48h, Ramp 각도·Berth Load Chart 상이
  → DAS 전용 스케줄 별도 작성 필요

---

원하시면 **이 스케줄로 엑셀(.xlsx) 파일을 즉시 생성**하거나, **Load‑out/Unload 소요를 2기 적재 기준으로 배율(DUR_*×2)** 적용한 현실 반영 버전도 제공하겠습니다.
