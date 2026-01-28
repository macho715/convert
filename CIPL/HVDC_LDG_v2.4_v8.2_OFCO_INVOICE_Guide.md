
# ◈ HVDC Logistics DATA-MAPPING + OCR v2.4 × LDG v8.2 Hybrid Guide — OFCO INVOICE 전용 (최종) ◈
*(2025-08-11 · Agent Mode 최적화 · OFCO INVOICE 패턴 반영 · KPI+Fail-safe+EA/Rate 규칙 강화 · COST MAIN 계층 구조 반영 · 개별 인보이스 분개 처리)*

## A. 목적
OFCO 발행 인보이스(PDF + Excel) 데이터를 **ChatGPT Agent Mode**에서  
**고정밀 OCR → 자동 검증 → COST MAIN/CENTER 매핑 → (최대) 3-Way 분개 → 집계-열 포맷 → 리포트(7+2)**까지  
원-스톱 처리하기 위한 **최종 통합 지침**.

---

## B. 구성
1. **DATA Mapping Rules** – Cost Center v2.5 (17-Step Regex) + OFCO Subject 패턴 + **COST MAIN 계층 구조**
2. **Invoice Scan OCR** – OCR v2.4 (LDG Ready) + KPI 기준  
3. **LDG v8.2 연동** – Agent Mode 트리거, Fail-safe(ZERO), Evidence 필수  
4. **PRICE CENTER 분개** – A/B/C 3-Way + **SAFEEN/ADP 특화 처리**
5. **EA 분해 알고리즘** – **SAFEEN(시간 기반) / ADP(수량 기반) / 일반** 타입별 분기
6. **개별 인보이스 분개** – **vendor_invoice_no별 그룹화 및 통합 처리**
7. **QC·Change-Mgmt** – Self-Check Loop, Dict 업데이트  
8. **보고 규격** – 7+2 Section Report  

---

## 1️⃣ DATA-MAPPING (OFCO Subject 패턴 보강 + COST MAIN 계층 구조)

### 1.1 COST MAIN → COST CENTER A → COST CENTER B → PRICE CENTER 계층 구조

**4단계 계층 구조:**
```
COST MAIN (Level 1: 최상위 분류)
    ↓
COST CENTER A (Level 2: 중간 분류)
    ↓
COST CENTER B (Level 3: 세부 분류)
    ↓
PRICE CENTER (Level 4: 최종 세부 항목)
```

**COST MAIN 분류 (3개):**
- **CONTRACT**: 계약 기반 수수료/서비스 (Agency Fee, Handling Fee, Pass Arrangement)
- **PORT HANDLING**: 항만 처리 관련 (Channel Transit, Port Dues, Bulk Material Handling)
- **AT COST**: 원가 기반 서비스 (Water Supply, Diesel, Forklift, Consumables)

**COST CENTER A 분류:**
- CONTRACT 계열: `CONTRACT`, `CONTRACT_MANPOWER`, `CONTRACT_EQUIPMENT`
- PORT HANDLING 계열: `PORT HANDLING CHARGE`, `PORT HANDLING`
- AT COST 계열: `AT COST`

**COST CENTER B 분류 (주요 패턴):**
- CONTRACT 계열: `AF FOR CC`, `AF FOR BA`, `AF FOR FW SA`, `AF FOR PASS ARRG`, `CONTRACT(AF FOR PTW ARRG)`, `OFCO HF`, `CONTRACT(YARD)`, `BULK CARGO_MANPOWER`, `BULK CARGO_EQUIPMENT`
- PORT HANDLING 계열: `CHANNEL TRANSIT CHARGES`, `PORT DUES & SERVICES CHARGES`, `BULK CARGO HANDLING CHARGES`, `PORT HANDLING CHARGE(GATE PASS)`, `Pilotage`
- AT COST 계열: `AT COST(WATER SUPPLY)`, `AT COST(DIESEL)`, `AT COST(FORKLIFT)`, `AT COST(CONSUMABLES)`

**PRICE CENTER 분류:**
- `cost_item_fields.JSON`의 필드명과 직접 연결
- **⚠️ 중요: `cost_item_fields.JSON`의 필드명은 절대 변경 금지**

### 1.2 Subject 패턴 매핑 (보강)

- **대상 시트:** `"OFCO INVOICE"` (또는 동일 구조 시트)
- **입력 열:**  
  - **BJ:** Total Amount(기준 금액)  
  - **BB:BI:** EA/Rate(최대 4쌍) + Amount(AED)  
  - **K:BA:** 코드/설명/Ref/VAT/통화 (EXT 행 가능)  
- **Cost Center 매핑:**  
  - HVDC v2.5 **17-Step Regex** 우선 적용  
  - **Subject 패턴 보강 (계층 구조 반영):**  
    - `"SAFEEN.*Channel.*Crossing"` → COST MAIN: PORT HANDLING, COST CENTER A: PORT HANDLING CHARGE, COST CENTER B: CHANNEL TRANSIT CHARGES, PRICE CENTER: CHANNEL TRANSIT CHARGES
    - `"ADP.*Port.*Dues"` → COST MAIN: PORT HANDLING, COST CENTER A: PORT HANDLING CHARGE, COST CENTER B: PORT DUES & SERVICES CHARGES, PRICE CENTER: PORT DUES
    - `"Cargo Clearance"` → COST MAIN: CONTRACT, COST CENTER A: CONTRACT, COST CENTER B: AF FOR CC, PRICE CENTER: AGENCY FEE FOR CARGO CLEARANCE
    - `"FW Supply"` / `"Arranging FW Supply"` → COST MAIN: CONTRACT, COST CENTER A: CONTRACT, COST CENTER B: AF FOR FW SA, PRICE CENTER: SUPPLY WATER 5000IG
    - `"Berthing Arrangement"` → COST MAIN: CONTRACT, COST CENTER A: CONTRACT(AF FOR BA), COST CENTER B: CONTRACT, PRICE CENTER: AGENCY FEE FOR BERTHING ARRANGEMENT
    - `"5000 IG FW"` / `"SUPPLY WATER 5001IG"` → COST MAIN: AT COST, COST CENTER A: AT COST, COST CENTER B: AT COST(WATER SUPPLY), PRICE CENTER: SUPPLY WATER 5000IG
- **집계-열 매핑:** RAW ↔ Column 불변 테이블 (cost_item_fields.JSON 필드명 사용)  

---

## 2️⃣ OCR (LDG Ready + LDG_PAYLOAD v2.4 호환성)

- 실행: `/ocr_basic {file} mode:LDG+`  
- 출력: `LDG_PAYLOAD v2.4` + `LDG_AUDIT`  
- **KPI:** MeanConf ≥ 0.92 · TableAcc ≥ 0.98 · NumericIntegrity = 1.00 · EntityMatch ≥ 0.98  
- **ZERO 모드 트리거:** 위 KPI 중 하나라도 FAIL, 핵심 식별자 결손, Evidence 상충·부재  
- **ZERO 로그 예시:**

```
| 단계 | 트리거 | 세부내역 | 조치 |
|---|---|---|---|
| ZERO-01 | NumericIntegrity FAIL | Line #12 합계 오차 | /ocr_retry |
| ZERO-02 | MeanConf<0.90 | p.3 저해상/경사 | /ocr_lowres_fix |
| ZERO-03 | Evidence 부족 | Ref 근거 부재 | web.run 재수집 |
```

- 품질 경고: 핵심 필드 conf < 0.90 → `[OCR ALERT]` 후 중단  

### 2.1 LDG_PAYLOAD v2.4 구조 개요

**주요 필드:**
- `invoice_header` (최상위 키, 문서의 `invoice_meta`와 다름)
- `invoice_number` (필드명, 문서의 `invoice_no`와 다름)
- `qty`, `unit_price`, `unit` (수량/단가/단위)
- USD + AED 이중 통화 구조 (`amount_excl_tax`, `amount_aed_excl_tax` 등)
- `calc_check`, `tax_check`, `total_check` (boolean 형식, 문서는 string)
- 이미 매핑된 Cost/Price Center 정보 포함
- **`vendor_invoice_no`** (vendor 인보이스 번호, 개별 분개용)
- **`vendor_invoice_refs`** (배열 형식, vendor_invoice_no 추출용)

**필드명 매핑:**
- `invoice_header` → `invoice_meta`
- `invoice_number` → `invoice_no`
- `qty` → `unit1`
- `unit_price` → `rate`
- `amount_excl_tax` (USD) → `amount_excl_tax_usd`
- `amount_aed_excl_tax` → `amount_excl_tax_aed`
- `calc_check` (boolean) → `calc_check` (string: "PASS"/"FAIL")
- **`vendor_invoice_refs[0]` → `vendor_invoice_no`** (그룹화용)

### 2.2 Parsed JSON vs LDG_PAYLOAD v2.4 차이점

**Parsed JSON (텍스트 추출 기반):**
- 최상위 키: `invoice_meta` (LDG는 `invoice_header`)
- 인보이스 번호: `invoice_no` (LDG는 `invoice_number`)
- Vendor 참조: `vendor_invoice_refs` 배열만 있음 (LDG는 `vendor_invoice_no` 필드 있음)
- **⚠️ 문제:** `vendor_invoices` 섹션이 없어 개별 인보이스 분개 불가
- **해결:** `enhance_parsed_json_with_vendor_invoices()` 함수로 개선 필요

**LDG_PAYLOAD v2.4 (OCR 엔진 기반):**
- 최상위 키: `invoice_header`
- 인보이스 번호: `invoice_number`
- Vendor 참조: `vendor_invoice_no` 필드 직접 포함
- **장점:** 이미 그룹화 가능한 구조
- **사용:** `group_lines_by_vendor_invoice()` 함수로 바로 그룹화 가능  

---

## 3️⃣ 매칭·검증 로직 (OFCO 전용 + SAFEEN/ADP 특화)

### 3.1 SAFEEN/ADP 인보이스 타입 식별

**SAFEEN 인보이스 식별:**
- Description에 `"SAFEEN INV-"` 포함
- Vendor Invoice No: `"INV-XXXXX"` 형식
- Port: `"Musaffah Channel"`

**ADP 인보이스 식별:**
- Description에 `"ADP INV-"` 또는 `"ADP INV0325"` 포함
- Vendor Invoice No: `"INV-XXXXX"` 또는 `"INV0325XXXXX"` 형식
- Port: `"Musaffah Port GC"`

### 3.2 개별 인보이스 분개 처리

**vendor_invoice_no별 그룹화:**
- 하나의 OFCO 인보이스에 여러 vendor 인보이스 포함 가능
- 각 vendor_invoice_no별로 그룹화하여 분개
- vendor_invoice_no가 null인 경우 OFCO 자체 서비스로 처리

**그룹화 구조:**
```json
{
  "vendor_invoices": {
    "SAFEEN INV-83291": {
      "vendor_invoice_no": "SAFEEN INV-83291",
      "invoice_type": "SAFEEN",
      "lines": [...],
      "total_amount_aed": 0.0,
      "total_amount_usd": 0.0,
      "total_tax_aed": 0.0,
      "total_tax_usd": 0.0
    },
    "ADP INV-84402": {
      "vendor_invoice_no": "ADP INV-84402",
      "invoice_type": "ADP",
      "lines": [...],
      "total_amount_aed": 0.0,
      "total_amount_usd": 0.0,
      "total_tax_aed": 0.0,
      "total_tax_usd": 0.0
    }
  },
  "unassigned_lines": [...]  // vendor_invoice_no가 null인 라인 (OFCO 자체 서비스)
}
```

**처리 프로세스:**
1. `vendor_invoice_refs` 배열에서 첫 번째 값을 `vendor_invoice_no`로 추출
2. `vendor_invoice_no`가 없으면 `invoice_type="OFCO"`로 분류
3. 각 vendor 인보이스별로 합계 계산 및 검증
4. `group_lines_by_vendor_invoice()` 함수로 그룹화
5. `process_ldg_payload_with_vendor_invoice_grouping()` 함수로 통합 처리

**Parsed JSON 개선 필요:**
- Parsed JSON은 `vendor_invoice_refs`만 있고 `vendor_invoice_no` 필드가 없음
- `enhance_parsed_json_with_vendor_invoices()` 함수로 개선 후 사용
- 개선 후에는 LDG_PAYLOAD와 동일한 구조로 처리 가능

### 3.3 cost_center_a 형식 정규화

**LDG 형식:** `"PORT HANDLING CHARGE(CHANNEL TRANSIT CHARGES)"`  
**문서 형식:** `cost_center_a="PORT HANDLING CHARGE"`, `cost_center_b="CHANNEL TRANSIT CHARGES"`  
→ 괄호 형식을 자동으로 분리하여 문서 구조에 맞게 정규화

### 3.4 매칭·검증 프로세스

1. **BJ 확인:** Total Amount 잠정 고정  
2. **PDF 매칭:** OFCO 인보이스 파일명/회전/설명(Subject)/금액 기반 → 합계 BJ ± 2%  
   - 실패 시 `"OFCO PORT INV"` 참조(EA/Rate 힌트) → **Ref: Row 기록**  
   - 충돌 시 PDF 우선, `[MISMATCH]` 표기  
3. **EA/Rate 분해 (타입별):**

   **SAFEEN 인보이스 (LDG_PAYLOAD 기반):**
   - qty × rate = amount 검증 (±2% 허용오차)
   - 원본 구조 보존 또는 단순화 모드 적용

   **ADP 인보이스:**
   - 원본 Qty/Unit Price 구조 보존
   - ±2% 허용오차 검증

   **일반 인보이스:**
   - EA × Rate 합 = Amount ± 0.01 AED 허용 (±2% 내)  
   - EA NaN → EA=1, Rate=Amount 처리  
   - 잔여 슬롯=0  

4. **다중 라인 합산 검증:** 동일 인보이스 번호 + Subject 그룹 합계 = PDF Total  
5. **K:BA 입력:** 매핑 규칙 자동 채움, 없으면 EXT 행 삽입  
6. **others 필드:** `[EXT] Others` 행 생성, M열에 근거 기재  
7. **미매핑:** 군집 Rate/설명으로 ±2% 맞으면 기입, 아니면 `[UNMAPPED]`  
8. **VAT 검증:** 
   - NaN → PDF 근거 확인 후 채움
   - 5%·0% 외 → [MISMATCH]
   - `vat_check`: |VAT_USD - Amount_USD×0.05| ≤ 0.01 USD
9. **정합성 검증:**
   - **calc_check:** |EA_Total - Total_AED| / Total_AED ≤ 2%
   - **vat_check:** |VAT_USD - Amount_USD×0.05| ≤ 0.01 USD
   - **pc_check:** |Σ(Price Center AMOUNT) - Total_AED| ≤ 1.00 AED
   - Σ(BB:BI) = BJ ± 2%  
   - VAT/환율 일관성 100%  
   - EXT 금액 집계 제외  

---

## 4️⃣ PRICE CENTER (3-Way + cost_item_fields.JSON 매핑)

- **A/B:** Tariff·키워드 기반  
- **C:** 수수료·Pass·Document  
- **규칙:**  
  - C=0 의심 시 재검토  
  - A 금액 > 원본 또는 B < 0 → 일부를 C로 이동  
- **합계:** A + B + C = Original_TOTAL, Diff=0 유지  
- **⚠️ 중요: `cost_item_fields.JSON`의 필드명은 절대 변경 금지**
  - 모든 Price Center 필드명은 `cost_item_fields.JSON`에서 가져옴
  - 동적 생성 시에도 기존 필드명 그대로 사용
  - `cost_item.JSON`은 메타데이터(description, unit 등) 제공용이며, 실제 필드명은 `cost_item_fields.JSON`에서 가져옴  

---

## 5️⃣ QC & Fail-safe

- **Self-Check Loop:** `/critic_mode on` → 최대 2회 Self-Fix  
- **ZERO 트리거:** KPI 미달, 핵심 식별자 결손, Evidence 부재·상충  
- **Evidence 필수:** PDF p./line 또는 참조시트 Row 기록 (형식: "p1,row1" 또는 "p1,bbox(x1,y1,x2,y2)")
- **EXT 행 정책:**
  - K:BA에 필요한 필드가 없으면 컬럼 추가 금지
  - 본행 아래 EXT 행 삽입으로 메타 기록
  - EXT 행은 금액 집계에서 제외  

---

## 6️⃣ 보고 규격 (7+2 Sections)
1) Auto Guard Summary(3줄)  
1.5) Risk Assessment(등급/동인/신뢰도)  
2) Discrepancy Table(Δ·허용오차·상태)  
3) Compliance Matrix(UAE 법규·링크)  
4) Auto-Fill(Freight/Insurance)  
5) Auto Action Hooks(명령·가이드)  
6) DEM/DET & Gate-Out Forecast  
7) Evidence & Citations  
8) Weak Spot & Improvements  
9) Changelog  

---

## 7️⃣ 명령 세트

- OCR: `/ocr_basic`, `/ocr_table`, `/ocr_lowres_fix`, `/ocr_retry`  
- 매핑: `/mapping run`, `/run pricecenter map`, `/mapping update pricecenter`  
- 비용: `/switch_mode COST-GUARD + /logi-master invoice-audit --AEDonly`  
- 규제: `/logi-master cert-chk`, `/ocr_certchk`  
- 배치: `/workflow bulk …`, `/export excel`  
- **LDG 처리:** `/process_ldg_payload_with_vendor_invoice_grouping` (개별 인보이스 분개)
- **JSON 개선:** `/enhance_parsed_json_with_vendor_invoices` (parsed JSON에 vendor_invoices 섹션 추가)
- **타입 식별:** `/identify_invoice_type` (SAFEEN/ADP/OFCO 자동 식별)  

---

## 8️⃣ 완료 조건

- UNMAPPED=0, 편차 ≤ ±2%, VAT/환율/합계 100% 정합성  
- 모든 수치 **소수점 2자리 고정**  
- 각 항목별 **Evidence 기재 필수**  
- 실패·보류 건은 사유·대응안 요약  
- **calc_check/vat_check/pc_check 모두 PASS**
- **cost_item_fields.JSON 필드명 일관성 유지 (절대 변경 금지)**

---

## 9️⃣ 추가 참고사항

### 9.1 cost_item_fields.JSON 필드명 불변 원칙

**⚠️ 가장 중요: `cost_item_fields.JSON`의 필드명은 절대 변경할 수 없습니다.**

- 이 필드명들은 기존 엑셀 구조 및 외부 시스템과 직접 연결되어 있습니다.
- 필드명 변경 시 기존 데이터 호환성이 깨질 수 있습니다.
- 모든 매핑 로직은 `cost_item_fields.JSON`의 기존 필드명을 그대로 사용해야 합니다.
- `cost_item.JSON`은 메타데이터(description, unit 등) 제공용이며, 실제 필드명은 `cost_item_fields.JSON`에서 가져옵니다.

### 9.2 COST-GUARD (Δ% 밴드 검증)

**목적:**
- 단가 급변 조기 감지
- 계약 요율 준수 여부 확인
- 비정상 요율 자동 플래깅

**밴드 분류:**
- PASS: |Δ%| ≤ 3%
- WARN: 3% < |Δ%| ≤ 10%
- HIGH: 10% < |Δ%| ≤ 20%
- CRITICAL: |Δ%| > 20%

### 9.3 PDF↔엑셀 매칭 로직

**매칭 기준:**
1. Subject 패턴 유사도
2. 금액 일치 (±2%)
3. Tariff Code/ID 일치

**충돌 처리:**
- PDF 우선: PDF 라인 정보를 기준으로 사용
- [MISMATCH] 표기 추가
- 근거 기록 필수

### 9.4 LDG_PAYLOAD v2.4 처리 프로세스

1. **LDG_PAYLOAD 로드:** `/ocr_basic {file} mode:LDG+` 실행
2. **vendor_invoice_no별 그룹화:** 개별 인보이스 분개를 위한 그룹화
   - `group_lines_by_vendor_invoice()` 함수로 vendor_invoice_no별 그룹화
   - 각 그룹별 합계 계산 및 타입 식별 (SAFEEN/ADP/OFCO)
   - `vendor_invoice_refs[0]` → `vendor_invoice_no` 추출
   - vendor_invoice_no가 null인 경우 `unassigned_lines`로 분류
3. **cost_center 정규화:** LDG 형식 → 문서 형식 변환
   - `normalize_cost_center_from_ldg()` 함수로 괄호 형식 분리
   - 예: `"PORT HANDLING CHARGE(CHANNEL TRANSIT CHARGES)"` → `cost_center_a="PORT HANDLING CHARGE"`, `cost_center_b="CHANNEL TRANSIT CHARGES"`
4. **cost_item_fields 매핑:** `cost_item_fields_GROUPL.JSON`으로 빠른 조회
   - `map_ldg_line_to_cost_item_fields()` 함수로 매핑
   - price_center 우선, 실패 시 description 기반 추정
5. **EA 분해 적용:** SAFEEN/ADP/일반 타입별 분기 처리
   - SAFEEN: `decompose_safeen_line_from_ldg()` (qty × rate 검증, ±2% 허용오차)
   - ADP: `decompose_adp_line()` (원본 Qty/Unit Price 구조 보존)
   - 일반: `decompose_to_ea_slots()` (기본 EA 분해 로직)
6. **검증 결과 집계:** calc_check, vat_check, pc_check 통합
   - 각 vendor 인보이스별 검증 결과 집계
   - 전체 인보이스 검증 결과 통합
7. **엑셀 삽입:** 표준 라인을 엑셀에 append
   - `process_ldg_payload_with_vendor_invoice_grouping()` 함수로 통합 처리
   - vendor 인보이스별로 순차 삽입 또는 일괄 삽입

### 9.5 Parsed JSON 개선 (vendor_invoices 섹션 추가) - 검증 완료

**검증 결과:**
- 합계 검증: PASS (Δ=0.0)
- `vendor_invoices` 섹션: 없음 (추가 필요)
- `vendor_invoice_no` 필드: 각 라인에 없음 (추가 필요)
- `invoice_type` 필드: 각 라인에 없음 (추가 필요)
- `vendor_invoice_refs` 배열: 있음 (추출 가능)

**개선 방법:**
1. `vendor_invoice_refs` 배열에서 첫 번째 값 추출 시도
2. **description에서 직접 추출** (개선된 패턴)
3. 각 라인에 `invoice_type` 필드 추가 (SAFEEN/ADP/OFCO 식별)
4. `vendor_invoices` 섹션 추가 (vendor 인보이스별 그룹화)
5. `unassigned_lines` 섹션 추가 (OFCO 자체 서비스 라인)
6. totals 재계산 (null 값 제외, AED/USD 모두)

**개선 스크립트 (검증 완료 버전):**

```python
import re

def extract_vendor_invoice_no_from_description(description: str) -> str:
    """
    description에서 vendor_invoice_no 추출 (개선된 버전)
    
    패턴:
    - SAFEEN: "SAFEEN INV-XXXXX", "SAFEEN INV-XXXXX-", "SAFEEN INV: XXXX"
    - ADP: "ADP INV-XXXXX", "ADP INV0325XXXXX", "ADP INV0325XXXXX -"
    """
    if not description:
        return None
    
    description_upper = description.upper()
    
    # SAFEEN 패턴 (다양한 형식 지원)
    safeen_patterns = [
        r'SAFEEN\s+INV-(\d+)',           # "SAFEEN INV-77262"
        r'SAFEEN\s+INV[:\s-]+(\d+)',    # "SAFEEN INV: 77262" 또는 "SAFEEN INV 77262"
        r'SAFEEN\s+INV-(\d+)-',         # "SAFEEN INV-77262-"
    ]
    
    for pattern in safeen_patterns:
        match = re.search(pattern, description_upper)
        if match:
            return f"SAFEEN INV-{match.group(1)}"
    
    # ADP 패턴 (다양한 형식 지원)
    adp_patterns = [
        r'ADP\s+INV-(\d+)',                    # "ADP INV-84402"
        r'ADP\s+INV(\d{11})',                  # "ADP INV03251101212"
        r'ADP\s+INV(\d{11})\s*-',              # "ADP INV03251101212 -"
        r'ADP\s+INV[:\s-]+(\d+)',              # "ADP INV: 84402"
    ]
    
    for pattern in adp_patterns:
        match = re.search(pattern, description_upper)
        if match:
            inv_num = match.group(1)
            if len(inv_num) == 11:
                return f"ADP INV{inv_num}"
            else:
                return f"ADP INV-{inv_num}"
    
    return None

def identify_invoice_type_from_description(description: str, vendor_invoice_no: str = None) -> str:
    """
    description과 vendor_invoice_no로 인보이스 타입 식별 (개선된 버전)
    """
    # vendor_invoice_no 우선 확인
    if vendor_invoice_no:
        vendor_upper = vendor_invoice_no.upper()
        if "SAFEEN" in vendor_upper:
            return "SAFEEN"
        elif "ADP" in vendor_upper:
            return "ADP"
    
    # description에서 직접 식별
    description_upper = description.upper() if description else ""
    
    if "SAFEEN INV-" in description_upper or "SAFEEN" in description_upper:
        return "SAFEEN"
    elif "ADP INV-" in description_upper or "ADP INV0325" in description_upper or "ADP" in description_upper:
        return "ADP"
    
    return "OFCO"

def enhance_parsed_json_with_vendor_invoices_fixed(parsed_json: dict) -> dict:
    """
    Parsed JSON에 vendor_invoices 섹션 추가 (검증 완료 버전)
    
    개선 사항:
    1. vendor_invoice_refs 배열 우선 추출
    2. description에서 직접 추출 (fallback)
    3. totals 재계산 (null 값 제외, AED/USD 모두)
    4. vendor_invoices 섹션 생성
    5. unassigned_lines 섹션 생성
    """
    lines = parsed_json.get("lines", [])
    vendor_invoices = {}
    unassigned_lines = []
    
    for line in lines:
        description = line.get("description", "")
        
        # 1. 기존 vendor_invoice_no 확인
        vendor_invoice_no = line.get("vendor_invoice_no")
        
        # 2. vendor_invoice_refs 배열에서 추출 (우선)
        if not vendor_invoice_no:
            vendor_refs = line.get("vendor_invoice_refs", [])
            if vendor_refs:
                vendor_invoice_no = vendor_refs[0]  # 첫 번째 값 사용
        
        # 3. description에서 직접 추출 (fallback)
        if not vendor_invoice_no:
            vendor_invoice_no = extract_vendor_invoice_no_from_description(description)
        
        # 4. invoice_type 식별
        invoice_type = identify_invoice_type_from_description(description, vendor_invoice_no)
        
        # 5. 라인 업데이트
        line["vendor_invoice_no"] = vendor_invoice_no
        line["invoice_type"] = invoice_type
        
        # 6. 그룹화
        if vendor_invoice_no:
            if vendor_invoice_no not in vendor_invoices:
                vendor_invoices[vendor_invoice_no] = {
                    "vendor_invoice_no": vendor_invoice_no,
                    "invoice_type": invoice_type,
                    "lines": [],
                    "totals": {
                        "amount_excl_vat_aed": 0.0,
                        "tax_amount_aed": 0.0,
                        "amount_incl_vat_aed": 0.0,
                        "amount_excl_vat_usd": 0.0,
                        "tax_amount_usd": 0.0,
                        "amount_incl_vat_usd": 0.0
                    },
                    "line_count": 0
                }
            
            vendor_invoices[vendor_invoice_no]["lines"].append(line)
            
            # 합계 계산 (null 값 제외, AED/USD 모두)
            amount_excl_aed = line.get("amount_excl_tax_aed") or 0.0
            tax_amount_aed = line.get("tax_amount_aed") or 0.0
            amount_incl_aed = line.get("total_incl_tax_aed") or 0.0
            amount_excl_usd = line.get("amount_excl_tax_usd") or 0.0
            tax_amount_usd = line.get("tax_amount_usd") or 0.0
            amount_incl_usd = line.get("total_incl_tax_usd") or 0.0
            
            vendor_invoices[vendor_invoice_no]["totals"]["amount_excl_vat_aed"] += amount_excl_aed
            vendor_invoices[vendor_invoice_no]["totals"]["tax_amount_aed"] += tax_amount_aed
            vendor_invoices[vendor_invoice_no]["totals"]["amount_incl_vat_aed"] += amount_incl_aed
            vendor_invoices[vendor_invoice_no]["totals"]["amount_excl_vat_usd"] += amount_excl_usd
            vendor_invoices[vendor_invoice_no]["totals"]["tax_amount_usd"] += tax_amount_usd
            vendor_invoices[vendor_invoice_no]["totals"]["amount_incl_vat_usd"] += amount_incl_usd
            vendor_invoices[vendor_invoice_no]["line_count"] += 1
        else:
            unassigned_lines.append(line)
    
    # 7. JSON 업데이트
    parsed_json["vendor_invoices"] = vendor_invoices
    parsed_json["unassigned_lines"] = unassigned_lines
    
    # 8. totals 재계산 (null 값 제외, AED/USD 모두)
    total_excl_aed = sum(
        line.get("amount_excl_tax_aed") or 0.0 
        for line in lines
    )
    total_tax_aed = sum(
        line.get("tax_amount_aed") or 0.0 
        for line in lines
    )
    total_incl_aed = sum(
        line.get("total_incl_tax_aed") or 0.0 
        for line in lines
    )
    total_excl_usd = sum(
        line.get("amount_excl_tax_usd") or 0.0 
        for line in lines
    )
    total_tax_usd = sum(
        line.get("tax_amount_usd") or 0.0 
        for line in lines
    )
    total_incl_usd = sum(
        line.get("total_incl_tax_usd") or 0.0 
        for line in lines
    )
    
    parsed_json["totals"] = {
        "sum_amount_excl_vat_aed": total_excl_aed,
        "sum_tax_amount_aed": total_tax_aed,
        "sum_amount_incl_vat_aed": total_incl_aed,
        "sum_amount_excl_vat_usd": total_excl_usd,
        "sum_tax_amount_usd": total_tax_usd,
        "sum_amount_incl_vat_usd": total_incl_usd,
        "line_count": len([l for l in lines if (l.get("amount_excl_tax_aed") or l.get("amount_excl_tax_usd")) is not None])
    }
    
    return parsed_json
```

**사용 예시:**
```python
import json

# Parsed JSON 로드
with open("OFCO-INV-0002054_parsed.json", "r", encoding="utf-8") as f:
    parsed_json = json.load(f)

# 개선 실행
enhanced_json = enhance_parsed_json_with_vendor_invoices_fixed(parsed_json)

# 통계 출력
print(f"총 라인 수: {len(enhanced_json['lines'])}")
print(f"Vendor 인보이스 수: {len(enhanced_json['vendor_invoices'])}")
print(f"Unassigned 라인 수: {len(enhanced_json['unassigned_lines'])}")

# SAFEEN/ADP 인보이스 확인
safeen_count = sum(1 for v in enhanced_json["vendor_invoices"].values() if v["invoice_type"] == "SAFEEN")
adp_count = sum(1 for v in enhanced_json["vendor_invoices"].values() if v["invoice_type"] == "ADP")
print(f"SAFEEN 인보이스 수: {safeen_count}")
print(f"ADP 인보이스 수: {adp_count}")

# Vendor 인보이스별 요약
for vendor_invoice_no, vendor_data in enhanced_json["vendor_invoices"].items():
    print(f"\n{vendor_invoice_no}:")
    print(f"  타입: {vendor_data['invoice_type']}")
    print(f"  라인 수: {vendor_data['line_count']}")
    print(f"  총액 (AED): {vendor_data['totals']['amount_excl_vat_aed']:,.2f}")
    print(f"  총액 (USD): {vendor_data['totals']['amount_excl_vat_usd']:,.2f}")

# 개선된 JSON 저장
with open("OFCO-INV-0002054_parsed_enhanced.json", "w", encoding="utf-8") as f:
    json.dump(enhanced_json, f, indent=2, ensure_ascii=False)

print("\n✅ 개선된 JSON 저장 완료")
```

**검증 완료 사항:**
- ✅ 합계 검증: PASS (Δ=0.0)
- ✅ vendor_invoice_refs 배열에서 추출 성공
- ✅ description에서 직접 추출 성공 (SAFEEN/ADP 패턴)
- ✅ vendor_invoices 섹션 생성
- ✅ totals 재계산 (null 값 제외, AED/USD 모두)

**개선 전후 비교:**
- **개선 전:** `vendor_invoice_refs` 배열만 있음, description에서 추출 실패, `vendor_invoices` 섹션에 SAFEEN 인보이스 누락
- **개선 후:** description에서 직접 추출, 모든 SAFEEN/ADP 인보이스가 `vendor_invoices` 섹션에 포함, totals 재계산 (AED/USD 모두)
- **결과:** LDG_PAYLOAD와 동일한 구조로 `process_ldg_payload_with_vendor_invoice_grouping()` 함수 사용 가능

---

**이 가이드는 OFCO 인보이스 엑셀 삽입 로직 상세 문서와 완전히 통합되었습니다.**  
