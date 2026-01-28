# OFCO 인보이스 엑셀 삽입 로직 상세 문서

## 목차
1. [전체 파이프라인 개요](#1-전체-파이프라인-개요)
2. [입력 데이터 준비](#2-입력-데이터-준비)
3. [표준 라인 생성](#3-표준-라인-생성)
4. [Cost/Price Center 매핑](#4-costprice-center-매핑)
   - 4.4 [SAFEEN/ADP 인보이스 특화 처리](#44-safeenadp-인보이스-특화-처리)
   - 4.4.4 [SAFEEN/ADP 인보이스 → cost_item.JSON 매핑 예시](#444-safeenadp-인보이스--cost_itemjson-매핑-예시)
5. [EA 분해 알고리즘](#5-ea-분해-알고리즘)
   - 5.1.1 [SAFEEN 인보이스 EA 분해](#511-safeen-인보이스-ea-분해)
   - 5.1.2 [ADP 인보이스 EA 분해](#512-adp-인보이스-ea-분해)
6. [Price Center 피벗 생성](#6-price-center-피벗-생성)
7. [검증 로직](#7-검증-로직)
8. [엑셀 삽입 규칙](#8-엑셀-삽입-규칙)
   - 8.5 [SAFEEN/ADP 인보이스 엑셀 삽입 특화 규칙](#85-safeenadp-인보이스-엑셀-삽입-특화-규칙)
9. [구현 코드 예시](#9-구현-코드-예시)
10. [주의사항 및 Best Practices](#10-주의사항-및-best-practices)
11. [JSON 파일 구조 및 관계](#11-json-파일-구조-및-관계)
12. [Risk Assessment 및 완화 전략](#12-risk-assessment-및-완화-전략)
13. [Evidence 및 EXT 행 정책](#13-evidence-및-ext-행-정책)
14. [PDF↔엑셀 매칭 로직](#14-pdf엑셀-매칭-로직)
15. [COST-GUARD (Δ% 밴드 검증)](#15-cost-guard-밴드-검증)

---

## ⚠️ 중요 제약사항

### cost_item_fields.JSON 필드명 불변 원칙

**`cost_item_fields.JSON`의 필드명은 절대 변경할 수 없습니다.**

- 이 필드명들은 기존 엑셀 구조 및 외부 시스템과 직접 연결되어 있습니다.
- 필드명 변경 시 기존 데이터 호환성이 깨질 수 있습니다.
- 모든 매핑 로직은 `cost_item_fields.JSON`의 기존 필드명을 그대로 사용해야 합니다.
- `cost_item.JSON`은 메타데이터(description, unit 등) 제공용이며, 실제 필드명은 `cost_item_fields.JSON`에서 가져옵니다.

---

## 1. 전체 파이프라인 개요

### 1.1 파이프라인 흐름도

```
[PDF/이미지 인보이스]
    ↓
[OCR 추출 (LDG+)]
    ↓
[PDF↔엑셀 매칭]
    ↓
[표준 라인 생성]
    ↓
[Subject → Cost/Price Center 매핑]
    ↓
[EA 분해 (최대 4 RatePair)]
    ↓
[Price Center 피벗 생성]
    ↓
[통화 변환 (AED ↔ USD)]
    ↓
[3종 검증 (calc/vat/pc)]
    ↓
[엑셀 삽입 (Sheet1 append)]
```

### 1.2 데이터 레이어 구조

```
Layer 1: 메타데이터 (Invoice No, Vessel, Date, etc.)
Layer 2: Cost Center A/B
Layer 3: Price Center (44개 × 2 = 88컬럼)
Layer 4: 통화 (AED/USD)
Layer 5: EA Slot (최대 4쌍)
Layer 6: 중간계산 (Qty × Rate)
Layer 7: 검증 (calc_check/vat_check/pc_check)
```

---

## 2. 입력 데이터 준비

### 2.1 OCR 출력 구조

```python
ocr_output = {
    "invoice_meta": {
        "invoice_no": "OFCO-INV-0002054",
        "vessel_name": "JOPETWIL 71",
        "rotation_no": "2503129579",
        "bol": "HVDC-AGI-GRM-J71-70",
        "port": "Musaffah Channel",
        "arrival_date": "05-Oct-2025",
        "departure_date": "05-Oct-2025 19:12",
        "currency": "AED",
        "exchange_rate": 3.6725,  # USD→AED (있는 경우)
        "grand_total_aed": 3291.25,
        "vat_total_aed": 0.00,
        "total_incl_vat_aed": 3291.25
    },
    "lines": [
        {
            "line_no": 1,
            "tariff_code": "6.1",
            "description": "Administration Fees Channel Transit Request",
            "unit1": None,  # 또는 1.000
            "unit2": None,
            "unit3": None,
            "rate": None,  # 또는 100.00
            "amount_excl_tax": 100.00,
            "vat_pct": 0,
            "vat_amount": 0.00,
            "total_incl_tax": 100.00,
            "evidence": "p1,row1"  # PDF 페이지/행 위치
        }
    ],
    "ocr_kpi": {
        "mean_confidence": 0.95,
        "table_accuracy": 0.99,
        "numeric_integrity": 1.00
    }
}
```

### 2.2 OCR KPI 게이트

```python
def validate_ocr_kpi(ocr_output: dict) -> bool:
    """
    OCR KPI 검증: 미달 시 ZERO 중단
    """
    kpi = ocr_output.get("ocr_kpi", {})
    
    if kpi.get("mean_confidence", 0) < 0.92:
        raise ValueError("OCR MeanConf < 0.92: ZERO 중단")
    
    if kpi.get("table_accuracy", 0) < 0.98:
        raise ValueError("OCR TableAcc < 0.98: ZERO 중단")
    
    if kpi.get("numeric_integrity", 0) < 1.00:
        raise ValueError("OCR NumericIntegrity < 1.00: ZERO 중단")
    
    return True
```

---

## 3. 표준 라인 생성

### 3.1 표준 라인 스키마

```python
standard_line_schema = {
    "invoice_no": str,           # 필수
    "line_no": int,              # 필수
    "tariff_id": str,            # 선택 (예: "6.1", "2.20", "201.3")
    "description": str,          # 필수
    "unit1": float,             # 선택 (톤/CBM/GT 등)
    "unit2": float,             # 선택
    "unit3": float,             # 선택
    "rate": float,              # 선택 (단가)
    "amount_excl_tax_aed": float,  # 필수
    "tax_rate_pct": float,      # 필수 (0 또는 5)
    "tax_amount_aed": float,    # 필수
    "total_incl_tax_aed": float,  # 필수
    "evidence": str,            # 필수 (예: "p1,row1")
    "cost_center_a": str,       # 매핑 후
    "cost_center_b": str,       # 매핑 후
    "price_center": str,        # 매핑 후
    "cost_item_code": str,      # 매핑 후
    "qty_field": str,           # 매핑 후
    "amount_field": str         # 매핑 후
}
```

### 3.2 표준 라인 생성 함수

```python
def create_standard_line(ocr_line: dict, invoice_meta: dict) -> dict:
    """
    OCR 라인을 표준 라인 형식으로 변환
    """
    standard_line = {
        "invoice_no": invoice_meta["invoice_no"],
        "line_no": ocr_line["line_no"],
        "tariff_id": ocr_line.get("tariff_code") or ocr_line.get("tariff_id"),
        "description": ocr_line["description"],
        "unit1": ocr_line.get("unit1") or 0.0,
        "unit2": ocr_line.get("unit2") or 0.0,
        "unit3": ocr_line.get("unit3") or 0.0,
        "rate": ocr_line.get("rate"),
        "amount_excl_tax_aed": ocr_line["amount_excl_tax"],
        "tax_rate_pct": ocr_line.get("vat_pct", 0),
        "tax_amount_aed": ocr_line.get("vat_amount", 0.0),
        "total_incl_tax_aed": ocr_line["total_incl_tax"],
        "evidence": ocr_line.get("evidence", "p1,row1")
    }
    
    return standard_line
```

---

## 4. Cost/Price Center 매핑

### 4.1 Subject 패턴 매핑 테이블

```python
SUBJECT_PATTERN_MAPPING = {
    # SAFEEN 관련
    "SAFEEN.*Channel.*Crossing": {
        "cost_center_a": "PORT HANDLING CHARGE",
        "cost_center_b": "CHANNEL TRANSIT CHARGES",
        "price_center": "CHANNEL TRANSIT CHARGES",
        "cost_item_code": "CHANNEL_CROSSING_CHARGES_FOR_VESSELS_WITH_1000_TO_3_001_GT"
    },
    "Administration Fees.*Channel.*Transit": {
        "cost_center_a": "PORT HANDLING CHARGE",
        "cost_center_b": "CHANNEL TRANSIT CHARGES",
        "price_center": "CHANNEL TRANSIT CHARGES",
        "cost_item_code": "CHANNEL_TRANSIT_CROSSING_REQUEST"
    },
    "Administration Fees.*Channel.*Shifting": {
        "cost_center_a": "PORT HANDLING CHARGE",
        "cost_center_b": "CHANNEL TRANSIT CHARGES",
        "price_center": "CHANNEL TRANSIT CHARGES",
        "cost_item_code": "CHANNEL_TRANSIT_CROSSING_REQUEST"
    },
    
    # ADP Port 관련
    "ADP.*Port.*Dues": {
        "cost_center_a": "PORT HANDLING CHARGE",
        "cost_center_b": "PORT DUES & SERVICES CHARGES",
        "price_center": "PORT DUES",
        "cost_item_code": "PORT_DUES_FOR_VESSELS_WITH_ABOVE_1_000_UP_TO_3_001GT"
    },
    
    # Document Processing
    "Document Processing Charge": {
        "cost_center_a": "PORT HANDLING CHARGE",
        "cost_center_b": "DOCUMENT PROCESSING",
        "price_center": "DOCUMENT PROCESSING CHARGE",
        "cost_item_code": "DOCUMENT_PROCESSING_CHARGE"
    },
    "VAT - Document Processing Charge": {
        "cost_center_a": "PORT HANDLING CHARGE",
        "cost_center_b": "DOCUMENT PROCESSING",
        "price_center": "DOCUMENT PROCESSING CHARGE",
        "cost_item_code": "DOCUMENT_PROCESSING_CHARGE"
    },
    
    # Bulk Material
    "Bulk Material.*Solids.*0-10.*Tons.*Direct Delivery": {
        "cost_center_a": "PORT HANDLING CHARGE",
        "cost_center_b": "BULK MATERIAL HANDLING",
        "price_center": "BULK MATERIAL (PHC)",
        "cost_item_code": "BULK_MATERIAL_SOLIDS_A_PARCEL_SIZE_0_10_001_TONS_DIRECT_DELIVERY"
    },
    
    # Agency Fees
    "Cargo Clearance": {
        "cost_center_a": "CONTRACT",
        "cost_center_b": "AF FOR CC",
        "price_center": "AGENCY FEE FOR CARGO CLEARANCE",
        "cost_item_code": "AGENCY_FEE_FOR_CARGO_CLEARANCE"
    },
    "Berthing Arrangement": {
        "cost_center_a": "CONTRACT(AF FOR BA)",
        "cost_center_b": "CONTRACT",
        "price_center": "AGENCY FEE FOR BERTHING ARRANGEMENT",
        "cost_item_code": "AGENCY_FEE_FOR_BERTHING_ARRANGEMENT"
    },
    "FW Supply|Arranging FW Supply": {
        "cost_center_a": "CONTRACT",
        "cost_center_b": "AF FOR FW SA",
        "price_center": "SUPPLY WATER 5000IG",
        "cost_item_code": "SUPPLY_WATER_5001IG"
    }
}
```

### 4.2 cost_item_fields.JSON 기반 필드명 매핑

**중요**: `cost_item_fields.JSON`의 필드명은 절대 변경하지 않습니다. 모든 필드명은 이 파일에서 가져와야 합니다.

```python
def map_subject_to_cost_item(
    subject: str, 
    cost_items: dict,
    cost_item_fields: dict,  # cost_item_fields.JSON 추가
    pattern_mapping: dict = SUBJECT_PATTERN_MAPPING
) -> dict:
    """
    Subject를 분석하여 Cost/Price Center 및 cost_item 필드명 매핑
    
    중요: cost_item_fields.JSON의 필드명은 절대 변경하지 않음
    """
    import re
    
    # 1. 패턴 매칭
    matched_pattern = None
    for pattern, mapping in pattern_mapping.items():
        if re.search(pattern, subject, re.IGNORECASE):
            matched_pattern = mapping
            break
    
    if not matched_pattern:
        # 기본값 (OTHERS)
        matched_pattern = {
            "cost_center_a": "OTHERS",
            "cost_center_b": "OTHERS",
            "price_center": "OTHERS",
            "cost_item_code": "OTHERS"
        }
    
    # 2. cost_item.JSON에서 메타데이터 가져오기
    cost_item_code = matched_pattern["cost_item_code"]
    cost_item = None
    
    for item in cost_items.get("cost_items", []):
        if item["code"] == cost_item_code:
            cost_item = item
            break
    
    # 3. cost_item_fields.JSON에서 필드명 가져오기 (절대 변경 금지)
    qty_field = None
    amount_field = None
    
    # cost_item_fields.JSON의 필드명을 그대로 사용
    for field in cost_item_fields.get("cost_item_fields", []):
        if field.endswith("_QTY") and cost_item_code in field:
            qty_field = field  # 기존 필드명 그대로 사용
        elif field.endswith("_AMOUNT") and cost_item_code in field:
            amount_field = field  # 기존 필드명 그대로 사용
    
    # 필드명을 찾지 못한 경우 (새로운 항목 추가 시)
    if not qty_field or not amount_field:
        # cost_item.JSON의 qty_field/amount_field 사용 (fallback)
        if cost_item:
            qty_field = cost_item.get("qty_field")
            amount_field = cost_item.get("amount_field")
        else:
            # 최후의 수단: 기본값 생성 (하지만 실제로는 cost_item_fields.JSON에 추가 필요)
            qty_field = f"{cost_item_code}_QTY"
            amount_field = f"{cost_item_code}_AMOUNT"
            print(f"⚠️ WARNING: {cost_item_code} 필드명을 cost_item_fields.JSON에서 찾을 수 없습니다.")
    
    # 4. 최종 매핑 결과 반환
    return {
        "cost_center_a": matched_pattern["cost_center_a"],
        "cost_center_b": matched_pattern["cost_center_b"],
        "price_center": matched_pattern["price_center"],
        "cost_item_code": cost_item_code,
        "qty_field": qty_field,  # cost_item_fields.JSON의 기존 필드명
        "amount_field": amount_field,  # cost_item_fields.JSON의 기존 필드명
        "unit": cost_item.get("unit", "건당") if cost_item else "건당"
    }
```

### 4.3 표준 라인에 매핑 적용

```python
def apply_cost_price_center_mapping(
    standard_line: dict,
    cost_items: dict,
    cost_item_fields: dict  # cost_item_fields.JSON 추가
) -> dict:
    """
    표준 라인에 Cost/Price Center 매핑 적용
    
    중요: cost_item_fields.JSON의 필드명은 절대 변경하지 않음
    """
    subject = standard_line["description"]
    mapping = map_subject_to_cost_item(subject, cost_items, cost_item_fields)
    
    # 표준 라인에 매핑 정보 추가
    standard_line.update({
        "cost_center_a": mapping["cost_center_a"],
        "cost_center_b": mapping["cost_center_b"],
        "price_center": mapping["price_center"],
        "cost_item_code": mapping["cost_item_code"],
        "qty_field": mapping["qty_field"],  # cost_item_fields.JSON의 필드명
        "amount_field": mapping["amount_field"],  # cost_item_fields.JSON의 필드명
        "unit": mapping["unit"]
    })
    
    return standard_line
```

### 4.4 SAFEEN/ADP 인보이스 특화 처리

SAFEEN과 ADP 인보이스는 각각 고유한 구조와 분해 방식을 가지고 있습니다.

#### 4.4.1 SAFEEN 인보이스 특징

**SAFEEN 인보이스 (Musaffah Channel):**
- **형식**: 시간 기반 요금 구조
- **Tariff Code**: 6.1 (Administration Fees), 6.6 (Channel Crossing)
- **특징**:
  - 시간 단위로 요금 계산 (예: "3 Hours")
  - Description에 시간 정보 포함 (예: "09-Oct-2025 04:12:00 to 09-Oct-2025 07:12:00 3 Hours")
  - VAT 0%
  - 단가(Rate) 정보가 명시되지 않음 (총액만 표시)

**SAFEEN 인보이스 처리 로직:**

```python
def identify_safeen_invoice(invoice_meta: dict, lines: list) -> bool:
    """
    SAFEEN 인보이스 식별
    """
    # 포트명으로 식별
    if "Musaffah Channel" in invoice_meta.get("port", ""):
        return True
    
    # Tariff Code로 식별
    for line in lines:
        tariff_code = line.get("tariff_code") or line.get("tariff_id")
        if tariff_code in ["6.1", "6.6"]:
            return True
    
    return False

def parse_safeen_line(ocr_line: dict) -> dict:
    """
    SAFEEN 인보이스 라인 파싱
    """
    description = ocr_line["description"]
    
    # 시간 정보 추출 (예: "3 Hours")
    import re
    hours_match = re.search(r'(\d+)\s*Hours?', description, re.IGNORECASE)
    hours = float(hours_match.group(1)) if hours_match else 1.0
    
    # 시간 범위 추출
    time_range_match = re.search(
        r'(\d{2}-\w{3}-\d{4}\s+\d{2}:\d{2}:\d{2})\s+to\s+(\d{2}-\w{3}-\d{4}\s+\d{2}:\d{2}:\d{2})',
        description
    )
    
    parsed_line = {
        "line_no": ocr_line["line_no"],
        "tariff_code": ocr_line.get("tariff_code"),
        "description": description,
        "unit1": hours,  # 시간을 Unit1에 저장
        "unit2": 0.0,
        "unit3": 0.0,
        "rate": None,  # SAFEEN은 Rate 정보 없음
        "amount_excl_tax": ocr_line["amount_excl_tax"],
        "vat_pct": ocr_line.get("vat_pct", 0),
        "vat_amount": ocr_line.get("vat_amount", 0.0),
        "total_incl_tax": ocr_line["total_incl_tax"],
        "evidence": ocr_line.get("evidence", "p1,row1"),
        "invoice_type": "SAFEEN",
        "time_range": {
            "start": time_range_match.group(1) if time_range_match else None,
            "end": time_range_match.group(2) if time_range_match else None
        } if time_range_match else None
    }
    
    return parsed_line
```

#### 4.4.2 ADP 인보이스 특징

**ADP 인보이스 (Musaffah Port GC):**
- **형식**: 수량 기반 요금 구조
- **Tariff ID**: 2.20 (Document Processing), 201.3 (Bulk Material)
- **특징**:
  - 수량(Unit 1, 2, 3)과 단가(Rate) 정보 명시
  - Description에 "VAT -" 접두사 포함
  - VAT 5% 표기 (실제 세액은 0.00인 경우 많음)
  - 계산식: Amount = Unit × Rate

**ADP 인보이스 처리 로직:**

```python
def identify_adp_invoice(invoice_meta: dict, lines: list) -> bool:
    """
    ADP 인보이스 식별
    """
    # 포트명으로 식별
    if "Musaffah Port GC" in invoice_meta.get("port", ""):
        return True
    
    # Tariff ID로 식별
    for line in lines:
        tariff_id = line.get("tariff_id") or line.get("tariff_code")
        if tariff_id in ["2.20", "201.3"]:
            return True
    
    # Description 패턴으로 식별
    for line in lines:
        desc = line.get("description", "")
        if "VAT -" in desc and ("Document Processing" in desc or "Bulk Material" in desc):
            return True
    
    return False

def parse_adp_line(ocr_line: dict) -> dict:
    """
    ADP 인보이스 라인 파싱
    """
    parsed_line = {
        "line_no": ocr_line["line_no"],
        "tariff_id": ocr_line.get("tariff_id") or ocr_line.get("tariff_code"),
        "description": ocr_line["description"],
        "unit1": ocr_line.get("unit1") or 0.0,
        "unit2": ocr_line.get("unit2") or 0.0,
        "unit3": ocr_line.get("unit3") or 0.0,
        "rate": ocr_line.get("rate"),
        "amount_excl_tax": ocr_line["amount_excl_tax"],
        "vat_pct": ocr_line.get("vat_pct", 5.0),  # ADP는 보통 5%
        "vat_amount": ocr_line.get("vat_amount", 0.0),
        "total_incl_tax": ocr_line["total_incl_tax"],
        "evidence": ocr_line.get("evidence", "p1,row1"),
        "invoice_type": "ADP"
    }
    
    # 계산 검증: Unit × Rate = Amount
    if parsed_line["unit1"] > 0 and parsed_line["rate"]:
        calculated = parsed_line["unit1"] * parsed_line["rate"]
        actual = parsed_line["amount_excl_tax"]
        diff = abs(calculated - actual)
        
        if diff > 0.01:  # 0.01 AED 허용오차
            print(f"⚠️ WARNING: Line {parsed_line['line_no']} 계산 불일치: {calculated} vs {actual}")
    
    return parsed_line
```

#### 4.4.3 SAFEEN/ADP 인보이스 통합 처리

```python
def process_safeen_adp_invoice(
    ocr_output: dict,
    cost_items: dict,
    cost_item_fields: dict
) -> list:
    """
    SAFEEN 또는 ADP 인보이스를 표준 라인으로 변환
    """
    invoice_meta = ocr_output["invoice_meta"]
    lines = ocr_output["lines"]
    
    # 인보이스 타입 식별
    is_safeen = identify_safeen_invoice(invoice_meta, lines)
    is_adp = identify_adp_invoice(invoice_meta, lines)
    
    standard_lines = []
    
    for ocr_line in lines:
        # SAFEEN 인보이스 처리
        if is_safeen:
            parsed_line = parse_safeen_line(ocr_line)
        # ADP 인보이스 처리
        elif is_adp:
            parsed_line = parse_adp_line(ocr_line)
        # 일반 인보이스 처리
        else:
            parsed_line = create_standard_line(ocr_line, invoice_meta)
        
        # Cost/Price Center 매핑 적용
        standard_line = apply_cost_price_center_mapping(
            parsed_line,
            cost_items,
            cost_item_fields
        )
        
        standard_lines.append(standard_line)
    
    return standard_lines
```

#### 4.4.4 SAFEEN/ADP 인보이스 → cost_item.JSON 매핑 예시

**SAFEEN 인보이스 매핑:**

```python
SAFEEN_COST_ITEM_MAPPING = {
    "6.1": {  # Administration Fees
        "cost_item_code": "CHANNEL_TRANSIT_CROSSING_REQUEST",
        "qty_field": "CHANNEL_TRANSIT_CROSSING_REQUEST_QTY",  # cost_item_fields.JSON
        "amount_field": "CHANNEL_TRANSIT_CROSSING_REQUEST_AMOUNT",  # cost_item_fields.JSON
        "unit": "건당",
        "description": "항만 채널 통과 허가 요청 수수료"
    },
    "6.6": {  # Channel Crossing
        "cost_item_code": "CHANNEL_CROSSING_CHARGES_FOR_VESSELS_WITH_1000_TO_3_001_GT",
        "qty_field": "CHANNEL_CROSSING_CHARGES_FOR_VESSELS_WITH_1000_TO_3_001_GT_QTY",
        "amount_field": "CHANNEL_CROSSING_CHARGES_FOR_VESSELS_WITH_1000_TO_3_001_GT_AMOUNT",
        "unit": "GT",
        "description": "1,000~3,001GT 선박 채널 통과 비용"
    }
}
```

**ADP 인보이스 매핑:**

```python
ADP_COST_ITEM_MAPPING = {
    "2.20": {  # Document Processing
        "cost_item_code": "DOCUMENT_PROCESSING_CHARGE",
        "qty_field": "DOCUMENT_PROCESSING_CHARGE_QTY",  # cost_item_fields.JSON
        "amount_field": "DOCUMENT_PROCESSING_CHARGE_AMOUNT",  # cost_item_fields.JSON
        "unit": "건당",
        "description": "선적/하역/통관 서류를 항만 시스템에 등록·처리하는 수수료"
    },
    "201.3": {  # Bulk Material
        "cost_item_code": "BULK_MATERIAL_SOLIDS_A_PARCEL_SIZE_0_10_001_TONS_DIRECT_DELIVERY",
        "qty_field": "BULK_MATERIAL_SOLIDS_A_PARCEL_SIZE_0_10_001_TONS_DIRECT_DELIVERY_QTY",
        "amount_field": "BULK_MATERIAL_SOLIDS_A_PARCEL_SIZE_0_10_001_TONS_DIRECT_DELIVERY_AMOUNT",
        "unit": "톤(MT)",
        "description": "0~10,001톤 벌크 고형 화물의 직배송 하역·이동 수수료"
    }
}
```

---

## 5. EA 분해 알고리즘

### 5.1 EA 분해 규칙

EA 분해는 인보이스 타입(SAFEEN/ADP/일반)에 따라 다른 방식으로 처리됩니다.

#### 5.1.1 SAFEEN 인보이스 EA 분해

**SAFEEN 인보이스는 시간 기반이므로 단순화 모드 사용:**

```python
def decompose_safeen_line(standard_line: dict) -> dict:
    """
    SAFEEN 인보이스 라인 EA 분해
    
    SAFEEN 특징:
    - 시간 기반 요금 (Rate 정보 없음)
    - EA_1 = 1 (건당)
    - Rate_1 = Total_Amount_AED (총액)
    """
    amount_aed = standard_line["amount_excl_tax_aed"]
    hours = standard_line.get("unit1", 1.0)  # 시간 정보
    
    # SAFEEN은 단순화 모드
    ea_slots = {
        "ea_1": 1.0,
        "rate_1": amount_aed,
        "amount_1_aed": amount_aed,
        "name_1": "건",
        "ea_2": 0.0, "rate_2": 0.0, "amount_2_aed": 0.0, "name_2": "",
        "ea_3": 0.0, "rate_3": 0.0, "amount_3_aed": 0.0, "name_3": "",
        "ea_4": 0.0, "rate_4": 0.0, "amount_4_aed": 0.0, "name_4": ""
    }
    
    # 시간 정보는 메타데이터로 보존
    standard_line["hours"] = hours
    standard_line.update(ea_slots)
    standard_line["ea_total_aed"] = amount_aed
    
    return standard_line
```

**SAFEEN 인보이스 예시:**

```python
# 예시: SAFEEN Channel Crossing (3 Hours, 3,091.25 AED)
line = {
    "description": "Channel Crossing 09-Oct-2025 04:12:00 to 09-Oct-2025 07:12:00 3 Hours",
    "tariff_code": "6.6",
    "unit1": 3.0,  # 3 Hours
    "rate": None,  # Rate 정보 없음
    "amount_excl_tax_aed": 3091.25
}

result = decompose_safeen_line(line)
# 결과:
# {
#   "ea_1": 1.0,
#   "rate_1": 3091.25,
#   "amount_1_aed": 3091.25,
#   "name_1": "건",
#   "hours": 3.0,
#   "ea_total_aed": 3091.25
# }
```

#### 5.1.2 ADP 인보이스 EA 분해

**ADP 인보이스는 수량 기반이므로 원본 구조 보존:**

```python
def decompose_adp_line(standard_line: dict) -> dict:
    """
    ADP 인보이스 라인 EA 분해
    
    ADP 특징:
    - 수량(Unit) × 단가(Rate) = 금액 구조
    - 원본 Qty/Unit Price 구조 보존
    - ±2% 허용오차 검증
    """
    amount_aed = standard_line["amount_excl_tax_aed"]
    unit1 = standard_line.get("unit1") or 0.0
    unit2 = standard_line.get("unit2") or 0.0
    unit3 = standard_line.get("unit3") or 0.0
    rate = standard_line.get("rate")
    
    # EA 슬롯 초기화
    ea_slots = {
        "ea_1": 0.0, "rate_1": 0.0, "amount_1_aed": 0.0, "name_1": "",
        "ea_2": 0.0, "rate_2": 0.0, "amount_2_aed": 0.0, "name_2": "",
        "ea_3": 0.0, "rate_3": 0.0, "amount_3_aed": 0.0, "name_3": "",
        "ea_4": 0.0, "rate_4": 0.0, "amount_4_aed": 0.0, "name_4": ""
    }
    
    # Unit1이 있는 경우 (주요 수량)
    if unit1 > 0 and rate:
        ea_slots["ea_1"] = unit1
        ea_slots["rate_1"] = rate
        ea_slots["amount_1_aed"] = unit1 * rate
        ea_slots["name_1"] = "톤" if "Bulk Material" in standard_line.get("description", "") else "건"
        
        # 검증: ±2% 허용오차
        calculated_total = ea_slots["amount_1_aed"]
        diff_pct = abs(calculated_total - amount_aed) / amount_aed if amount_aed > 0 else 0
        
        if diff_pct > 0.02:
            # 허용오차 초과 시 단순화 모드로 전환
            ea_slots = {
                "ea_1": 1.0,
                "rate_1": amount_aed,
                "amount_1_aed": amount_aed,
                "name_1": "건",
                "ea_2": 0.0, "rate_2": 0.0, "amount_2_aed": 0.0, "name_2": "",
                "ea_3": 0.0, "rate_3": 0.0, "amount_3_aed": 0.0, "name_3": "",
                "ea_4": 0.0, "rate_4": 0.0, "amount_4_aed": 0.0, "name_4": ""
            }
            print(f"⚠️ WARNING: Line {standard_line['line_no']} 허용오차 초과 ({diff_pct:.2%}), 단순화 모드 적용")
    
    # Unit1이 없거나 Rate가 없는 경우 (단순화)
    else:
        ea_slots["ea_1"] = 1.0
        ea_slots["rate_1"] = amount_aed
        ea_slots["amount_1_aed"] = amount_aed
        ea_slots["name_1"] = "건"
    
    # EA 총합 계산
    ea_total = (
        ea_slots["amount_1_aed"] +
        ea_slots["amount_2_aed"] +
        ea_slots["amount_3_aed"] +
        ea_slots["amount_4_aed"]
    )
    
    standard_line.update(ea_slots)
    standard_line["ea_total_aed"] = ea_total
    
    return standard_line
```

**ADP 인보이스 예시:**

```python
# 예시 1: ADP Bulk Material (738톤 × 6.50 AED/톤)
line = {
    "description": "VAT - Bulk Material - Solids a) Parcel Size 0-10,000 Tons Direct Delivery",
    "tariff_id": "201.3",
    "unit1": 738.000,
    "unit2": 0.000,
    "unit3": 0.000,
    "rate": 6.50,
    "amount_excl_tax_aed": 4797.00
}

result = decompose_adp_line(line)
# 결과:
# {
#   "ea_1": 738.000,
#   "rate_1": 6.50,
#   "amount_1_aed": 4797.00,
#   "name_1": "톤",
#   "ea_total_aed": 4797.00
# }

# 예시 2: ADP Document Processing (단순화)
line = {
    "description": "VAT - Document Processing Charge (Bulk)",
    "tariff_id": "2.20",
    "unit1": 0.000,
    "unit2": 0.000,
    "unit3": 1.000,
    "rate": 35.00,
    "amount_excl_tax_aed": 35.00
}

result = decompose_adp_line(line)
# 결과:
# {
#   "ea_1": 1.0,
#   "rate_1": 35.00,
#   "amount_1_aed": 35.00,
#   "name_1": "건",
#   "ea_total_aed": 35.00
# }
```

#### 5.1.3 통합 EA 분해 함수

```python
def decompose_to_ea_slots(standard_line: dict) -> dict:
    """
    표준 라인을 최대 4개의 EA RatePair로 분해
    
    SAFEEN/ADP 인보이스 타입에 따라 분기 처리
    
    규칙:
    1. 원본 Qty/Unit Price 구조 보존 우선 (ADP)
    2. 단순화 필요 시: EA_1=1, Rate_1=Total_Amount_AED (SAFEEN)
    3. 최대 4쌍 (EA_1~4, Rate_1~4, Amount_1~4)
    4. 잔여 슬롯은 0으로 채움
    """
    invoice_type = standard_line.get("invoice_type")
    
    # SAFEEN 인보이스
    if invoice_type == "SAFEEN":
        return decompose_safeen_line(standard_line)
    
    # ADP 인보이스
    elif invoice_type == "ADP":
        return decompose_adp_line(standard_line)
    
    # 일반 인보이스 (기존 로직)
    else:
        amount_aed = standard_line["amount_excl_tax_aed"]
        unit1 = standard_line.get("unit1") or 0.0
        unit2 = standard_line.get("unit2") or 0.0
        unit3 = standard_line.get("unit3") or 0.0
        rate = standard_line.get("rate")
        
        # EA 슬롯 초기화
        ea_slots = {
            "ea_1": 0.0, "rate_1": 0.0, "amount_1_aed": 0.0, "name_1": "",
            "ea_2": 0.0, "rate_2": 0.0, "amount_2_aed": 0.0, "name_2": "",
            "ea_3": 0.0, "rate_3": 0.0, "amount_3_aed": 0.0, "name_3": "",
            "ea_4": 0.0, "rate_4": 0.0, "amount_4_aed": 0.0, "name_4": ""
        }
        
        # 케이스 1: 원본 Qty/Unit Price 구조가 있는 경우
        if unit1 > 0 and rate:
            # 단일 EA로 분해
            ea_slots["ea_1"] = unit1
            ea_slots["rate_1"] = rate
            ea_slots["amount_1_aed"] = unit1 * rate
            ea_slots["name_1"] = standard_line.get("unit", "건")
            
            # 검증: ±2% 허용오차
            calculated_total = ea_slots["amount_1_aed"]
            diff_pct = abs(calculated_total - amount_aed) / amount_aed if amount_aed > 0 else 0
            
            if diff_pct > 0.02:
                # 허용오차 초과 시 단순화 모드로 전환
                ea_slots = {
                    "ea_1": 1.0,
                    "rate_1": amount_aed,
                    "amount_1_aed": amount_aed,
                    "name_1": "건",
                    "ea_2": 0.0, "rate_2": 0.0, "amount_2_aed": 0.0, "name_2": "",
                    "ea_3": 0.0, "rate_3": 0.0, "amount_3_aed": 0.0, "name_3": "",
                    "ea_4": 0.0, "rate_4": 0.0, "amount_4_aed": 0.0, "name_4": ""
                }
        
        # 케이스 2: Qty/Unit Price 구조가 없는 경우 (단순화)
        else:
            ea_slots["ea_1"] = 1.0
            ea_slots["rate_1"] = amount_aed
            ea_slots["amount_1_aed"] = amount_aed
            ea_slots["name_1"] = "건"
        
        # EA 총합 계산
        ea_total = (
            ea_slots["amount_1_aed"] +
            ea_slots["amount_2_aed"] +
            ea_slots["amount_3_aed"] +
            ea_slots["amount_4_aed"]
        )
        
        # 표준 라인에 EA 정보 추가
        standard_line.update(ea_slots)
        standard_line["ea_total_aed"] = ea_total
        
        return standard_line
```

### 5.2 EA 분해 예시

```python
# 예시 1: Bulk Material (542톤 × 6.50 AED/톤)
line = {
    "description": "VAT - Bulk Material - Solids a) Parcel Size 0-10,000 Tons",
    "unit1": 542.000,
    "rate": 6.50,
    "amount_excl_tax_aed": 3523.00
}

result = decompose_to_ea_slots(line)
# 결과:
# {
#   "ea_1": 542.000,
#   "rate_1": 6.50,
#   "amount_1_aed": 3523.00,
#   "name_1": "톤",
#   "ea_total_aed": 3523.00
# }

# 예시 2: Document Processing (단순화)
line = {
    "description": "Document Processing Charge",
    "unit1": 0.0,
    "rate": None,
    "amount_excl_tax_aed": 35.00
}

result = decompose_to_ea_slots(line)
# 결과:
# {
#   "ea_1": 1.0,
#   "rate_1": 35.00,
#   "amount_1_aed": 35.00,
#   "name_1": "건",
#   "ea_total_aed": 35.00
# }
```

---

## 6. Price Center 피벗 생성

### 6.1 Long → Wide Format 변환

```python
def create_price_center_pivot(
    standard_lines: list,
    cost_items: dict,
    cost_item_fields: dict  # cost_item_fields.JSON 추가
) -> pd.DataFrame:
    """
    표준 라인 리스트를 Price Center 피벗(wide format)으로 변환
    
    결과: 44개 Price Center × 2 (QTY/AMOUNT) = 88컬럼
    
    중요: cost_item_fields.JSON의 필드명은 절대 변경하지 않음
    """
    import pandas as pd
    
    # Long format 데이터프레임 생성
    df_long = pd.DataFrame(standard_lines)
    
    # Price Center별 QTY 집계
    pivot_qty = df_long.pivot_table(
        index="invoice_no",
        columns="price_center",
        values="unit1",  # 또는 적절한 QTY 필드
        aggfunc="sum",
        fill_value=0.0
    )
    
    # Price Center별 AMOUNT 집계
    pivot_amount = df_long.pivot_table(
        index="invoice_no",
        columns="price_center",
        values="amount_excl_tax_aed",
        aggfunc="sum",
        fill_value=0.0
    )
    
    # 컬럼명 정규화 (공백→_, 대문자)
    def normalize_column_name(name):
        return name.replace(" ", "_").upper()
    
    pivot_qty.columns = [f"{normalize_column_name(col)}_QTY" for col in pivot_qty.columns]
    pivot_amount.columns = [f"{normalize_column_name(col)}_AMOUNT" for col in pivot_amount.columns]
    
    # QTY와 AMOUNT 병합
    pivot_wide = pd.concat([pivot_qty, pivot_amount], axis=1)
    
    # cost_item_fields.JSON 기반 필드명으로 재매핑
    pivot_wide = remap_columns_with_cost_items(pivot_wide, cost_items, cost_item_fields)
    
    return pivot_wide
```

### 6.2 cost_item_fields.JSON 기반 컬럼명 재매핑

**중요**: `cost_item_fields.JSON`의 필드명은 절대 변경하지 않습니다. 모든 컬럼명 매핑은 이 파일의 기존 필드명을 사용합니다.

```python
def remap_columns_with_cost_items(
    pivot_df: pd.DataFrame,
    cost_items: dict,
    cost_item_fields: dict  # cost_item_fields.JSON 추가
) -> pd.DataFrame:
    """
    Price Center 컬럼명을 cost_item_fields.JSON의 필드명으로 재매핑
    
    중요: cost_item_fields.JSON의 필드명은 절대 변경하지 않음
    """
    field_mapping = {}
    
    # cost_item_fields.JSON의 필드명을 기준으로 매핑 생성
    # (절대 변경하지 않고 그대로 사용)
    cost_item_fields_list = cost_item_fields.get("cost_item_fields", [])
    
    # cost_item.JSON과 cost_item_fields.JSON 매칭
    for item in cost_items.get("cost_items", []):
        code = item["code"]
        
        # cost_item_fields.JSON에서 해당 code의 필드명 찾기
        qty_field = None
        amount_field = None
        
        for field in cost_item_fields_list:
            if field.endswith("_QTY") and code in field:
                qty_field = field  # 기존 필드명 그대로 사용
            elif field.endswith("_AMOUNT") and code in field:
                amount_field = field  # 기존 필드명 그대로 사용
        
        if qty_field and amount_field:
            # Price Center 이름 정규화
            price_center_normalized = code.replace("_", " ").upper()
            
            # pivot_df의 컬럼명과 매칭
            for col in pivot_df.columns:
                if col.endswith("_QTY"):
                    pc_name = col.rsplit("_QTY", 1)[0]
                    if pc_name.replace("_", " ").upper() == price_center_normalized:
                        field_mapping[col] = qty_field  # cost_item_fields.JSON의 기존 필드명
                
                elif col.endswith("_AMOUNT"):
                    pc_name = col.rsplit("_AMOUNT", 1)[0]
                    if pc_name.replace("_", " ").upper() == price_center_normalized:
                        field_mapping[col] = amount_field  # cost_item_fields.JSON의 기존 필드명
    
    # 컬럼명 재매핑 (cost_item_fields.JSON의 필드명으로)
    pivot_df = pivot_df.rename(columns=field_mapping)
    
    return pivot_df
```

---

## 7. 검증 로직

### 7.1 calc_check (EA 합계 vs Total)

```python
def validate_calc_check(standard_line: dict) -> dict:
    """
    EA 합계와 Total(AED) 일치 검증
    
    기준: |EA_Total - Total_AED| / Total_AED ≤ 2%
    """
    ea_total = standard_line.get("ea_total_aed", 0.0)
    total_aed = standard_line["amount_excl_tax_aed"]
    
    if total_aed == 0:
        return {
            "calc_check": "PASS",
            "calc_diff": 0.0,
            "calc_diff_pct": 0.0
        }
    
    diff = abs(ea_total - total_aed)
    diff_pct = diff / total_aed
    
    status = "PASS" if diff_pct <= 0.02 else "WARN"
    
    return {
        "calc_check": status,
        "calc_diff": diff,
        "calc_diff_pct": diff_pct,
        "ea_total": ea_total,
        "total_aed": total_aed
    }
```

### 7.2 vat_check (VAT vs Amount×5%)

```python
def validate_vat_check(standard_line: dict) -> dict:
    """
    VAT 검증
    
    기준: |VAT_USD - Amount_USD×0.05| ≤ 0.01 USD
    """
    amount_usd = standard_line.get("amount_usd", 0.0)
    vat_usd = standard_line.get("vat_amount_usd", 0.0)
    tax_rate_pct = standard_line.get("tax_rate_pct", 0.0)
    
    if tax_rate_pct == 0:
        # VAT 0%인 경우
        expected_vat = 0.0
    else:
        expected_vat = amount_usd * (tax_rate_pct / 100.0)
    
    diff = abs(vat_usd - expected_vat)
    
    # 허용오차: ±0.01 USD
    status = "PASS" if diff <= 0.01 else "WARN"
    
    return {
        "vat_check": status,
        "vat_diff": diff,
        "vat_usd": vat_usd,
        "expected_vat_usd": expected_vat,
        "tax_rate_pct": tax_rate_pct
    }
```

### 7.3 pc_check (Price Center 합계 vs Total)

```python
def validate_pc_check(
    invoice_lines: list,
    invoice_total_aed: float
) -> dict:
    """
    Price Center AMOUNT 합계와 Total(AED) 일치 검증
    
    기준: |Σ(Price Center AMOUNT) - Total_AED| ≤ 1.00 AED
    """
    # Price Center별 AMOUNT 합계
    pc_amount_sum = sum(
        line["amount_excl_tax_aed"]
        for line in invoice_lines
    )
    
    diff = abs(pc_amount_sum - invoice_total_aed)
    
    status = "PASS" if diff <= 1.00 else "WARN"
    
    return {
        "pc_check": status,
        "pc_amount_sum": pc_amount_sum,
        "invoice_total_aed": invoice_total_aed,
        "pc_diff": diff
    }
```

### 7.4 통합 검증 함수

```python
def run_all_validations(
    standard_lines: list,
    invoice_meta: dict
) -> dict:
    """
    모든 검증 실행
    """
    validation_results = {
        "line_validations": [],
        "invoice_validation": {}
    }
    
    # 라인별 검증
    for line in standard_lines:
        line_validation = {
            "line_no": line["line_no"],
            "calc_check": validate_calc_check(line),
            "vat_check": validate_vat_check(line)
        }
        validation_results["line_validations"].append(line_validation)
    
    # 인보이스 전체 검증
    validation_results["invoice_validation"] = validate_pc_check(
        standard_lines,
        invoice_meta["grand_total_aed"]
    )
    
    return validation_results
```

---

## 8. 엑셀 삽입 규칙

### 8.1 엑셀 시트 구조

```
Sheet1 구조:
- 행 1: 헤더 (고정)
- 행 2~N: 기존 데이터
- 행 N+1~N+M: 신규 인보이스 라인 (append)
```

### 8.2 삽입할 컬럼 매핑

```python
EXCEL_COLUMN_MAPPING = {
    # 메타데이터
    "Invoice No": "invoice_no",
    "Line No": "line_no",
    "Vessel": "vessel_name",
    "Rotation No": "rotation_no",
    "BOL": "bol",
    "Port": "port",
    "Arrival Date": "arrival_date",
    "Departure Date": "departure_date",
    
    # 라인 정보
    "Tariff ID": "tariff_id",
    "Description": "description",
    "Unit 1": "unit1",
    "Unit 2": "unit2",
    "Unit 3": "unit3",
    "Rate": "rate",
    
    # 금액 (AED)
    "Amount Excl TAX (AED)": "amount_excl_tax_aed",
    "TAX Rate (%)": "tax_rate_pct",
    "TAX Amount (AED)": "tax_amount_aed",
    "Total Amount Incl TAX (AED)": "total_incl_tax_aed",
    
    # 금액 (USD) - 환율 적용
    "Amount Excl TAX (USD)": "amount_excl_tax_usd",
    "TAX Amount (USD)": "vat_amount_usd",
    "Total Amount Incl TAX (USD)": "total_incl_tax_usd",
    
    # Cost/Price Center
    "Cost Center A": "cost_center_a",
    "Cost Center B": "cost_center_b",
    "Price Center": "price_center",
    
    # EA Slot
    "EA_1": "ea_1",
    "Rate_1": "rate_1",
    "Amount_1 (AED)": "amount_1_aed",
    "Name_1": "name_1",
    "EA_2": "ea_2",
    "Rate_2": "rate_2",
    "Amount_2 (AED)": "amount_2_aed",
    "Name_2": "name_2",
    "EA_3": "ea_3",
    "Rate_3": "rate_3",
    "Amount_3 (AED)": "amount_3_aed",
    "Name_3": "name_3",
    "EA_4": "ea_4",
    "Rate_4": "rate_4",
    "Amount_4 (AED)": "amount_4_aed",
    "Name_4": "name_4",
    "EA Total (AED)": "ea_total_aed",
    
    # Price Center 필드 (cost_item_fields.JSON 기반)
    # 중요: cost_item_fields.JSON의 필드명을 그대로 사용 (절대 변경 금지)
    # 동적 생성: 각 Price Center별로 _QTY, _AMOUNT 컬럼은 cost_item_fields.JSON에서 가져옴
    
    # 검증
    "calc_check": "calc_check",
    "calc_diff": "calc_diff",
    "vat_check": "vat_check",
    "vat_diff": "vat_diff",
    "pc_check": "pc_check",
    
    # Evidence
    "Evidence": "evidence"
}
```

### 8.3 통화 변환

```python
def convert_aed_to_usd(
    standard_lines: list,
    exchange_rate: float = 3.6725
) -> list:
    """
    AED 금액을 USD로 변환
    """
    for line in standard_lines:
        amount_aed = line["amount_excl_tax_aed"]
        vat_aed = line["tax_amount_aed"]
        total_aed = line["total_incl_tax_aed"]
        
        line["amount_excl_tax_usd"] = amount_aed / exchange_rate
        line["vat_amount_usd"] = vat_aed / exchange_rate
        line["total_incl_tax_usd"] = total_aed / exchange_rate
        line["exchange_rate"] = exchange_rate
    
    return standard_lines
```

### 8.4 엑셀 삽입 함수

**중요**: 엑셀 삽입 시 `cost_item_fields.JSON`의 필드명을 그대로 사용합니다.

```python
def insert_lines_to_excel(
    excel_path: str,
    standard_lines: list,
    cost_item_fields: dict,  # cost_item_fields.JSON 추가
    sheet_name: str = "Sheet1"
) -> None:
    """
    표준 라인을 엑셀에 삽입 (append)
    
    중요: cost_item_fields.JSON의 필드명은 절대 변경하지 않음
    """
    import pandas as pd
    from openpyxl import load_workbook
    
    # 기존 엑셀 읽기
    df_existing = pd.read_excel(excel_path, sheet_name=sheet_name)
    
    # 표준 라인을 데이터프레임으로 변환
    df_new = pd.DataFrame(standard_lines)
    
    # 컬럼명 매핑 적용
    df_new = df_new.rename(columns={
        v: k for k, v in EXCEL_COLUMN_MAPPING.items()
        if v in df_new.columns
    })
    
    # cost_item_fields.JSON의 필드명을 엑셀 컬럼명으로 사용
    # (Price Center 필드명은 cost_item_fields.JSON의 필드명 그대로 사용)
    for field in cost_item_fields.get("cost_item_fields", []):
        if field in df_new.columns:
            # 필드명을 그대로 유지 (변경하지 않음)
            pass
    
    # 기존 데이터와 병합
    df_combined = pd.concat([df_existing, df_new], ignore_index=True)
    
    # 엑셀에 저장
    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_combined.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print(f"✅ {len(standard_lines)}개 라인을 엑셀에 삽입 완료")
    print(f"   총 행 수: {len(df_combined)}")
```

### 8.5 SAFEEN/ADP 인보이스 엑셀 삽입 특화 규칙

#### 8.5.1 SAFEEN 인보이스 엑셀 입력 방식

**SAFEEN 인보이스 엑셀 입력 필드:**

```python
SAFEEN_EXCEL_MAPPING = {
    # 기본 필드
    "Invoice No": "invoice_no",
    "Line No": "line_no",
    "Vessel": "vessel_name",
    "Rotation No": "rotation_no",
    "Port": "port",  # "Musaffah Channel"
    "Arrival Date": "arrival_date",
    "Departure Date": "departure_date",
    
    # SAFEEN 특화 필드
    "Tariff Code": "tariff_code",  # 6.1 또는 6.6
    "Description": "description",
    "Hours": "hours",  # 시간 정보 (예: 3.0)
    "Time Range Start": "time_range.start",  # 시간 범위 시작
    "Time Range End": "time_range.end",  # 시간 범위 종료
    
    # 금액 (AED)
    "Amount Excl TAX (AED)": "amount_excl_tax_aed",
    "TAX Rate (%)": "tax_rate_pct",  # SAFEEN은 보통 0%
    "TAX Amount (AED)": "tax_amount_aed",
    "Total Amount Incl TAX (AED)": "total_incl_tax_aed",
    
    # Cost/Price Center
    "Cost Center A": "cost_center_a",
    "Cost Center B": "cost_center_b",
    "Price Center": "price_center",
    
    # EA Slot (SAFEEN은 단순화 모드)
    "EA_1": "ea_1",  # 항상 1.0
    "Rate_1": "rate_1",  # Total_Amount_AED
    "Amount_1 (AED)": "amount_1_aed",
    "Name_1": "name_1",  # "건"
    
    # cost_item_fields.JSON 필드명
    # CHANNEL_TRANSIT_CROSSING_REQUEST_QTY / _AMOUNT
    # CHANNEL_CROSSING_CHARGES_FOR_VESSELS_WITH_1000_TO_3_001_GT_QTY / _AMOUNT
}
```

**SAFEEN 인보이스 엑셀 입력 예시:**

```python
# SAFEEN Channel Crossing 라인
safeen_line = {
    "invoice_no": "OFCO-INV-XXXXX",
    "line_no": 2,
    "vessel_name": "JOPETWIL 71",
    "rotation_no": "2503129927",
    "port": "Musaffah Channel",
    "tariff_code": "6.6",
    "description": "Channel Crossing 09-Oct-2025 04:12:00 to 09-Oct-2025 07:12:00 3 Hours",
    "hours": 3.0,
    "time_range": {
        "start": "09-Oct-2025 04:12:00",
        "end": "09-Oct-2025 07:12:00"
    },
    "amount_excl_tax_aed": 3091.25,
    "tax_rate_pct": 0,
    "tax_amount_aed": 0.00,
    "total_incl_tax_aed": 3091.25,
    "cost_center_a": "PORT HANDLING CHARGE",
    "cost_center_b": "CHANNEL TRANSIT CHARGES",
    "price_center": "CHANNEL TRANSIT CHARGES",
    "cost_item_code": "CHANNEL_CROSSING_CHARGES_FOR_VESSELS_WITH_1000_TO_3_001_GT",
    "qty_field": "CHANNEL_CROSSING_CHARGES_FOR_VESSELS_WITH_1000_TO_3_001_GT_QTY",
    "amount_field": "CHANNEL_CROSSING_CHARGES_FOR_VESSELS_WITH_1000_TO_3_001_GT_AMOUNT",
    "ea_1": 1.0,
    "rate_1": 3091.25,
    "amount_1_aed": 3091.25,
    "name_1": "건",
    "ea_total_aed": 3091.25,
    "evidence": "p1,row2"
}

# 엑셀 입력 시
# CHANNEL_CROSSING_CHARGES_FOR_VESSELS_WITH_1000_TO_3_001_GT_QTY = 1.0
# CHANNEL_CROSSING_CHARGES_FOR_VESSELS_WITH_1000_TO_3_001_GT_AMOUNT = 3091.25
```

#### 8.5.2 ADP 인보이스 엑셀 입력 방식

**ADP 인보이스 엑셀 입력 필드:**

```python
ADP_EXCEL_MAPPING = {
    # 기본 필드
    "Invoice No": "invoice_no",
    "Line No": "line_no",
    "Vessel": "vessel_name",
    "Rotation No": "rotation_no",
    "BOL": "bol",
    "Port": "port",  # "Musaffah Port GC"
    "Arrival Date": "arrival_date",
    "Sailing Date": "sailing_date",
    
    # ADP 특화 필드
    "Tariff ID": "tariff_id",  # 2.20 또는 201.3
    "Description": "description",
    "Unit 1": "unit1",  # 수량 (톤 등)
    "Unit 2": "unit2",
    "Unit 3": "unit3",
    "Rate": "rate",  # 단가
    
    # 금액 (AED)
    "Amount Excl TAX (AED)": "amount_excl_tax_aed",
    "TAX Rate (%)": "tax_rate_pct",  # ADP는 보통 5%
    "TAX Amount (AED)": "tax_amount_aed",  # 실제는 0.00인 경우 많음
    "Total Amount Incl TAX (AED)": "total_incl_tax_aed",
    
    # Cost/Price Center
    "Cost Center A": "cost_center_a",
    "Cost Center B": "cost_center_b",
    "Price Center": "price_center",
    
    # EA Slot (ADP는 원본 구조 보존)
    "EA_1": "ea_1",  # Unit1 값
    "Rate_1": "rate_1",  # Rate 값
    "Amount_1 (AED)": "amount_1_aed",  # EA_1 × Rate_1
    "Name_1": "name_1",  # "톤" 또는 "건"
    
    # cost_item_fields.JSON 필드명
    # DOCUMENT_PROCESSING_CHARGE_QTY / _AMOUNT
    # BULK_MATERIAL_SOLIDS_A_PARCEL_SIZE_0_10_001_TONS_DIRECT_DELIVERY_QTY / _AMOUNT
}
```

**ADP 인보이스 엑셀 입력 예시:**

```python
# ADP Bulk Material 라인
adp_line = {
    "invoice_no": "OFCO-INV-XXXXX",
    "line_no": 2,
    "vessel_name": "JOPETWIL 71",
    "rotation_no": "2503129927",
    "bol": "HVDC-AGI-GRM-J71-72",
    "port": "Musaffah Port GC",
    "tariff_id": "201.3",
    "description": "VAT - Bulk Material - Solids a) Parcel Size 0-10,000 Tons Direct Delivery",
    "unit1": 738.000,
    "unit2": 0.000,
    "unit3": 0.000,
    "rate": 6.50,
    "amount_excl_tax_aed": 4797.00,
    "tax_rate_pct": 5.0,
    "tax_amount_aed": 0.00,
    "total_incl_tax_aed": 4797.00,
    "cost_center_a": "PORT HANDLING CHARGE",
    "cost_center_b": "BULK MATERIAL HANDLING",
    "price_center": "BULK MATERIAL (PHC)",
    "cost_item_code": "BULK_MATERIAL_SOLIDS_A_PARCEL_SIZE_0_10_001_TONS_DIRECT_DELIVERY",
    "qty_field": "BULK_MATERIAL_SOLIDS_A_PARCEL_SIZE_0_10_001_TONS_DIRECT_DELIVERY_QTY",
    "amount_field": "BULK_MATERIAL_SOLIDS_A_PARCEL_SIZE_0_10_001_TONS_DIRECT_DELIVERY_AMOUNT",
    "ea_1": 738.000,
    "rate_1": 6.50,
    "amount_1_aed": 4797.00,
    "name_1": "톤",
    "ea_total_aed": 4797.00,
    "evidence": "p1,row2"
}

# 엑셀 입력 시
# BULK_MATERIAL_SOLIDS_A_PARCEL_SIZE_0_10_001_TONS_DIRECT_DELIVERY_QTY = 738.000
# BULK_MATERIAL_SOLIDS_A_PARCEL_SIZE_0_10_001_TONS_DIRECT_DELIVERY_AMOUNT = 4797.00
```

#### 8.5.3 SAFEEN/ADP 인보이스 통합 엑셀 삽입 함수

```python
def insert_safeen_adp_lines_to_excel(
    excel_path: str,
    standard_lines: list,
    cost_item_fields: dict,
    sheet_name: str = "Sheet1"
) -> None:
    """
    SAFEEN/ADP 인보이스를 엑셀에 삽입
    
    중요: cost_item_fields.JSON의 필드명은 절대 변경하지 않음
    """
    import pandas as pd
    
    # 기존 엑셀 읽기
    df_existing = pd.read_excel(excel_path, sheet_name=sheet_name)
    
    # 표준 라인을 데이터프레임으로 변환
    df_new = pd.DataFrame(standard_lines)
    
    # 인보이스 타입별 매핑 적용
    excel_mapping = {}
    if standard_lines:
        invoice_type = standard_lines[0].get("invoice_type")
        
        if invoice_type == "SAFEEN":
            excel_mapping.update(SAFEEN_EXCEL_MAPPING)
        elif invoice_type == "ADP":
            excel_mapping.update(ADP_EXCEL_MAPPING)
        else:
            excel_mapping.update(EXCEL_COLUMN_MAPPING)
    
    # 컬럼명 매핑 적용
    df_new = df_new.rename(columns={
        v: k for k, v in excel_mapping.items()
        if v in df_new.columns
    })
    
    # cost_item_fields.JSON의 필드명을 엑셀 컬럼명으로 사용
    # Price Center 필드명은 cost_item_fields.JSON에서 가져온 필드명 그대로 사용
    for idx, line in enumerate(standard_lines):
        qty_field = line.get("qty_field")
        amount_field = line.get("amount_field")
        
        if qty_field and qty_field in cost_item_fields.get("cost_item_fields", []):
            # cost_item_fields.JSON의 필드명을 그대로 사용
            if qty_field not in df_new.columns:
                # QTY 값 설정 (ADP는 unit1, SAFEEN은 1.0)
                if line.get("invoice_type") == "ADP":
                    df_new[qty_field] = 0.0
                    df_new.loc[idx, qty_field] = line.get("unit1", 0.0)
                else:  # SAFEEN
                    df_new[qty_field] = 0.0
                    df_new.loc[idx, qty_field] = 1.0
            else:
                # 기존 컬럼이 있으면 값만 업데이트
                if line.get("invoice_type") == "ADP":
                    df_new.loc[idx, qty_field] = line.get("unit1", 0.0)
                else:  # SAFEEN
                    df_new.loc[idx, qty_field] = 1.0
        
        if amount_field and amount_field in cost_item_fields.get("cost_item_fields", []):
            # cost_item_fields.JSON의 필드명을 그대로 사용
            if amount_field not in df_new.columns:
                df_new[amount_field] = 0.0
                df_new.loc[idx, amount_field] = line.get("amount_excl_tax_aed", 0.0)
            else:
                df_new.loc[idx, amount_field] = line.get("amount_excl_tax_aed", 0.0)
    
    # 기존 데이터와 병합
    df_combined = pd.concat([df_existing, df_new], ignore_index=True)
    
    # 엑셀에 저장
    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_combined.to_excel(writer, sheet_name=sheet_name, index=False)
    
    invoice_type = standard_lines[0].get("invoice_type", "일반") if standard_lines else "일반"
    print(f"✅ {len(standard_lines)}개 라인을 엑셀에 삽입 완료 (타입: {invoice_type})")
    print(f"   총 행 수: {len(df_combined)}")
```

---

## 9. 구현 코드 예시

### 9.1 전체 파이프라인 통합 함수

```python
def process_invoice_to_excel(
    ocr_output: dict,
    cost_items: dict,
    cost_item_fields: dict,  # cost_item_fields.JSON 추가
    excel_path: str,
    exchange_rate: float = 3.6725
) -> dict:
    """
    OCR 출력을 엑셀 삽입까지 완료하는 전체 파이프라인
    
    중요: cost_item_fields.JSON의 필드명은 절대 변경하지 않음
    """
    # Step 1: OCR KPI 검증
    validate_ocr_kpi(ocr_output)
    
    # Step 2: 표준 라인 생성 (SAFEEN/ADP 인보이스 특화 처리)
    invoice_meta = ocr_output["invoice_meta"]
    
    # SAFEEN/ADP 인보이스 통합 처리
    standard_lines = process_safeen_adp_invoice(
        ocr_output,
        cost_items,
        cost_item_fields
    )
    
    # Step 3: Cost/Price Center 매핑 (이미 process_safeen_adp_invoice에서 처리됨)
    # 추가 매핑이 필요한 경우에만 수행
    
    # Step 4: EA 분해 (SAFEEN/ADP 타입별 분기 처리)
    for line in standard_lines:
        line = decompose_to_ea_slots(line)
    
    # Step 5: 통화 변환
    standard_lines = convert_aed_to_usd(standard_lines, exchange_rate)
    
    # Step 6: 검증
    validation_results = run_all_validations(standard_lines, invoice_meta)
    
    # Step 7: 엑셀 삽입 (SAFEEN/ADP 인보이스 특화 처리)
    # 인보이스 타입에 따라 적절한 삽입 함수 사용
    if standard_lines:
        invoice_type = standard_lines[0].get("invoice_type")
        if invoice_type in ["SAFEEN", "ADP"]:
            insert_safeen_adp_lines_to_excel(
                excel_path,
                standard_lines,
                cost_item_fields
            )
        else:
            insert_lines_to_excel(
                excel_path,
                standard_lines,
                cost_item_fields
            )
    
    return {
        "status": "SUCCESS",
        "invoice_no": invoice_meta["invoice_no"],
        "line_count": len(standard_lines),
        "validation_results": validation_results
    }
```

### 9.2 사용 예시

```python
# cost_item.JSON 로드
import json
with open("cost_item.JSON", "r", encoding="utf-8") as f:
    cost_items = json.load(f)

# cost_item_fields.JSON 로드 (필드명 절대 변경 금지)
with open("cost_item_fields.JSON", "r", encoding="utf-8") as f:
    cost_item_fields = json.load(f)

# OCR 출력 (예시)
ocr_output = {
    "invoice_meta": {
        "invoice_no": "OFCO-INV-0002054",
        "vessel_name": "JOPETWIL 71",
        "rotation_no": "2503129579",
        "grand_total_aed": 3291.25,
        "vat_total_aed": 0.00
    },
    "lines": [
        {
            "line_no": 1,
            "tariff_code": "6.1",
            "description": "Administration Fees Channel Transit Request",
            "amount_excl_tax": 100.00,
            "vat_pct": 0,
            "vat_amount": 0.00,
            "total_incl_tax": 100.00,
            "evidence": "p1,row1"
        }
        # ... 더 많은 라인
    ],
    "ocr_kpi": {
        "mean_confidence": 0.95,
        "table_accuracy": 0.99,
        "numeric_integrity": 1.00
    }
}

# 파이프라인 실행
result = process_invoice_to_excel(
    ocr_output=ocr_output,
    cost_items=cost_items,
    cost_item_fields=cost_item_fields,  # 추가
    excel_path="OFCO INVOICE 2025.xlsx",
    exchange_rate=3.6725
)

print(result)
```

---

## 10. 주의사항 및 Best Practices

### 10.1 Evidence 필수
- 모든 라인에 `evidence` 필드 포함 (예: "p1,row1")
- PDF 원문 위치 추적 가능하도록

### 10.2 EA 분해 전략
- 원본 Qty/Unit Price 구조 보존 우선
- ±2% 허용오차 초과 시 단순화 모드 전환
- 대량수량 라인 환산 오차 방지

### 10.3 검증 플래그 관리
- calc_check/vat_check/pc_check 결과를 엑셀에 컬럼으로 추가
- WARN 발생 시 수동 검토 필요

### 10.4 Price Center 필드명 일관성
- **`cost_item_fields.JSON`의 필드명은 절대 변경하지 않음** (가장 중요)
- 모든 필드명은 `cost_item_fields.JSON`에서 가져와야 함
- `cost_item.JSON`은 메타데이터(description, unit 등) 제공용
- 동적 생성 시에도 `cost_item_fields.JSON`의 기존 필드명을 그대로 사용
- Price Center 이름 정규화는 매핑을 위한 중간 단계일 뿐, 최종 필드명은 `cost_item_fields.JSON`에서 가져옴

---

## 11. JSON 파일 구조 및 관계

### 11.1 파일 역할 분리

**cost_item.JSON (메타데이터 + 필드명):**
- 역할: 비용 항목의 완전한 메타데이터 제공
- 포함: code, description, unit, condition, qty_field, amount_field
- 용도: Subject 매핑, 설명 표시, 단위 정보
- 항목 수: 42개

**cost_item_fields.JSON (필드명 목록 - 불변):**
- 역할: 엑셀 컬럼명으로 사용되는 필드명 목록
- 포함: 84개 필드명 (42개 항목 × 2)
- 용도: 엑셀 삽입 시 컬럼명 검증, 필드명 참조
- ⚠️ 중요: 절대 변경 금지

**cost_item_fields_GROUPL.JSON (그룹화 매핑):**
- 역할: code → qty_field/amount_field 빠른 매핑
- 포함: code, qty_field, amount_field만
- 용도: 빠른 조회, 매핑 테이블 생성
- 항목 수: 42개

### 11.2 필드명 매핑 우선순위

1. **최우선**: `cost_item_fields.JSON`의 필드명 (절대 변경 금지)
2. **2순위**: `cost_item.JSON`의 qty_field/amount_field (검증용)
3. **3순위**: `cost_item_fields_GROUPL.JSON` (빠른 조회용)

### 11.3 필드명 검증 프로세스

```python
def get_cost_item_field_names(
    cost_item_code: str,
    cost_items: dict,
    cost_item_fields: dict
) -> dict:
    """
    cost_item_code로 필드명 가져오기
    
    우선순위:
    1. cost_item_fields.JSON에서 직접 검색
    2. cost_item.JSON의 qty_field/amount_field 사용
    
    중요: cost_item_fields.JSON의 필드명은 절대 변경하지 않음
    """
    # cost_item_fields.JSON에서 검색 (최우선)
    qty_field = None
    amount_field = None
    
    for field in cost_item_fields.get("cost_item_fields", []):
        if field.endswith("_QTY") and cost_item_code in field:
            qty_field = field  # cost_item_fields.JSON의 필드명 그대로
        elif field.endswith("_AMOUNT") and cost_item_code in field:
            amount_field = field  # cost_item_fields.JSON의 필드명 그대로
    
    # cost_item.JSON에서 fallback
    if not qty_field or not amount_field:
        for item in cost_items.get("cost_items", []):
            if item["code"] == cost_item_code:
                qty_field = item.get("qty_field")
                amount_field = item.get("amount_field")
                break
    
    return {
        "qty_field": qty_field,
        "amount_field": amount_field
    }
```

### 11.4 JSON 파일 일관성 검증

```python
def validate_json_consistency(
    cost_items: dict,
    cost_item_fields: dict,
    cost_item_fields_group: dict
) -> dict:
    """
    세 JSON 파일 간 일관성 검증
    
    중요: cost_item_fields.JSON의 필드명은 절대 변경하지 않음
    """
    validation_results = {
        "status": "PASS",
        "errors": [],
        "warnings": []
    }
    
    # 1. cost_item.JSON과 cost_item_fields.JSON 일치 검증
    cost_item_fields_list = set(cost_item_fields.get("cost_item_fields", []))
    
    for item in cost_items.get("cost_items", []):
        code = item["code"]
        qty_field = item["qty_field"]
        amount_field = item["amount_field"]
        
        # cost_item_fields.JSON에 필드명이 있는지 확인
        if qty_field not in cost_item_fields_list:
            validation_results["errors"].append({
                "type": "MISSING_QTY_FIELD",
                "code": code,
                "field": qty_field,
                "message": f"{qty_field}가 cost_item_fields.JSON에 없습니다"
            })
            validation_results["status"] = "FAIL"
        
        if amount_field not in cost_item_fields_list:
            validation_results["errors"].append({
                "type": "MISSING_AMOUNT_FIELD",
                "code": code,
                "field": amount_field,
                "message": f"{amount_field}가 cost_item_fields.JSON에 없습니다"
            })
            validation_results["status"] = "FAIL"
    
    # 2. cost_item_fields_GROUPL.JSON과 일치 검증
    for item in cost_item_fields_group.get("cost_items", []):
        code = item["code"]
        qty_field = item["qty_field"]
        amount_field = item["amount_field"]
        
        # cost_item.JSON에서 찾기
        cost_item = None
        for ci in cost_items.get("cost_items", []):
            if ci["code"] == code:
                cost_item = ci
                break
        
        if cost_item:
            if cost_item["qty_field"] != qty_field:
                validation_results["warnings"].append({
                    "type": "QTY_FIELD_MISMATCH",
                    "code": code,
                    "cost_item": cost_item["qty_field"],
                    "group": qty_field
                })
            
            if cost_item["amount_field"] != amount_field:
                validation_results["warnings"].append({
                    "type": "AMOUNT_FIELD_MISMATCH",
                    "code": code,
                    "cost_item": cost_item["amount_field"],
                    "group": amount_field
                })
    
    return validation_results
```

### 11.5 통합 검증 스크립트 사용 예시

```python
# validate_json_consistency.py
import json

def main():
    # JSON 파일 로드
    with open("cost_item.JSON", "r", encoding="utf-8") as f:
        cost_items = json.load(f)
    
    with open("cost_item_fields.JSON", "r", encoding="utf-8") as f:
        cost_item_fields = json.load(f)
    
    with open("cost_item_fields_GROUPL.JSON", "r", encoding="utf-8") as f:
        cost_item_fields_group = json.load(f)
    
    # 일관성 검증
    results = validate_json_consistency(
        cost_items,
        cost_item_fields,
        cost_item_fields_group
    )
    
    # 결과 출력
    if results["status"] == "PASS":
        print("✅ 모든 JSON 파일이 일관성 있게 구성되어 있습니다.")
    else:
        print("❌ 일관성 오류 발견:")
        for error in results["errors"]:
            print(f"  - {error['message']}")
    
    if results["warnings"]:
        print("⚠️ 경고:")
        for warning in results["warnings"]:
            print(f"  - {warning}")

if __name__ == "__main__":
    main()
```

---

## 12. Risk Assessment 및 완화 전략

### 12.1 주요 리스크 및 완화책

| 리스크 | 발생 지점 | 영향 | 완화책 |
|--------|----------|------|--------|
| OCR 표 누락/오인식 | OCR/테이블 파서 | 라인 누락·수치 불일치 | KPI 게이트(MeanConf≥0.92, TableAcc≥0.98, NumericIntegrity=1.00) 미달 시 **ZERO 중단 → /ocr_retry** |
| 키 매칭 실패 | PDF↔엑셀 매칭 | 다른 인보이스가 붙거나 누락 | Multi-Key 흡수+휴리스틱(예: ±7d) 등 운영 가드 |
| Subject 패턴 분류 오류 | Cost Center 매핑 | Cost/PriceCenter 오배정 | Cost Center v2.5 + Subject 패턴 우선순위 테이블 적용 |
| EA 분해 모호 | EA Slot | calc_check FAIL 증가 | "최우선 PDF, 보조 근거(참조시트)" + 잔여 슬롯 0 + ±2% 상한 |
| VAT 표기 상충 | VAT 검증 | vat_check 경고/오류 | VAT 0/5% 외는 MISMATCH, "세율>0 & 세액=0"은 WARN 처리 권고 |

### 12.2 리스크 완화 구현

```python
def apply_risk_mitigation(
    ocr_output: dict,
    standard_lines: list
) -> dict:
    """
    리스크 완화 전략 적용
    """
    risk_report = {
        "ocr_risks": [],
        "mapping_risks": [],
        "ea_risks": [],
        "vat_risks": []
    }
    
    # OCR 리스크 검증
    kpi = ocr_output.get("ocr_kpi", {})
    if kpi.get("mean_confidence", 0) < 0.92:
        risk_report["ocr_risks"].append({
            "type": "LOW_CONFIDENCE",
            "severity": "HIGH",
            "action": "ZERO_STOP",
            "message": "OCR MeanConf < 0.92: 재처리 필요"
        })
    
    # Subject 매핑 리스크
    for line in standard_lines:
        if line.get("cost_item_code") == "OTHERS":
            risk_report["mapping_risks"].append({
                "line_no": line["line_no"],
                "type": "UNMAPPED_SUBJECT",
                "severity": "MEDIUM",
                "action": "MANUAL_REVIEW",
                "subject": line.get("description", "")
            })
    
    # EA 분해 리스크
    for line in standard_lines:
        calc_check = line.get("calc_check", "PASS")
        if calc_check == "WARN":
            risk_report["ea_risks"].append({
                "line_no": line["line_no"],
                "type": "EA_DECOMPOSITION_WARN",
                "severity": "LOW",
                "action": "REVIEW",
                "calc_diff_pct": line.get("calc_diff_pct", 0)
            })
    
    # VAT 리스크
    for line in standard_lines:
        vat_check = line.get("vat_check", "PASS")
        if vat_check == "WARN":
            risk_report["vat_risks"].append({
                "line_no": line["line_no"],
                "type": "VAT_MISMATCH",
                "severity": "MEDIUM",
                "action": "REVIEW",
                "vat_diff": line.get("vat_diff", 0)
            })
    
    return risk_report
```

---

## 13. Evidence 및 EXT 행 정책

### 13.1 Evidence 필수 요구사항

모든 표준 라인에는 `evidence` 필드가 필수입니다. 이는 PDF 원문 위치를 추적하기 위한 것입니다.

```python
def add_evidence_to_line(
    standard_line: dict,
    pdf_page: int,
    table_row: int,
    bbox: dict = None
) -> dict:
    """
    표준 라인에 evidence 추가
    
    형식:
    - 기본: "p{페이지},row{표행}"
    - Bbox 좌표: "p{페이지},bbox(x1,y1,x2,y2)"
    
    예시:
    - "p1,row2"
    - "p1,bbox(100,200,300,250)"
    """
    if bbox:
        evidence = f"p{pdf_page},bbox({bbox['x1']},{bbox['y1']},{bbox['x2']},{bbox['y2']})"
    else:
        evidence = f"p{pdf_page},row{table_row}"
    
    standard_line["evidence"] = evidence
    return standard_line
```

### 13.2 Evidence 검증

```python
def validate_evidence(standard_lines: list) -> dict:
    """
    모든 라인의 evidence 필수 검증
    """
    missing_evidence = []
    
    for line in standard_lines:
        if not line.get("evidence"):
            missing_evidence.append({
                "line_no": line.get("line_no"),
                "invoice_no": line.get("invoice_no"),
                "status": "MISSING_EVIDENCE"
            })
    
    return {
        "status": "PASS" if len(missing_evidence) == 0 else "FAIL",
        "missing_count": len(missing_evidence),
        "missing_lines": missing_evidence
    }
```

### 13.3 EXT 행 정책

EXT(확장행)는 메타데이터를 기록하기 위한 특수 행입니다.

**규칙:**
- K:BA에 필요한 필드가 없으면 **컬럼 추가 금지**
- 본행 아래 **EXT 행 삽입**으로 메타 기록
- EXT 행은 **금액 집계에서 제외** (메타데이터)

```python
def create_ext_line(base_line: dict, metadata: dict) -> dict:
    """
    EXT(확장행) 생성
    
    규칙:
    - K:BA에 필요한 필드가 없으면 컬럼 추가 금지
    - 본행 아래 EXT 행 삽입으로 메타 기록
    - EXT 행은 금액 집계에서 제외
    """
    ext_line = base_line.copy()
    ext_line["line_type"] = "EXT"
    ext_line["line_no"] = f"{base_line['line_no']}_EXT"
    
    # 금액 집계 제외
    ext_line["amount_excl_tax_aed"] = 0.0
    ext_line["tax_amount_aed"] = 0.0
    ext_line["total_incl_tax_aed"] = 0.0
    
    # 메타데이터 추가
    ext_line.update(metadata)
    ext_line["evidence"] = base_line.get("evidence", "")
    
    return ext_line

def filter_ext_lines_for_aggregation(standard_lines: list) -> list:
    """
    금액 집계 시 EXT 행 제외
    """
    return [
        line for line in standard_lines
        if line.get("line_type") != "EXT"
    ]
```

### 13.4 EXT 행 사용 예시

```python
# 예시: 추가 메타데이터가 필요한 경우
base_line = {
    "invoice_no": "OFCO-INV-0002054",
    "line_no": 1,
    "description": "Channel Crossing",
    "amount_excl_tax_aed": 3091.25,
    "evidence": "p1,row2"
}

# EXT 행 생성 (추가 메타데이터)
ext_line = create_ext_line(
    base_line,
    {
        "tariff_reference": "SAFEEN Tariff 6.6",
        "approval_code": "APP-2025-001",
        "notes": "Special handling required"
    }
)

# 결과:
# {
#   "line_type": "EXT",
#   "line_no": "1_EXT",
#   "amount_excl_tax_aed": 0.0,  # 집계 제외
#   "tariff_reference": "SAFEEN Tariff 6.6",
#   ...
# }
```

---

## 14. PDF↔엑셀 매칭 로직

### 14.1 매칭 전략

PDF 라인과 엑셀 참조 시트를 매칭하는 과정은 다음 규칙을 따릅니다:

- **기준금액(BJ=Total Amount)**을 "잠정 고정"
- **PDF(파일명/회전/Subject/금액)** 기반으로 매칭
- **BJ ±2%** 이내를 목표
- 충돌 시 **PDF 우선** + `[MISMATCH]` 표기
- 실패 시 보조 레퍼런스를 참조하되 **근거 기록 필수**

### 14.2 매칭 알고리즘

```python
def match_pdf_to_excel(
    pdf_lines: list,
    excel_reference: pd.DataFrame,
    tolerance_pct: float = 0.02
) -> dict:
    """
    PDF 라인을 엑셀 참조 시트와 매칭
    
    매칭 기준:
    1. Subject 패턴 유사도
    2. 금액 일치 (±2%)
    3. Tariff Code/ID 일치
    """
    matches = []
    unmatched_pdf_lines = []
    
    for pdf_line in pdf_lines:
        pdf_amount = pdf_line.get("amount_excl_tax", 0.0)
        pdf_subject = pdf_line.get("description", "")
        pdf_tariff = pdf_line.get("tariff_code") or pdf_line.get("tariff_id")
        
        best_match = None
        best_score = 0.0
        
        # 엑셀 참조 시트에서 매칭 후보 찾기
        for idx, excel_row in excel_reference.iterrows():
            excel_amount = excel_row.get("Amount Excl TAX (AED)", 0.0)
            excel_subject = excel_row.get("Description", "")
            excel_tariff = excel_row.get("Tariff ID") or excel_row.get("Tariff Code")
            
            # 점수 계산
            score = 0.0
            
            # 1. 금액 일치도 (±2%)
            if excel_amount > 0:
                amount_diff_pct = abs(pdf_amount - excel_amount) / excel_amount
                if amount_diff_pct <= tolerance_pct:
                    score += 0.5
                else:
                    continue  # 금액 차이가 너무 크면 제외
            
            # 2. Subject 유사도
            if pdf_subject and excel_subject:
                subject_similarity = calculate_text_similarity(pdf_subject, excel_subject)
                score += subject_similarity * 0.3
            
            # 3. Tariff 일치
            if pdf_tariff and excel_tariff and pdf_tariff == excel_tariff:
                score += 0.2
            
            if score > best_score:
                best_score = score
                best_match = {
                    "excel_row": excel_row.to_dict(),
                    "excel_index": idx,
                    "score": score
                }
        
        # 매칭 결과 기록
        if best_match and best_score >= 0.5:  # 최소 임계값
            matches.append({
                "pdf_line": pdf_line,
                "excel_match": best_match,
                "status": "MATCHED",
                "match_score": best_score
            })
        else:
            unmatched_pdf_lines.append({
                "pdf_line": pdf_line,
                "status": "UNMATCHED",
                "reason": "No suitable match found"
            })
    
    match_rate = len(matches) / len(pdf_lines) if pdf_lines else 0.0
    
    return {
        "matches": matches,
        "unmatched": unmatched_pdf_lines,
        "match_rate": match_rate,
        "status": "PASS" if match_rate >= 0.95 else "WARN"  # 95% 이상 매칭 목표
    }

def calculate_text_similarity(text1: str, text2: str) -> float:
    """
    두 텍스트의 유사도 계산 (간단한 버전)
    """
    from difflib import SequenceMatcher
    return SequenceMatcher(None, text1.lower(), text2.lower()).ratio()
```

### 14.3 매칭 충돌 처리

```python
def handle_matching_conflict(
    pdf_line: dict,
    excel_matches: list,
    conflict_resolution: str = "PDF_PRIORITY"
) -> dict:
    """
    매칭 충돌 처리
    
    규칙:
    - PDF 우선: PDF 라인 정보를 기준으로 사용
    - [MISMATCH] 표기 추가
    - 근거 기록 필수
    """
    if conflict_resolution == "PDF_PRIORITY":
        result = pdf_line.copy()
        result["matching_status"] = "[MISMATCH]"
        result["matching_evidence"] = {
            "pdf_amount": pdf_line.get("amount_excl_tax"),
            "excel_candidates": [
                {
                    "amount": m.get("excel_row", {}).get("Amount Excl TAX (AED)"),
                    "score": m.get("score")
                }
                for m in excel_matches
            ],
            "resolution": "PDF_PRIORITY"
        }
        return result
```

---

## 15. COST-GUARD (Δ% 밴드 검증)

### 15.1 COST-GUARD 개요

COST-GUARD는 기준요율 대비 실제 요율의 차이(Δ%)를 계산하여 **PASS/WARN/HIGH/CRITICAL** 밴드로 분류하는 검증 시스템입니다.

**목적:**
- 단가 급변 조기 감지
- 계약 요율 준수 여부 확인
- 비정상 요율 자동 플래깅

### 15.2 기준요율 대비 Δ% 계산

```python
def calculate_cost_guard_band(
    invoice_line: dict,
    reference_rate: float,
    tolerance_pct: float = 0.03
) -> dict:
    """
    기준요율 대비 Δ% 밴드 계산
    
    밴드 분류:
    - PASS: |Δ%| ≤ 3%
    - WARN: 3% < |Δ%| ≤ 10%
    - HIGH: 10% < |Δ%| ≤ 20%
    - CRITICAL: |Δ%| > 20%
    
    기준요율 소스:
    - 계약 요율 테이블 (contract)
    - 시장 요율 (market)
    - 특수 요율 (special)
    """
    actual_rate = invoice_line.get("rate")
    if not actual_rate or not reference_rate or reference_rate == 0:
        return {
            "band": "UNKNOWN",
            "delta_pct": None,
            "actual_rate": actual_rate,
            "reference_rate": reference_rate,
            "status": "SKIP"
        }
    
    delta_pct = ((actual_rate - reference_rate) / reference_rate) * 100
    abs_delta = abs(delta_pct)
    
    # 밴드 분류
    if abs_delta <= 3:
        band = "PASS"
        severity = "LOW"
    elif abs_delta <= 10:
        band = "WARN"
        severity = "MEDIUM"
    elif abs_delta <= 20:
        band = "HIGH"
        severity = "HIGH"
    else:
        band = "CRITICAL"
        severity = "CRITICAL"
    
    return {
        "band": band,
        "delta_pct": delta_pct,
        "abs_delta_pct": abs_delta,
        "actual_rate": actual_rate,
        "reference_rate": reference_rate,
        "severity": severity,
        "status": "CHECKED"
    }
```

### 15.3 기준요율 조회

```python
def get_reference_rate(
    invoice_line: dict,
    tariff_mapping: dict,
    rate_tables: dict
) -> float:
    """
    기준요율 조회
    
    우선순위:
    1. 계약 요율 테이블 (contract)
    2. 시장 요율 (market)
    3. 특수 요율 (special)
    """
    tariff_id = invoice_line.get("tariff_id") or invoice_line.get("tariff_code")
    price_center = invoice_line.get("price_center")
    unit = invoice_line.get("unit", "건당")
    
    # 계약 요율 테이블에서 조회
    contract_rate = rate_tables.get("contract", {}).get(
        f"{tariff_id}_{price_center}_{unit}"
    )
    
    if contract_rate:
        return contract_rate
    
    # 시장 요율에서 조회
    market_rate = rate_tables.get("market", {}).get(
        f"{price_center}_{unit}"
    )
    
    if market_rate:
        return market_rate
    
    # 특수 요율에서 조회
    special_rate = rate_tables.get("special", {}).get(
        f"{tariff_id}_{price_center}"
    )
    
    return special_rate  # None일 수 있음
```

### 15.4 COST-GUARD 통합 검증

```python
def run_cost_guard_validation(
    standard_lines: list,
    rate_tables: dict,
    tariff_mapping: dict
) -> dict:
    """
    모든 라인에 대해 COST-GUARD 검증 실행
    """
    validation_results = {
        "line_validations": [],
        "summary": {
            "PASS": 0,
            "WARN": 0,
            "HIGH": 0,
            "CRITICAL": 0,
            "UNKNOWN": 0
        }
    }
    
    for line in standard_lines:
        # 기준요율 조회
        reference_rate = get_reference_rate(
            line,
            tariff_mapping,
            rate_tables
        )
        
        if not reference_rate:
            validation_results["line_validations"].append({
                "line_no": line["line_no"],
                "cost_guard": {
                    "band": "UNKNOWN",
                    "status": "NO_REFERENCE_RATE"
                }
            })
            validation_results["summary"]["UNKNOWN"] += 1
            continue
        
        # COST-GUARD 계산
        cost_guard = calculate_cost_guard_band(line, reference_rate)
        
        validation_results["line_validations"].append({
            "line_no": line["line_no"],
            "cost_guard": cost_guard
        })
        
        # 요약 통계 업데이트
        band = cost_guard.get("band", "UNKNOWN")
        if band in validation_results["summary"]:
            validation_results["summary"][band] += 1
    
    return validation_results
```

### 15.5 COST-GUARD 사용 예시

```python
# 기준요율 테이블 예시
rate_tables = {
    "contract": {
        "6.6_CHANNEL_TRANSIT_CHARGES_GT": 3091.25,
        "201.3_BULK_MATERIAL (PHC)_톤(MT)": 6.50
    },
    "market": {
        "CHANNEL_TRANSIT_CHARGES_GT": 3100.00,
        "BULK_MATERIAL (PHC)_톤(MT)": 6.55
    }
}

# 인보이스 라인
line = {
    "line_no": 2,
    "tariff_code": "6.6",
    "price_center": "CHANNEL_TRANSIT_CHARGES",
    "rate": 3200.00,  # 실제 요율
    "unit": "GT"
}

# COST-GUARD 검증
cost_guard = calculate_cost_guard_band(line, 3091.25)

# 결과:
# {
#   "band": "WARN",
#   "delta_pct": 3.52,
#   "abs_delta_pct": 3.52,
#   "actual_rate": 3200.00,
#   "reference_rate": 3091.25,
#   "severity": "MEDIUM",
#   "status": "CHECKED"
# }
```

---

이 문서는 OFCO 인보이스 엑셀 삽입 로직의 구현 가이드입니다. 추가 질문이나 세부 사항이 필요하면 알려주세요.