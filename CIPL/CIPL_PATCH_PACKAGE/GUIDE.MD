문서 간 관계 분석 중. 통합 가이드 작성 중.


[2 tools called]


## LDG CIPL Builder 통합 실행 가이드

### 전체 워크플로우

```
사용자 입력 (Text/PDF/이미지)
    │
    ▼
┌─────────────────────────────────────────────────────────┐
│  GPT: 데이터 추출 및 검증                               │
│  - LDG_PAYLOAD v2.4 생성                                │
│  - LDG_AUDIT v2.4 생성                                  │
│  - Auto Guard Summary                                   │
└─────────────────────────────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────────────────────────────┐
│  데이터 변환: LDG_PAYLOAD → Commons Mode JSON           │
│  - commons: {pol, pod, carrier, invoice_no, ...}        │
│  - static_parties: {shipper, consignee, notify, ...}   │
│  - ci_rider_items: [] (선택적, 자동 생성 가능)          │
│  - pl_rider_items: [] (선택적, 자동 생성 가능)          │
└─────────────────────────────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────────────────────────────┐
│  프리필 어댑터 적용 (필수키 누락 대응)                    │
│  - 필수키 누락 감지                                       │
│  - 빈칸/0.00으로 채움 (HallucinationBan 준수)            │
│  - warnings[]에 "가정:" 기록                             │
└─────────────────────────────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────────────────────────────┐
│  파이프라인 실행: make_cipl_set.py                      │
│  - CIPL.py.make_4page_data_dicts() 호출                 │
│  - 4개 시트 생성                                         │
└─────────────────────────────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────────────────────────────┐
│  출력: CIPL_{Invoice_No}_{YYYY-MM-DD}.xlsx             │
└─────────────────────────────────────────────────────────┘
```

### GPT 실행 단계별 가이드

#### Step 1: 입력 데이터 추출

```python
# LDG_PAYLOAD v2.4 구조로 추출
{
  "meta": {
    "source_file": "...",
    "doc_hash": "...",
    "ocr_version": "v2.4",
    "mean_conf": 0.95,  # ≥ 0.92 목표
    "table_acc": 98.5,  # ≥ 98.00% 목표
    "numeric_integrity": 1.00,  # 목표
    "pages": 1,
    "created_at": "2026-01-14T10:30:00+04:00"  # Asia/Dubai
  },
  "data": {
    "DocType": "Commercial Invoice",
    "Invoice_No": "HVDC-ADOPT-SCT-0159",
    "Invoice_Date": "26-Dec-25",
    "Shipper": "Samsung C&T Corporation",
    "Consignee": "Abu Dhabi Offshore Power Transmission Company Limited LLC",
    "POL": "BUSAN",
    "POD": "KHALIFA PORT, ABU DHABI",
    "Carrier": "HMM RAON 0022W",
    "Sailing_On": "26-Dec-25",
    "Project_No": "AD164",
    "PO_No": "5000802593",
    "HS_Code": "7303.00.1090",
    "Country_of_Origin": "KOREA",
    "Packages": 5,
    "Net_Weight_kg": 11330.0,
    "Gross_Weight_kg": 12330.0,
    "CBM": 19.239,
    "Freight_USD": 0.0,
    "Insurance_USD": 0.0,
    "Total_CIF_USD": 26483.04,
    "Manufacturer_Name": "SANGDONG INDUSTRIES CO., LTD.",
    "Manufacturer_Address_1": "54, NEUNGHEODAE-RO 595BEON-GIL, NAMDONG-GU, INCHEON",
    "Manufacturer_Address_2": "REPUBLIC OF KOREA",
    "Item_Description": "STEEL CONDUIT (14TH)"
  },
  "warnings": [
    "가정: Freight와 Insurance가 0으로 추출됨 (CIF 조건 확인 필요)"
  ]
}
```

#### Step 2: Commons Mode JSON 변환

```python
# LDG_PAYLOAD → Commons Mode 변환
{
  "commons": {
    "pol": "BUSAN",
    "pod": "KHALIFA PORT, ABU DHABI",
    "carrier": "HMM RAON 0022W",
    "sailing_on": "26-Dec-25",
    "invoice_no": "HVDC-ADOPT-SCT-0159",
    "invoice_date": "26-Dec-25",
    "hs_code": "7303.00.1090",
    "manufacturer_name": "SANGDONG INDUSTRIES CO., LTD.",
    "manufacturer_address_1": "54, NEUNGHEODAE-RO 595BEON-GIL, NAMDONG-GU, INCHEON",
    "manufacturer_address_2": "REPUBLIC OF KOREA",
    "item_description": "STEEL CONDUIT (14TH)",
    "package_no": "5 PACKAGES",
    "net_weight_kg": 11330.0,
    "gross_weight_kg": 12330.0,
    "dimension_cbm": 19.239,
    "item": "STEEL CONDUIT (14TH)",
    "quantity_pkgs": 5,
    "freight_usd": 0.0,
    "insurance_usd": 0.0,
    "total_cif_usd": 26483.04,
    "country_of_origin": "KOREA",
    # Rider 관련 (선택적)
    "ci_qty": 4536,
    "ci_unit": "M",
    "pl_qty": 4536,
    "pl_unit": "M"
  },
  "static_parties": {
    "shipper_name": "Samsung C&T Corporation",
    "shipper_addr1": "26, SANGIL-RO 6-GIL, GANGDONG-GU, SEOUL,",
    "shipper_addr2": "REPUBLIC OF KOREA (05288)",
    "consignee_name1": "Abu Dhabi Offshore Power Transmission Company",
    "consignee_name2": "Limited LLC",
    "consignee_addr1": "1301 & 1304 Al Wahda City, Commercial City Tower",
    "consignee_addr2": "Level-13, hazza Bin Zayed Street, Abu Dhabi, UAE",
    "consignee_addr3": "P.O. Box No: 108708",
    "notify_name": "DSV SOLUTIONS PJSC",
    "notify_addr1": "M19 Mussafah 2nd round about after Al Jabber Mussafah",
    "notify_addr2": "Abu Dhabi, UAE. P.O.Box 93971",
    "project_no": "AD164",
    "project_name": "Independent Subsea HVDC System Project (Project Lightning), UAE",
    "po_no": "5000802593",
    "country_of_origin": "KOREA",
    "pol_country": "KOREA",
    "pod_country": "U.A.E."
  },
  "ci_rider_items": [],  # 비어있으면 자동 생성
  "pl_rider_items": []   # 비어있으면 자동 생성
}
```

#### Step 2.5: 프리필 어댑터 적용 (필수키 누락 대응)

```python
# 필수키 누락 시 프리필 적용 (HallucinationBan 준수)
def apply_prefill_adapter(commons_payload: dict, warnings: list) -> dict:
    """
    필수키 누락 시 프리필 적용
    - 문자열 필드: "" (빈칸)
    - 숫자 필드: 0.00 (CBM은 0.000)
    - warnings에 "가정: 필수키 누락으로 빈칸/0.00 처리" 기록
    """
    REQUIRED_COMMONS = [
        "pol", "pod", "carrier", "sailing_on", "invoice_no", "invoice_date",
        "hs_code", "manufacturer_name", "manufacturer_address_1", 
        "manufacturer_address_2", "item_description", "package_no",
        "net_weight_kg", "gross_weight_kg", "dimension_cbm", "item",
        "quantity_pkgs", "freight_usd", "insurance_usd", "total_cif_usd"
    ]
    
    REQUIRED_PARTIES = [
        "shipper_name", "shipper_addr1", "shipper_addr2",
        "consignee_name1", "consignee_name2", "consignee_addr1",
        "consignee_addr2", "consignee_addr3", "notify_name",
        "notify_addr1", "notify_addr2", "project_no", "project_name",
        "po_no", "country_of_origin"
    ]
    
    # Commons 프리필
    commons = commons_payload.get("commons", {})
    for key in REQUIRED_COMMONS:
        if key not in commons:
            if key == "dimension_cbm":
                commons[key] = 0.000
            elif key in ["net_weight_kg", "gross_weight_kg", "freight_usd", 
                         "insurance_usd", "total_cif_usd"]:
                commons[key] = 0.00
            elif key == "quantity_pkgs":
                commons[key] = 0
            else:
                commons[key] = ""
            warnings.append(f"가정: commons.{key} 누락으로 빈칸/0.00 처리")
    
    # Static parties 프리필
    parties = commons_payload.get("static_parties", {})
    for key in REQUIRED_PARTIES:
        if key not in parties:
            parties[key] = ""
            warnings.append(f"가정: static_parties.{key} 누락으로 빈칸 처리")
    
    return commons_payload

# 사용 예시
if missing_required_keys(commons_payload):
    commons_payload = apply_prefill_adapter(commons_payload, warnings)
    # Excel 생성은 계속 진행 (Style-first 유지, 빈칸으로 표시)
```

#### Step 3: 파이프라인 실행

```python
# make_cipl_set.py 실행
import subprocess
from pathlib import Path

# Commons Mode JSON 파일 저장
json_path = Path("temp_commons_input.json")
json_path.write_text(json.dumps(commons_payload, ensure_ascii=False, indent=2))

# 파이프라인 실행
result = subprocess.run(
    [
        "python",
        "CIPL/CIPL_PATCH_PACKAGE/make_cipl_set.py",
        "--in", str(json_path),
        "--out", f"CIPL_{invoice_no}_{date}.xlsx"
    ],
    cwd=Path("CONVERT"),
    capture_output=True,
    text=True
)

if result.returncode == 0:
    print(f"✅ Excel 생성 완료: {result.stdout}")
else:
    print(f"❌ 오류: {result.stderr}")
```

### 필드 매핑 가이드

| LDG_PAYLOAD | Commons Mode | 필수 여부 | 비고 |
|-------------|--------------|-----------|------|
| `data.Invoice_No` | `commons.invoice_no` | ✅ 필수 | |
| `data.Invoice_Date` | `commons.invoice_date` | ✅ 필수 | |
| `data.POL` | `commons.pol` | ✅ 필수 | |
| `data.POD` | `commons.pod` | ✅ 필수 | |
| `data.Carrier` | `commons.carrier` | ✅ 필수 | |
| `data.Sailing_On` | `commons.sailing_on` | ✅ 필수 | |
| `data.HS_Code` | `commons.hs_code` | ✅ 필수 | |
| `data.Manufacturer_Name` | `commons.manufacturer_name` | ✅ 필수 | |
| `data.Item_Description` | `commons.item_description` | ✅ 필수 | |
| `data.Packages` | `commons.quantity_pkgs` | ✅ 필수 | |
| `data.Net_Weight_kg` | `commons.net_weight_kg` | ✅ 필수 | |
| `data.Gross_Weight_kg` | `commons.gross_weight_kg` | ✅ 필수 | |
| `data.CBM` | `commons.dimension_cbm` | ✅ 필수 | 3 decimals |
| `data.Total_CIF_USD` | `commons.total_cif_usd` | ✅ 필수 | |
| `data.Freight_USD` | `commons.freight_usd` | ✅ 필수 | |
| `data.Insurance_USD` | `commons.insurance_usd` | ✅ 필수 | |
| `data.Shipper` | `static_parties.shipper_name` | ✅ 필수 | |
| `data.Consignee` | `static_parties.consignee_name1/2` | ✅ 필수 | |
| `data.Project_No` | `static_parties.project_no` | ✅ 필수 | |
| `data.PO_No` | `static_parties.po_no` | ✅ 필수 | |
| `data.Country_of_Origin` | `static_parties.country_of_origin` | ✅ 필수 | |

### ZERO Fail-Safe 처리

```python
# ZERO 조건 발생 시 처리
if zero_condition:
    # 1. LDG_AUDIT에 기록
    audit["zero_flags"] = {
        "active": True,
        "reasons": ["Template cannot be loaded"],
        "ZERO_FAILSAFE_LOG": {
            "reason": "make_cipl_set.py 모듈 로드 실패",
            "risk": "Excel 생성 불가, JSON만 제공",
            "requested_data": "올바른 경로 확인 필요",
            "next_action": "수동으로 make_cipl_set.py 경로 확인",
            "metrics_at_fail": {
                "mean_conf": 0.95,
                "table_acc": 98.5
            }
        }
    }
    
    # 2. 최선의 노력으로 계속 진행
    # - 부분 데이터로 Commons Mode JSON 생성
    # - NULL 값으로 채우고 warnings에 기록
    # - Excel 생성 시도 (실패해도 JSON은 제공)
```

### 출력 파일명 규칙

```python
# 파일명 생성 로직
if invoice_no:
    filename = f"CIPL_{invoice_no}_{date}.xlsx"
elif bl_no:
    filename = f"CIPL_{bl_no}_{date}.xlsx"
else:
    from datetime import datetime
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"CIPL_UNK_{timestamp}.xlsx"
```

### 검증 및 품질 체크

```python
# LDG_AUDIT 생성
audit = {
    "kpi": {
        "mean_conf": 0.95,  # ≥ 0.92 목표
        "table_acc": 98.5,  # ≥ 98.00% 목표
        "entity_match": 1.00
    },
    "cross_doc": {
        "summary": "All documents consistent",
        "findings": []
    },
    "zero_flags": {
        "active": False,
        "reasons": []
    },
    "warnings": [
        "가정: Freight와 Insurance가 0으로 추출됨"
    ]
}
```

### 통합 실행 예시

```python
# GPT가 실행해야 할 전체 프로세스
def generate_cipl_excel(user_input):
    # 1. 데이터 추출
    ldg_payload = extract_from_input(user_input)
    
    # 2. 검증
    audit = validate_data(ldg_payload)
    
    # 3. Commons Mode 변환
    commons_json = convert_to_commons_mode(ldg_payload)
    
    # 3.5. 프리필 어댑터 적용 (필수키 누락 대응)
    warnings = ldg_payload.get("warnings", [])
    commons_json = apply_prefill_adapter(commons_json, warnings)
    audit["warnings"] = warnings
    
    # 4. 파이프라인 실행
    excel_path = run_pipeline(commons_json)
    
    # 5. 결과 반환
    return {
        "guard_summary": generate_guard_summary(audit),
        "ldg_payload": ldg_payload,
        "ldg_audit": audit,
        "excel_path": excel_path
    }
```

### 주의사항

1. 경로: `CIPL_PATCH_PACKAGE/make_cipl_set.py` 사용
2. 모드: Commons Mode 우선 (자동 Rider 생성)
3. 단위: CBM은 3 decimals 유지
4. 가정: 모든 가정은 `warnings`에 "가정:" 접두사로 기록
5. ZERO: Excel 생성 실패해도 JSON은 제공
6. **프리필 어댑터**: 필수키 누락 시 빈칸/0.00으로 채우고 Excel 생성 계속 진행 (Style-first 유지)

### 프리필 어댑터 정책

- 목적: "기본 정보만으로도 생성" 요구 충족
- 원칙: HallucinationBan 준수 (값 추정 금지, 빈칸 유지)
- 처리:
  - 문자열 필드: `""` (빈칸)
  - 숫자 필드: `0.00` (CBM은 `0.000`)
  - 정수 필드: `0`
- 효과: 서식은 100% 유지, 데이터는 빈칸으로 표시
- 기록: 모든 프리필은 `warnings[]`에 "가정:" 접두사로 기록

이 가이드에 따라 GPT는 입력에서 Excel까지 전체 프로세스를 실행할 수 있습니다.
### CIPL_INPUT JSON Schema

- Schema file: `CIPL_PATCH_PACKAGE/cipl_input_schema_v2.4.json`
- Scope: Commons Mode input (`commons`, `static_parties`, optional rider items)
- Notes: `ci_rider_items.total_price` is used for totals; `amount` is allowed but ignored.

#### Schema Mode: Relaxed (GPT Extraction Friendly)

The schema uses `additionalProperties: true` to allow flexibility for GPT extraction:

- **Additional fields allowed**: GPT can include extra fields extracted from documents (e.g., `container_no`, `vessel_name`, `eta_date`) without validation errors
- **Required fields still enforced**: All 20 `commons` and 15 `static_parties` required fields must be present (or filled by prefill adapter)
- **Type validation maintained**: Number fields still enforce `minimum: 0`, string fields remain string type
- **Rider items flexible**: Additional properties in rider items are allowed for future extensions

This relaxed mode supports the GPTS goal of **maximizing CIPL generation from scanned files and text inputs** by accepting all extracted data while maintaining core validation.
