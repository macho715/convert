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


### CIPL_INPUT JSON Schema (권장)

- 스키마 파일(권장): `CIPL/CIPL_PATCH_PACKAGE/cipl_input_schema_v2.4.1.json`
- 목적: Commons Mode JSON을 **키 누락/타입 오류** 없이 정규화하여 `make_cipl_set.py` 실행 실패(KeyError)를 최소화
- 운영 권장 플로우:
  1) (선택) LDG_PAYLOAD v2.4 생성
  2) Commons 변환
  3) **PrefillAdapter**로 필수 키 존재 보장 + NULL/빈값 기본값 채움
  4) JSON Schema 검증(본 스키마)
  5) Excel AutoBuild 실행

> 주의: Knowledge 파일은 텍스트 기반으로 처리되므로(특히 이미지/로고), 빌더 실행에 필요한 `*.py`/`*.png`는 **Code Interpreter 실행 환경에 실제 파일로 존재**하도록 구성 필요.

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

이 가이드에 따라 GPT는 입력에서 Excel까지 전체 프로세스를 실행할 수 있습니다.