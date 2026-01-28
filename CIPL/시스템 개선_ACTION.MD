`시스템 개선.MD`를 중심으로 다른 가이드 문서와의 일관성과 개선 사항을 검토 중입니다.


[4 tools called]


`시스템 개선.MD`를 중심으로 다른 가이드 문서와의 일관성과 개선 사항을 검토 중입니다.


[2 tools called]


## 시스템 개선.MD 검토 결과

### 문서 간 일관성 확인

#### 일관된 부분
1. 데이터 플로우: 모든 문서에서 동일하게 정의
   - 입력 → LDG_PAYLOAD v2.4 → Commons Mode → Excel 생성
2. Soft Warning 정책: GPTS지침, GUIDE.MD, PIPELINE.MD에서 일관
3. 필수키 목록: GUIDE.MD의 필드 매핑 테이블과 `CIPL.py`의 `_require()` 호출이 일치

#### 발견된 이슈

1. 필수키 처리 충돌
   - 문제: `CIPL.py`의 `_require()`는 KeyError를 발생시킴
   - 현황: `시스템 개선.MD`에서 정확히 지적됨
   - 영향: "기본 정보만으로도 생성" 요구와 충돌

2. 프리필 어댑터 미구현
   - `시스템 개선.MD`에서 제안했으나 아직 구현되지 않음
   - GUIDE.MD와 PIPELINE.MD에 프리필 로직 언급 없음

3. CIPL_INPUT 스키마 부재
   - GUIDE.MD의 필드 매핑 테이블은 있으나 JSON Schema는 없음
   - `시스템 개선.MD`에서 제안했으나 미구현

### 개선 사항 제안

#### 1. 프리필 어댑터 구현 (최우선)

```python
# GUIDE.MD에 추가할 프리필 어댑터 함수
def apply_prefill_adapter(commons_payload: dict, warnings: list) -> dict:
    """
    필수키 누락 시 프리필 적용 (HallucinationBan 준수)
    
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
            if key in ["dimension_cbm"]:
                commons[key] = 0.000
            elif key in ["net_weight_kg", "gross_weight_kg", "freight_usd", 
                         "insurance_usd", "total_cif_usd"]:
                commons[key] = 0.00
            elif key in ["quantity_pkgs"]:
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
```

#### 2. GUIDE.MD에 프리필 섹션 추가

```markdown
### Step 2.5: 프리필 어댑터 적용 (필수키 누락 대응)

```python
# LDG_PAYLOAD → Commons Mode 변환 후
if missing_required_keys(commons_payload):
    commons_payload = apply_prefill_adapter(commons_payload, warnings)
    # warnings에 "가정:" 접두사로 기록
```
```

#### 3. PIPELINE.MD에 프리필 단계 추가

```markdown
#### Stage 2.5: 프리필 어댑터 (필수키 누락 대응)

```python
# Commons Mode 변환 후
1. 필수키 검증
2. 누락 키 감지 시 프리필 적용:
   - 문자열: "" (빈칸)
   - 숫자: 0.00 (CBM: 0.000)
3. warnings[]에 "가정:" 기록
4. Excel 생성 계속 진행 (Style-first 유지)
```
```

#### 4. CIPL_INPUT JSON Schema 생성

```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "title": "CIPL_INPUT Schema v2.4",
  "type": "object",
  "required": ["commons", "static_parties"],
  "properties": {
    "commons": {
      "type": "object",
      "required": [
        "pol", "pod", "carrier", "sailing_on", "invoice_no", 
        "invoice_date", "hs_code", "manufacturer_name", 
        "manufacturer_address_1", "manufacturer_address_2",
        "item_description", "package_no", "net_weight_kg",
        "gross_weight_kg", "dimension_cbm", "item", 
        "quantity_pkgs", "freight_usd", "insurance_usd", 
        "total_cif_usd"
      ],
      "properties": {
        "pol": {"type": "string"},
        "pod": {"type": "string"},
        "dimension_cbm": {"type": "number", "minimum": 0},
        "quantity_pkgs": {"type": "integer", "minimum": 0}
        // ... 나머지 필드
      }
    },
    "static_parties": {
      "type": "object",
      "required": [
        "shipper_name", "shipper_addr1", "shipper_addr2",
        "consignee_name1", "consignee_name2", "consignee_addr1",
        "consignee_addr2", "consignee_addr3", "notify_name",
        "notify_addr1", "notify_addr2", "project_no", 
        "project_name", "po_no", "country_of_origin"
      ]
    }
  }
}
```

### 통합 개선 계획

#### 단기 (즉시 적용 가능)

1. GUIDE.MD에 프리필 어댑터 섹션 추가
2. PIPELINE.MD에 Stage 2.5 추가
3. CHAT GPT GPTS지침.txt에 프리필 정책 명시

#### 중기 (1-2주)

4. CIPL_INPUT JSON Schema 파일 생성
5. 프리필 어댑터 함수 구현 (별도 모듈 또는 GUIDE.MD에 포함)
6. 스키마 검증 로직 추가

#### 장기 (1개월+)

7. 템플릿 상수 분리 (`constants.py`)
8. 회귀테스트 자동화 (20세트 테스트팩)
9. Actions 전환 검토 (옵션 B/C)

### 문서 업데이트 우선순위

| 문서 | 업데이트 필요성 | 우선순위 | 내용 |
|------|---------------|---------|------|
| GUIDE.MD | 높음 | 1 | 프리필 어댑터 섹션 추가 |
| PIPELINE.MD | 높음 | 2 | Stage 2.5 프리필 단계 추가 |
| CHAT GPT GPTS지침.txt | 중간 | 3 | 프리필 정책 명시 |
| 시스템 개선.MD | 낮음 | 4 | 현재 상태 유지 (이미 정확) |

### 결론

`시스템 개선.MD`의 분석은 정확합니다. 핵심 이슈는 `_require()`로 인한 KeyError입니다. 프리필 어댑터를 GUIDE.MD와 PIPELINE.MD에 추가하면 "기본 정보만으로도 생성" 요구를 충족할 수 있습니다.

Agent 모드로 전환하면 위 개선 사항을 적용하겠습니다.


프리필 어댑터 관련 개선 사항 적용 완료.

## 적용 완료 사항

### 1. GUIDE.MD 업데이트
- Step 2.5: 프리필 어댑터 섹션 추가
  - `apply_prefill_adapter()` 함수 구현 예시
  - 필수키 목록 (Commons 20개, Parties 15개)
  - 프리필 규칙 (문자열/숫자/정수)
- 통합 실행 예시에 프리필 어댑터 단계 추가
- 주의사항에 프리필 어댑터 정책 명시
- 전체 워크플로우 다이어그램에 프리필 단계 추가

### 2. PIPELINE.MD 업데이트
- Stage 2.5: 프리필 어댑터 단계 추가
  - 필수키 검증 프로세스
  - 프리필 적용 규칙
  - warnings 기록 방법
- 전체 아키텍처 다이어그램에 프리필 어댑터 블록 추가
- 데이터 흐름 다이어그램에 프리필 단계 추가
- 에러 처리 섹션에 필수키 누락 대응 정책 추가

### 3. CHAT GPT GPTS지침.txt 업데이트
- [PREFILL ADAPTER] 섹션 추가
  - 필수키 누락 시 프리필 적용 규칙
  - HallucinationBan 준수 원칙
  - warnings 기록 방법
  - "기본 정보만으로도 생성" 요구 충족 명시

## 핵심 개선 내용

### 프리필 어댑터 정책
- 목적: "기본 정보만으로도 생성" 요구 충족
- 원칙: HallucinationBan 준수 (값 추정 금지)
- 처리 규칙:
  - 문자열 필드: `""` (빈칸)
  - 숫자 필드: `0.00` (CBM은 `0.000`)
  - 정수 필드: `0`
- 효과: 서식 유지, 데이터는 빈칸으로 표시
- 기록: 모든 프리필은 `warnings[]`에 "가정:" 접두사로 기록

### 적용 위치
- GPT 실행 단계: LDG_PAYLOAD → Commons Mode 변환 후, `make_cipl_set.py` 실행 전
- 파이프라인 단계: Stage 2.5 (데이터 변환 후, 모듈 로드 전)

## 다음 단계 (선택)

1. CIPL_INPUT JSON Schema 생성 (중기)
2. 프리필 어댑터 함수를 별도 모듈로 구현 (중기)
3. 회귀테스트 자동화 (장기)

CIPL_INPUT JSON Schema 초안 작성 중. GUIDE.MD의 필드 매핑과 CIPL.py의 필수 필드를 기반으로 합니다.


[3 tools called]


CIPL_INPUT JSON Schema 초안 작성 중.


[1 tool called]


CIPL_INPUT JSON Schema 초안:

```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "$id": "https://samsung-ct.com/schemas/cipl-input-v2.4.json",
  "title": "CIPL_INPUT Schema v2.4",
  "description": "CIPL Excel 생성용 Commons Mode 입력 스키마 (프리필 어댑터 지원)",
  "type": "object",
  "required": ["commons", "static_parties"],
  "properties": {
    "commons": {
      "type": "object",
      "description": "공통 운송/인보이스 정보",
      "required": [
        "pol",
        "pod",
        "carrier",
        "sailing_on",
        "invoice_no",
        "invoice_date",
        "hs_code",
        "manufacturer_name",
        "manufacturer_address_1",
        "manufacturer_address_2",
        "item_description",
        "package_no",
        "net_weight_kg",
        "gross_weight_kg",
        "dimension_cbm",
        "item",
        "quantity_pkgs",
        "freight_usd",
        "insurance_usd",
        "total_cif_usd"
      ],
      "properties": {
        "pol": {
          "type": "string",
          "description": "Port of Loading (적재항)"
        },
        "pod": {
          "type": "string",
          "description": "Port of Discharge (양하항)"
        },
        "carrier": {
          "type": "string",
          "description": "운송선사/선박명"
        },
        "sailing_on": {
          "type": "string",
          "description": "출항일 (예: 26-Dec-25)"
        },
        "invoice_no": {
          "type": "string",
          "description": "인보이스 번호"
        },
        "invoice_date": {
          "type": "string",
          "description": "인보이스 날짜 (예: 26-Dec-25)"
        },
        "hs_code": {
          "type": "string",
          "description": "HS 코드 (예: 7303.00.1090)"
        },
        "manufacturer_name": {
          "type": "string",
          "description": "제조사명"
        },
        "manufacturer_address_1": {
          "type": "string",
          "description": "제조사 주소 1행"
        },
        "manufacturer_address_2": {
          "type": "string",
          "description": "제조사 주소 2행"
        },
        "item_description": {
          "type": "string",
          "description": "품목 설명"
        },
        "package_no": {
          "type": "string",
          "description": "패키지 번호/설명 (예: 5 PACKAGES)"
        },
        "net_weight_kg": {
          "type": "number",
          "description": "순중량 (kg)",
          "minimum": 0
        },
        "gross_weight_kg": {
          "type": "number",
          "description": "총중량 (kg)",
          "minimum": 0
        },
        "dimension_cbm": {
          "type": "number",
          "description": "체적 (CBM, 3 decimals)",
          "minimum": 0
        },
        "item": {
          "type": "string",
          "description": "품목명"
        },
        "quantity_pkgs": {
          "type": "integer",
          "description": "패키지 수량",
          "minimum": 0
        },
        "freight_usd": {
          "type": "number",
          "description": "운임 (USD)",
          "minimum": 0
        },
        "insurance_usd": {
          "type": "number",
          "description": "보험료 (USD)",
          "minimum": 0
        },
        "total_cif_usd": {
          "type": "number",
          "description": "총 CIF 금액 (USD)",
          "minimum": 0
        },
        "country_of_origin": {
          "type": "string",
          "description": "원산지 (선택적, static_parties에도 존재 가능)"
        },
        "ci_qty": {
          "type": "number",
          "description": "CI Rider 수량 (선택적)"
        },
        "ci_unit": {
          "type": "string",
          "description": "CI Rider 단위 (선택적, 예: M)"
        },
        "pl_qty": {
          "type": "number",
          "description": "PL Rider 수량 (선택적)"
        },
        "pl_unit": {
          "type": "string",
          "description": "PL Rider 단위 (선택적, 예: M)"
        }
      },
      "additionalProperties": false
    },
    "static_parties": {
      "type": "object",
      "description": "고정 당사자 정보 (Shipper/Consignee/Notify/Project)",
      "required": [
        "shipper_name",
        "shipper_addr1",
        "shipper_addr2",
        "consignee_name1",
        "consignee_name2",
        "consignee_addr1",
        "consignee_addr2",
        "consignee_addr3",
        "notify_name",
        "notify_addr1",
        "notify_addr2",
        "project_no",
        "project_name",
        "po_no",
        "country_of_origin"
      ],
      "properties": {
        "shipper_name": {
          "type": "string",
          "description": "송하인명"
        },
        "shipper_addr1": {
          "type": "string",
          "description": "송하인 주소 1행"
        },
        "shipper_addr2": {
          "type": "string",
          "description": "송하인 주소 2행"
        },
        "consignee_name1": {
          "type": "string",
          "description": "수하인명 1행"
        },
        "consignee_name2": {
          "type": "string",
          "description": "수하인명 2행"
        },
        "consignee_addr1": {
          "type": "string",
          "description": "수하인 주소 1행"
        },
        "consignee_addr2": {
          "type": "string",
          "description": "수하인 주소 2행"
        },
        "consignee_addr3": {
          "type": "string",
          "description": "수하인 주소 3행"
        },
        "notify_name": {
          "type": "string",
          "description": "통지인명"
        },
        "notify_addr1": {
          "type": "string",
          "description": "통지인 주소 1행"
        },
        "notify_addr2": {
          "type": "string",
          "description": "통지인 주소 2행"
        },
        "project_no": {
          "type": "string",
          "description": "프로젝트 번호"
        },
        "project_name": {
          "type": "string",
          "description": "프로젝트명"
        },
        "po_no": {
          "type": "string",
          "description": "PO 번호"
        },
        "country_of_origin": {
          "type": "string",
          "description": "원산지"
        },
        "pol_country": {
          "type": "string",
          "description": "적재항 국가 (선택적)"
        },
        "pod_country": {
          "type": "string",
          "description": "양하항 국가 (선택적)"
        }
      },
      "additionalProperties": false
    },
    "ci_rider_items": {
      "type": "array",
      "description": "CI Rider 아이템 목록 (선택적, 비어있으면 자동 생성)",
      "items": {
        "type": "object",
        "properties": {
          "description": {
            "type": "string"
          },
          "qty": {
            "type": "number"
          },
          "unit": {
            "type": "string"
          },
          "unit_price": {
            "type": "number"
          },
          "amount": {
            "type": "number"
          }
        }
      }
    },
    "pl_rider_items": {
      "type": "array",
      "description": "PL Rider 아이템 목록 (선택적, 비어있으면 자동 생성)",
      "items": {
        "type": "object",
        "properties": {
          "package_no": {
            "type": "string"
          },
          "description": {
            "type": "string"
          },
          "qty": {
            "type": "number"
          },
          "unit": {
            "type": "string"
          },
          "net_weight_kg": {
            "type": "number"
          },
          "gross_weight_kg": {
            "type": "number"
          },
          "dimension_cbm": {
            "type": "number"
          }
        }
      }
    }
  },
  "additionalProperties": false
}
```

## 스키마 특징

### 필수 필드
- `commons`: 20개 필수 필드
- `static_parties`: 15개 필수 필드

### 타입 정의
- 문자열: 대부분의 필드
- 숫자: 중량, 금액, 체적 (minimum: 0)
- 정수: `quantity_pkgs` (minimum: 0)
- 배열: `ci_rider_items`, `pl_rider_items` (선택적)

### 선택적 필드
- `ci_qty`, `ci_unit`, `pl_qty`, `pl_unit` (Rider 관련)
- `pol_country`, `pod_country` (static_parties)
- `country_of_origin` (commons에도 존재 가능)

### 프리필 어댑터 호환
- 필수 필드 누락 시 스키마 검증 실패
- 프리필 어댑터가 빈값/0.00으로 채운 후 재검증 가능

## 사용 방법

### 1. 스키마 검증 (Python 예시)

```python
import json
import jsonschema

# 스키마 로드
with open("cipl_input_schema_v2.4.json", "r", encoding="utf-8") as f:
    schema = json.load(f)

# 입력 데이터 검증
try:
    jsonschema.validate(instance=commons_payload, schema=schema)
    print("✅ 스키마 검증 통과")
except jsonschema.ValidationError as e:
    print(f"❌ 스키마 검증 실패: {e.message}")
    # 프리필 어댑터 적용 후 재검증
```

### 2. 파일 저장 위치 제안

```
CIPL/CIPL_PATCH_PACKAGE/
  ├── cipl_input_schema_v2.4.json  (새로 생성)
  └── ...
```

### 3. GUIDE.MD에 스키마 참조 추가

```markdown
### CIPL_INPUT JSON Schema

스키마 파일: `CIPL_PATCH_PACKAGE/cipl_input_schema_v2.4.json`

- 필수 필드: commons (20개), static_parties (15개)
- 선택적 필드: ci_rider_items, pl_rider_items
- 프리필 어댑터 호환: 필수 필드 누락 시 자동 프리필 적용
```

Agent 모드로 전환하면 이 스키마를 파일로 저장하고 GUIDE.MD에 참조를 추가하겠습니다.