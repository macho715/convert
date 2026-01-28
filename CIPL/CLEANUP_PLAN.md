# CIPL 폴더 정리 계획

**작성일**: 2026-01-14  
**목적**: 중복 파일 제거 및 최신 상태 유지  
**상태**: ✅ 완료

---

## 📊 중복 파일 분석

### 1. Python 소스 파일 중복

| 파일명 | 위치 | 상태 | 조치 |
|--------|------|------|------|
| `COMMERCIAL INVOICE.PY` | `CIPL_LEGACY/` | ⚠️ 레거시 | ✅ 이동 완료 |
| `COMMERCIAL INVOICE.PY` | `CIPL_PATCH_PACKAGE/` | ✅ 최신 | **활성 유지** |
| `PACKING LIST.PY` | `CIPL_LEGACY/` | ⚠️ 레거시 | ✅ 이동 완료 |
| `PACKING LIST.PY` | `CIPL_PATCH_PACKAGE/` | ✅ 최신 | **활성 유지** |
| `CI RIDER.PY` | `CIPL_LEGACY/` | ⚠️ 레거시 | ✅ 이동 완료 |
| `CI RIDER.PY` | `CIPL_PATCH_PACKAGE/` | ✅ 최신 | **활성 유지** |
| `PACKING LIST ATTACHED RIDER.PY` | `CIPL_LEGACY/` | ⚠️ 레거시 | ✅ 이동 완료 |
| `PACKING LIST ATTACHED RIDER.PY` | `CIPL_PATCH_PACKAGE/` | ✅ 최신 | **활성 유지** |
| `CIPL.PY` | `CIPL_LEGACY/` | ⚠️ 레거시 | ✅ 이동 완료 |
| `CIPL.py` | `CIPL_PATCH_PACKAGE/` | ✅ 최신 | **활성 유지** |
| `make_cipl_set.py` | `CIPL_LEGACY/` | ⚠️ 레거시 | ✅ 이동 완료 |
| `make_cipl_set.py` | `CIPL_PATCH_PACKAGE/` | ✅ 최신 | **활성 유지** |

### 2. 테스트/임시 Excel 파일 (삭제 완료)

**CIPL_PATCH_PACKAGE/** 폴더의 다음 파일들은 삭제 완료:
- ✅ `CIPL_D5E5_LINE.xlsx`
- ✅ `CIPL_FINAL_V2.xlsx`, `CIPL_FINAL_V3.xlsx`
- ✅ `CIPL_HEADER_FIXED.xlsx`, `CIPL_HEADER_SHIFTED.xlsx`
- ✅ `CIPL_ROW45_FIXED.xlsx`, `CIPL_ROW45_CORRECTED.xlsx`
- ✅ `CIPL_TEST.xlsx`, `CIPL_UPDATED.xlsx`, `CIPL_PATCHED.xlsx`
- ✅ `CIPL_OUTPUT_PATCHED.xlsx`, `CIPL_PROJECT_BOX_FIXED.xlsx`
- ✅ `VERIFY_OPTIMIZED.xlsx`
- ✅ `~$VERIFY_OPTIMIZED_FINAL.xlsx` (임시 파일)

**보관 중**:
- `CIPL_FINAL.xlsx` (최종 출력 예시)
- `CIPL_HVDC-ADOPT-SCT-0159.xlsx` (실제 프로젝트 출력)
- `VERIFY_BASELINE.xlsx`, `VERIFY_OPTIMIZED_FINAL.xlsx` (검증용, 선택적 보관)

### 3. 기타 중복 파일 (삭제 완료)

- ✅ `CIPL_PATCH_PACKAGE.zip` (CIPL/ 폴더) - 삭제 완료
- ✅ `COMMERCIAL_INVOICE_TEST.xlsx` (CIPL/) - 삭제 완료
- ✅ `voyage_input_sample_full.json` (CIPL/) - 삭제 완료 (CIPL_PATCH_PACKAGE/의 것만 유지)

### 4. 남은 파일

**CIPL/** 폴더:
- `CIPL_HVDC-ADOPT-SCT-0159.xlsx` - CIPL_PATCH_PACKAGE/에도 있음 (수동 확인 후 삭제 권장)
- `voyage_input.json` - 확장 모드 입력 파일 (보관)
- `Samsung_Logistics_Docs_HighFidelity.xlsx` - 참고용 (보관)

---

## ✅ 정리 실행 완료

### Phase 1: 레거시 파일 보관 ✅
- ✅ `CIPL_LEGACY/` 폴더 생성
- ✅ 레거시 Python 파일 이동 완료

### Phase 2: 테스트 파일 정리 ✅
- ✅ 테스트 Excel 파일 삭제 완료
- ✅ 임시 파일 삭제 완료

### Phase 3: 중복 파일 제거 ✅
- ✅ 중복 파일 삭제 완료

### Phase 4: 추가 파일 정리 ✅
- ✅ 임시 파일 삭제 완료

---

## 📁 최종 구조

```
CIPL/
├── CIPL_PATCH_PACKAGE/          # ✅ 활성 버전 (최신)
│   ├── COMMERCIAL INVOICE.PY
│   ├── PACKING LIST.PY
│   ├── CI RIDER.PY
│   ├── PACKING LIST ATTACHED RIDER.PY
│   ├── CIPL.py
│   ├── make_cipl_set.py
│   ├── voyage_input_sample_full.json
│   ├── voyage_input_commons_TEMPLATE.json
│   ├── SAMSUNGLOGO.PNG
│   ├── LDG_PAYLOAD_PRL-MIR-032-A2_2026-01-06.json
│   ├── CIPL_FINAL.xlsx
│   ├── CIPL_HVDC-ADOPT-SCT-0159.xlsx
│   ├── VERIFY_BASELINE.xlsx
│   ├── VERIFY_OPTIMIZED_FINAL.xlsx
│   ├── OPTIMIZATION_SUMMARY.md
│   ├── compare_outputs.py
│   └── verify_visual_output.py
│
├── CIPL_LEGACY/                 # ⚠️ 레거시 (참고용)
│   ├── COMMERCIAL INVOICE.PY
│   ├── PACKING LIST.PY
│   ├── CI RIDER.PY
│   ├── PACKING LIST ATTACHED RIDER.PY
│   ├── CIPL.PY
│   └── make_cipl_set.py
│
├── Samsung_Logistics_Docs_HighFidelity.xlsx  # 참고용
├── voyage_input.json            # 확장 모드 입력 (보관)
├── CIPL_HVDC-ADOPT-SCT-0159.xlsx  # 수동 확인 후 삭제 권장
├── UPDATE_LOG.md                # 업데이트 로그
└── CLEANUP_PLAN.md              # 이 문서
```

---

## ⚠️ 추가 확인 사항

### 최적화 적용 상태
| 파일 | 최적화 적용 | 비고 |
|------|-------------|------|
| `COMMERCIAL INVOICE.PY` | ✅ 적용됨 | excel_helpers 사용 |
| `PACKING LIST.PY` | ✅ 적용됨 | excel_helpers 사용 |
| `CI RIDER.PY` | ❌ 미적용 | 향후 적용 예정 |
| `PACKING LIST ATTACHED RIDER.PY` | ❌ 미적용 | 향후 적용 예정 |

### 의존성
- `excel_helpers.py`: `CONVERT/excel_helpers.py` 위치
  - 모든 최적화된 파일이 이 모듈에 의존
  - ROOT_DIR 경로 설정으로 접근

### 입력 파일 형식 차이
- **확장 모드** (`voyage_input.json`):
  ```json
  {
    "ci_p1": {...},
    "ci_rider_p2": {...},
    "pl_p1": {...},
    "pl_rider_p2": {...}
  }
  ```
- **Commons 모드** (`voyage_input_sample_full.json`):
  ```json
  {
    "commons": {...},
    "static_parties": {...},
    "ci_rider_items": [...],
    "pl_rider_items": [...]
  }
  ```
  - `CIPL.py`의 `make_4page_data_dicts()`로 변환됨

---

## ✅ 정리 완료 체크리스트

- [x] 전체 폴더 백업 (권장)
- [x] 레거시 파일 이동
- [x] 테스트 파일 삭제
- [x] 중복 파일 제거
- [x] 추가 파일 정리 (임시 파일 삭제)
- [x] 최종 구조 확인
- [x] 문서 생성 (UPDATE_LOG.md, CLEANUP_PLAN.md)

---

**정리 완료일**: 2026-01-14  
**상태**: ✅ 완료

