# Vessel Stability Booklet - Excel to Python

Vessel Stability Booklet Excel 파일의 모든 계산 함수를 Python으로 구현한 프로젝트입니다.

## 프로젝트 개요

이 프로젝트는 선박 안정성 계산서(Stability Booklet) Excel 파일의 모든 계산 로직을 Python으로 변환하여 구현한 것입니다. Excel 파일의 모든 시트에서 사용되는 함수들을 Python 클래스와 함수로 구현하여, Excel 없이도 동일한 계산을 수행할 수 있습니다.

## 주요 기능

### 구현된 Excel 함수

#### 1. Volum 시트 함수
- `calculate_weight()` - 중량 계산 (Volume × Density)
- `calculate_l_moment()` - 종향 모멘트 계산 (Weight × LCG)
- `calculate_v_moment()` - 수직 모멘트 계산 (Weight × VCG)
- `calculate_t_moment()` - 횡향 모멘트 계산 (Weight × TCG)
- `calculate_percentage()` - 용적 비율 계산
- `calculate_subtotal()` - Sub Total 계산
- `calculate_total_displacement()` - 최종 배수량 및 중심 계산

#### 2. Hydrostatic 시트 함수
- `calculate_bg()` - BG 계산 (LCB - LCG)
- `calculate_trim()` - Trim 계산 ((∆ × BG) / MTC)
- `calculate_diff()` - Diff 계산 (Above - Below)
- `calculate_interpolation_factor()` - 보간 계수 계산
- `interpolate_hydrostatic_data()` - Hydrostatic 데이터 보간
- `calculate_lost_gm()` - Lost GM 계산 (FSM / ∆)
- `calculate_vcg_corrected()` - VCG Corrected 계산
- `calculate_tan_list()` - Tan List 계산
- `interpolate_hydrostatic_by_draft()` - Draft에 따른 수정 데이터 보간

#### 3. GZ Curve 시트 함수
- `calculate_righting_arm()` - 복원팔 계산 (GZ(KN) - KG × Sin(Heel))
- `calculate_area_simpsons()` - Simpson's rule로 GZ 곡선 아래 면적 계산
- `interpolate_gz_between_displacements()` - 배수량에 따른 GZ 보간
- `interpolate_gz_complete()` - 완전한 GZ 보간 로직

#### 4. Trim = 0 시트 함수
- `interpolate_hydrostatic_by_draft()` - Draft 보간
- `get_displacement_by_draft()` - Draft로 배수량 찾기
- `get_mtc_by_draft()` - Draft로 MTC 찾기

## 설치 및 사용법

### 요구사항

```bash
pip install pandas numpy openpyxl xlrd
```

### 사용 예제

```python
from src.vessel_stability_functions import (
    StabilityCalculator,
    VesselParticulars,
    load_excel_data,
    extract_particulars_from_sheet,
    extract_hydrostatic_from_sheet
)

# Excel 파일 로드
data = load_excel_data("data/1.Vessel Stability Booklet.xls")

# 데이터 추출
particulars = extract_particulars_from_sheet(data['PRINCIPAL PARTICULARS'])
hydrostatic = extract_hydrostatic_from_sheet(data['Hydrostatic'])

# 계산기 생성
calculator = StabilityCalculator(particulars)

# BG 계산
bg = calculator.calculate_bg(hydrostatic.lcb, hydrostatic.lcg)
print(f"BG = {bg:.6f} m")

# Trim 계산
trim = calculator.calculate_trim(hydrostatic.displacement, bg, hydrostatic.mtc)
print(f"Trim = {trim:.6f} m")

# Lost GM 계산
lost_gm = calculator.calculate_lost_gm(hydrostatic.fsm, hydrostatic.displacement)
print(f"Lost GM = {lost_gm:.6f} m")
```

## 파일 구조

```
vessel_stability_python/
├── README.md                          # 프로젝트 개요
├── docs/                              # 문서 폴더
│   ├── implementation_guide.md        # 구현 가이드
│   ├── function_reference.md         # 함수 참조 문서
│   ├── validation_report.md          # 검증 결과 리포트
│   └── test_results.md               # 테스트 결과 문서
├── src/                               # 소스 코드
│   ├── vessel_stability_functions.py # 메인 구현 파일
│   ├── excel_to_python_stability.py  # 초기 버전 (참고용)
│   └── analyze_excel_functions.py    # 분석 스크립트
├── tests/                             # 테스트 파일
│   └── test_excel_functions.py       # 단위 테스트
├── validation/                        # 검증 스크립트
│   ├── validate_stability_calculations.py  # 전체 검증
│   └── validate_hydrostatic_detailed.py    # Hydrostatic 상세 검증
└── data/                              # 데이터 파일
    └── 1.Vessel Stability Booklet.xls      # 원본 Excel 파일
```

## 테스트 실행

```bash
# 단위 테스트 실행
python tests/test_excel_functions.py

# 전체 검증 실행
python validation/validate_stability_calculations.py

# Hydrostatic 시트 상세 검증
python validation/validate_hydrostatic_detailed.py
```

## 검증 결과

- ✅ BG 계산: Python과 Excel 완전 일치
- ✅ Lost GM 계산: 완전 일치
- ✅ VCG Corrected 계산: 완전 일치
- ✅ GM 계산: 완전 일치
- ✅ Tan List 계산: 완전 일치
- ✅ Diff 계산: 완전 일치
- ✅ 모든 단위 테스트 통과 (24/24)

## 구현된 시트

1. **PRINCIPAL PARTICULARS** - 선박 주요 제원
2. **Volum** - 탱크 용적 및 중량 계산
3. **Hydrostatic** - 수정 데이터 및 보간
4. **GZ Curve** - GZ 곡선 보간
5. **Trim = 0** - 수정 표 보간

## 주요 클래스

### StabilityCalculator
모든 Excel 함수를 구현한 메인 계산기 클래스입니다.

### VesselParticulars
선박 주요 제원을 저장하는 데이터 클래스입니다.

### HydrostaticData
수정 데이터를 저장하는 데이터 클래스입니다.

### GZData
GZ 곡선 데이터를 저장하는 데이터 클래스입니다.

## 빠른 시작

프로젝트 루트에서 예제 실행:

```bash
python example_usage.py
```

## 참고 문서

- `docs/USAGE_GUIDE.md` - **상세 사용 가이드** (추천!)
- `docs/implementation_guide.md` - 구현 가이드 및 과정
- `docs/function_reference.md` - 모든 함수의 API 문서
- `docs/validation_report.md` - 검증 결과 상세 리포트
- `docs/test_results.md` - 테스트 결과 문서

## 라이선스

이 프로젝트는 내부 사용을 위한 것입니다.

## 작성자

Vessel Stability Booklet Excel 함수 Python 구현 프로젝트

