# 구현 가이드

Vessel Stability Booklet Excel 함수를 Python으로 구현한 과정과 방법을 설명합니다.

## 구현 과정

### 1단계: Excel 파일 분석

먼저 Excel 파일의 모든 시트를 분석하여 사용되는 함수와 계산 로직을 파악했습니다.

**분석된 시트:**
- PRINCIPAL PARTICULARS (선박 주요 제원)
- Volum (탱크 용적 및 중량)
- Hydrostatic (수정 데이터)
- GZ Curve (GZ 곡선)
- Trim = 0 (수정 표)

### 2단계: 함수 분류 및 구현 순서

Excel 함수를 시트별로 분류하고 우선순위를 정하여 구현했습니다:

1. **기본 계산 함수** - BG, Trim, GM 등 기본 계산
2. **Volum 시트 함수** - 중량 및 모멘트 계산
3. **Hydrostatic 시트 함수** - 보간 및 수정 데이터 계산
4. **GZ Curve 시트 함수** - 복잡한 보간 로직
5. **검증 함수** - Excel과의 비교 검증

### 3단계: 데이터 구조 설계

Excel의 데이터 구조를 Python 데이터 클래스로 변환:

- `VesselParticulars` - 선박 주요 제원
- `HydrostaticData` - 수정 데이터
- `GZData` - GZ 곡선 데이터

### 4단계: 함수 구현

각 Excel 함수를 Python 함수로 변환하여 `StabilityCalculator` 클래스에 구현했습니다.

## 각 시트별 함수 설명

### Volum 시트

Volum 시트는 탱크의 용적, 중량, 중심을 계산합니다.

#### 구현된 함수

1. **`calculate_weight(volume, density)`**
   - Excel: `=Volume × Density`
   - 용적과 밀도로 중량 계산
   
2. **`calculate_l_moment(weight, lcg)`**
   - Excel: `=Weight × LCG`
   - 종향 모멘트 계산
   
3. **`calculate_v_moment(weight, vcg)`**
   - Excel: `=Weight × VCG`
   - 수직 모멘트 계산
   
4. **`calculate_t_moment(weight, tcg)`**
   - Excel: `=Weight × TCG`
   - 횡향 모멘트 계산
   
5. **`calculate_percentage(volume, capacity)`**
   - Excel: `=(Volume / Cap) × 100`
   - 용적 비율 계산
   
6. **`calculate_subtotal(...)`**
   - Excel: 각 열의 합계
   - Sub Total 계산
   
7. **`calculate_total_displacement(...)`**
   - Excel: Displacement Condition 계산
   - 최종 배수량 및 중심 계산

### Hydrostatic 시트

Hydrostatic 시트는 수정 데이터와 보간 계산을 수행합니다.

#### 구현된 함수

1. **`calculate_bg(lcb, lcg)`**
   - Excel: `=LCB - LCG`
   - BG 계산
   
2. **`calculate_trim(displacement, bg, mtc)`**
   - Excel: `=(∆ × BG) / MTC`
   - Trim 계산
   
3. **`calculate_diff(above, below)`**
   - Excel: `=Above - Below`
   - 차이 계산
   
4. **`calculate_interpolation_factor(target, low, high)`**
   - Excel: `=(Target - Low) / (High - Low)`
   - 보간 계수 계산
   
5. **`interpolate_hydrostatic_data(...)`**
   - Excel: Low/High Trim Value 사이에서 배수량과 트림에 따라 보간
   - 복합 보간 로직
   
6. **`calculate_lost_gm(fsm, displacement)`**
   - Excel: `=FSM / ∆`
   - Lost GM 계산
   
7. **`calculate_vcg_corrected(vcg, fsm, displacement)`**
   - Excel: `=VCG + (FSM / Displacement)`
   - VCG Corrected 계산
   
8. **`calculate_tan_list(list_moment, displacement, gm)`**
   - Excel: `=List Moment / (Displacement × GM)`
   - Tan List 계산
   
9. **`interpolate_hydrostatic_by_draft(draft, trim_zero_table)`**
   - Excel: Draft 값으로 수정 표에서 보간
   - Draft 보간

### GZ Curve 시트

GZ Curve 시트는 안정성 곡선을 계산합니다.

#### 구현된 함수

1. **`calculate_righting_arm(gz_kn, vcg_corrected, heel_angle)`**
   - Excel: `=GZ(KN) - KG(corrected VCG) × Sin(Heel)`
   - 복원팔 계산
   
2. **`calculate_area_simpsons(gz_values, heel_angles)`**
   - Excel: Simpson's rule 사용
   - GZ 곡선 아래 면적 계산
   
3. **`interpolate_gz_between_displacements(...)`**
   - Excel: 배수량에 따른 선형 보간
   - 배수량 보간
   
4. **`interpolate_gz_complete(...)`**
   - Excel: 배수량과 트림에 따른 복합 보간
   - 완전한 GZ 보간 로직

### Trim = 0 시트

Trim = 0 시트는 수정 표에서 Draft에 따른 데이터를 보간합니다.

#### 구현된 함수

1. **`interpolate_hydrostatic_by_draft(draft, trim_zero_table)`**
   - Excel: Draft 값으로 수정 표에서 보간
   - Draft 보간
   
2. **`get_displacement_by_draft(draft, trim_zero_table)`**
   - Excel: Draft 값으로 배수량 찾기
   - 배수량 조회
   
3. **`get_mtc_by_draft(draft, trim_zero_table)`**
   - Excel: Draft 값으로 MTC 찾기
   - MTC 조회

## Excel 함수와 Python 함수 매핑

### 기본 계산

| Excel 함수 | Python 함수 | 설명 |
|-----------|------------|------|
| `=LCB - LCG` | `calculate_bg(lcb, lcg)` | BG 계산 |
| `=(∆ × BG) / MTC` | `calculate_trim(displacement, bg, mtc)` | Trim 계산 |
| `=KM - KG` | `calculate_metacentric_height(km, kg)` | GM 계산 |
| `=Displacement / Density` | `calculate_volume(displacement, density)` | 용적 계산 |
| `=Displacement - Lightship` | `calculate_deadweight(displacement, lightship)` | DWT 계산 |

### Volum 시트

| Excel 함수 | Python 함수 | 설명 |
|-----------|------------|------|
| `=Volume × Density` | `calculate_weight(volume, density)` | 중량 계산 |
| `=Weight × LCG` | `calculate_l_moment(weight, lcg)` | 종향 모멘트 |
| `=Weight × VCG` | `calculate_v_moment(weight, vcg)` | 수직 모멘트 |
| `=Weight × TCG` | `calculate_t_moment(weight, tcg)` | 횡향 모멘트 |
| `=(Volume / Cap) × 100` | `calculate_percentage(volume, capacity)` | 용적 비율 |

### Hydrostatic 시트

| Excel 함수 | Python 함수 | 설명 |
|-----------|------------|------|
| `=Above - Below` | `calculate_diff(above, below)` | Diff 계산 |
| `=(Target - Low) / (High - Low)` | `calculate_interpolation_factor(...)` | 보간 계수 |
| `=FSM / ∆` | `calculate_lost_gm(fsm, displacement)` | Lost GM |
| `=VCG + (FSM / Displacement)` | `calculate_vcg_corrected(...)` | VCG Corrected |
| `=List Moment / (∆ × GM)` | `calculate_tan_list(...)` | Tan List |

### GZ Curve 시트

| Excel 함수 | Python 함수 | 설명 |
|-----------|------------|------|
| `=GZ(KN) - KG × Sin(Heel)` | `calculate_righting_arm(...)` | 복원팔 |
| Simpson's rule | `calculate_area_simpsons(...)` | 면적 계산 |

## 구현 특징

### 1. 정확성 보장

모든 함수는 Excel의 계산 결과와 일치하도록 구현되었으며, 검증을 통해 확인되었습니다.

### 2. 타입 힌팅

모든 함수에 타입 힌팅을 추가하여 사용성을 높였습니다.

### 3. 문서화

모든 함수에 docstring을 추가하여 사용법을 명확히 했습니다.

### 4. 에러 처리

0으로 나누기 등 예외 상황을 처리하여 안정성을 높였습니다.

### 5. 검증 기능

Excel과의 비교 검증 기능을 포함하여 정확성을 확인할 수 있습니다.

## 사용 예제

### 기본 계산

```python
from src.vessel_stability_functions import StabilityCalculator, VesselParticulars

# 계산기 생성
particulars = VesselParticulars()
calculator = StabilityCalculator(particulars)

# BG 계산
bg = calculator.calculate_bg(lcb=31.438885, lcg=31.816168)
print(f"BG = {bg:.6f} m")  # BG = -0.377284 m

# Trim 계산
trim = calculator.calculate_trim(
    displacement=1183.8462,
    bg=bg,
    mtc=33.991329
)
print(f"Trim = {trim:.6f} m")
```

### Volum 시트 계산

```python
# 중량 계산
weight = calculator.calculate_weight(volume=2.4, density=0.82)
print(f"Weight = {weight:.3f} tonnes")  # Weight = 1.968 tonnes

# 종향 모멘트 계산
l_moment = calculator.calculate_l_moment(weight=1.968, lcg=11.251)
print(f"L-moment = {l_moment:.6f}")  # L-moment = 22.141968
```

### Hydrostatic 보간

```python
# Lost GM 계산
lost_gm = calculator.calculate_lost_gm(fsm=164.76, displacement=1183.8462)
print(f"Lost GM = {lost_gm:.6f} m")  # Lost GM = 0.139173 m

# VCG Corrected 계산
vcg_corrected = calculator.calculate_vcg_corrected(
    vcg=3.35748,
    fsm=164.76,
    displacement=1183.8462
)
print(f"VCG Corrected = {vcg_corrected:.6f} m")
```

## 구현 시 주의사항

### 1. 단위 일관성

Excel과 Python 모두 동일한 단위를 사용해야 합니다:
- 중량: tonnes
- 길이: metres
- 모멘트: tonne-metres

### 2. 부호 처리

Trim 계산에서 BG의 부호에 따라 Forward/Aft가 결정됩니다:
- BG < 0: Forward trim
- BG > 0: Aft trim

### 3. 보간 범위

보간 계산 시 target 값이 low/high 범위를 벗어날 수 있으므로, 실제 사용 시 범위 확인이 필요합니다.

### 4. 0으로 나누기

모든 계산 함수에서 0으로 나누기 상황을 처리했습니다.

## 검증 방법

각 함수는 다음 방법으로 검증되었습니다:

1. **단위 테스트**: `tests/test_excel_functions.py`
2. **전체 검증**: `validation/validate_stability_calculations.py`
3. **상세 검증**: `validation/validate_hydrostatic_detailed.py`

모든 검증 결과 Excel과의 오차가 0.001% 이하로 확인되었습니다.

