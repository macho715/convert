# 사용 가이드

Vessel Stability Booklet Python 구현의 사용 방법을 상세히 설명합니다.

## 빠른 시작

### 1. 기본 사용법

```python
import sys
from pathlib import Path

# src 폴더를 경로에 추가
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from vessel_stability_functions import (
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

# 계산 실행
bg = calculator.calculate_bg(hydrostatic.lcb, hydrostatic.lcg)
print(f"BG = {bg:.6f} m")
```

### 2. 프로젝트 루트에서 실행

프로젝트 루트(`vessel_stability_python/`)에서 실행하는 경우:

```python
from src.vessel_stability_functions import StabilityCalculator, VesselParticulars

# 계산기 생성
particulars = VesselParticulars()
calculator = StabilityCalculator(particulars)

# 계산 실행
bg = calculator.calculate_bg(lcb=31.438885, lcg=31.816168)
print(f"BG = {bg:.6f} m")
```

## 상세 사용 예제

### 예제 1: 기본 계산

```python
from src.vessel_stability_functions import (
    StabilityCalculator,
    VesselParticulars
)

# 선박 제원 설정
particulars = VesselParticulars(
    length_oa=64.0,
    length_bp=60.302,
    moulded_breadth=14.0,
    moulded_depth=3.65,
    draft_loaded=2.691,
    lightship_weight=770.162,
    lightship_lcg=26.349,
    lightship_vcg=3.884
)

# 계산기 생성
calculator = StabilityCalculator(particulars)

# BG 계산
bg = calculator.calculate_bg(lcb=31.438885, lcg=31.816168)
print(f"BG = {bg:.6f} m")

# Trim 계산
trim = calculator.calculate_trim(
    displacement=1183.8462,
    bg=bg,
    mtc=33.991329
)
print(f"Trim = {trim:.6f} m")

# Lost GM 계산
lost_gm = calculator.calculate_lost_gm(
    fsm=164.76,
    displacement=1183.8462
)
print(f"Lost GM = {lost_gm:.6f} m")
```

### 예제 2: Volum 시트 계산

```python
from src.vessel_stability_functions import StabilityCalculator, VesselParticulars

calculator = StabilityCalculator(VesselParticulars())

# 탱크 데이터
volume = 2.4  # m³
density = 0.82  # t/m³
lcg = 11.251  # m
vcg = 2.825  # m
tcg = -6.247  # m
capacity = 3.5  # m³

# 중량 계산
weight = calculator.calculate_weight(volume, density)
print(f"Weight = {weight:.3f} tonnes")

# 모멘트 계산
l_moment = calculator.calculate_l_moment(weight, lcg)
v_moment = calculator.calculate_v_moment(weight, vcg)
t_moment = calculator.calculate_t_moment(weight, tcg)

print(f"L-moment = {l_moment:.6f}")
print(f"V-moment = {v_moment:.6f}")
print(f"T-moment = {t_moment:.6f}")

# 용적 비율 계산
percentage = calculator.calculate_percentage(volume, capacity)
print(f"Percentage = {percentage:.2f}%")
```

### 예제 3: 여러 탱크의 Sub Total 계산

```python
from src.vessel_stability_functions import StabilityCalculator, VesselParticulars

calculator = StabilityCalculator(VesselParticulars())

# 여러 탱크 데이터
tanks = [
    {"volume": 2.4, "density": 0.82, "lcg": 11.251, "vcg": 2.825, "tcg": -6.247, "capacity": 3.5},
    {"volume": 2.4, "density": 0.82, "lcg": 11.251, "vcg": 2.825, "tcg": 6.247, "capacity": 3.5},
    {"volume": 4.8, "density": 0.82, "lcg": 12.287, "vcg": 0.669, "tcg": 0, "capacity": 15.8},
]

weights = []
l_moments = []
v_moments = []
t_moments = []
volumes = []
capacities = []
fsm_values = []

for tank in tanks:
    weight = calculator.calculate_weight(tank["volume"], tank["density"])
    l_moment = calculator.calculate_l_moment(weight, tank["lcg"])
    v_moment = calculator.calculate_v_moment(weight, tank["vcg"])
    t_moment = calculator.calculate_t_moment(weight, tank["tcg"])
    
    weights.append(weight)
    l_moments.append(l_moment)
    v_moments.append(v_moment)
    t_moments.append(t_moment)
    volumes.append(tank["volume"])
    capacities.append(tank["capacity"])
    fsm_values.append(0)  # 예시

# Sub Total 계산
subtotal = calculator.calculate_subtotal(
    weights, l_moments, v_moments, t_moments,
    volumes, capacities, fsm_values
)

print("Sub Total:")
print(f"  Total Weight: {subtotal['total_weight']:.3f} tonnes")
print(f"  Total L-moment: {subtotal['total_l_moment']:.6f}")
print(f"  Total V-moment: {subtotal['total_v_moment']:.6f}")
print(f"  Total T-moment: {subtotal['total_t_moment']:.6f}")
```

### 예제 4: 최종 배수량 계산

```python
from src.vessel_stability_functions import StabilityCalculator, VesselParticulars

calculator = StabilityCalculator(VesselParticulars())

# 경하중량 데이터
light_ship_weight = 770.16
light_ship_lcg = 26.349
light_ship_vcg = 3.884
light_ship_tcg = -0.004

# 탱크 Sub Total 데이터
subtotal_weight = 413.6862
subtotal_l_moment = 17362.445
subtotal_v_moment = 893.524587
subtotal_t_moment = -25.398553

# 최종 배수량 계산
result = calculator.calculate_total_displacement(
    light_ship_weight, light_ship_lcg, light_ship_vcg, light_ship_tcg,
    subtotal_weight, subtotal_l_moment, subtotal_v_moment, subtotal_t_moment
)

print("최종 배수량 및 중심:")
print(f"  Displacement: {result['displacement']:.4f} tonnes")
print(f"  LCG: {result['lcg']:.6f} m")
print(f"  VCG: {result['vcg']:.6f} m")
print(f"  TCG: {result['tcg']:.6f} m")
```

### 예제 5: Hydrostatic 보간

```python
from src.vessel_stability_functions import StabilityCalculator, VesselParticulars

calculator = StabilityCalculator(VesselParticulars())

# Low Trim 데이터
low_trim_data = {
    'trim_value': 1.29,
    'disp_below': 1695.066,
    'disp_above': 1711.945,
    'draft_below': 2.54,
    'draft_above': 2.56,
    'lcf_below': 29.243,
    'lcf_above': 29.221,
    'lcb_below': 30.924,
    'lcb_above': 30.907,
    'vcb_below': 1.407,
    'vcb_above': 1.419,
    'kmt_below': 9.052,
    'kmt_above': 9.008,
    'mtc_below': 41.23,
    'mtc_above': 41.469,
    'tcp_below': 8.431,
    'tcp_above': 8.447
}

# High Trim 데이터
high_trim_data = {
    'trim_value': 2.11,
    'disp_below': 1641.88,
    'disp_above': 1658.777,
    # ... 기타 데이터
}

# 보간 실행
target_displacement = 1183.8462
target_trim = -0.1314

result = calculator.interpolate_hydrostatic_data(
    target_displacement,
    low_trim_data,
    high_trim_data,
    target_trim
)

print("보간된 수정 데이터:")
print(f"  Draft: {result['draft']:.6f} m")
print(f"  LCF: {result['lcf']:.6f} m")
print(f"  LCB: {result['lcb']:.6f} m")
print(f"  MTC: {result['mtc']:.6f} t-m")
```

### 예제 6: GZ Curve 계산

```python
from src.vessel_stability_functions import StabilityCalculator, VesselParticulars
import math

calculator = StabilityCalculator(VesselParticulars())

# GZ 데이터
gz_kn = 1.976047
vcg_corrected = 3.218307
heel_angle = 10.0

# 복원팔 계산
righting_arm = calculator.calculate_righting_arm(
    gz_kn, vcg_corrected, heel_angle
)
print(f"Righting Arm at {heel_angle}° = {righting_arm:.6f} m")

# GZ 곡선 면적 계산 (Simpson's rule)
gz_values = [0, 1.416061, 2.404653, 2.292553, 2.058209, 1.699501, 1.101626]
heel_angles = [0, 10, 20, 30, 40, 50, 60]

area = calculator.calculate_area_simpsons(gz_values, heel_angles)
print(f"GZ Curve Area = {area:.6f}")
```

### 예제 7: Excel 파일에서 데이터 로드

```python
from src.vessel_stability_functions import (
    load_excel_data,
    extract_particulars_from_sheet,
    extract_hydrostatic_from_sheet,
    StabilityCalculator
)

# Excel 파일 로드
file_path = "data/1.Vessel Stability Booklet.xls"
data = load_excel_data(file_path)

# 데이터 추출
particulars = extract_particulars_from_sheet(data.get('PRINCIPAL PARTICULARS'))
hydrostatic = extract_hydrostatic_from_sheet(data.get('Hydrostatic'))

print("선박 제원:")
print(f"  Length (O.A.): {particulars.length_oa} m")
print(f"  Length (B.P.): {particulars.length_bp} m")
print(f"  Moulded Breadth: {particulars.moulded_breadth} m")

print("\n수정 데이터:")
print(f"  Displacement: {hydrostatic.displacement:.4f} tonnes")
print(f"  LCG: {hydrostatic.lcg:.6f} m")
print(f"  LCB: {hydrostatic.lcb:.6f} m")

# 계산 실행
calculator = StabilityCalculator(particulars)
bg = calculator.calculate_bg(hydrostatic.lcb, hydrostatic.lcg)
print(f"\nBG = {bg:.6f} m")
```

## 파일별 사용 방법

### vessel_stability_functions.py (메인 구현 파일)

이 파일은 모든 Excel 함수를 구현한 핵심 파일입니다.

**주요 클래스:**
- `StabilityCalculator` - 모든 계산 함수를 포함
- `VesselParticulars` - 선박 제원 데이터 클래스
- `HydrostaticData` - 수정 데이터 클래스
- `GZData` - GZ 곡선 데이터 클래스

**주요 함수:**
- `load_excel_data()` - Excel 파일 로드
- `extract_particulars_from_sheet()` - 선박 제원 추출
- `extract_hydrostatic_from_sheet()` - 수정 데이터 추출

### analyze_excel_functions.py (분석 스크립트)

Excel 파일의 구조를 분석하는 스크립트입니다.

```bash
python src/analyze_excel_functions.py
```

이 스크립트는:
- 모든 시트 목록 출력
- 각 시트의 크기 확인
- 수식 추출 시도

### excel_to_python_stability.py (초기 버전)

초기 구현 버전으로 참고용입니다. 최신 버전은 `vessel_stability_functions.py`를 사용하세요.

## 실행 방법

### 방법 1: 프로젝트 루트에서 실행

```python
# vessel_stability_python/ 폴더에서 실행
from src.vessel_stability_functions import StabilityCalculator, VesselParticulars

calculator = StabilityCalculator(VesselParticulars())
```

### 방법 2: src 폴더를 경로에 추가

```python
import sys
from pathlib import Path

# src 폴더를 경로에 추가
sys.path.insert(0, str(Path("vessel_stability_python/src")))

from vessel_stability_functions import StabilityCalculator, VesselParticulars
```

### 방법 3: 직접 import (절대 경로)

```python
import sys
sys.path.insert(0, r"C:\Users\SAMSUNG\Downloads\CONVERT\vessel_stability_python\src")

from vessel_stability_functions import StabilityCalculator, VesselParticulars
```

## 주의사항

### 1. 경로 설정

프로젝트 루트에서 실행할 때는 `src.vessel_stability_functions`로 import하고, 다른 위치에서 실행할 때는 `sys.path`에 경로를 추가해야 합니다.

### 2. 데이터 파일 경로

Excel 파일 경로는 실행 위치에 따라 조정해야 합니다:
- 프로젝트 루트에서: `"data/1.Vessel Stability Booklet.xls"`
- 다른 위치에서: 절대 경로 사용

### 3. 의존성 패키지

다음 패키지가 필요합니다:
```bash
pip install pandas numpy openpyxl xlrd
```

## 완전한 예제 스크립트

```python
"""
Vessel Stability Calculator 사용 예제
"""

import sys
from pathlib import Path

# 프로젝트 루트를 경로에 추가
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from src.vessel_stability_functions import (
    StabilityCalculator,
    VesselParticulars,
    load_excel_data,
    extract_particulars_from_sheet,
    extract_hydrostatic_from_sheet
)

def main():
    """메인 함수"""
    # Excel 파일 로드
    file_path = "data/1.Vessel Stability Booklet.xls"
    data = load_excel_data(file_path)
    
    # 데이터 추출
    particulars = extract_particulars_from_sheet(data.get('PRINCIPAL PARTICULARS'))
    hydrostatic = extract_hydrostatic_from_sheet(data.get('Hydrostatic'))
    
    # 계산기 생성
    calculator = StabilityCalculator(particulars)
    
    # 계산 실행
    bg = calculator.calculate_bg(hydrostatic.lcb, hydrostatic.lcg)
    trim = calculator.calculate_trim(hydrostatic.displacement, bg, hydrostatic.mtc)
    lost_gm = calculator.calculate_lost_gm(hydrostatic.fsm, hydrostatic.displacement)
    
    # 결과 출력
    print("계산 결과:")
    print(f"  BG = {bg:.6f} m")
    print(f"  Trim = {trim:.6f} m")
    print(f"  Lost GM = {lost_gm:.6f} m")

if __name__ == "__main__":
    main()
```

이 스크립트를 `vessel_stability_python/` 폴더에 저장하고 실행하면 됩니다.


