# 함수 참조 문서

Vessel Stability Booklet Python 구현의 모든 함수에 대한 API 문서입니다.

## StabilityCalculator 클래스

모든 Excel 함수를 구현한 메인 계산기 클래스입니다.

### 초기화

```python
calculator = StabilityCalculator(particulars: VesselParticulars)
```

**파라미터:**
- `particulars` (VesselParticulars): 선박 주요 제원

---

## 기본 계산 함수

### calculate_bg

BG (Longitudinal Center of Buoyancy - Center of Gravity) 계산

```python
def calculate_bg(self, lcb: float, lcg: float) -> float
```

**파라미터:**
- `lcb` (float): 종향 부력 중심 (m)
- `lcg` (float): 종향 무게 중심 (m)

**반환값:**
- `float`: BG 값 (m)

**Excel 수식:** `=LCB - LCG`

**예제:**
```python
bg = calculator.calculate_bg(lcb=31.438885, lcg=31.816168)
# 결과: -0.377284 m
```

---

### calculate_trim

Trim 계산

```python
def calculate_trim(self, displacement: float, bg: float, mtc: float) -> float
```

**파라미터:**
- `displacement` (float): 배수량 (tonnes)
- `bg` (float): BG 값 (m)
- `mtc` (float): Moment to Change Trim (t-m)

**반환값:**
- `float`: Trim 값 (m, 절댓값)

**Excel 수식:** `=(∆ × |BG|) / MTC`

**예제:**
```python
trim = calculator.calculate_trim(
    displacement=1183.8462,
    bg=-0.377284,
    mtc=33.991329
)
# 결과: 0.1314 m
```

---

### calculate_metacentric_height

초심고(GM) 계산

```python
def calculate_metacentric_height(self, km: float, kg: float) -> float
```

**파라미터:**
- `km` (float): 횡향 초심고 (m)
- `kg` (float): 무게 중심 높이 (m)

**반환값:**
- `float`: GM 값 (m)

**Excel 수식:** `=KM - KG`

**예제:**
```python
gm = calculator.calculate_metacentric_height(km=10.384642, kg=3.35748)
# 결과: 7.027162 m
```

---

### calculate_volume

용적 계산

```python
def calculate_volume(self, displacement: float, density: float = 1.025) -> float
```

**파라미터:**
- `displacement` (float): 배수량 (tonnes)
- `density` (float, optional): 밀도 (t/m³), 기본값 1.025

**반환값:**
- `float`: 용적 (m³)

**Excel 수식:** `=Displacement / Density`

**예제:**
```python
volume = calculator.calculate_volume(displacement=1183.8462)
# 결과: 1154.972 m³
```

---

### calculate_deadweight

적화중량(DWT) 계산

```python
def calculate_deadweight(self, displacement: float, lightship: float) -> float
```

**파라미터:**
- `displacement` (float): 배수량 (tonnes)
- `lightship` (float): 경하중량 (tonnes)

**반환값:**
- `float`: 적화중량 (tonnes)

**Excel 수식:** `=Displacement - Lightship`

**예제:**
```python
dwt = calculator.calculate_deadweight(displacement=1183.8462, lightship=770.162)
# 결과: 413.684 tonnes
```

---

## Volum 시트 함수

### calculate_weight

중량 계산

```python
def calculate_weight(self, volume: float, density: float) -> float
```

**파라미터:**
- `volume` (float): 용적 (m³)
- `density` (float): 밀도 (t/m³)

**반환값:**
- `float`: 중량 (tonnes)

**Excel 수식:** `=Volume × Density`

**예제:**
```python
weight = calculator.calculate_weight(volume=2.4, density=0.82)
# 결과: 1.968 tonnes
```

---

### calculate_l_moment

종향 모멘트 계산

```python
def calculate_l_moment(self, weight: float, lcg: float) -> float
```

**파라미터:**
- `weight` (float): 중량 (tonnes)
- `lcg` (float): 종향 무게 중심 (m)

**반환값:**
- `float`: 종향 모멘트 (tonne-metres)

**Excel 수식:** `=Weight × LCG`

**예제:**
```python
l_moment = calculator.calculate_l_moment(weight=1.968, lcg=11.251)
# 결과: 22.141968 tonne-metres
```

---

### calculate_v_moment

수직 모멘트 계산

```python
def calculate_v_moment(self, weight: float, vcg: float) -> float
```

**파라미터:**
- `weight` (float): 중량 (tonnes)
- `vcg` (float): 수직 무게 중심 (m)

**반환값:**
- `float`: 수직 모멘트 (tonne-metres)

**Excel 수식:** `=Weight × VCG`

**예제:**
```python
v_moment = calculator.calculate_v_moment(weight=1.968, vcg=2.825)
# 결과: 5.5596 tonne-metres
```

---

### calculate_t_moment

횡향 모멘트 계산

```python
def calculate_t_moment(self, weight: float, tcg: float) -> float
```

**파라미터:**
- `weight` (float): 중량 (tonnes)
- `tcg` (float): 횡향 무게 중심 (m)

**반환값:**
- `float`: 횡향 모멘트 (tonne-metres)

**Excel 수식:** `=Weight × TCG`

**예제:**
```python
t_moment = calculator.calculate_t_moment(weight=1.968, tcg=-6.247)
# 결과: -12.294096 tonne-metres
```

---

### calculate_percentage

용적 비율 계산

```python
def calculate_percentage(self, volume: float, capacity: float) -> float
```

**파라미터:**
- `volume` (float): 용적 (m³)
- `capacity` (float): 용량 (m³)

**반환값:**
- `float`: 비율 (%)

**Excel 수식:** `=(Volume / Cap) × 100`

**예제:**
```python
percentage = calculator.calculate_percentage(volume=2.4, capacity=3.5)
# 결과: 68.5714%
```

---

### calculate_subtotal

Sub Total 계산

```python
def calculate_subtotal(self,
                      weights: List[float],
                      l_moments: List[float],
                      v_moments: List[float],
                      t_moments: List[float],
                      volumes: List[float],
                      capacities: List[float],
                      fsm_values: List[float]) -> Dict[str, float]
```

**파라미터:**
- `weights` (List[float]): 중량 리스트
- `l_moments` (List[float]): 종향 모멘트 리스트
- `v_moments` (List[float]): 수직 모멘트 리스트
- `t_moments` (List[float]): 횡향 모멘트 리스트
- `volumes` (List[float]): 용적 리스트
- `capacities` (List[float]): 용량 리스트
- `fsm_values` (List[float]): FSM 리스트

**반환값:**
- `Dict[str, float]`: Sub Total 딕셔너리
  - `total_volume`: 총 용적
  - `total_capacity`: 총 용량
  - `total_weight`: 총 중량
  - `total_l_moment`: 총 종향 모멘트
  - `total_v_moment`: 총 수직 모멘트
  - `total_t_moment`: 총 횡향 모멘트
  - `total_fsm`: 총 FSM

---

### calculate_total_displacement

최종 배수량 및 중심 계산

```python
def calculate_total_displacement(self,
                                light_ship_weight: float,
                                light_ship_lcg: float,
                                light_ship_vcg: float,
                                light_ship_tcg: float,
                                subtotal_weight: float,
                                subtotal_l_moment: float,
                                subtotal_v_moment: float,
                                subtotal_t_moment: float) -> Dict[str, float]
```

**파라미터:**
- `light_ship_weight` (float): 경하중량 (tonnes)
- `light_ship_lcg` (float): 경하 LCG (m)
- `light_ship_vcg` (float): 경하 VCG (m)
- `light_ship_tcg` (float): 경하 TCG (m)
- `subtotal_weight` (float): 탱크 중량 합계 (tonnes)
- `subtotal_l_moment` (float): 탱크 종향 모멘트 합계
- `subtotal_v_moment` (float): 탱크 수직 모멘트 합계
- `subtotal_t_moment` (float): 탱크 횡향 모멘트 합계

**반환값:**
- `Dict[str, float]`: 최종 배수량 및 중심
  - `displacement`: 배수량 (tonnes)
  - `lcg`: 종향 무게 중심 (m)
  - `vcg`: 수직 무게 중심 (m)
  - `tcg`: 횡향 무게 중심 (m)

---

## Hydrostatic 시트 함수

### calculate_diff

차이 계산

```python
def calculate_diff(self, above_value: float, below_value: float) -> float
```

**파라미터:**
- `above_value` (float): Above 값
- `below_value` (float): Below 값

**반환값:**
- `float`: 차이 (Above - Below)

**Excel 수식:** `=Above - Below`

---

### calculate_interpolation_factor

보간 계수 계산

```python
def calculate_interpolation_factor(self,
                                 target_value: float,
                                 low_value: float,
                                 high_value: float) -> float
```

**파라미터:**
- `target_value` (float): 목표 값
- `low_value` (float): 낮은 값
- `high_value` (float): 높은 값

**반환값:**
- `float`: 보간 계수 (0~1 사이, 범위 밖일 수 있음)

**Excel 수식:** `=(Target - Low) / (High - Low)`

---

### interpolate_hydrostatic_data

Hydrostatic 데이터 보간

```python
def interpolate_hydrostatic_data(self,
                                displacement: float,
                                low_trim_data: Dict[str, float],
                                high_trim_data: Dict[str, float],
                                target_trim: float) -> Dict[str, float]
```

**파라미터:**
- `displacement` (float): 목표 배수량
- `low_trim_data` (Dict): 낮은 트림 데이터
- `high_trim_data` (Dict): 높은 트림 데이터
- `target_trim` (float): 목표 트림

**반환값:**
- `Dict[str, float]`: 보간된 수정 데이터

---

### calculate_lost_gm

Lost GM 계산

```python
def calculate_lost_gm(self, fsm: float, displacement: float) -> float
```

**파라미터:**
- `fsm` (float): Free Surface Moment
- `displacement` (float): 배수량 (tonnes)

**반환값:**
- `float`: Lost GM (m)

**Excel 수식:** `=FSM / ∆`

**예제:**
```python
lost_gm = calculator.calculate_lost_gm(fsm=164.76, displacement=1183.8462)
# 결과: 0.139173 m
```

---

### calculate_vcg_corrected

FSM 보정된 VCG 계산

```python
def calculate_vcg_corrected(self,
                           vcg: float,
                           fsm: float,
                           displacement: float) -> float
```

**파라미터:**
- `vcg` (float): VCG (m)
- `fsm` (float): Free Surface Moment
- `displacement` (float): 배수량 (tonnes)

**반환값:**
- `float`: VCG Corrected (m)

**Excel 수식:** `=VCG + (FSM / Displacement)`

---

### calculate_tan_list

Tan List 계산

```python
def calculate_tan_list(self,
                      list_moment: float,
                      displacement: float,
                      gm: float) -> float
```

**파라미터:**
- `list_moment` (float): List Moment
- `displacement` (float): 배수량 (tonnes)
- `gm` (float): GM (m)

**반환값:**
- `float`: Tan List

**Excel 수식:** `=List Moment / (Displacement × GM)`

---

## GZ Curve 시트 함수

### calculate_righting_arm

복원팔 계산

```python
def calculate_righting_arm(self,
                          gz_kn: float,
                          vcg_corrected: float,
                          heel_angle_deg: float) -> float
```

**파라미터:**
- `gz_kn` (float): GZ(KN) 값
- `vcg_corrected` (float): FSM 보정된 VCG (KG)
- `heel_angle_deg` (float): 경사각 (도)

**반환값:**
- `float`: 복원팔 (Righting Arm)

**Excel 수식:** `=GZ(KN) - KG(corrected VCG) × Sin(Heel)`

---

### calculate_area_simpsons

Simpson's rule로 GZ 곡선 아래 면적 계산

```python
def calculate_area_simpsons(self,
                           gz_values: List[float],
                           heel_angles: List[float]) -> float
```

**파라미터:**
- `gz_values` (List[float]): GZ 값 리스트
- `heel_angles` (List[float]): 경사각 리스트 (도)

**반환값:**
- `float`: 면적 (GZ 곡선 아래 면적)

---

### interpolate_gz_complete

완전한 GZ 보간 로직

```python
def interpolate_gz_complete(self,
                           target_displacement: float,
                           target_trim: float,
                           low_trim: float,
                           high_trim: float,
                           low_trim_gz_below: List[float],
                           low_trim_gz_above: List[float],
                           high_trim_gz_below: List[float],
                           high_trim_gz_above: List[float],
                           low_trim_disp_below: float,
                           low_trim_disp_above: float,
                           high_trim_disp_below: float,
                           high_trim_disp_above: float,
                           heel_angles: List[float]) -> List[float]
```

**파라미터:**
- 배수량과 트림에 따른 복합 보간 파라미터들

**반환값:**
- `List[float]`: 최종 보간된 GZ(KN) 값들

---

## 데이터 클래스

### VesselParticulars

선박 주요 제원

```python
@dataclass
class VesselParticulars:
    length_oa: float           # Length (O.A.) (m)
    length_bp: float           # Length (B.P.) (m)
    moulded_breadth: float     # Moulded Breadth (m)
    moulded_depth: float       # Moulded Depth (m)
    draft_loaded: float        # Draft Loaded (m)
    lightship_weight: float    # Lightship weight (tonnes)
    lightship_lcg: float       # LCG (m)
    lightship_vcg: float       # VCG (m)
```

### HydrostaticData

수정 데이터

```python
@dataclass
class HydrostaticData:
    displacement: float    # 배수량 (tonnes)
    lcg: float            # LCG (m)
    vcg: float            # VCG (m)
    tcg: float            # TCG (m)
    fsm: float            # Free Surface Moment
    mtc: float            # Moment to Change Trim (t-m)
    draft: float          # Draft (m)
    lcb: float            # LCB (m)
    trim: float           # Trim (m)
    lbp: float            # Length Between Perpendiculars (m)
    # ... 기타 필드
```

### GZData

GZ 곡선 데이터

```python
@dataclass
class GZData:
    heel_angles: List[float] = [0, 10, 20, 30, 40, 50, 60]
    low_trim: float
    high_trim: float
    gz_low_below: List[float]
    gz_low_above: List[float]
    gz_high_below: List[float]
    gz_high_above: List[float]
```

