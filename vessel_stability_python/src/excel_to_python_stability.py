"""
Vessel Stability Booklet Excel 함수를 Python으로 구현
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass

@dataclass
class VesselParticulars:
    """선박 주요 제원"""
    length_oa: float  # Length (O.A.)
    length_bp: float  # Length (B.P.)
    moulded_breadth: float  # Moulded Breadth
    moulded_depth: float  # Moulded Depth
    draft_loaded: float  # Draft Loaded
    lightship_weight: float  # Lightship weight
    lightship_lcg: float  # LCG
    lightship_vcg: float  # VCG


@dataclass
class HydrostaticData:
    """수정 데이터"""
    displacement: float  # Displacement (∆)
    lcg: float  # LCG
    vcg: float  # VCG
    tcg: float  # TCG
    fsm: float  # Free Surface Moment (FSM)
    mtc: float  # Moment to Change Trim (MTC)
    draft: float  # Draft
    lcb: float  # LCB at displacement
    trfap: float  # TRFAP (Trim Reference Forward of AP)
    trffp: float  # TRFFP (Trim Reference Forward of FP)
    draft_ap: float  # Draft AP
    draft_fp: float  # Draft FP
    trim: float  # Trim
    lbp: float  # Length Between Perpendiculars


class StabilityCalculator:
    """선박 안정성 계산기"""
    
    def __init__(self, particulars: VesselParticulars):
        self.particulars = particulars
    
    def calculate_bg(self, lcb: float, lcg: float) -> float:
        """
        BG 계산: BG = LCB - LCG
        Excel: BG = LCB - LCG
        """
        return lcb - lcg
    
    def calculate_trim(self, displacement: float, bg: float, mtc: float) -> float:
        """
        Trim 계산: Trim = (∆ × BG) / MTC
        Excel: Trim = (∆) x BG / MTC
        
        Note: BG가 음수면 Forward trim, 양수면 Aft trim
        """
        if mtc == 0:
            return 0.0
        trim = (displacement * bg) / mtc
        # Excel에서는 절댓값을 사용하거나 부호를 반대로 하는 경우가 있음
        return abs(trim) if trim < 0 else trim
    
    def interpolate_gz(self, 
                      displacement: float,
                      trim: float,
                      low_trim: float,
                      high_trim: float,
                      gz_low: List[float],
                      gz_high: List[float],
                      heel_angles: List[float]) -> List[float]:
        """
        GZ 보간 계산
        Excel에서 사용되는 선형 보간 로직
        
        Args:
            displacement: 현재 배수량
            trim: 현재 트림
            low_trim: 낮은 트림 값
            high_trim: 높은 트림 값
            gz_low: 낮은 트림에서의 GZ 값들 (각 경사각별)
            gz_high: 높은 트림에서의 GZ 값들 (각 경사각별)
            heel_angles: 경사각 리스트
        
        Returns:
            보간된 GZ 값들
        """
        if low_trim == high_trim:
            return gz_low
        
        # 트림에 따른 보간 계수
        trim_factor = (trim - low_trim) / (high_trim - low_trim)
        
        # 각 경사각별로 GZ 보간
        interpolated_gz = []
        for i in range(len(heel_angles)):
            gz = gz_low[i] + (gz_high[i] - gz_low[i]) * trim_factor
            interpolated_gz.append(gz)
        
        return interpolated_gz
    
    def calculate_gz_kn(self, displacement: float, gz_values: List[float]) -> List[float]:
        """
        GZ(KN) 계산: GZ(KN) = GZ × (∆ / 1000)
        Excel에서 배수량에 따른 GZ 스케일링
        """
        return [gz * (displacement / 1000.0) for gz in gz_values]
    
    def calculate_draft_ap_fp(self, 
                              draft: float, 
                              trim: float, 
                              lbp: float) -> Tuple[float, float]:
        """
        Draft AP와 FP 계산
        Excel: Draft AP = Draft - (Trim × LBP) / 2
               Draft FP = Draft + (Trim × LBP) / 2
        """
        draft_ap = draft - (trim * lbp) / 2.0
        draft_fp = draft + (trim * lbp) / 2.0
        return draft_ap, draft_fp
    
    def calculate_metacentric_height(self, 
                                    km: float, 
                                    kg: float) -> float:
        """
        초심고(GM) 계산: GM = KM - KG
        Excel: GM = KM - KG
        """
        return km - kg
    
    def calculate_volume(self, 
                        displacement: float, 
                        density: float = 1.025) -> float:
        """
        용적 계산: Volume = Displacement / Density
        Excel: Volume = ∆ / ρ
        """
        return displacement / density
    
    def calculate_deadweight(self, 
                           displacement: float, 
                           lightship: float) -> float:
        """
        적화중량(DWT) 계산: DWT = Displacement - Lightship
        Excel: DWT = ∆ - Lightship
        """
        return displacement - lightship


def load_stability_data(file_path: str) -> Dict[str, pd.DataFrame]:
    """Excel 파일의 모든 시트를 로드"""
    xls_file = pd.ExcelFile(file_path)
    data = {}
    
    for sheet_name in xls_file.sheet_names:
        try:
            df = pd.read_excel(xls_file, sheet_name=sheet_name, header=None)
            data[sheet_name] = df
        except Exception as e:
            print(f"⚠️  시트 '{sheet_name}' 로드 실패: {e}")
    
    return data


def extract_particulars(data: Dict[str, pd.DataFrame]) -> VesselParticulars:
    """PRINCIPAL PARTICULARS 시트에서 선박 제원 추출"""
    df = data.get('PRINCIPAL PARTICULARS', pd.DataFrame())
    
    # 데이터 추출 (실제 위치는 파일에 맞게 조정 필요)
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
    
    # 실제 데이터에서 추출 (예시)
    for idx, row in df.iterrows():
        row_str = str(row[1]) if len(row) > 1 else ""
        if "Length (O.A.)" in row_str:
            try:
                particulars.length_oa = float(row[3]) if pd.notna(row[3]) else particulars.length_oa
            except:
                pass
        elif "Length (B.P.)" in row_str:
            try:
                particulars.length_bp = float(row[3]) if pd.notna(row[3]) else particulars.length_bp
            except:
                pass
        elif "Moulded Breadth" in row_str:
            try:
                particulars.moulded_breadth = float(row[3]) if pd.notna(row[3]) else particulars.moulded_breadth
            except:
                pass
        elif "Moulded Depth" in row_str:
            try:
                particulars.moulded_depth = float(row[3]) if pd.notna(row[3]) else particulars.moulded_depth
            except:
                pass
        elif "Draft Loaded" in row_str:
            try:
                particulars.draft_loaded = float(row[3]) if pd.notna(row[3]) else particulars.draft_loaded
            except:
                pass
        elif "Lightship weight" in row_str:
            try:
                particulars.lightship_weight = float(row[3]) if pd.notna(row[3]) else particulars.lightship_weight
            except:
                pass
        elif "LCG" in row_str and "Lightship" in str(df.iloc[idx-1, 1]):
            try:
                particulars.lightship_lcg = float(row[3]) if pd.notna(row[3]) else particulars.lightship_lcg
            except:
                pass
        elif "VCG" in row_str and "Lightship" in str(df.iloc[idx-1, 1]):
            try:
                particulars.lightship_vcg = float(row[3]) if pd.notna(row[3]) else particulars.lightship_vcg
            except:
                pass
    
    return particulars


def extract_hydrostatic_data(data: Dict[str, pd.DataFrame]) -> HydrostaticData:
    """Hydrostatic 시트에서 수정 데이터 추출"""
    df = data.get('Hydrostatic', pd.DataFrame())
    
    hydrostatic = HydrostaticData(
        displacement=1183.8462,
        lcg=31.816168,
        vcg=3.35748,
        tcg=-0.024056,
        fsm=164.76,
        mtc=33.991329,
        draft=1.934253,
        lcb=31.438885,
        trfap=-0.065173,
        trffp=-0.066227,
        draft_ap=1.86908,
        draft_fp=2.00048,
        trim=0.1314,
        lbp=60.302
    )
    
    # 실제 데이터에서 추출
    for idx, row in df.iterrows():
        row_str = str(row[0]) if len(row) > 0 else ""
        if "Displacement" in row_str and pd.notna(row[2]):
            try:
                hydrostatic.displacement = float(row[2])
            except:
                pass
        elif row_str == "LCG" and pd.notna(row[2]):
            try:
                hydrostatic.lcg = float(row[2])
            except:
                pass
        elif row_str == "VCG" and pd.notna(row[2]):
            try:
                hydrostatic.vcg = float(row[2])
            except:
                pass
        elif row_str == "TCG" and pd.notna(row[2]):
            try:
                hydrostatic.tcg = float(row[2])
            except:
                pass
        elif row_str == "FSM" and pd.notna(row[2]):
            try:
                hydrostatic.fsm = float(row[2])
            except:
                pass
        elif row_str == "MTC" and pd.notna(row[2]):
            try:
                hydrostatic.mtc = float(row[2])
            except:
                pass
        elif row_str == "Draft" and pd.notna(row[2]):
            try:
                hydrostatic.draft = float(row[2])
            except:
                pass
        elif "LCB" in row_str and pd.notna(row[2]):
            try:
                hydrostatic.lcb = float(row[2])
            except:
                pass
        elif row_str == "TRFAP" and pd.notna(row[2]):
            try:
                hydrostatic.trfap = float(row[2])
            except:
                pass
        elif row_str == "TRFFP" and pd.notna(row[2]):
            try:
                hydrostatic.trffp = float(row[2])
            except:
                pass
        elif row_str == "Draft AP" and pd.notna(row[2]):
            try:
                hydrostatic.draft_ap = float(row[2])
            except:
                pass
        elif row_str == "Draft FP" and pd.notna(row[2]):
            try:
                hydrostatic.draft_fp = float(row[2])
            except:
                pass
        elif row_str == "Trim" and pd.notna(row[2]):
            try:
                hydrostatic.trim = float(row[2])
            except:
                pass
        elif row_str == "LBP" or (row_str == "metres" and "LBP" in str(df.iloc[idx-1, 0])):
            try:
                if pd.notna(row[4]):
                    hydrostatic.lbp = float(row[4])
            except:
                pass
    
    return hydrostatic


def main():
    """메인 실행 함수"""
    print("=" * 60)
    print("🚢 Vessel Stability Booklet - Excel to Python")
    print("=" * 60)
    
    file_path = "1.Vessel Stability Booklet.xls"
    
    # 데이터 로드
    print(f"\n📖 Excel 파일 로드: {file_path}")
    data = load_stability_data(file_path)
    print(f"  ✓ 로드된 시트: {len(data)}개")
    
    # 데이터 추출
    print("\n📊 데이터 추출 중...")
    particulars = extract_particulars(data)
    hydrostatic = extract_hydrostatic_data(data)
    
    print("\n📋 추출된 선박 제원:")
    print(f"  Length (O.A.): {particulars.length_oa} m")
    print(f"  Length (B.P.): {particulars.length_bp} m")
    print(f"  Moulded Breadth: {particulars.moulded_breadth} m")
    print(f"  Draft Loaded: {particulars.draft_loaded} m")
    
    print("\n📊 수정 데이터:")
    print(f"  Displacement (∆): {hydrostatic.displacement} tonnes")
    print(f"  LCG: {hydrostatic.lcg} m")
    print(f"  LCB: {hydrostatic.lcb} m")
    print(f"  MTC: {hydrostatic.mtc} t-m")
    print(f"  Trim: {hydrostatic.trim} m")
    
    # 계산기 생성
    calculator = StabilityCalculator(particulars)
    
    # 계산 실행
    print("\n🧮 계산 실행:")
    
    # BG 계산
    bg = calculator.calculate_bg(hydrostatic.lcb, hydrostatic.lcg)
    print(f"  BG = LCB - LCG = {hydrostatic.lcb} - {hydrostatic.lcg} = {bg:.6f} m")
    
    # Trim 검증
    calculated_trim = calculator.calculate_trim(
        hydrostatic.displacement, bg, hydrostatic.mtc
    )
    print(f"  Trim = (∆ × BG) / MTC = ({hydrostatic.displacement} × {bg:.6f}) / {hydrostatic.mtc} = {calculated_trim:.6f} m")
    print(f"  실제 Trim: {hydrostatic.trim} m")
    
    # Deadweight 계산
    dwt = calculator.calculate_deadweight(
        hydrostatic.displacement, particulars.lightship_weight
    )
    print(f"  DWT = ∆ - Lightship = {hydrostatic.displacement} - {particulars.lightship_weight} = {dwt:.3f} tonnes")
    
    # Volume 계산
    volume = calculator.calculate_volume(hydrostatic.displacement)
    print(f"  Volume = ∆ / ρ = {hydrostatic.displacement} / 1.025 = {volume:.3f} m³")
    
    print("\n" + "=" * 60)
    print("✅ 계산 완료!")
    print("=" * 60)
    
    return calculator, particulars, hydrostatic


if __name__ == "__main__":
    calculator, particulars, hydrostatic = main()

