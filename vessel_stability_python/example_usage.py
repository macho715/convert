"""
Vessel Stability Calculator 사용 예제
프로젝트 루트(vessel_stability_python/)에서 실행
"""

from src.vessel_stability_functions import (
    StabilityCalculator,
    VesselParticulars,
    load_excel_data,
    extract_particulars_from_sheet,
    extract_hydrostatic_from_sheet
)


def main():
    """메인 예제 함수"""
    print("=" * 60)
    print("Vessel Stability Calculator - 사용 예제")
    print("=" * 60)
    
    # 방법 1: Excel 파일에서 데이터 로드
    print("\n[방법 1] Excel 파일에서 데이터 로드")
    print("-" * 60)
    
    try:
        file_path = "data/1.Vessel Stability Booklet.xls"
        data = load_excel_data(file_path)
        
        particulars = extract_particulars_from_sheet(data.get('PRINCIPAL PARTICULARS'))
        hydrostatic = extract_hydrostatic_from_sheet(data.get('Hydrostatic'))
        
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
        
    except FileNotFoundError:
        print("⚠️  Excel 파일을 찾을 수 없습니다. 방법 2를 사용하세요.")
    
    # 방법 2: 직접 값 입력
    print("\n[방법 2] 직접 값 입력")
    print("-" * 60)
    
    calculator = StabilityCalculator(VesselParticulars())
    
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
    lost_gm = calculator.calculate_lost_gm(fsm=164.76, displacement=1183.8462)
    print(f"Lost GM = {lost_gm:.6f} m")
    
    # VCG Corrected 계산
    vcg_corrected = calculator.calculate_vcg_corrected(
        vcg=3.35748,
        fsm=164.76,
        displacement=1183.8462
    )
    print(f"VCG Corrected = {vcg_corrected:.6f} m")
    
    # Volum 시트 계산 예제
    print("\n[방법 3] Volum 시트 계산 예제")
    print("-" * 60)
    
    volume = 2.4
    density = 0.82
    weight = calculator.calculate_weight(volume, density)
    print(f"Weight = {weight:.3f} tonnes (Volume={volume} m³ × Density={density} t/m³)")
    
    l_moment = calculator.calculate_l_moment(weight, lcg=11.251)
    print(f"L-moment = {l_moment:.6f} (Weight={weight} × LCG=11.251)")
    
    v_moment = calculator.calculate_v_moment(weight, vcg=2.825)
    print(f"V-moment = {v_moment:.6f} (Weight={weight} × VCG=2.825)")
    
    print("\n" + "=" * 60)
    print("✅ 예제 실행 완료!")
    print("=" * 60)
    print("\n더 자세한 사용법은 docs/USAGE_GUIDE.md를 참고하세요.")


if __name__ == "__main__":
    main()


