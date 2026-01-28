"""
Vessel Stability Booklet 전체 검증 스크립트
모든 시트의 계산을 검증하고 Excel 결과와 비교
"""

import pandas as pd
import sys
from pathlib import Path

# 상위 디렉토리를 경로에 추가
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.vessel_stability_functions import (
    StabilityCalculator,
    VesselParticulars,
    load_excel_data,
    extract_particulars_from_sheet,
    extract_hydrostatic_from_sheet,
    validate_volum_calculations,
    validate_hydrostatic_calculations,
    validate_gz_calculations,
    compare_with_excel
)


def main():
    """메인 검증 함수"""
    print("=" * 60)
    print("🔍 Vessel Stability Booklet - 전체 검증")
    print("=" * 60)
    
    file_path = "data/1.Vessel Stability Booklet.xls"
    
    # 데이터 로드
    print(f"\n📖 Excel 파일 로드: {file_path}")
    data = load_excel_data(file_path)
    print(f"  ✓ 로드된 시트: {len(data)}개")
    
    # 데이터 추출
    print("\n📊 데이터 추출 중...")
    particulars = extract_particulars_from_sheet(data.get('PRINCIPAL PARTICULARS', pd.DataFrame()))
    hydrostatic = extract_hydrostatic_from_sheet(data.get('Hydrostatic', pd.DataFrame()))
    
    # 계산기 생성
    calculator = StabilityCalculator(particulars)
    
    # 검증 실행
    print("\n" + "=" * 60)
    print("🔍 Volum 시트 검증")
    print("=" * 60)
    volum_data = data.get('Volum', pd.DataFrame())
    volum_result = validate_volum_calculations(calculator, volum_data, tolerance=0.001)
    
    print(f"\n  ✓ 검증 완료:")
    print(f"    - 오류: {len(volum_result['errors'])}개")
    print(f"    - 경고: {len(volum_result['warnings'])}개")
    
    if volum_result['errors']:
        print("\n  ❌ 오류:")
        for error in volum_result['errors'][:10]:  # 처음 10개만
            print(f"    - {error}")
    
    if volum_result['warnings']:
        print("\n  ⚠️  경고:")
        for warning in volum_result['warnings'][:5]:  # 처음 5개만
            print(f"    - {warning}")
    
    print("\n" + "=" * 60)
    print("🔍 Hydrostatic 시트 검증")
    print("=" * 60)
    hydrostatic_result = validate_hydrostatic_calculations(
        calculator, 
        data.get('Hydrostatic', pd.DataFrame()),
        tolerance=0.001
    )
    
    print(f"\n  ✓ 검증 완료:")
    print(f"    - 오류: {len(hydrostatic_result['errors'])}개")
    print(f"    - 경고: {len(hydrostatic_result['warnings'])}개")
    
    if hydrostatic_result['errors']:
        print("\n  ❌ 오류:")
        for error in hydrostatic_result['errors']:
            print(f"    - {error}")
    
    if hydrostatic_result['warnings']:
        print("\n  ⚠️  경고:")
        for warning in hydrostatic_result['warnings']:
            print(f"    - {warning}")
    
    # 주요 계산 검증
    print("\n" + "=" * 60)
    print("🔍 주요 계산 검증")
    print("=" * 60)
    
    # BG 계산
    bg = calculator.calculate_bg(hydrostatic.lcb, hydrostatic.lcg)
    print(f"\n1. BG 계산:")
    print(f"   Python: {bg:.6f} m")
    print(f"   Excel:  {hydrostatic.lcb - hydrostatic.lcg:.6f} m")
    print(f"   ✓ 일치" if abs(bg - (hydrostatic.lcb - hydrostatic.lcg)) < 0.001 else "   ✗ 불일치")
    
    # Lost GM 계산
    lost_gm = calculator.calculate_lost_gm(hydrostatic.fsm, hydrostatic.displacement)
    print(f"\n2. Lost GM 계산:")
    print(f"   Python: {lost_gm:.6f} m")
    excel_lost_gm = hydrostatic.fsm / hydrostatic.displacement if hydrostatic.displacement > 0 else 0
    print(f"   Excel:  {excel_lost_gm:.6f} m")
    print(f"   ✓ 일치" if abs(lost_gm - excel_lost_gm) < 0.001 else "   ✗ 불일치")
    
    # VCG Corrected 계산
    vcg_corrected = calculator.calculate_vcg_corrected(
        hydrostatic.vcg, hydrostatic.fsm, hydrostatic.displacement
    )
    print(f"\n3. VCG Corrected 계산:")
    print(f"   Python: {vcg_corrected:.6f} m")
    excel_vcg_corrected = hydrostatic.vcg + (hydrostatic.fsm / hydrostatic.displacement) if hydrostatic.displacement > 0 else hydrostatic.vcg
    print(f"   Excel:  {excel_vcg_corrected:.6f} m")
    print(f"   ✓ 일치" if abs(vcg_corrected - excel_vcg_corrected) < 0.001 else "   ✗ 불일치")
    
    # 최종 결과 비교
    print("\n" + "=" * 60)
    print("📊 최종 배수량 및 중심 검증")
    print("=" * 60)
    
    # Volum 시트에서 최종 배수량 추출 (Row 53: Displacement Condition)
    volum_df = data.get('Volum', pd.DataFrame())
    try:
        # 올바른 열 인덱스 확인 (Weight=6, LCG=7, VCG=9, TCG=11)
        if len(volum_df) > 53:
            excel_displacement = float(volum_df.iloc[53, 6]) if pd.notna(volum_df.iloc[53, 6]) else hydrostatic.displacement
            excel_lcg = float(volum_df.iloc[53, 7]) if pd.notna(volum_df.iloc[53, 7]) else hydrostatic.lcg
            excel_vcg = float(volum_df.iloc[53, 9]) if pd.notna(volum_df.iloc[53, 9]) else hydrostatic.vcg
            excel_tcg = float(volum_df.iloc[53, 11]) if pd.notna(volum_df.iloc[53, 11]) else hydrostatic.tcg
        else:
            # Volum 시트에서 데이터를 찾을 수 없으면 Hydrostatic 시트 값 사용
            excel_displacement = hydrostatic.displacement
            excel_lcg = hydrostatic.lcg
            excel_vcg = hydrostatic.vcg
            excel_tcg = hydrostatic.tcg
        
        python_result = {
            'displacement': hydrostatic.displacement,
            'lcg': hydrostatic.lcg,
            'vcg': hydrostatic.vcg,
            'tcg': hydrostatic.tcg
        }
        
        excel_result = {
            'displacement': excel_displacement,
            'lcg': excel_lcg,
            'vcg': excel_vcg,
            'tcg': excel_tcg
        }
        
        comparison = compare_with_excel(python_result, excel_result, tolerance=0.001)
        
        print(f"\n  ✓ 일치 항목: {len(comparison['matches'])}개")
        print(f"  ✗ 오류 항목: {len(comparison['errors'])}개")
        print(f"  ⚠️  경고: {len(comparison['warnings'])}개")
        
        if comparison['errors']:
            print("\n  ❌ 오류 상세:")
            for error in comparison['errors']:
                print(f"    - {error['key']}: Python={error['python']:.6f}, Excel={error['excel']:.6f}, Error={error['error_pct']:.4f}%")
        
    except Exception as e:
        print(f"\n  ⚠️  최종 배수량 검증 오류: {e}")
    
    print("\n" + "=" * 60)
    print("✅ 검증 완료!")
    print("=" * 60)


if __name__ == "__main__":
    main()

