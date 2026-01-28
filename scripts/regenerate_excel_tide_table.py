"""
조석표 Excel 파일을 올바른 구조로 재생성
구조: 행(세로) = 날짜, 열(가로) = 시간대
"""

import pandas as pd
from pathlib import Path
from datetime import datetime

def regenerate_excel_file(excel_path: str):
    """Excel 파일을 읽어서 올바른 구조로 재생성"""
    print("=" * 60)
    print("🔄 Excel 파일 재생성")
    print("=" * 60)
    
    # Excel 파일 읽기
    print(f"\n📖 Excel 파일 읽기: {excel_path}")
    df = pd.read_excel(excel_path, index_col=0)
    
    print(f"  ✓ 원본 데이터 크기: {df.shape[0]}행 x {df.shape[1]}열")
    print(f"  ✓ 인덱스 (첫 3개): {list(df.index[:3])}")
    print(f"  ✓ 컬럼 (첫 3개): {list(df.columns[:3])}")
    
    # 구조 확인 및 전치
    first_index = str(df.index[0])
    first_col = str(df.columns[0])
    
    # 날짜 형식인지 확인 (Timestamp, Dec, 날짜 형식 등)
    is_date_index = any([
        'Dec' in first_index,
        '2025' in first_index,
        isinstance(df.index[0], pd.Timestamp),
        '01-' in first_index or '02-' in first_index
    ])
    
    # 시간 형식인지 확인
    is_time_col = any([
        ':' in first_col,
        '0:00' in first_col,
        '00:00' in first_col
    ])
    
    if is_date_index and is_time_col:
        # 올바른 구조: 행=날짜, 열=시간대
        print(f"\n  ✓ 구조 확인: 행=날짜, 열=시간대 (전치 불필요)")
    elif ':' in first_index or '0:00' in first_index:
        # 잘못된 구조: 행=시간대, 열=날짜 → 전치 필요
        print(f"\n  ⚠️  전치 필요: 현재 구조는 행=시간대, 열=날짜")
        df = df.T
        print(f"  ✓ 전치 완료")
    else:
        print(f"\n  ✓ 구조 확인: 행=날짜, 열=시간대 (전치 불필요)")
    
    # 최종 구조 확인
    print(f"\n📊 최종 데이터 구조:")
    print(f"  ✓ 행(세로): 날짜 - {len(df)}개 ({df.index[0]} ~ {df.index[-1]})")
    print(f"  ✓ 열(가로): 시간대 - {len(df.columns)}개 ({df.columns[0]} ~ {df.columns[-1]})")
    
    # Excel 파일 다시 저장
    output_path = Path(excel_path)
    print(f"\n💾 Excel 파일 저장: {output_path}")
    
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 메인 시트: 조석표 데이터
            df.to_excel(writer, sheet_name='Tide_Table', index=True)
            
            # 추가 시트: 요약 통계
            summary_data = {
                '항목': ['총 날짜', '총 시간대', '최고 조석 (m)', '최저 조석 (m)', '평균 조석 (m)'],
                '값': [
                    len(df),
                    len(df.columns),
                    f"{df.max().max():.2f}",
                    f"{df.min().min():.2f}",
                    f"{df.mean().mean():.2f}"
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            # 추가 시트: 일별 요약 (최고/최저 조석 시간)
            daily_summary = []
            for date in df.index:
                row = df.loc[date]
                max_val = row.max()
                min_val = row.min()
                max_time = row.idxmax()
                min_time = row.idxmin()
                daily_summary.append({
                    '날짜': date,
                    '최고 조석 (m)': f"{max_val:.2f}",
                    '최고 조석 시간': max_time,
                    '최저 조석 (m)': f"{min_val:.2f}",
                    '최저 조석 시간': min_time
                })
            daily_summary_df = pd.DataFrame(daily_summary)
            daily_summary_df.to_excel(writer, sheet_name='Daily_Summary', index=False)
        
        print(f"  ✓ 저장 완료!")
        print(f"\n📋 생성된 시트:")
        print(f"  - Tide_Table: 메인 조석표 데이터")
        print(f"  - Summary: 전체 요약 통계")
        print(f"  - Daily_Summary: 일별 최고/최저 조석 정보")
        
    except PermissionError:
        print(f"\n❌ 오류: Excel 파일이 열려있어 저장할 수 없습니다.")
        print(f"   파일을 닫고 다시 실행하세요.")
        return 1
    except Exception as e:
        print(f"\n❌ 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return 1
    
    print("\n" + "=" * 60)
    print("✅ 완료!")
    print("=" * 60)
    
    return 0

if __name__ == "__main__":
    excel_path = "tide_extracted/December_Tide_Table_2025.xlsx"
    exit(regenerate_excel_file(excel_path))

