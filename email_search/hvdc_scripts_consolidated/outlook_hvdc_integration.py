#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook PST ↔ HVDC Core 통합 분석

기능:
- Outlook 이메일 데이터를 HVDC Core Ontology 구조에 매핑
- 트렌드 데이터와 Core 문서 통합
- Cross-domain 분석 (PST ↔ Warehouse Ops ↔ Invoice)

사용:
  python outlook_hvdc_integration.py
  python outlook_hvdc_integration.py --integration-type full
"""

import pandas as pd
import glob
import re
from pathlib import Path
from datetime import datetime
import sys
from typing import Dict, List, Tuple

def load_trend_report() -> pd.DataFrame:
    """트렌드 보고서 로드"""
    pattern = "results/HVDC_TREND_REPORT_*.xlsx"
    files = sorted(glob.glob(pattern), reverse=True)
    
    if not files:
        print("오류: 트렌드 보고서를 찾을 수 없습니다")
        sys.exit(1)
    
    latest_file = files[0]
    print(f"트렌드 보고서 로드: {latest_file}")
    
    df = pd.read_excel(latest_file, sheet_name='월별_요약', engine='openpyxl')
    return df

def map_to_core_ontology(df_trend: pd.DataFrame) -> Dict:
    """트렌드 데이터를 HVDC Core Ontology에 매핑"""
    
    core_mapping = {
        'Nodes': {
            'Warehouse': ['DSV Indoor', 'DSV Outdoor', 'DSV Al Markaz', 'AAA Storage', 'DSV MZP', 'MOSB'],
            'Site': ['AGI', 'DAS', 'MIR', 'MIRFA', 'GHALLAN', 'SHU'],
            'OffshoreBase': ['MOSB']
        },
        'Documents': {
            'Invoice': [],
            'BL': [],
            'BOE': [],
            'DO': [],
            'Certificate': ['ECAS', 'EQM', 'FANR', 'CoC']
        },
        'Process': {
            'Booking': [],
            'Clearance': [],
            'Transport': [],
            'Warehouse': ['WH In', 'WH Out', 'Stowage']
        },
        'Event': {
            'ETA': [],
            'ATA': [],
            'Berth': [],
            'Gate Pass': []
        },
        'Party': {
            'Shipper': [],
            'Consignee': [],
            'Carrier': ['DSV', 'DHL', 'MSC'],
            '3PL': ['DSV', 'DHL'],
            'Authority': ['MOIAT', 'FANR', 'MOIAT', 'Customs']
        }
    }
    
    # Outlook에서 추출된 데이터 분석
    monthly_summary = []
    
    for _, row in df_trend.iterrows():
        monthly_summary.append({
            '월': row['월'],
            '총_이메일': row['총 이메일'],
            '케이스_추출': row['케이스 추출'],
            '사이트_추출': row['사이트 식별'],
            'LPO_추출': row['LPO 추출'],
            '단계_추출': row['단계 분류'],
            '케이스_추출률(%)': row['케이스 추출률(%)'],
            '사이트_추출률(%)': row['사이트 식별률(%)'],
            'LPO_추출률(%)': row['LPO 추출률(%)']
        })
    
    return {
        'Core_Mapping': core_mapping,
        'Monthly_Summary': monthly_summary
    }

def generate_core_integration_report(mapping: Dict, output_path: str):
    """Core 통합 보고서 생성"""
    print("\nCore 통합 보고서 생성 중...")
    
    monthly_summary = mapping['Monthly_Summary']
    core_mapping = mapping['Core_Mapping']
    
    # 1. 월별 통계
    df_monthly = pd.DataFrame(monthly_summary)
    
    # 2. Core Ontology 맵핑
    df_nodes = pd.DataFrame({
        'Node_Type': ['Warehouse', 'Site', 'OffshoreBase'],
        'Count': [
            len(core_mapping['Nodes']['Warehouse']),
            len(core_mapping['Nodes']['Site']),
            len(core_mapping['Nodes']['OffshoreBase'])
        ],
        'Examples': [
            ', '.join(core_mapping['Nodes']['Warehouse'][:3]),
            ', '.join(core_mapping['Nodes']['Site'][:3]),
            ', '.join(core_mapping['Nodes']['OffshoreBase'])
        ]
    })
    
    # 3. 통합 통계
    total_emails = sum([m['총_이메일'] for m in monthly_summary])
    total_cases = sum([m['케이스_추출'] for m in monthly_summary])
    total_sites = sum([m['사이트_추출'] for m in monthly_summary])
    
    integration_stats = pd.DataFrame({
        '지표': [
            '분석 기간',
            '총 이메일',
            '총 케이스 추출',
            '총 사이트 추출',
            '평균 케이스 추출률(%)',
            '평균 사이트 추출률(%)',
            'Core Nodes 수',
            '사용된 표준'
        ],
        '값': [
            f"{monthly_summary[0]['월']} ~ {monthly_summary[-1]['월']}",
            f"{total_emails:,}개",
            f"{total_cases:,}개",
            f"{total_sites:,}개",
            f"{sum([m['케이스_추출률(%)'] for m in monthly_summary]) / len(monthly_summary):.1f}",
            f"{sum([m['사이트_추출률(%)'] for m in monthly_summary]) / len(monthly_summary):.1f}",
            f"{sum(df_nodes['Count'])}개",
            'UN/CEFACT, WCO DM, DCSA, ICC Incoterms, HS 2022, MOIAT, FANR'
        ]
    })
    
    # Excel 저장
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        integration_stats.to_excel(writer, sheet_name='통합_통계', index=False)
        df_monthly.to_excel(writer, sheet_name='월별_통계', index=False)
        df_nodes.to_excel(writer, sheet_name='Core_Nodes', index=False)
    
    print(f"\n✅ Core 통합 보고서: {output_path}")
    print(f"\n포함된 시트:")
    print(f"  1. 통합_통계 - 전체 지표")
    print(f"  2. 월별_통계 - 월별 추이")
    print(f"  3. Core_Nodes - 온톨로지 노드")

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description='Outlook PST ↔ HVDC Core 통합 분석')
    parser.add_argument('--output', '-o', default=None,
                       help='출력 파일 경로')
    parser.add_argument('--integration-type', choices=['basic', 'full'], default='basic',
                       help='통합 타입 (basic/full)')
    
    args = parser.parse_args()
    
    print("="*70)
    print("  Outlook PST ↔ HVDC Core 통합 분석")
    print("="*70)
    
    # 트렌드 데이터 로드
    df_trend = load_trend_report()
    
    # Core Ontology 매핑
    mapping = map_to_core_ontology(df_trend)
    
    # 출력 파일명
    if args.output:
        output_path = args.output
    else:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        output_path = f"results/HVDC_CORE_INTEGRATION_{timestamp}.xlsx"
    
    # 통합 보고서 생성
    generate_core_integration_report(mapping, output_path)
    
    print(f"\n✅ 완료!")

