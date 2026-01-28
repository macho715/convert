#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGI TR Schedule MD/JSON 파일을 CSV 형식으로 변환
"""

import json
import csv
import os
import sys
from datetime import datetime

def convert_json_to_csv(json_file_path):
    """JSON 파일을 여러 CSV 파일로 변환"""
    
    # JSON 파일 읽기
    with open(json_file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    script_dir = os.path.dirname(os.path.abspath(json_file_path))
    base_name = "agi tr schedule"
    
    csv_files = []
    
    # 1. 전체 일정 통합 CSV (항차별 상세 일정 + AGI 설치 일정)
    integrated_csv = os.path.join(script_dir, f"{base_name}_통합일정.csv")
    csv_files.append(integrated_csv)
    
    with open(integrated_csv, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        
        # 헤더
        writer.writerow([
            "항차", "구분", "날짜", "활동", "위치", "주요 작업", 
            "TR Unit", "SPMT 세트", "적재 위치"
        ])
        
        # 각 항차별 데이터
        for voyage in data['voyages']:
            voyage_num = voyage['voyage_number']
            cargo_info = voyage['cargo']
            spmt_set = cargo_info['spmt_set']
            
            # 상세 일정
            for schedule in voyage.get('detailed_schedule', []):
                # TR Unit 추출
                units_str = ", ".join(cargo_info['units'])
                
                writer.writerow([
                    f"{voyage_num}차",
                    "운송 일정",
                    schedule.get('date', ''),
                    schedule.get('activity', ''),
                    schedule.get('location', ''),
                    schedule.get('work', ''),
                    units_str,
                    spmt_set,
                    ", ".join(cargo_info.get('loading_positions', []))
                ])
            
            # AGI 설치 일정
            for install in voyage.get('installation_schedule_agi', []):
                units_str = ", ".join(cargo_info['units'])
                work = install.get('work', '')
                
                # TR Unit 추출 (work에서)
                unit_match = None
                for unit in cargo_info['units']:
                    if unit in work:
                        unit_match = unit
                        break
                
                writer.writerow([
                    f"{voyage_num}차",
                    "AGI 설치",
                    install.get('date', ''),
                    "설치 작업",
                    "AGI",
                    work,
                    unit_match or units_str,
                    spmt_set,
                    ", ".join(cargo_info.get('loading_positions', []))
                ])
    
    print(f"✅ 생성: {os.path.basename(integrated_csv)}")
    
    # 2. 항차별 상세 일정 CSV
    transport_schedule_csv = os.path.join(script_dir, f"{base_name}_항차별운송일정.csv")
    csv_files.append(transport_schedule_csv)
    
    with open(transport_schedule_csv, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        
        writer.writerow([
            "항차", "운송 물량", "SPMT 세트", "날짜", "활동", "위치", "주요 작업"
        ])
        
        for voyage in data['voyages']:
            voyage_num = voyage['voyage_number']
            cargo_info = voyage['cargo']
            units_str = ", ".join(cargo_info['units'])
            spmt_set = cargo_info['spmt_set']
            
            for schedule in voyage.get('detailed_schedule', []):
                writer.writerow([
                    f"{voyage_num}차",
                    units_str,
                    spmt_set,
                    schedule.get('date', ''),
                    schedule.get('activity', ''),
                    schedule.get('location', ''),
                    schedule.get('work', '')
                ])
    
    print(f"✅ 생성: {os.path.basename(transport_schedule_csv)}")
    
    # 3. AGI 설치 일정 CSV
    installation_csv = os.path.join(script_dir, f"{base_name}_AGI설치일정.csv")
    csv_files.append(installation_csv)
    
    with open(installation_csv, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        
        writer.writerow([
            "항차", "TR Unit", "날짜", "작업 내용", "기간(일)", "종료일", "SPMT 세트", "적재 위치"
        ])
        
        for voyage in data['voyages']:
            voyage_num = voyage['voyage_number']
            cargo_info = voyage['cargo']
            spmt_set = cargo_info['spmt_set']
            
            for install in voyage.get('installation_schedule_agi', []):
                work = install.get('work', '')
                
                # TR Unit 추출
                unit_match = None
                for unit in cargo_info['units']:
                    if unit in work:
                        unit_match = unit
                        break
                
                if not unit_match:
                    unit_match = ", ".join(cargo_info['units'])
                
                writer.writerow([
                    f"{voyage_num}차",
                    unit_match,
                    install.get('date', ''),
                    work,
                    install.get('duration_days', ''),
                    install.get('end_date', ''),
                    spmt_set,
                    ", ".join(cargo_info.get('loading_positions', []))
                ])
    
    print(f"✅ 생성: {os.path.basename(installation_csv)}")
    
    # 4. LCT 운송 현황 CSV
    lct_summary_csv = os.path.join(script_dir, f"{base_name}_LCT운송현황.csv")
    csv_files.append(lct_summary_csv)
    
    with open(lct_summary_csv, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        
        writer.writerow([
            "항차", "운송 물량", "출항일", "입항일", "운송 소요(일)", 
            "복귀일", "총 소요일", "비고"
        ])
        
        for stat in data.get('lct_transport_summary', {}).get('voyage_statistics', []):
            writer.writerow([
                stat.get('voyage', ''),
                stat.get('cargo', ''),
                stat.get('departure_date', ''),
                stat.get('arrival_date', ''),
                stat.get('transport_duration_days', ''),
                stat.get('return_date', '') or '-',
                stat.get('total_duration_days', ''),
                stat.get('note', '')
            ])
    
    print(f"✅ 생성: {os.path.basename(lct_summary_csv)}")
    
    # 5. SPMT 운영 현황 CSV
    spmt_summary_csv = os.path.join(script_dir, f"{base_name}_SPMT운영현황.csv")
    csv_files.append(spmt_summary_csv)
    
    with open(spmt_summary_csv, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        
        writer.writerow([
            "SPMT 세트", "모빌라이제이션일", "운영 시작일", "운영 종료일", 
            "운영 기간(일)", "담당 Units", "Port 복귀일", "디모빌라이제이션일", "비고"
        ])
        
        spmt_ops = data.get('spmt_operations_summary', {})
        
        # 1st Set
        first_set = spmt_ops.get('1st_set', {})
        op_period = first_set.get('operation_period', {})
        demob_period = first_set.get('demobilization_period', {})
        port_returns = first_set.get('port_return_dates', [])
        
        writer.writerow([
            "1st Set",
            first_set.get('mobilization_date', ''),
            op_period.get('start_date', ''),
            op_period.get('end_date', ''),
            op_period.get('duration_days', ''),
            ", ".join(first_set.get('assigned_units', [])),
            ", ".join([str(d) for d in port_returns]),
            f"{demob_period.get('start_date', '')} ~ {demob_period.get('end_date', '')}" if demob_period else '',
            first_set.get('reuse_info', '')
        ])
        
        # 2nd Set
        second_set = spmt_ops.get('2nd_set', {})
        op_period2 = second_set.get('operation_period', {})
        port_returns2 = second_set.get('port_return_dates', [])
        
        writer.writerow([
            "2nd Set",
            second_set.get('mobilization_date', ''),
            op_period2.get('start_date', ''),
            op_period2.get('end_date', ''),
            op_period2.get('duration_days', ''),
            ", ".join(second_set.get('assigned_units', [])),
            ", ".join([str(d) for d in port_returns2]),
            second_set.get('demobilization_date', ''),
            second_set.get('reuse_info', '')
        ])
    
    print(f"✅ 생성: {os.path.basename(spmt_summary_csv)}")
    
    # 6. 프로젝트 요약 CSV
    summary_csv = os.path.join(script_dir, f"{base_name}_프로젝트요약.csv")
    csv_files.append(summary_csv)
    
    with open(summary_csv, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        
        writer.writerow(["항목", "내용"])
        
        exec_summary = data.get('executive_summary', {})
        project_summary = data.get('project_summary', {})
        
        writer.writerow(["총 항차", exec_summary.get('total_voyages', '')])
        writer.writerow(["운송 물량", ", ".join(exec_summary.get('voyage_cargo', []))])
        writer.writerow(["총 프로젝트 기간(일)", project_summary.get('total_project_duration_days', '')])
        writer.writerow(["프로젝트 시작일", exec_summary.get('duration_period', {}).get('start_date', '')])
        writer.writerow(["프로젝트 종료일", exec_summary.get('duration_period', {}).get('end_date', '')])
        writer.writerow(["LCT 운송 횟수", exec_summary.get('lct_transport_count', '')])
        writer.writerow(["병렬 운영 시작일", exec_summary.get('parallel_operation_period', {}).get('start_date', '')])
        writer.writerow(["병렬 운영 종료일", exec_summary.get('parallel_operation_period', {}).get('end_date', '')])
        writer.writerow(["병렬 운영 기간(일)", exec_summary.get('parallel_operation_period', {}).get('duration_days', '')])
        writer.writerow(["순차 운영 기간(일)", project_summary.get('time_savings', {}).get('sequential_days', '')])
        writer.writerow(["병렬 운영 기간(일)", project_summary.get('time_savings', {}).get('parallel_days', '')])
        writer.writerow(["시간 절약(일)", project_summary.get('time_savings', {}).get('savings_days', '')])
        writer.writerow(["시간 절약률(%)", project_summary.get('time_savings', {}).get('savings_percent', '')])
        writer.writerow(["SPMT 활용도(%)", project_summary.get('spmt_utilization_percent', '')])
        writer.writerow(["LCT 운송 효율", project_summary.get('lct_transport_efficiency', '')])
        writer.writerow(["총 TR Units", project_summary.get('total_tr_units', '')])
        writer.writerow(["완료일", project_summary.get('completion_date', '')])
    
    print(f"✅ 생성: {os.path.basename(summary_csv)}")
    
    return csv_files


def main():
    """메인 함수"""
    # Windows 콘솔 UTF-8 인코딩 설정
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except:
            pass

    # 파일 경로 설정
    script_dir = os.path.dirname(os.path.abspath(__file__))
    json_file = "agi tr schedule.json"
    json_path = os.path.join(script_dir, json_file)

    if not os.path.exists(json_path):
        print(f"❌ 오류: {json_path} 파일을 찾을 수 없습니다.")
        print(f"   먼저 convert_agi_tr_schedule_to_json.py를 실행하여 JSON 파일을 생성해주세요.")
        return 1

    print(f"📄 JSON 파일 읽는 중: {json_path}")
    
    # CSV 변환
    csv_files = convert_json_to_csv(json_path)
    
    print(f"\n✅ 변환 완료! 총 {len(csv_files)}개의 CSV 파일이 생성되었습니다.")
    print("\n생성된 CSV 파일:")
    for csv_file in csv_files:
        print(f"  - {os.path.basename(csv_file)}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
