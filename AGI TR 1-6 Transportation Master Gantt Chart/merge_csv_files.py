#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGI TR Schedule 모든 데이터를 하나의 통합 테이블로 변환
"""

import csv
import os
import sys
import json

def create_unified_csv():
    """모든 데이터를 하나의 통합 테이블로 생성"""
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    base_name = "agi tr schedule"
    json_file = f"{base_name}.json"
    json_path = os.path.join(script_dir, json_file)
    output_file = os.path.join(script_dir, f"{base_name}_통합단일.csv")
    
    if not os.path.exists(json_path):
        print(f"❌ 오류: {json_path} 파일을 찾을 수 없습니다.")
        return 1
    
    # JSON 파일 읽기
    print(f"📄 JSON 파일 읽는 중: {json_path}")
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    print(f"💾 통합 CSV 파일 생성 중: {os.path.basename(output_file)}")
    with open(output_file, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        
        # 통합 헤더
        writer.writerow([
            "데이터 타입", "항차", "TR Unit", "날짜", "활동", "위치", 
            "주요 작업", "SPMT 세트", "적재 위치", "기간(일)", "종료일", "비고"
        ])
        
        # 1. 프로젝트 요약 정보
        exec_summary = data.get('executive_summary', {})
        project_summary = data.get('project_summary', {})
        
        writer.writerow([
            "프로젝트 요약", "", "", exec_summary.get('duration_period', {}).get('start_date', ''), 
            "프로젝트 시작", "", f"총 항차: {exec_summary.get('total_voyages', '')}회", 
            "", "", project_summary.get('total_project_duration_days', ''), 
            exec_summary.get('duration_period', {}).get('end_date', ''), 
            f"병렬 운영 절약: {project_summary.get('time_savings', {}).get('savings_days', '')}일"
        ])
        
        # 2. LCT 운송 현황
        for stat in data.get('lct_transport_summary', {}).get('voyage_statistics', []):
            writer.writerow([
                "LCT 운송 현황", stat.get('voyage', ''), stat.get('cargo', ''),
                stat.get('departure_date', ''), "LCT 운송", 
                f"{stat.get('departure_date', '')} → {stat.get('arrival_date', '')}",
                f"운송 소요: {stat.get('transport_duration_days', '')}일",
                "", "", stat.get('transport_duration_days', ''), stat.get('return_date', '') or '-',
                stat.get('note', '') or f"총 소요: {stat.get('total_duration_days', '')}일"
            ])
        
        # 3. SPMT 운영 현황
        spmt_ops = data.get('spmt_operations_summary', {})
        for spmt_set in ['1st_set', '2nd_set']:
            spmt_data = spmt_ops.get(spmt_set, {})
            if spmt_data:
                op_period = spmt_data.get('operation_period', {})
                writer.writerow([
                    "SPMT 운영", "", ", ".join(spmt_data.get('assigned_units', [])),
                    op_period.get('start_date', ''), "SPMT 운영",
                    "", f"{spmt_set.replace('_', ' ').title()} 운영",
                    spmt_set.replace('_', ' ').title(), "", op_period.get('duration_days', ''),
                    op_period.get('end_date', ''), 
                    f"모빌: {spmt_data.get('mobilization_date', '')}"
                ])
        
        # 4. 항차별 상세 일정 (운송 + 설치)
        for voyage in data['voyages']:
            voyage_num = voyage['voyage_number']
            cargo_info = voyage['cargo']
            spmt_set = cargo_info['spmt_set']
            units_str = ", ".join(cargo_info['units'])
            positions_str = ", ".join(cargo_info.get('loading_positions', []))
            
            # 운송 일정
            for schedule in voyage.get('detailed_schedule', []):
                writer.writerow([
                    "운송 일정", f"{voyage_num}차", units_str,
                    schedule.get('date', ''), schedule.get('activity', ''),
                    schedule.get('location', ''), schedule.get('work', ''),
                    spmt_set, positions_str, "", "", ""
                ])
            
            # AGI 설치 일정
            for install in voyage.get('installation_schedule_agi', []):
                work = install.get('work', '')
                # TR Unit 추출
                unit_match = None
                for unit in cargo_info['units']:
                    if unit in work:
                        unit_match = unit
                        break
                
                writer.writerow([
                    "AGI 설치", f"{voyage_num}차", unit_match or units_str,
                    install.get('date', ''), "설치 작업", "AGI", work,
                    spmt_set, positions_str, install.get('duration_days', ''),
                    install.get('end_date', ''), ""
                ])
    
    print(f"✅ 통합 단일 CSV 파일 생성 완료: {os.path.basename(output_file)}")
    return 0


if __name__ == "__main__":
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except:
            pass
    sys.exit(create_unified_csv())
