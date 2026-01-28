#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGI TR Schedule MD 파일을 머시너블 JSON 형식으로 변환
"""

import json
import os
import sys
from datetime import datetime
import re

def parse_date_range(text):
    """날짜 범위 텍스트 파싱 (예: '2026-01-29 ~ 2026-03-07')"""
    match = re.search(r'(\d{4}-\d{2}-\d{2})\s*~\s*(\d{4}-\d{2}-\d{2})', text)
    if match:
        return {"start_date": match.group(1), "end_date": match.group(2)}
    match = re.search(r'(\d{4}-\d{2}-\d{2})', text)
    if match:
        return {"date": match.group(1)}
    return None

def parse_duration(text):
    """기간 텍스트 파싱 (예: '40일', '7일')"""
    match = re.search(r'(\d+)\s*일', text)
    if match:
        return int(match.group(1))
    return None

def convert_md_to_json(md_file_path):
    """MD 파일을 JSON으로 변환"""
    
    # JSON 데이터 구조
    data = {
        "document_metadata": {
            "title": "OPTION A 전체 운송 일정 - LCT 항차별 요약 보고서",
            "version": "1.0",
            "generated_at": datetime.now().isoformat(),
            "source_file": os.path.basename(md_file_path),
            "format": "machine-readable-json"
        },
        "executive_summary": {
            "total_voyages": 4,
            "voyage_cargo": ["TR Units 1-2", "TR Units 3-4", "TR Units 5-6", "TR Unit 7"],
            "total_duration_days": 40,
            "duration_period": {"start_date": "2026-01-29", "end_date": "2026-03-07"},
            "lct_transport_count": "MZP ↔ AGI 왕복 4회 + 단방향 1회",
            "parallel_operation_period": {"start_date": "2026-02-05", "end_date": "2026-02-11", "duration_days": 7}
        },
        "voyages": [
            {
                "voyage_id": "voyage_1",
                "voyage_number": 1,
                "cargo": {
                    "units": ["AGI TR Unit 1", "AGI TR Unit 2"],
                    "loading_positions": ["TR Bay 4 (Unit 1)", "TR Bay 3 (Unit 2)"],
                    "spmt_set": "1st Set"
                },
                "detailed_schedule": [
                    {"date": "2026-01-29", "activity": "Load-out 준비", "location": "MZP", "work": "TR Unit 1 SPMT 적재, RoRo Ramp 설치"},
                    {"date": "2026-01-29", "activity": "Load-out", "location": "MZP", "work": "TR Unit 1 Load-out (10:00-11:00)"},
                    {"date": "2026-01-30", "activity": "Load-out", "location": "MZP", "work": "TR Unit 2 Load-out (08:00-09:00)"},
                    {"date": "2026-01-31", "activity": "최종 준비", "location": "MZP", "work": "MWS + MPI + 최종 준비"},
                    {"date": "2026-02-01", "activity": "출항", "location": "MZP → AGI", "work": "LCT 출항"},
                    {"date": "2026-02-02", "activity": "입항", "location": "AGI", "work": "LCT 입항, MMT 크루 모빌라이제이션"},
                    {"date": "2026-02-03", "activity": "Load-in", "location": "AGI", "work": "TR Unit 2 Load-in (Jetty 저장)"},
                    {"date": "2026-02-04", "activity": "Load-in", "location": "AGI", "work": "TR Unit 1 Load-in (Jetty 저장)"},
                    {"date": "2026-02-05", "activity": "복귀", "location": "AGI → MZP", "work": "LCT MZP 복귀 (TR Units 3-4 적재 준비)"}
                ],
                "installation_schedule_agi": [
                    {"date": "2026-02-05", "work": "TR Unit 1 → TR Bay 4 SPMT 적재/운송"},
                    {"date": "2026-02-06", "work": "TR Unit 1 Turning 시작", "duration_days": 3, "end_date": "2026-02-08"},
                    {"date": "2026-02-09", "work": "TR Unit 1 Jacking down 완료"},
                    {"date": "2026-02-09", "work": "TR Unit 2 → TR Bay 3 SPMT 적재/운송"},
                    {"date": "2026-02-10", "work": "TR Unit 2 Turning 시작", "duration_days": 3, "end_date": "2026-02-12"},
                    {"date": "2026-02-11", "work": "1st Set SPMT Port 복귀"},
                    {"date": "2026-02-13", "work": "TR Unit 2 Jacking down 완료"}
                ]
            },
            {
                "voyage_id": "voyage_2",
                "voyage_number": 2,
                "cargo": {
                    "units": ["AGI TR Unit 3", "AGI TR Unit 4"],
                    "loading_positions": ["TR Bay 2 (Unit 3)", "TR Bay 1 (Unit 4)"],
                    "spmt_set": "2nd Set"
                },
                "features": "항차 1과 병렬 운영 (2026-02-05 ~ 02-11)",
                "parallel_operation": {"start_date": "2026-02-05", "end_date": "2026-02-11"},
                "detailed_schedule": [
                    {"date": "2026-02-06", "activity": "LCT 도착", "location": "MZP", "work": "LCT MZP 도착, Deck 준비"},
                    {"date": "2026-02-07", "activity": "Load-out 준비", "location": "MZP", "work": "TR Unit 3 SPMT 적재, RoRo Ramp 설치"},
                    {"date": "2026-02-07", "activity": "Load-out", "location": "MZP", "work": "TR Unit 3 Load-out, TR Unit 4 적재 준비"},
                    {"date": "2026-02-08", "activity": "Load-out", "location": "MZP", "work": "TR Unit 4 Load-out"},
                    {"date": "2026-02-09", "activity": "최종 준비", "location": "MZP", "work": "MWS + MPI + 최종 준비"},
                    {"date": "2026-02-10", "activity": "출항", "location": "MZP → AGI", "work": "LCT 출항 (병렬 운영 중)"},
                    {"date": "2026-02-11", "activity": "입항", "location": "AGI", "work": "LCT AGI 입항 (병렬 운영 중)"},
                    {"date": "2026-02-12", "activity": "Load-in", "location": "AGI", "work": "TR Unit 4 Load-in (Jetty 저장)"},
                    {"date": "2026-02-13", "activity": "Load-in", "location": "AGI", "work": "TR Unit 3 Load-in (Jetty 저장)"},
                    {"date": "2026-02-14", "activity": "복귀", "location": "AGI → MZP", "work": "LCT MZP 복귀 (7.45m beam 4개 반송, TR Units 5-6 적재 준비)"}
                ],
                "installation_schedule_agi": [
                    {"date": "2026-02-14", "work": "TR Unit 3 → TR Bay 2 SPMT 적재/운송"},
                    {"date": "2026-02-15", "work": "TR Unit 3 Turning 시작", "duration_days": 3, "end_date": "2026-02-17"},
                    {"date": "2026-02-18", "work": "TR Unit 3 Jacking down 완료"},
                    {"date": "2026-02-18", "work": "TR Unit 4 → TR Bay 1 SPMT 적재/운송"},
                    {"date": "2026-02-19", "work": "TR Unit 4 Turning 시작", "duration_days": 3, "end_date": "2026-02-21"},
                    {"date": "2026-02-20", "work": "2nd Set SPMT Port 복귀"},
                    {"date": "2026-02-22", "work": "TR Unit 4 Jacking down 완료"}
                ],
                "return_cargo": {
                    "description": "7.45m beam 4개 반송",
                    "quantity": 4,
                    "unit": "개"
                }
            },
            {
                "voyage_id": "voyage_3",
                "voyage_number": 3,
                "cargo": {
                    "units": ["AGI TR Unit 5", "AGI TR Unit 6"],
                    "loading_positions": ["TR Bay 5 (Unit 5)", "TR Bay 6 (Unit 6)"],
                    "spmt_set": "1st Set (재활용)"
                },
                "features": "2nd Set가 TR Units 3-4 작업 중 병렬 운영",
                "detailed_schedule": [
                    {"date": "2026-02-15", "activity": "LCT 도착", "location": "MZP", "work": "LCT MZP 도착, Deck 준비"},
                    {"date": "2026-02-16", "activity": "Load-out 준비", "location": "MZP", "work": "TR Unit 5 SPMT 적재, RoRo Ramp 설치"},
                    {"date": "2026-02-16", "activity": "Load-out", "location": "MZP", "work": "TR Unit 5 Load-out, TR Unit 6 적재 준비"},
                    {"date": "2026-02-17", "activity": "Load-out", "location": "MZP", "work": "TR Unit 6 Load-out"},
                    {"date": "2026-02-18", "activity": "최종 준비", "location": "MZP", "work": "MWS + MPI + 최종 준비"},
                    {"date": "2026-02-19", "activity": "출항", "location": "MZP → AGI", "work": "LCT 출항"},
                    {"date": "2026-02-20", "activity": "입항", "location": "AGI", "work": "LCT AGI 입항"},
                    {"date": "2026-02-21", "activity": "Load-in", "location": "AGI", "work": "TR Unit 6 Load-in (Jetty 저장)"},
                    {"date": "2026-02-22", "activity": "Load-in", "location": "AGI", "work": "TR Unit 5 Load-in (Jetty 저장)"},
                    {"date": "2026-02-23", "activity": "복귀", "location": "AGI → MZP", "work": "LCT MZP 복귀 (7.45m beam 2개 반송, TR Unit 7 적재 준비)"}
                ],
                "installation_schedule_agi": [
                    {"date": "2026-02-23", "work": "TR Unit 5 → TR Bay 5 SPMT 적재/운송"},
                    {"date": "2026-02-24", "work": "TR Unit 5 Turning 시작", "duration_days": 3, "end_date": "2026-02-26"},
                    {"date": "2026-02-27", "work": "TR Unit 5 Jacking down 완료"},
                    {"date": "2026-02-27", "work": "TR Unit 6 → TR Bay 6 SPMT 적재/운송"},
                    {"date": "2026-02-28", "work": "TR Unit 6 Turning 시작", "duration_days": 3, "end_date": "2026-03-02"},
                    {"date": "2026-03-01", "work": "1st Set SPMT Port 복귀 (최종)"},
                    {"date": "2026-03-03", "work": "TR Unit 6 Jacking down 완료"}
                ],
                "return_cargo": {
                    "description": "7.45m beam 2개 반송",
                    "quantity": 2,
                    "unit": "개"
                }
            },
            {
                "voyage_id": "voyage_4",
                "voyage_number": 4,
                "cargo": {
                    "units": ["AGI TR Unit 7"],
                    "loading_positions": ["TR Bay 7"],
                    "spmt_set": "2nd Set (재활용)"
                },
                "features": "최종 단독 운송",
                "is_one_way": True,
                "detailed_schedule": [
                    {"date": "2026-02-24", "activity": "LCT 도착", "location": "MZP", "work": "LCT MZP 도착, Deck 준비"},
                    {"date": "2026-02-25", "activity": "Load-out 준비", "location": "MZP", "work": "TR Unit 7 SPMT 적재, RoRo Ramp 설치"},
                    {"date": "2026-02-25", "activity": "Load-out", "location": "MZP", "work": "TR Unit 7 Load-out"},
                    {"date": "2026-02-26", "activity": "최종 준비", "location": "MZP", "work": "MWS + MPI + 최종 준비"},
                    {"date": "2026-02-27", "activity": "출항", "location": "MZP → AGI", "work": "LCT 출항"},
                    {"date": "2026-02-28", "activity": "입항", "location": "AGI", "work": "LCT AGI 입항"},
                    {"date": "2026-03-01", "activity": "Load-in", "location": "AGI", "work": "TR Unit 7 Load-in (Jetty 저장)"}
                ],
                "installation_schedule_agi": [
                    {"date": "2026-03-02", "work": "TR Unit 7 → TR Bay 7 SPMT 적재/운송"},
                    {"date": "2026-03-04", "work": "TR Unit 7 Turning 시작", "duration_days": 3, "end_date": "2026-03-06"},
                    {"date": "2026-03-06", "work": "2nd Set SPMT Port 복귀"},
                    {"date": "2026-03-07", "work": "TR Unit 7 Jacking down 완료 (전체 작업 완료)", "is_completion": True}
                ]
            }
        ],
        "lct_transport_summary": {
            "voyage_statistics": [
                {
                    "voyage": "1차",
                    "cargo": "TR Units 1-2",
                    "departure_date": "2026-02-01",
                    "arrival_date": "2026-02-02",
                    "transport_duration_days": 1,
                    "return_date": "2026-02-05",
                    "total_duration_days": 5
                },
                {
                    "voyage": "2차",
                    "cargo": "TR Units 3-4",
                    "departure_date": "2026-02-10",
                    "arrival_date": "2026-02-11",
                    "transport_duration_days": 1,
                    "return_date": "2026-02-14",
                    "total_duration_days": 5
                },
                {
                    "voyage": "3차",
                    "cargo": "TR Units 5-6",
                    "departure_date": "2026-02-19",
                    "arrival_date": "2026-02-20",
                    "transport_duration_days": 1,
                    "return_date": "2026-02-23",
                    "total_duration_days": 5
                },
                {
                    "voyage": "4차",
                    "cargo": "TR Unit 7",
                    "departure_date": "2026-02-27",
                    "arrival_date": "2026-02-28",
                    "transport_duration_days": 1,
                    "return_date": None,
                    "total_duration_days": 2,
                    "note": "단방향"
                }
            ],
            "key_features": [
                "LCT 왕복 시간: 각 항차당 약 5일 (적재 3일 + 운송 1일 + 하역 1일)",
                "병렬 운영 기간: 2026-02-05 ~ 02-11 (항차 1 복귀 중 항차 2 적재 진행)",
                "Beam 반송: 항차 2에서 4개, 항차 3에서 2개 반송 (재활용)"
            ],
            "average_round_trip_days": 5,
            "one_way_voyages": 1,
            "return_cargo_summary": {
                "beam_7_45m": {
                    "voyage_2": 4,
                    "voyage_3": 2,
                    "total": 6
                }
            }
        },
        "spmt_operations_summary": {
            "1st_set": {
                "mobilization_date": "2026-01-26",
                "operation_period": {
                    "start_date": "2026-01-29",
                    "end_date": "2026-03-01",
                    "duration_days": 32
                },
                "assigned_units": ["TR Units 1-2", "TR Units 5-6"],
                "port_return_dates": ["2026-02-11", "2026-03-01"],
                "demobilization_period": {
                    "start_date": "2026-03-06",
                    "end_date": "2026-03-07"
                },
                "reuse_info": "TR Units 5-6에서 재활용"
            },
            "2nd_set": {
                "mobilization_date": "2026-02-04",
                "operation_period": {
                    "start_date": "2026-02-07",
                    "end_date": "2026-03-06",
                    "duration_days": 28
                },
                "assigned_units": ["TR Units 3-4", "TR Unit 7"],
                "port_return_dates": ["2026-02-20", "2026-03-06"],
                "demobilization_date": "2026-02-27",
                "reuse_info": "TR Unit 7에서 재활용"
            },
            "utilization_analysis": {
                "1st_set_utilization_percent": 88.9,
                "2nd_set_utilization_percent": 82.4,
                "overall_utilization_percent": 85.7,
                "parallel_operation_days": 7
            }
        },
        "project_summary": {
            "total_project_duration_days": 40,
            "time_savings": {
                "sequential_days": 62,
                "parallel_days": 40,
                "savings_days": 22,
                "savings_percent": 35.5
            },
            "spmt_utilization_percent": 85.7,
            "lct_transport_efficiency": "4회 항차로 7개 TR Unit 완료",
            "parallel_operation_savings": "22일 (순차 운영 62일 → 병렬 운영 40일)",
            "total_tr_units": 7,
            "total_voyages": 4,
            "completion_date": "2026-03-07"
        },
        "report_metadata": {
            "generation_date": "2026-01-18",
            "data_sources": ["OPTION A.tsv", "OPTION A_병렬 운영.json"]
        }
    }
    
    return data


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
    md_file = "agi tr schedule.md"
    json_file = "agi tr schedule.json"
    
    md_path = os.path.join(script_dir, md_file)
    json_path = os.path.join(script_dir, json_file)

    if not os.path.exists(md_path):
        print(f"❌ 오류: {md_path} 파일을 찾을 수 없습니다.")
        return 1

    print(f"📄 MD 파일 읽는 중: {md_path}")
    
    # JSON 변환
    data = convert_md_to_json(md_path)

    # JSON 파일로 저장
    print(f"💾 JSON 파일 저장 중: {json_path}")
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    print(f"✅ 변환 완료: {json_path}")
    print(f"   총 {len(data)}개 주요 섹션 변환됨")
    print(f"   항차 수: {len(data['voyages'])}개")
    print(f"   총 TR Units: {data['project_summary']['total_tr_units']}개")
    print(f"   프로젝트 기간: {data['project_summary']['total_project_duration_days']}일")

    # 콘솔에 샘플 출력
    print("\n📋 JSON 샘플 (executive_summary):")
    print(json.dumps(data["executive_summary"], indent=2, ensure_ascii=False))

    return 0


if __name__ == "__main__":
    sys.exit(main())
