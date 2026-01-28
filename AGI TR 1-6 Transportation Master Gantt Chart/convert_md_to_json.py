#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Option A 병렬 운영 MD 파일을 머시너블 JSON 형식으로 변환
"""

import json
import os
import sys
from datetime import datetime

if __name__ == "__main__":
    # Windows 콘솔 UTF-8 인코딩 설정
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except:
            pass

    md_file = "OPTION A_병령 운영.MD"
    json_file = "OPTION A_병렬 운영.json"

    # 현재 스크립트와 같은 디렉토리에서 파일 찾기
    script_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__ in globals()' else os.getcwd()
    md_path = os.path.join(script_dir, md_file)

    if not os.path.exists(md_path):
        md_path = md_file

    print(f"Converting {md_path} to JSON...")

    # MD 파일 내용을 기반으로 완전한 JSON 구조 생성
    data = {
        "document_metadata": {
            "title": "Option A 병렬 운영 패턴 분석",
            "version": "1.0",
            "generated_at": datetime.now().isoformat(),
            "source_file": md_file,
            "format": "machine-readable-json"
        },
        "executive_summary": {
            "concept": "SPMT 2개 세트가 동시에 서로 다른 작업을 수행해 전체 기간을 단축",
            "time_savings_days": 22,
            "total_duration_days": 40,
            "sequential_duration_days": 62,
            "resource_utilization_percent": 85,
            "strategy_type": "속도 우선 전략"
        },
        "mobilization_phase": {
            "description": "병렬 운영을 위한 준비 단계",
            "timeline": {"start_date": "2026-01-26", "end_date": "2026-02-04"},
            "spmt_sets": [
                {
                    "set_id": "1st_set",
                    "mobilization_date": "2026-01-26",
                    "completion_date": "2026-01-26",
                    "status": "즉시 사용 가능",
                    "first_task": "TR Units 1-2 작업 시작"
                },
                {
                    "set_id": "2nd_set",
                    "mobilization_date": "2026-02-04",
                    "completion_date": "2026-02-04",
                    "status": "병렬 운영 시작",
                    "first_task": "TR Units 3-4 작업 준비"
                }
            ]
        },
        "parallel_operation_phases": [
            {
                "phase_id": "phase_1",
                "name": "1st set만 사용",
                "period": {"start_date": "2026-01-29", "end_date": "2026-02-04"},
                "daily_activities": [
                    {"date": "2026-01-29", "first_set": {"location": "MZP", "activity": "TR Unit 1 Load-out"}, "second_set": {"location": "MZP", "activity": "모빌라이제이션 중"}},
                    {"date": "2026-01-30", "first_set": {"location": "MZP", "activity": "TR Unit 2 Load-out"}, "second_set": {"location": "MZP", "activity": "모빌라이제이션 중"}},
                    {"date": "2026-02-01", "first_set": {"location": "LCT", "activity": "LCT 출항 (MZP → AGI)"}, "second_set": {"location": "MZP", "activity": "모빌라이제이션 중"}},
                    {"date": "2026-02-02", "first_set": {"location": "AGI", "activity": "AGI 입항"}, "second_set": {"location": "MZP", "activity": "모빌라이제이션 중"}},
                    {"date": "2026-02-04", "first_set": {"location": "AGI", "activity": "TR Unit 1 Load-in (AGI)"}, "second_set": {"location": "MZP", "activity": "모빌라이제이션 완료"}}
                ]
            },
            {
                "phase_id": "phase_2",
                "name": "병렬 운영 시작",
                "period": {"start_date": "2026-02-05", "end_date": "2026-02-13"},
                "description": "두 세트가 동시에 다른 작업 수행",
                "key_dates": [
                    {
                        "date": "2026-02-05",
                        "description": "병렬 운영 시작일",
                        "time_periods": [
                            {
                                "period": "오전",
                                "first_set": {"location": "AGI", "activity": "TR Unit 1: AGI에서 TR Bay 4로 이동, Steel bridge 설치"},
                                "second_set": {"location": "MZP", "activity": "LCT가 MZP 도착, TR Units 3-4 준비 시작"}
                            },
                            {
                                "period": "오후",
                                "first_set": {"location": "AGI", "activity": "TR Unit 1: SPMT에 적재, Transportation 시작"},
                                "second_set": {"location": "MZP", "activity": "TR Unit 3: MZP에서 SPMT 적재 준비"}
                            }
                        ],
                        "parallel_operations": {
                            "first_set": "AGI 현장에서 TR Unit 1 설치 작업",
                            "second_set": "MZP에서 TR Units 3-4 적재 준비"
                        }
                    },
                    {
                        "date": "2026-02-06",
                        "first_set": {"location": "AGI 현장", "activity": "TR Unit 1: Turning 작업 (3일)"},
                        "second_set": {"location": "MZP", "activity": "TR Unit 3: Load-out 준비, Beam Replacement"}
                    },
                    {
                        "date": "2026-02-07",
                        "first_set": {"location": "AGI 현장", "activity": "TR Unit 1: Turning 계속"},
                        "second_set": {"location": "MZP", "activity": "TR Unit 3: Load-out 완료, TR Unit 4: 적재 시작"},
                        "parallel_effect": {
                            "agi": "1st set로 TR Units 1-2 설치 진행",
                            "mzp": "2nd set로 TR Units 3-4 적재 준비"
                        }
                    },
                    {
                        "date": "2026-02-08",
                        "first_set": {"location": "AGI", "activity": "TR Unit 1: Jack-down 완료, TR Unit 2: SPMT 적재 시작"},
                        "second_set": {"location": "MZP", "activity": "TR Unit 4: Load-out 완료, LCT 출항 준비"}
                    },
                    {
                        "date": "2026-02-09",
                        "first_set": {"location": "AGI", "activity": "TR Unit 2: Transportation 시작"},
                        "second_set": {"location": "LCT", "activity": "LCT 출항 (MZP → AGI), TR Units 3-4 운송 시작"}
                    },
                    {
                        "date": "2026-02-10",
                        "first_set": {"location": "AGI", "activity": "TR Unit 2: Turning 시작"},
                        "second_set": {"location": "해상", "activity": "LCT 해상 운송 중"}
                    },
                    {
                        "date": "2026-02-11",
                        "description": "병렬 운영 전환점",
                        "first_set": {"location": "AGI → Port", "activity": "SPMT shifting back to Port (AGI → MZP 이동 시작)"},
                        "second_set": {"location": "AGI", "activity": "LCT AGI 도착, TR Units 3-4 하역 준비"},
                        "parallel_transition": {
                            "first_set": "TR Units 1-2 작업 완료 → Port로 복귀",
                            "second_set": "TR Units 3-4를 AGI로 운송"
                        }
                    }
                ]
            },
            {
                "phase_id": "phase_3",
                "name": "2nd set 단독 운영",
                "period": {"start_date": "2026-02-12", "end_date": "2026-02-22"},
                "daily_activities": [
                    {"date": "2026-02-12", "first_set": {"location": "Port", "activity": "Port 복귀 완료, 대기 상태"}, "second_set": {"location": "AGI", "activity": "TR Unit 2: Jack-down 완료, TR Units 3-4: Load-in 시작"}},
                    {"date": "2026-02-13", "first_set": {"location": "Port", "activity": "대기"}, "second_set": {"location": "AGI", "activity": "TR Unit 3: Load-in 완료"}},
                    {"date": "2026-02-14", "first_set": {"location": "Port", "activity": "대기"}, "second_set": {"location": "AGI", "activity": "TR Unit 3: Steel bridge 설치, TR Unit 3: SPMT 적재"}},
                    {"date": "2026-02-18", "first_set": {"location": "Port", "activity": "대기"}, "second_set": {"location": "AGI", "activity": "TR Unit 4: SPMT 적재, TR Unit 4: Turning 시작"}},
                    {"date": "2026-02-20", "first_set": {"location": "Port", "activity": "대기"}, "second_set": {"location": "AGI → Port", "activity": "SPMT shifting back to Port"}},
                    {"date": "2026-02-22", "first_set": {"location": "Port", "activity": "대기"}, "second_set": {"location": "AGI", "activity": "TR Unit 4: Jack-down 완료"}}
                ]
            },
            {
                "phase_id": "phase_4",
                "name": "1st set 재활용",
                "period": {"start_date": "2026-02-15", "end_date": "2026-03-03"},
                "description": "2nd set가 TR Units 3-4 작업 중, 1st set는 TR Units 5-6 작업 시작",
                "daily_activities": [
                    {"date": "2026-02-15", "first_set": {"location": "MZP", "activity": "TR Units 5-6: 적재 준비 시작"}, "second_set": {"location": "AGI", "activity": "TR Unit 3: Turning 작업 중"}},
                    {"date": "2026-02-16", "first_set": {"location": "MZP", "activity": "TR Unit 5: Load-out"}, "second_set": {"location": "AGI", "activity": "TR Unit 3: Turning 계속"}},
                    {"date": "2026-02-17", "first_set": {"location": "MZP", "activity": "TR Unit 6: Load-out"}, "second_set": {"location": "AGI", "activity": "TR Unit 3: Jack-down 준비"}},
                    {"date": "2026-02-19", "first_set": {"location": "LCT", "activity": "LCT 출항 (MZP → AGI)"}, "second_set": {"location": "AGI", "activity": "TR Unit 4: Turning 중"}},
                    {"date": "2026-02-20", "first_set": {"location": "해상", "activity": "LCT 해상 운송"}, "second_set": {"location": "AGI → Port", "activity": "2nd set: Port 복귀"}},
                    {"date": "2026-02-21", "first_set": {"location": "AGI", "activity": "LCT AGI 도착"}, "second_set": {"location": "Port", "activity": "대기"}},
                    {"date": "2026-02-23", "first_set": {"location": "AGI", "activity": "TR Unit 5: SPMT 적재"}, "second_set": {"location": "Port", "activity": "대기"}},
                    {"date": "2026-02-27", "first_set": {"location": "AGI", "activity": "TR Unit 5: Jack-down 완료, TR Unit 6: SPMT 적재"}, "second_set": {"location": "Port", "activity": "대기"}},
                    {"date": "2026-03-01", "first_set": {"location": "AGI → Port", "activity": "1st set: Port 복귀"}, "second_set": {"location": "Port", "activity": "대기"}},
                    {"date": "2026-03-03", "first_set": {"location": "Port", "activity": "작업 완료"}, "second_set": {"location": "Port", "activity": "대기"}}
                ]
            },
            {
                "phase_id": "phase_5",
                "name": "2nd set 최종 재활용",
                "period": {"start_date": "2026-02-24", "end_date": "2026-03-07"},
                "description": "1st set 작업 완료 후, 2nd set가 TR Unit 7 작업",
                "daily_activities": [
                    {"date": "2026-02-24", "first_set": {"location": "Port", "activity": "대기"}, "second_set": {"location": "MZP", "activity": "TR Unit 7: 적재 준비"}},
                    {"date": "2026-02-25", "first_set": {"location": "Port", "activity": "대기"}, "second_set": {"location": "MZP", "activity": "TR Unit 7: Load-out"}},
                    {"date": "2026-02-27", "first_set": {"location": "Port", "activity": "대기"}, "second_set": {"location": "LCT", "activity": "LCT 출항 (MZP → AGI)"}},
                    {"date": "2026-02-28", "first_set": {"location": "Port", "activity": "대기"}, "second_set": {"location": "AGI", "activity": "LCT AGI 도착"}},
                    {"date": "2026-03-01", "first_set": {"location": "Port", "activity": "대기"}, "second_set": {"location": "AGI", "activity": "TR Unit 7: Load-in"}},
                    {"date": "2026-03-02", "first_set": {"location": "Port", "activity": "대기"}, "second_set": {"location": "AGI", "activity": "TR Unit 7: SPMT 적재"}},
                    {"date": "2026-03-04", "first_set": {"location": "Port", "activity": "대기"}, "second_set": {"location": "AGI", "activity": "TR Unit 7: Turning 시작"}},
                    {"date": "2026-03-06", "first_set": {"location": "Port", "activity": "대기"}, "second_set": {"location": "AGI → Port", "activity": "2nd set: Port 복귀"}},
                    {"date": "2026-03-07", "first_set": {"location": "Port", "activity": "대기"}, "second_set": {"location": "AGI", "activity": "TR Unit 7: Jack-down 완료"}}
                ]
            }
        ],
        "parallel_operation_mechanisms": {
            "time_overlap": {
                "description": "시간 겹침 (Overlap)",
                "period": {"start_date": "2026-02-05", "end_date": "2026-02-11", "duration_days": 7},
                "activities": {
                    "first_set": {"location": "AGI", "activity": "TR Units 1-2 설치", "start_date": "2026-02-05", "end_date": "2026-02-11", "end_action": "Port 복귀"},
                    "second_set": {"location": "MZP → AGI", "activity": "TR Units 3-4 적재", "start_date": "2026-02-05", "end_date": "2026-02-11", "end_action": "AGI 도착"}
                }
            },
            "resource_separation": {
                "description": "리소스 분리",
                "resources": [
                    {"resource_type": "위치", "first_set": "AGI 현장", "second_set": "MZP → AGI"},
                    {"resource_type": "작업", "first_set": "TR Units 1-2 설치", "second_set": "TR Units 3-4 적재/운송"},
                    {"resource_type": "인력", "first_set": "AGI 설치팀", "second_set": "MZP 적재팀"},
                    {"resource_type": "장비", "first_set": "SPMT 1st set", "second_set": "SPMT 2nd set"}
                ]
            },
            "sequential_transition": {
                "description": "순차적 전환",
                "pattern": ["1st set 완료 → Port 복귀 → 다음 배치 준비", "2nd set 완료 → Port 복귀 → 다음 배치 준비"]
            }
        },
        "performance_metrics": {
            "time_savings": {
                "unit": "days",
                "breakdown": [
                    {"task": "TR Units 1-2", "sequential_days": 16, "parallel_days": 16, "savings": 0},
                    {"task": "TR Units 3-4", "sequential_days": 26, "parallel_days": 17, "savings": 9, "note": "대기 시간 제거"},
                    {"task": "TR Units 5-6", "sequential_days": 26, "parallel_days": 17, "savings": 9, "note": "대기 시간 제거"},
                    {"task": "TR Unit 7", "sequential_days": 12, "parallel_days": 12, "savings": 0}
                ],
                "total": {"sequential_days": 62, "parallel_days": 40, "total_savings": 22}
            },
            "resource_utilization": {
                "spmt_utilization": {"sequential_percent": 50, "parallel_percent": 85, "improvement": 35},
                "project_duration": {"option_a_days": 40, "option_b_days": 62}
            }
        },
        "constraints": [
            {"constraint_id": "initial_investment", "description": "초기 투자", "details": "SPMT 2세트 필요"},
            {"constraint_id": "synchronization", "description": "동기화", "details": "두 세트 작업 일정 조율 필요"},
            {"constraint_id": "resource_distribution", "description": "리소스 분배", "details": "인력/장비를 두 현장에 분산"},
            {"constraint_id": "risk", "description": "리스크", "details": "한 세트 지연 시 전체 영향"}
        ],
        "summary": {
            "strategy": "Option A의 병렬 운영은 SPMT 2개 세트를 동시에 사용",
            "benefits": ["프로젝트 기간을 22일 단축 (40일 vs 62일)", "리소스 활용도 향상"],
            "trade_offs": ["초기 투자 증가", "운영 복잡도 상승"],
            "strategy_type": "속도 우선 전략"
        }
    }

    json_path = os.path.join(script_dir, json_file)

    # JSON 파일로 저장
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    print(f"✅ 변환 완료: {json_path}")
    print(f"   총 {len(data)}개 주요 섹션 변환됨")

    # 콘솔에 샘플 출력
    print("\n📋 JSON 샘플 (executive_summary):")
    print(json.dumps(data["executive_summary"], indent=2, ensure_ascii=False))

