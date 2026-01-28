"""
타임라인 및 네트워크 분석 생성기
"""
from __future__ import annotations
from typing import List, Dict, Any
import pandas as pd
from datetime import datetime
import json


def create_timeline_data(data: List[Dict[str, Any]]) -> pd.DataFrame:
    """
    이메일 데이터에서 타임라인 데이터 생성
    
    Args:
        data: 이메일 데이터 리스트
        
    Returns:
        pd.DataFrame: 타임라인 데이터
    """
    timeline_data = []
    
    for item in data:
        timeline_item = {
            'date': item.get('date', ''),
            'subject': item.get('subject', ''),
            'cases': ', '.join([h['value'] for h in item.get('cases', [])]),
            'sites': ', '.join(item.get('sites', [])),
            'lpos': ', '.join(item.get('lpos', [])),
            'sender': item.get('sender', ''),
            'folder': item.get('folder', '')
        }
        timeline_data.append(timeline_item)
    
    return pd.DataFrame(timeline_data)


def create_network_data(data: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    네트워크 분석 데이터 생성
    
    Args:
        data: 이메일 데이터 리스트
        
    Returns:
        Dict[str, Any]: 네트워크 분석 결과
    """
    # 케이스별 연결 분석
    case_connections = {}
    site_connections = {}
    
    for item in data:
        cases = [h['value'] for h in item.get('cases', [])]
        sites = item.get('sites', [])
        
        # 케이스 연결
        for case in cases:
            if case not in case_connections:
                case_connections[case] = {
                    'sites': set(),
                    'lpos': set(),
                    'emails': []
                }
            
            case_connections[case]['sites'].update(sites)
            case_connections[case]['lpos'].update(item.get('lpos', []))
            case_connections[case]['emails'].append({
                'subject': item.get('subject', ''),
                'date': item.get('date', ''),
                'folder': item.get('folder', '')
            })
        
        # 사이트 연결
        for site in sites:
            if site not in site_connections:
                site_connections[site] = {
                    'cases': set(),
                    'emails': []
                }
            
            site_connections[site]['cases'].update(cases)
            site_connections[site]['emails'].append({
                'subject': item.get('subject', ''),
                'date': item.get('date', ''),
                'folder': item.get('folder', '')
            })
    
    # 집합을 리스트로 변환
    for case_data in case_connections.values():
        case_data['sites'] = list(case_data['sites'])
        case_data['lpos'] = list(case_data['lpos'])
    
    for site_data in site_connections.values():
        site_data['cases'] = list(site_data['cases'])
    
    return {
        'case_connections': case_connections,
        'site_connections': site_connections,
        'total_cases': len(case_connections),
        'total_sites': len(site_connections),
        'total_emails': len(data)
    }


def create_summary_stats(data: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    요약 통계 생성
    
    Args:
        data: 이메일 데이터 리스트
        
    Returns:
        Dict[str, Any]: 요약 통계
    """
    total_emails = len(data)
    
    # 케이스 통계
    all_cases = []
    for item in data:
        all_cases.extend([h['value'] for h in item.get('cases', [])])
    
    unique_cases = list(set(all_cases))
    
    # 사이트 통계
    all_sites = []
    for item in data:
        all_sites.extend(item.get('sites', []))
    
    unique_sites = list(set(all_sites))
    
    # LPO 통계
    all_lpos = []
    for item in data:
        all_lpos.extend(item.get('lpos', []))
    
    unique_lpos = list(set(all_lpos))
    
    return {
        'total_emails': total_emails,
        'unique_cases': len(unique_cases),
        'unique_sites': len(unique_sites),
        'unique_lpos': len(unique_lpos),
        'case_list': unique_cases,
        'site_list': unique_sites,
        'lpo_list': unique_lpos
    }
