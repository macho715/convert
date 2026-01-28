"""
폴더 스캔 CLI - 기존 email_folder_scanner 래핑
"""
from __future__ import annotations
from pathlib import Path
from typing import List, Dict, Any
import logging
import json
from datetime import datetime
from ..scanner.fs_scanner import scan_folder
from ..parser.subject_parser import parse_folder_title
from ..core.config import EMAIL_ROOT
from ..core.errors import ScanError


def scan_email_folders(email_root: Path = None) -> Dict[str, Any]:
    """
    이메일 폴더를 스캔하고 분석하는 메인 함수
    
    Args:
        email_root: 이메일 루트 폴더 (기본값: EMAIL_ROOT)
        
    Returns:
        Dict[str, Any]: 스캔 결과
    """
    if email_root is None:
        email_root = EMAIL_ROOT
    
    logging.info(f"이메일 폴더 스캔 시작: {email_root}")
    
    try:
        # 1. 폴더 스캔
        files = scan_folder(email_root)
        logging.info(f"스캔된 파일 수: {len(files)}")
        
        # 2. 폴더별 분석
        folder_analysis = {}
        for file_path in files:
            folder_name = str(file_path.parent)
            
            if folder_name not in folder_analysis:
                folder_analysis[folder_name] = {
                    'file_count': 0,
                    'cases': [],
                    'sites': [],
                    'lpos': [],
                    'phases': []
                }
            
            folder_analysis[folder_name]['file_count'] += 1
            
            # 폴더명에서 메타데이터 추출
            parsed_data = parse_folder_title(folder_name)
            folder_analysis[folder_name]['cases'].extend([h['value'] for h in parsed_data['cases']])
            folder_analysis[folder_name]['sites'].extend(parsed_data['sites'])
            folder_analysis[folder_name]['lpos'].extend(parsed_data['lpos'])
            folder_analysis[folder_name]['phases'].extend(parsed_data['phases'])
        
        # 3. 중복 제거
        for folder_data in folder_analysis.values():
            folder_data['cases'] = list(set(folder_data['cases']))
            folder_data['sites'] = list(set(folder_data['sites']))
            folder_data['lpos'] = list(set(folder_data['lpos']))
            folder_data['phases'] = list(set(folder_data['phases']))
        
        # 4. 통계 생성
        total_folders = len(folder_analysis)
        total_files = len(files)
        
        all_cases = []
        all_sites = []
        all_lpos = []
        
        for folder_data in folder_analysis.values():
            all_cases.extend(folder_data['cases'])
            all_sites.extend(folder_data['sites'])
            all_lpos.extend(folder_data['lpos'])
        
        unique_cases = list(set(all_cases))
        unique_sites = list(set(all_sites))
        unique_lpos = list(set(all_lpos))
        
        result = {
            'scan_timestamp': datetime.now().isoformat(),
            'email_root': str(email_root),
            'total_folders': total_folders,
            'total_files': total_files,
            'unique_cases': len(unique_cases),
            'unique_sites': len(unique_sites),
            'unique_lpos': len(unique_lpos),
            'case_list': unique_cases,
            'site_list': unique_sites,
            'lpo_list': unique_lpos,
            'folder_analysis': folder_analysis
        }
        
        logging.info(f"폴더 스캔 완료: {total_folders}개 폴더, {total_files}개 파일")
        return result
        
    except Exception as e:
        logging.error(f"폴더 스캔 실패: {e}")
        raise


def save_scan_results(results: Dict[str, Any], output_path: Path = None) -> Path:
    """
    스캔 결과를 JSON 파일로 저장
    
    Args:
        results: 스캔 결과
        output_path: 출력 파일 경로
        
    Returns:
        Path: 저장된 파일 경로
    """
    if output_path is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = Path(f"email_scan_results_{timestamp}.json")
    
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    
    return output_path


def main():
    """CLI 진입점"""
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    
    try:
        results = scan_email_folders()
        output_path = save_scan_results(results)
        
        print(f"✅ 폴더 스캔 완료!")
        print(f"📁 총 폴더: {results['total_folders']}개")
        print(f"📄 총 파일: {results['total_files']}개")
        print(f"🎯 고유 케이스: {results['unique_cases']}개")
        print(f"🏗️ 고유 사이트: {results['unique_sites']}개")
        print(f"📋 고유 LPO: {results['unique_lpos']}개")
        print(f"💾 결과 저장: {output_path}")
        
    except Exception as e:
        print(f"❌ 오류 발생: {e}")
        return 1
    
    return 0


if __name__ == "__main__":
    exit(main())
