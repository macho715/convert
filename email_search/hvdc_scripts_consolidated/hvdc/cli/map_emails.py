"""
이메일 매핑 CLI - 기존 comprehensive_email_mapper 래핑
"""
from __future__ import annotations
from pathlib import Path
from typing import List, Dict, Any
import logging
from ..scanner.fs_scanner import scan_folder
from ..scanner.email_reader import read_email_file
from ..parser.subject_parser import parse_subject
from ..report.excel import create_excel_report
from ..report.timeline import create_timeline_data, create_network_data, create_summary_stats
from ..core.config import EMAIL_ROOT
from ..core.errors import ScanError, IoError


def map_emails_to_ontology(email_root: Path = None) -> Dict[str, Any]:
    """
    이메일을 온톨로지에 매핑하는 메인 함수
    
    Args:
        email_root: 이메일 루트 폴더 (기본값: EMAIL_ROOT)
        
    Returns:
        Dict[str, Any]: 매핑 결과
    """
    if email_root is None:
        email_root = EMAIL_ROOT
    
    logging.info(f"이메일 폴더 스캔 시작: {email_root}")
    
    try:
        # 1. 폴더 스캔
        files = scan_folder(email_root)
        logging.info(f"스캔된 파일 수: {len(files)}")
        
        # 2. 이메일 데이터 추출
        email_data = []
        for file_path in files:
            try:
                email_content = read_email_file(file_path)
                parsed_data = parse_subject(email_content['subject'])
                
                email_item = {
                    'file_path': str(file_path),
                    'subject': email_content['subject'],
                    'sender': email_content['sender'],
                    'date': email_content['date'],
                    'folder': str(file_path.parent),
                    'cases': parsed_data['cases'],
                    'sites': parsed_data['sites'],
                    'lpos': parsed_data['lpos'],
                    'phases': parsed_data['phases']
                }
                email_data.append(email_item)
                
            except Exception as e:
                logging.warning(f"파일 처리 실패: {file_path} - {e}")
                continue
        
        # 3. 보고서 생성
        excel_path = create_excel_report(email_data)
        timeline_df = create_timeline_data(email_data)
        network_data = create_network_data(email_data)
        summary_stats = create_summary_stats(email_data)
        
        result = {
            'total_files': len(files),
            'processed_emails': len(email_data),
            'excel_path': str(excel_path),
            'timeline_data': timeline_df,
            'network_data': network_data,
            'summary_stats': summary_stats
        }
        
        logging.info(f"매핑 완료: {len(email_data)}개 이메일 처리")
        return result
        
    except Exception as e:
        logging.error(f"이메일 매핑 실패: {e}")
        raise


def main():
    """CLI 진입점"""
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    
    try:
        result = map_emails_to_ontology()
        print(f"✅ 이메일 매핑 완료!")
        print(f"📁 처리된 파일: {result['total_files']}개")
        print(f"📧 처리된 이메일: {result['processed_emails']}개")
        print(f"📊 엑셀 보고서: {result['excel_path']}")
        print(f"🎯 고유 케이스: {result['summary_stats']['unique_cases']}개")
        print(f"🏗️ 고유 사이트: {result['summary_stats']['unique_sites']}개")
        
    except Exception as e:
        print(f"❌ 오류 발생: {e}")
        return 1
    
    return 0


if __name__ == "__main__":
    exit(main())
