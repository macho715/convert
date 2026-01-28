#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HVDC 프로젝트 설정 및 설치 스크립트
"""

import subprocess
import sys
import os
from pathlib import Path

def install_requirements():
    """필수 패키지 설치"""
    print("📦 필수 패키지 설치 중...")
    
    try:
        subprocess.check_call([
            sys.executable, "-m", "pip", "install", "-r", "requirements.txt"
        ])
        print("✅ 패키지 설치 완료")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ 패키지 설치 실패: {e}")
        return False

def create_directories():
    """필요한 디렉토리 생성"""
    print("📁 디렉토리 구조 생성 중...")
    
    directories = [
        "data",
        "output", 
        "logs",
        "config",
        "utils",
        "tests"
    ]
    
    for directory in directories:
        Path(directory).mkdir(exist_ok=True)
        print(f"  ✅ {directory}/")
    
    print("✅ 디렉토리 생성 완료")

def create_config_files():
    """설정 파일 생성"""
    print("⚙️ 설정 파일 생성 중...")
    
    # config/settings.py
    config_content = '''#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HVDC 프로젝트 설정 파일
"""

import os
from pathlib import Path

# 기본 경로 설정
BASE_DIR = Path(__file__).parent.parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
LOGS_DIR = BASE_DIR / "logs"

# 이메일 폴더 경로
EMAIL_FOLDER_PATH = r"C:\\Users\\SAMSUNG\\Documents\\EMAIL"

# 케이스 번호 패턴
CASE_PATTERNS = {
    'hvdc_adopt': r'HVDC-ADOPT-([A-Z]+)-([A-Z0-9\-]+)',
    'hvdc_project': r'HVDC-([A-Z]+)-([A-Z]+)-([A-Z0-9\-]+)',
    'parentheses': r'\\(([^)]+)\\)',
    'jptw_grm': r'\\[HVDC-AGI\\].*?(JPTW-(\\d+))\\s*/\\s*(GRM-(\\d+))',
    'colon_format': r':\\s*([A-Z]+-[A-Z]+-[A-Z]+\\d+-[A-Z]+\\d+)'
}

# 날짜 패턴
DATE_PATTERNS = [
    r'\\d{4}-\\d{2}-\\d{2}',  # YYYY-MM-DD
    r'\\d{4}\\.\\d{2}\\.\\d{2}',  # YYYY.MM.DD
    r'\\d{2}-\\d{2}-\\d{4}',  # MM-DD-YYYY
    r'\\d{2}\\.\\d{2}\\.\\d{4}',  # MM.DD.YYYY
    r'\\d{4}/\\d{2}/\\d{2}',  # YYYY/MM/DD
    r'\\d{2}/\\d{2}/\\d{4}',  # MM/DD/YYYY
    r'\\d{4}\\d{2}\\d{2}'  # YYYYMMDD
]

# 사이트 코드
SITE_CODES = ['DAS', 'AGI', 'MIR', 'MIRFA', 'GHALLAN', 'SHU']

# 로깅 설정
LOGGING_CONFIG = {
    'version': 1,
    'disable_existing_loggers': False,
    'formatters': {
        'standard': {
            'format': '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        },
    },
    'handlers': {
        'default': {
            'level': 'INFO',
            'formatter': 'standard',
            'class': 'logging.StreamHandler',
        },
        'file': {
            'level': 'INFO',
            'formatter': 'standard',
            'class': 'logging.FileHandler',
            'filename': str(LOGS_DIR / 'hvdc.log'),
            'mode': 'a',
        },
    },
    'loggers': {
        '': {
            'handlers': ['default', 'file'],
            'level': 'INFO',
            'propagate': False
        }
    }
}
'''
    
    with open("config/settings.py", "w", encoding="utf-8") as f:
        f.write(config_content)
    
    print("  ✅ config/settings.py")
    
    # utils/__init__.py
    with open("utils/__init__.py", "w", encoding="utf-8") as f:
        f.write("# HVDC 프로젝트 유틸리티 모듈")
    
    print("  ✅ utils/__init__.py")
    
    print("✅ 설정 파일 생성 완료")

def main():
    """메인 설정 함수"""
    print("🚀 HVDC 프로젝트 설정 시작")
    print("="*50)
    
    # 1. 디렉토리 생성
    create_directories()
    
    # 2. 설정 파일 생성
    create_config_files()
    
    # 3. 패키지 설치
    if install_requirements():
        print("\n🎉 HVDC 프로젝트 설정 완료!")
        print("\n📋 다음 단계:")
        print("1. python run_all_scripts.py --interactive  # 대화형 실행")
        print("2. python run_all_scripts.py --all          # 전체 스크립트 실행")
        print("3. python run_all_scripts.py --required     # 필수 스크립트만 실행")
    else:
        print("\n❌ 설정 중 오류가 발생했습니다.")
        print("수동으로 다음 명령어를 실행하세요:")
        print("pip install -r requirements.txt")

if __name__ == "__main__":
    main()
