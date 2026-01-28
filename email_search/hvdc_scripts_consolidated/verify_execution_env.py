#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HVDC PST 스캐너 실행 환경 검증 스크립트

실행 전 필수 환경 요구사항을 자동으로 확인합니다.
"""

import sys
import os
from pathlib import Path
import subprocess

# 색상 코드 (Windows 호환)
class Colors:
    GREEN = '\033[92m'
    RED = '\033[91m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    END = '\033[0m'
    BOLD = '\033[1m'

def print_header(text):
    """헤더 출력"""
    print(f"\n{Colors.BOLD}{Colors.BLUE}{'='*60}{Colors.END}")
    print(f"{Colors.BOLD}{Colors.BLUE}{text}{Colors.END}")
    print(f"{Colors.BOLD}{Colors.BLUE}{'='*60}{Colors.END}\n")

def print_success(text):
    """성공 메시지 출력"""
    print(f"{Colors.GREEN}✓ {text}{Colors.END}")

def print_error(text):
    """오류 메시지 출력"""
    print(f"{Colors.RED}✗ {text}{Colors.END}")

def print_warning(text):
    """경고 메시지 출력"""
    print(f"{Colors.YELLOW}⚠ {text}{Colors.END}")

def check_python_version():
    """Python 버전 확인"""
    print("1. Python 버전 확인...")
    version = sys.version_info
    if version >= (3, 11):
        print_success(f"Python {version.major}.{version.minor}.{version.micro} (요구사항: 3.11+)")
        return True
    else:
        print_error(f"Python {version.major}.{version.minor}.{version.micro} (요구사항: 3.11+)")
        return False

def check_module(module_name, install_command=None):
    """Python 모듈 확인"""
    try:
        __import__(module_name)
        print_success(f"{module_name} 모듈 설치됨")
        return True
    except ImportError:
        print_error(f"{module_name} 모듈이 설치되지 않음")
        if install_command:
            print_warning(f"설치 명령: {install_command}")
        return False

def check_pypff():
    """pypff 모듈 확인 (특별 처리)"""
    try:
        import pypff
        print_success("pypff 모듈 설치됨 (libpff-python)")
        return True
    except ImportError:
        print_error("pypff 모듈이 설치되지 않음")
        print_warning("설치 명령: pip install libpff-python")
        return False

def check_directory(path, create_if_missing=True):
    """디렉토리 확인"""
    dir_path = Path(path)
    if dir_path.exists() and dir_path.is_dir():
        print_success(f"'{path}' 디렉토리 존재")
        return True
    else:
        if create_if_missing:
            try:
                dir_path.mkdir(parents=True, exist_ok=True)
                print_success(f"'{path}' 디렉토리 생성됨")
                return True
            except Exception as e:
                print_error(f"'{path}' 디렉토리 생성 실패: {e}")
                return False
        else:
            print_error(f"'{path}' 디렉토리 없음")
            return False

def check_pst_file(pst_path):
    """PST 파일 경로 확인"""
    print(f"3. PST 파일 경로 확인...")
    pst_file = Path(pst_path)
    if pst_file.exists() and pst_file.is_file():
        size_mb = pst_file.stat().st_size / (1024 * 1024)
        print_success(f"PST 파일 존재: {pst_path}")
        print(f"  파일 크기: {size_mb:.2f} MB")
        return True
    else:
        print_error(f"PST 파일 없음: {pst_path}")
        print_warning("quick_run_2025_06.bat 또는 run_scanner.bat에서 경로 확인 필요")
        return False

def check_outlook_running():
    """Outlook 실행 여부 확인"""
    try:
        result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq outlook.exe'],
                              capture_output=True, text=True)
        if 'outlook.exe' in result.stdout:
            print_warning("Outlook이 실행 중입니다 (스캔 전 자동 종료됨)")
            return True
        else:
            print_success("Outlook이 실행 중이 아닙니다")
            return True
    except Exception as e:
        print_warning(f"Outlook 상태 확인 불가: {e}")
        return True

def main():
    """메인 함수"""
    print_header("HVDC PST 스캐너 환경 검증")
    
    checks = []
    
    # Python 버전 확인
    checks.append(("Python 버전", check_python_version()))
    
    # 필수 모듈 확인
    print("\n2. 필수 모듈 확인...")
    required_modules = [
        ("pandas", "pip install pandas"),
        ("numpy", "pip install numpy"),
        ("openpyxl", "pip install openpyxl"),
    ]
    
    for module, install_cmd in required_modules:
        checks.append((f"{module} 모듈", check_module(module, install_cmd)))
    
    # pypff 확인 (특별 처리)
    checks.append(("pypff 모듈", check_pypff()))
    
    # 디렉토리 확인
    print("\n4. 디렉토리 구조 확인...")
    directories = [
        "results",
        "output",
        "output/logs",
        "output/data",
        "output/reports",
    ]
    
    for directory in directories:
        checks.append((f"{directory} 디렉토리", check_directory(directory)))
    
    # PST 파일 확인
    pst_path = r"C:\Users\SAMSUNG\Documents\Outlook 파일\minkyu.cha@samsung.comswe - outlook.RECOVERED.20251002-092839.pst"
    checks.append(("PST 파일", check_pst_file(pst_path)))
    
    # Outlook 상태 확인
    print("\n5. Outlook 상태 확인...")
    checks.append(("Outlook 상태", check_outlook_running()))
    
    # 결과 요약
    print_header("검증 결과 요약")
    
    passed = sum(1 for _, result in checks if result)
    total = len(checks)
    
    print(f"\n총 {total}개 항목 중 {passed}개 통과 ({passed*100//total}%)\n")
    
    failed_checks = [name for name, result in checks if not result]
    
    if failed_checks:
        print_error("다음 항목을 확인하세요:")
        for check in failed_checks:
            print(f"  - {check}")
        print("\n실행 계획 문서 참조: docs/EXECUTION_PLAN.md")
        return 1
    else:
        print_success("모든 검증 항목 통과!")
        print("\n다음 단계:")
        print("  1. quick_run_2025_06.bat 실행 (2025년 6월 데이터)")
        print("  2. 또는 run_scanner.bat 실행 (자동 모드)")
        return 0

if __name__ == "__main__":
    sys.exit(main())

