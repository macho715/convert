#!/usr/bin/env python3
"""
CONVERT 폴더 정리 분석기

임시 파일 스캔, 중복 파일 탐지, 폴더 구조 분석을 수행합니다.
기본적으로 dry-run 모드로 실행되며, 실제 삭제는 명시적 승인 후에만 수행합니다.
"""
import argparse
import json
import os
import re
import subprocess
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

# 임시 파일 패턴
TEMP_PATTERNS = [
    r"__pycache__",
    r"~\$.*\.xlsx?$",  # Excel 임시 파일
    r"\.pyc$",
    r"\.DS_Store$",
    r"Thumbs\.db$",
    r"\.tmp$",
    r"\.log$",  # 로그 파일 (선택적)
]

# 보호 대상 파일/폴더
PROTECTED_PATTERNS = [
    r"\.git",
    r"AGENTS\.md$",
    r"README\.md$",
    r"requirements\.txt$",
    r"pyproject\.toml$",
    r"setup\.cfg$",
    r"\.env$",
    r"\.env\.example$",
]

# 핵심 문서/설정 파일
CRITICAL_FILES = {
    "AGENTS.md",
    "README.md",
    "requirements.txt",
    "pyproject.toml",
    "setup.cfg",
    ".gitignore",
}


def is_git_tracked(filepath: str, root: str) -> bool:
    """Git 추적 파일인지 확인"""
    try:
        result = subprocess.run(
            ["git", "ls-files", "--error-unmatch", filepath],
            cwd=root,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        return result.returncode == 0
    except (subprocess.CalledProcessError, FileNotFoundError):
        # Git이 없거나 오류 발생 시 보수적으로 True 반환 (보호)
        return True


def matches_pattern(filename: str, patterns: list) -> bool:
    """파일명이 패턴 목록과 일치하는지 확인"""
    for pattern in patterns:
        if re.search(pattern, filename, re.IGNORECASE):
            return True
    return False


def is_protected(filepath: str, root: str) -> bool:
    """파일이 보호 대상인지 확인"""
    filename = os.path.basename(filepath)
    
    # 핵심 파일 직접 확인
    if filename in CRITICAL_FILES:
        return True
    
    # 패턴 매칭
    if matches_pattern(filename, PROTECTED_PATTERNS):
        return True
    
    # Git 추적 파일
    if is_git_tracked(filepath, root):
        return True
    
    return False


def is_temp_file(filepath: str) -> bool:
    """임시 파일인지 확인"""
    filename = os.path.basename(filepath)
    dirname = os.path.basename(os.path.dirname(filepath))
    
    # __pycache__ 디렉토리
    if dirname == "__pycache__":
        return True
    
    # 파일명 패턴 매칭
    if matches_pattern(filename, TEMP_PATTERNS):
        return True
    
    return False


def get_file_size(filepath: str) -> int:
    """파일 크기 반환 (바이트)"""
    try:
        return os.path.getsize(filepath)
    except OSError:
        return 0


def format_size(size: int) -> str:
    """파일 크기를 읽기 쉬운 형식으로 변환"""
    for unit in ["B", "KB", "MB", "GB"]:
        if size < 1024.0:
            return f"{size:.1f} {unit}"
        size /= 1024.0
    return f"{size:.1f} TB"


def scan_directory(root: str, exclude_dirs: set = None):
    """디렉토리 스캔 및 파일 분류"""
    if exclude_dirs is None:
        exclude_dirs = {".git", ".venv", "node_modules", "dist", "build", "__pycache__"}
    
    results = {
        "temp_files": [],
        "duplicates": defaultdict(list),
        "protected": [],
        "stats": {
            "total_files": 0,
            "temp_count": 0,
            "duplicate_count": 0,
            "protected_count": 0,
            "total_size": 0,
            "temp_size": 0,
        },
    }
    
    root_path = Path(root).resolve()
    
    for dirpath, dirnames, filenames in os.walk(root):
        # 제외 디렉토리 스킵
        parts = set(Path(dirpath).parts)
        if any(part in exclude_dirs for part in parts):
            continue
        
        for filename in filenames:
            filepath = os.path.join(dirpath, filename)
            rel_path = os.path.relpath(filepath, root)
            
            # 보호 대상 확인
            if is_protected(filepath, root):
                results["protected"].append({
                    "path": rel_path,
                    "reason": "protected",
                    "size": get_file_size(filepath),
                })
                results["stats"]["protected_count"] += 1
                continue
            
            results["stats"]["total_files"] += 1
            file_size = get_file_size(filepath)
            results["stats"]["total_size"] += file_size
            
            # 임시 파일 확인
            if is_temp_file(filepath):
                results["temp_files"].append({
                    "path": rel_path,
                    "size": file_size,
                    "type": "temp",
                    "risk": "LOW",
                })
                results["stats"]["temp_count"] += 1
                results["stats"]["temp_size"] += file_size
            else:
                # 중복 파일 탐지 (파일명 기준)
                results["duplicates"][filename].append({
                    "path": rel_path,
                    "size": file_size,
                })
    
    # 중복 파일 정리 (2개 이상인 경우만)
    duplicates_final = {
        name: paths
        for name, paths in results["duplicates"].items()
        if len(paths) > 1
    }
    results["duplicates"] = duplicates_final
    results["stats"]["duplicate_count"] = sum(len(paths) - 1 for paths in duplicates_final.values())
    
    return results


def generate_report(results: dict, root: str, dry_run: bool = True) -> dict:
    """리포트 생성"""
    report = {
        "generated_at": datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
        "root": os.path.abspath(root),
        "dry_run": dry_run,
        "summary": {
            "total_files": results["stats"]["total_files"],
            "temp_files": results["stats"]["temp_count"],
            "duplicate_files": results["stats"]["duplicate_count"],
            "protected_files": results["stats"]["protected_count"],
            "total_size": results["stats"]["total_size"],
            "temp_size": results["stats"]["temp_size"],
        },
        "temp_files": results["temp_files"],
        "duplicates": {
            name: paths
            for name, paths in results["duplicates"].items()
        },
        "protected": results["protected"],
    }
    return report


def print_markdown_report(report: dict):
    """Markdown 형식 리포트 출력"""
    print("# 폴더 정리 분석 리포트\n")
    print(f"**생성 시간**: {report['generated_at']}")
    print(f"**대상 경로**: {report['root']}")
    print(f"**모드**: {'DRY-RUN (실제 변경 없음)' if report['dry_run'] else 'EXECUTE (실제 변경)'}\n")
    
    # 요약
    s = report["summary"]
    print("## 요약\n")
    print(f"- 총 파일 수: {s['total_files']:,}")
    print(f"- 임시 파일: {s['temp_files']:,}개 ({format_size(s['temp_size'])})")
    print(f"- 중복 파일: {s['duplicate_files']:,}개")
    print(f"- 보호된 파일: {s['protected_count']:,}개\n")
    
    # 임시 파일
    if report["temp_files"]:
        print("## 임시 파일 (안전 삭제 가능)\n")
        print("| 경로 | 크기 | 위험도 |")
        print("| --- | --- | --- |")
        for item in sorted(report["temp_files"], key=lambda x: x["size"], reverse=True)[:50]:
            print(f"| `{item['path']}` | {format_size(item['size'])} | {item['risk']} |")
        if len(report["temp_files"]) > 50:
            print(f"\n*총 {len(report['temp_files'])}개 중 상위 50개만 표시*\n")
    
    # 중복 파일
    if report["duplicates"]:
        print("## 중복 파일 (검토 필요)\n")
        print("| 파일명 | 경로 수 | 경로 목록 |")
        print("| --- | --- | --- |")
        for name, paths in sorted(report["duplicates"].items(), key=lambda x: len(x[1]), reverse=True)[:20]:
            path_list = ", ".join([f"`{p['path']}`" for p in paths[:3]])
            if len(paths) > 3:
                path_list += f", ... (총 {len(paths)}개)"
            print(f"| `{name}` | {len(paths)} | {path_list} |")
        if len(report["duplicates"]) > 20:
            print(f"\n*총 {len(report['duplicates'])}개 중 상위 20개만 표시*\n")
    
    # 보호된 파일
    if report["protected"]:
        print("## 보호된 파일 (삭제 불가)\n")
        print(f"*총 {len(report['protected'])}개 파일이 보호 목록에 포함되어 있습니다.*\n")


def main():
    ap = argparse.ArgumentParser(
        description="CONVERT 폴더 정리 분석기 (안전 우선, dry-run 기본)"
    )
    ap.add_argument("--root", default=".", help="스캔할 루트 디렉토리")
    ap.add_argument("--out", default="", help="JSON 리포트 출력 경로")
    ap.add_argument(
        "--dry-run",
        action="store_true",
        default=True,
        help="Dry-run 모드 (실제 변경 없음, 기본값)",
    )
    ap.add_argument(
        "--execute",
        action="store_true",
        help="실제 실행 모드 (명시적 지정 필요)",
    )
    ap.add_argument(
        "--confirm",
        action="store_true",
        help="확인 없이 실행 (위험, 권장하지 않음)",
    )
    
    args = ap.parse_args()
    
    root = os.path.abspath(args.root)
    dry_run = not args.execute
    
    if args.execute and not args.confirm:
        print("⚠️  WARNING: 실제 실행 모드는 위험합니다.")
        print("실제로 파일을 삭제하려면 --confirm 플래그도 함께 지정하세요.")
        print("현재는 dry-run 모드로 실행합니다.\n")
        dry_run = True
    
    # Windows 콘솔 인코딩 설정
    if sys.platform == "win32":
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    
    # 스캔 수행
    print("📁 폴더 스캔 중...")
    results = scan_directory(root)
    
    # 리포트 생성
    report = generate_report(results, root, dry_run)
    
    # Markdown 리포트 출력
    print_markdown_report(report)
    
    # JSON 리포트 저장
    if args.out:
        os.makedirs(os.path.dirname(args.out) if os.path.dirname(args.out) else ".", exist_ok=True)
        with open(args.out, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)
        print(f"\n✅ JSON 리포트 저장: {args.out}")
    
    # 실행 모드인 경우 (현재는 구현하지 않음, 안전을 위해)
    if args.execute and args.confirm:
        print("\n[WARNING] 실제 삭제 기능은 안전을 위해 구현되지 않았습니다.")
        print("삭제가 필요하면 리포트를 검토한 후 수동으로 수행하세요.")
    
    # 종료 코드
    if results["stats"]["temp_count"] > 0 or results["stats"]["duplicate_count"] > 0:
        print(f"\n[INFO] 정리 가능한 항목이 발견되었습니다.")
        if dry_run:
            print("   실제 정리를 수행하려면 --execute --confirm을 사용하세요.")
        sys.exit(0)
    else:
        print("\n[OK] 정리할 항목이 없습니다.")
        sys.exit(0)


if __name__ == "__main__":
    main()
