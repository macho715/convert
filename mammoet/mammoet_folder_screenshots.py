"""
MAMMOET 폴더 구조 스크린샷 생성 프로그램
- 전체 폴더 구조 트리 이미지
- 개별 폴더별 내용 이미지
"""

import os
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
from typing import List, Dict, Tuple
from datetime import datetime
import textwrap

class FolderScreenshotGenerator:
    def __init__(self, base_folder: str, output_dir: str = None):
        self.base_folder = Path(base_folder)
        if output_dir is None:
            self.output_dir = self.base_folder.parent / "screenshots"
        else:
            self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        # 폰트 설정 (Windows 기본 폰트)
        try:
            self.font_large = ImageFont.truetype("arial.ttf", 24)
            self.font_medium = ImageFont.truetype("arial.ttf", 18)
            self.font_small = ImageFont.truetype("arial.ttf", 14)
            self.font_tiny = ImageFont.truetype("arial.ttf", 12)
        except:
            # 폰트 로드 실패 시 기본 폰트 사용
            self.font_large = ImageFont.load_default()
            self.font_medium = ImageFont.load_default()
            self.font_small = ImageFont.load_default()
            self.font_tiny = ImageFont.load_default()
    
    def get_folder_structure(self) -> List[Dict]:
        """폴더 구조 수집"""
        folders = []
        
        for item in sorted(self.base_folder.iterdir()):
            if item.is_dir() and item.name[0].isdigit():
                files = []
                for file in sorted(item.iterdir()):
                    if file.is_file():
                        files.append({
                            'name': file.name,
                            'size': file.stat().st_size,
                            'ext': file.suffix.lower()
                        })
                
                folders.append({
                    'name': item.name,
                    'path': str(item),
                    'files': files,
                    'file_count': len(files)
                })
        
        return folders
    
    def get_text_size(self, text: str, font: ImageFont) -> Tuple[int, int]:
        """텍스트 크기 계산"""
        # PIL의 textbbox 사용 (최신 버전)
        try:
            bbox = ImageDraw.Draw(Image.new('RGB', (1, 1))).textbbox((0, 0), text, font=font)
            return bbox[2] - bbox[0], bbox[3] - bbox[1]
        except:
            # 구버전 호환
            draw = ImageDraw.Draw(Image.new('RGB', (1, 1)))
            return draw.textsize(text, font=font)
    
    def format_file_size(self, size: int) -> str:
        """파일 크기 포맷팅"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024.0:
                return f"{size:.1f} {unit}"
            size /= 1024.0
        return f"{size:.1f} TB"
    
    def create_folder_tree_image(self, folders: List[Dict]) -> Image.Image:
        """전체 폴더 구조 트리 이미지 생성"""
        # 이미지 크기 계산
        padding = 20
        line_height = 30
        header_height = 80
        footer_height = 40
        
        # 텍스트 라인 생성
        lines = []
        lines.append(f"MAMMOET Mina Zayed Manpower - 2026 - Part 1")
        lines.append(f"폴더 구조 스크린샷 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append("")
        
        for folder in folders:
            folder_name = folder['name']
            file_count = folder['file_count']
            lines.append(f"[FOLDER] {folder_name} ({file_count}개 파일)")
            
            # 파일 목록 (최대 5개만 표시)
            for i, file_info in enumerate(folder['files'][:5]):
                file_name = file_info['name']
                file_size = self.format_file_size(file_info['size'])
                ext = file_info['ext']
                
                # 파일명이 길면 자르기
                if len(file_name) > 60:
                    file_name = file_name[:57] + "..."
                
                lines.append(f"   ├─ {file_name} ({file_size})")
            
            if len(folder['files']) > 5:
                lines.append(f"   └─ ... 외 {len(folder['files']) - 5}개 파일")
            
            lines.append("")
        
        # 이미지 크기 계산
        max_width = 0
        for line in lines:
            width, _ = self.get_text_size(line, self.font_small)
            max_width = max(max_width, width)
        
        total_height = header_height + (len(lines) * line_height) + footer_height
        img_width = max(1200, max_width + padding * 2)
        img_height = total_height
        
        # 이미지 생성
        img = Image.new('RGB', (img_width, img_height), color='white')
        draw = ImageDraw.Draw(img)
        
        # 배경 그리기
        draw.rectangle([0, 0, img_width, header_height], fill='#2c3e50')
        draw.rectangle([0, header_height, img_width, img_height], fill='#ecf0f1')
        
        # 헤더 텍스트
        title = "MAMMOET 폴더 구조"
        title_width, title_height = self.get_text_size(title, self.font_large)
        draw.text(
            ((img_width - title_width) // 2, 20),
            title,
            fill='white',
            font=self.font_large
        )
        
        subtitle = f"{len(folders)}개 폴더 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        subtitle_width, _ = self.get_text_size(subtitle, self.font_small)
        draw.text(
            ((img_width - subtitle_width) // 2, 50),
            subtitle,
            fill='#bdc3c7',
            font=self.font_small
        )
        
        # 본문 텍스트
        y = header_height + padding
        for line in lines:
            if line.startswith("[FOLDER]"):
                # 폴더명은 굵게
                draw.text((padding, y), line, fill='#2c3e50', font=self.font_medium)
            elif line.startswith("   "):
                # 파일명은 작게
                draw.text((padding, y), line, fill='#34495e', font=self.font_tiny)
            elif line:
                # 일반 텍스트
                draw.text((padding, y), line, fill='#34495e', font=self.font_small)
            
            y += line_height
        
        # 푸터
        footer_text = f"생성일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        footer_width, _ = self.get_text_size(footer_text, self.font_tiny)
        draw.text(
            (img_width - footer_width - padding, img_height - 30),
            footer_text,
            fill='#7f8c8d',
            font=self.font_tiny
        )
        
        return img
    
    def create_individual_folder_image(self, folder: Dict) -> Image.Image:
        """개별 폴더 내용 이미지 생성"""
        padding = 30
        line_height = 35
        header_height = 100
        footer_height = 50
        
        # 헤더 정보
        folder_name = folder['name']
        file_count = folder['file_count']
        
        # 파일 목록 생성
        lines = []
        lines.append("파일 목록:")
        lines.append("")
        
        total_size = 0
        file_types = {}
        
        for file_info in folder['files']:
            file_name = file_info['name']
            file_size = file_info['size']
            ext = file_info['ext']
            
            total_size += file_size
            
            # 파일 타입 카운트
            if ext:
                file_types[ext] = file_types.get(ext, 0) + 1
            
            # 파일명이 길면 자르기
            display_name = file_name
            if len(display_name) > 70:
                display_name = display_name[:67] + "..."
            
            size_str = self.format_file_size(file_size)
            lines.append(f"  [FILE] {display_name}")
            lines.append(f"     크기: {size_str} | 확장자: {ext or '없음'}")
            lines.append("")
        
        # 통계 정보
        stats_lines = []
        stats_lines.append("")
        stats_lines.append("=" * 50)
        stats_lines.append("통계 정보:")
        stats_lines.append(f"  총 파일 수: {file_count}개")
        stats_lines.append(f"  총 크기: {self.format_file_size(total_size)}")
        stats_lines.append(f"  파일 타입:")
        for ext, count in sorted(file_types.items()):
            stats_lines.append(f"    {ext}: {count}개")
        
        lines.extend(stats_lines)
        
        # 이미지 크기 계산
        max_width = 0
        for line in lines:
            width, _ = self.get_text_size(line, self.font_small)
            max_width = max(max_width, width)
        
        total_height = header_height + (len(lines) * line_height) + footer_height
        img_width = max(1400, max_width + padding * 2)
        img_height = min(3000, total_height)  # 최대 높이 제한
        
        # 이미지 생성
        img = Image.new('RGB', (img_width, img_height), color='white')
        draw = ImageDraw.Draw(img)
        
        # 배경
        draw.rectangle([0, 0, img_width, header_height], fill='#3498db')
        draw.rectangle([0, header_height, img_width, img_height], fill='#ecf0f1')
        
        # 헤더
        title = f"[FOLDER] {folder_name}"
        title_width, title_height = self.get_text_size(title, self.font_large)
        draw.text(
            ((img_width - title_width) // 2, 25),
            title,
            fill='white',
            font=self.font_large
        )
        
        subtitle = f"{file_count}개 파일"
        subtitle_width, _ = self.get_text_size(subtitle, self.font_medium)
        draw.text(
            ((img_width - subtitle_width) // 2, 60),
            subtitle,
            fill='#ecf0f1',
            font=self.font_medium
        )
        
        # 본문
        y = header_height + padding
        for line in lines:
            if line.startswith("  [FILE]"):
                # 파일명
                draw.text((padding, y), line, fill='#2c3e50', font=self.font_medium)
            elif line.startswith("[FILE]"):
                # 파일명 (들여쓰기 없음)
                draw.text((padding, y), line, fill='#2c3e50', font=self.font_medium)
            elif line.startswith("     크기:"):
                # 파일 정보
                draw.text((padding, y), line, fill='#7f8c8d', font=self.font_tiny)
            elif line.startswith("="):
                # 구분선
                draw.text((padding, y), line, fill='#95a5a6', font=self.font_small)
            elif line.startswith("통계 정보:") or line.startswith("  총") or line.startswith("  파일 타입:"):
                # 통계 헤더
                draw.text((padding, y), line, fill='#2c3e50', font=self.font_medium)
            elif line.startswith("    "):
                # 통계 세부
                draw.text((padding, y), line, fill='#34495e', font=self.font_small)
            elif line:
                # 일반 텍스트
                draw.text((padding, y), line, fill='#34495e', font=self.font_small)
            
            y += line_height
            
            # 최대 높이 초과 시 중단
            if y > img_height - footer_height:
                draw.text((padding, y), "... (내용이 너무 많아 일부만 표시)", 
                         fill='#e74c3c', font=self.font_small)
                break
        
        # 푸터
        footer_text = f"생성일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        footer_width, _ = self.get_text_size(footer_text, self.font_tiny)
        draw.text(
            (img_width - footer_width - padding, img_height - 35),
            footer_text,
            fill='#7f8c8d',
            font=self.font_tiny
        )
        
        return img
    
    def generate_all_screenshots(self):
        """모든 스크린샷 생성"""
        print("=" * 80)
        print("[SCREENSHOT] MAMMOET 폴더 스크린샷 생성 시작")
        print("=" * 80)
        
        # 폴더 구조 수집
        print("\n[1] 폴더 구조 수집 중...")
        folders = self.get_folder_structure()
        print(f"   [OK] {len(folders)}개 폴더 발견")
        
        # 전체 폴더 구조 이미지
        print("\n[2] 전체 폴더 구조 이미지 생성 중...")
        tree_img = self.create_folder_tree_image(folders)
        tree_path = self.output_dir / "00_folder_structure_overview.png"
        tree_img.save(tree_path, 'PNG', quality=95)
        print(f"   [OK] 저장 완료: {tree_path}")
        
        # 개별 폴더 이미지
        print("\n[3] 개별 폴더 이미지 생성 중...")
        for i, folder in enumerate(folders, 1):
            print(f"   [{i}/{len(folders)}] {folder['name']} 처리 중...")
            folder_img = self.create_individual_folder_image(folder)
            
            # 파일명에서 특수문자 제거
            safe_name = folder['name'].replace(':', '_').replace('/', '_').replace('\\', '_')
            folder_path = self.output_dir / f"{i:02d}_{safe_name}.png"
            folder_img.save(folder_path, 'PNG', quality=95)
            print(f"      [OK] 저장 완료: {folder_path.name}")
        
        print("\n" + "=" * 80)
        print("[SUCCESS] 모든 스크린샷 생성 완료")
        print(f"   출력 디렉토리: {self.output_dir}")
        print("=" * 80)
        
        # 파일명 목록 생성
        file_list = ['00_folder_structure_overview.png']
        for i, folder in enumerate(folders, 1):
            safe_name = folder['name'].replace(':', '_').replace('/', '_').replace('\\', '_')
            file_list.append(f"{i:02d}_{safe_name}.png")
        
        return {
            'total_folders': len(folders),
            'output_dir': str(self.output_dir),
            'files': file_list
        }

def main():
    """메인 함수"""
    base_folder = r"C:\Users\SAMSUNG\Downloads\CONVERT\mammoet\Mammoet Mina Zayed Manpower - 2026 - Part 1"
    output_dir = r"C:\Users\SAMSUNG\Downloads\CONVERT\mammoet\screenshots"
    
    generator = FolderScreenshotGenerator(base_folder, output_dir)
    result = generator.generate_all_screenshots()
    
    print(f"\n생성된 파일:")
    for file in result['files']:
        print(f"  - {file}")

if __name__ == "__main__":
    main()