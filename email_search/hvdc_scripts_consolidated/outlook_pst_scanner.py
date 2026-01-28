#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook PST Scanner v5.0 (libpst 기반)
개별 폴더 선택 + 날짜 필터링 + 대용량 최적화

🔒 PST 안전 스캔 가이드:
- 완전한 읽기 전용 접근 (libpst 기반)
- PST 파일 절대 수정 안 함
- Outlook 프로세스 불필요
- 대용량 PST 처리 가능 (60GB+)

⚠️ 사용 전 확인사항:
1. Outlook 자동 종료 (스크립트가 자동 처리)
2. PST 파일 백업 권장
3. 충분한 디스크 공간 (결과 파일용)

📊 출력 형식:
- 파일명: OUTLOOK_YYYYMM.xlsx
- 위치: results/ 폴더
- 시트: 전체_이메일, 폴더별_통계, 발신자별_통계

🚀 빠른 실행:
  python outlook_pst_scanner.py --pst "경로" --start 2025-06-01 --end 2025-06-30 --folders all --auto
  
📖 상세 가이드: docs/PST_SAFETY_GUIDE.md
"""

import sys
import os
from pathlib import Path
import pandas as pd
from datetime import datetime, timedelta
import time
import re
import argparse
import subprocess

try:
    import pypff  # libpst Python 바인딩
except ImportError:
    print("❌ pypff 모듈이 설치되지 않았습니다")
    print("설치 방법: pip install libpff-python")
    sys.exit(1)


class FolderSelectPSTScanner:
    """
    폴더 선택 PST 스캐너
    - 폴더 목록 표시 + 선택 UI
    - 날짜 필터링
    - 메모리 최적화
    """
    
    def __init__(self, start_date=None, end_date=None, 
                 max_body_length=500, batch_size=1000):
        """
        Args:
            start_date: 시작 날짜 (datetime 객체)
            end_date: 종료 날짜 (datetime 객체)
            max_body_length: 본문 최대 길이
            batch_size: 배치 저장 크기
        """
        self.pst_file = None
        self.email_data = []
        self.folder_list = []  # 전체 폴더 목록
        
        # 날짜 필터링
        self.start_date = start_date
        self.end_date = end_date
        if end_date:
            self.end_date = end_date.replace(hour=23, minute=59, second=59)
        
        # 최적화 설정
        self.max_body_length = max_body_length
        self.batch_size = batch_size
        
        # 통계
        self.total_scanned = 0
        self.total_matched = 0
        self.total_skipped = 0
        self.start_time = None
        self.last_report_time = None
    
    def close_outlook(self):
        """Outlook 프로세스 강제 종료"""
        print("\n🔄 Outlook 종료 중...")
        try:
            result = subprocess.run(['taskkill', '/F', '/IM', 'outlook.exe', '/T'], 
                                  capture_output=True, text=True)
            if "성공" in result.stdout or "SUCCESS" in result.stdout:
                print("✅ Outlook이 종료되었습니다")
            else:
                print("ℹ️ Outlook이 실행 중이 아닙니다")
            time.sleep(2)  # 프로세스 완전 종료 대기
            return True
        except Exception as e:
            print(f"⚠️ Outlook 종료 시도 중 오류: {e}")
            return False
        
    def open_pst_readonly(self, pst_path):
        """PST 파일 열기 (읽기 전용)"""
        # Outlook 종료 (2차 안전장치)
        self.close_outlook()
        
        print(f"\n📂 PST 파일 열기: {pst_path}")
        
        try:
            self.pst_file = pypff.file()
            self.pst_file.open(pst_path)
            
            print(f"✅ PST 파일 열림")
            try:
                root_folder = self.pst_file.get_root_folder()
                print(f"   루트 폴더: {root_folder.name if root_folder else '(알 수 없음)'}")
            except:
                print(f"   루트 폴더: (접근 불가)")
            
            return True
            
        except Exception as e:
            print(f"❌ PST 파일 열기 실패: {e}")
            return False
    
    def list_all_folders(self, folder, path="", depth=0):
        """
        모든 폴더를 재귀적으로 탐색하여 목록 생성
        Returns: [(index, name, path, num_messages, folder_object)]
        """
        current_path = f"{path}/{folder.name}" if path else folder.name
        
        try:
            num_messages = folder.get_number_of_sub_messages()
        except:
            num_messages = 0
        
        # 현재 폴더 추가
        folder_info = {
            'index': len(self.folder_list),
            'name': folder.name,
            'path': current_path,
            'messages': num_messages,
            'depth': depth,
            'folder_obj': folder
        }
        self.folder_list.append(folder_info)
        
        # 하위 폴더 재귀 탐색
        try:
            num_subfolders = folder.get_number_of_sub_folders()
            for i in range(num_subfolders):
                try:
                    subfolder = folder.get_sub_folder(i)
                    self.list_all_folders(subfolder, current_path, depth + 1)
                except Exception as e:
                    pass
        except Exception as e:
            pass
    
    def display_folder_menu(self):
        """폴더 목록을 보기 좋게 표시"""
        print("\n" + "="*70)
        print("📁 PST 폴더 목록")
        print("="*70)
        
        print(f"\n{'번호':<6} {'폴더명':<40} {'메시지':<10}")
        print("-" * 70)
        
        for folder in self.folder_list:
            indent = "  " * folder['depth']
            folder_name = f"{indent}{folder['name']}"
            
            # 긴 이름 줄이기
            if len(folder_name) > 38:
                folder_name = folder_name[:35] + "..."
            
            print(f"{folder['index']:<6} {folder_name:<40} {folder['messages']:>8,}개")
        
        print("-" * 70)
        total_messages = sum(f['messages'] for f in self.folder_list)
        print(f"{'총계':<6} {'':<40} {total_messages:>8,}개")
        print("="*70)
    
    def select_folders(self):
        """
        사용자가 폴더를 선택
        Returns: 선택된 폴더 인덱스 리스트
        """
        print("\n📌 폴더 선택 방법:")
        print("   - 단일 선택: 3")
        print("   - 복수 선택: 1,3,5")
        print("   - 범위 선택: 1-5")
        print("   - 조합 가능: 1,3-5,7")
        print("   - 전체 선택: all")
        print("   - 특정 폴더 제외: all,-2,-5 (전체에서 2,5번 제외)")
        
        while True:
            user_input = input("\n폴더 선택: ").strip().lower()
            
            if not user_input:
                print("⚠️  입력이 없습니다. 다시 입력해주세요.")
                continue
            
            try:
                selected = set()
                
                if user_input == 'all':
                    # 전체 선택
                    selected = set(range(len(self.folder_list)))
                elif user_input.startswith('all,'):
                    # 전체에서 제외
                    selected = set(range(len(self.folder_list)))
                    exclude_part = user_input[4:]  # "all," 제거
                    exclude_indices = self._parse_selection(exclude_part)
                    selected -= set(exclude_indices)
                else:
                    # 일반 선택
                    selected = set(self._parse_selection(user_input))
                
                # 유효성 검사
                max_index = len(self.folder_list) - 1
                invalid = [i for i in selected if i < 0 or i > max_index]
                
                if invalid:
                    print(f"⚠️  잘못된 번호: {invalid}")
                    print(f"   유효 범위: 0 ~ {max_index}")
                    continue
                
                if not selected:
                    print("⚠️  선택된 폴더가 없습니다.")
                    continue
                
                # 선택 확인
                selected_list = sorted(list(selected))
                print(f"\n✅ 선택된 폴더 ({len(selected_list)}개):")
                for idx in selected_list[:10]:  # 최대 10개만 표시
                    folder = self.folder_list[idx]
                    print(f"   [{idx}] {folder['name']} ({folder['messages']:,}개)")
                
                if len(selected_list) > 10:
                    print(f"   ... 외 {len(selected_list)-10}개")
                
                total_msgs = sum(self.folder_list[i]['messages'] for i in selected_list)
                print(f"\n   총 메시지: {total_msgs:,}개")
                
                confirm = input("\n이대로 진행하시겠습니까? (y/n): ").strip().lower()
                if confirm == 'y':
                    return selected_list
                else:
                    print("\n다시 선택해주세요.")
                    
            except Exception as e:
                print(f"⚠️  입력 오류: {e}")
                continue
    
    def _parse_selection(self, selection_str):
        """
        선택 문자열 파싱
        "1,3,5-7" -> [1, 3, 5, 6, 7]
        "-2,-5" -> [-2, -5] (제외용)
        """
        indices = []
        
        parts = selection_str.split(',')
        for part in parts:
            part = part.strip()
            
            if '-' in part and not part.startswith('-'):
                # 범위: 1-5
                start, end = part.split('-')
                indices.extend(range(int(start), int(end) + 1))
            else:
                # 단일 또는 제외: 3 또는 -3
                indices.append(int(part))
        
        return indices
    
    def is_date_in_range(self, dt):
        """날짜가 필터 범위 내인지 확인"""
        if not dt:
            return False
        
        if self.start_date and dt < self.start_date:
            return False
        
        if self.end_date and dt > self.end_date:
            return False
        
        return True
    
    def extract_message_data(self, message):
        """메시지 데이터 추출 (날짜 필터링 포함)"""
        try:
            # 날짜 확인
            delivery_time = None
            creation_time = None
            
            try:
                delivery_time = message.delivery_time
            except:
                pass
            
            try:
                creation_time = message.creation_time
            except:
                pass
            
            # 날짜 필터링
            if self.start_date or self.end_date:
                check_date = delivery_time or creation_time
                if not self.is_date_in_range(check_date):
                    return None
            
            # 제목
            subject = ''
            try:
                subject = message.subject or '(제목 없음)'
            except:
                subject = '(제목 없음)'
            
            # 발신자
            sender_name = ''
            try:
                sender_name = message.sender_name or ''
            except:
                sender_name = ''
            
            # 이메일 주소
            sender_email = ''
            recipient_to = ''
            try:
                headers = message.transport_headers or ''
                from_match = re.search(r'From:\s*([^\r\n]+)', headers, re.IGNORECASE)
                if from_match:
                    sender_email = from_match.group(1).strip()
                to_match = re.search(r'To:\s*([^\r\n]+)', headers, re.IGNORECASE)
                if to_match:
                    recipient_to = to_match.group(1).strip()
            except:
                pass
            
            # 크기 및 첨부파일
            size = 0
            num_attachments = 0
            try:
                size = message.size or 0
            except:
                pass
            try:
                num_attachments = message.number_of_attachments or 0
            except:
                pass
            
            # 본문 (길이 제한)
            plain_body = ''
            html_body = ''
            try:
                body = message.plain_text_body
                if body:
                    body_str = body.decode('utf-8', errors='ignore') if isinstance(body, bytes) else body
                    plain_body = body_str[:self.max_body_length] if self.max_body_length else body_str
            except:
                pass
            
            try:
                html = message.html_body
                if html:
                    html_str = html.decode('utf-8', errors='ignore') if isinstance(html, bytes) else html
                    html_body = html_str[:self.max_body_length] if self.max_body_length else html_str
            except:
                pass
            
            data = {
                'Subject': subject,
                'SenderName': sender_name,
                'SenderEmail': sender_email,
                'RecipientTo': recipient_to,
                'DeliveryTime': delivery_time,
                'CreationTime': creation_time,
                'Size': size,
                'HasAttachments': num_attachments > 0,
                'AttachmentCount': num_attachments,
                'PlainTextBody': plain_body,
                'HTMLBody': html_body,
            }
            
            # 첨부파일 이름
            attachments = []
            if num_attachments > 0:
                for i in range(num_attachments):
                    try:
                        attachment = message.get_attachment(i)
                        att_name = attachment.name or f'attachment_{i}'
                        attachments.append(att_name)
                    except:
                        attachments.append(f'unknown_attachment_{i}')
            data['AttachmentNames'] = '; '.join(attachments)
            
            return data
            
        except Exception as e:
            return None
    
    def print_progress(self, force=False):
        """진행 상황 출력"""
        now = time.time()
        
        if not force and self.last_report_time:
            if now - self.last_report_time < 10:
                return
        
        self.last_report_time = now
        elapsed = now - self.start_time
        speed = self.total_scanned / elapsed if elapsed > 0 else 0
        
        print(f"\n⏳ 진행 상황:")
        print(f"   스캔: {self.total_scanned:,}개")
        print(f"   매칭: {self.total_matched:,}개")
        print(f"   스킵: {self.total_skipped:,}개")
        print(f"   속도: {speed:.1f} 메시지/초")
        print(f"   경과: {elapsed/60:.1f}분")
        print(f"   메모리: {len(self.email_data):,}개")
    
    def save_batch(self, output_file, mode='a'):
        """배치 저장"""
        if not self.email_data:
            return
        
        try:
            df = pd.DataFrame(self.email_data)
            
            # 컬럼 순서 표준화 (HVDC Analyzer 호환)
            column_order = [
                'Subject', 'SenderName', 'SenderEmail', 'RecipientTo',
                'DeliveryTime', 'CreationTime',
                'Size', 'HasAttachments', 'AttachmentCount', 'AttachmentNames',
                'FolderPath', 'PlainTextBody', 'HTMLBody'
            ]
            
            # 존재하는 컬럼만 순서대로 선택
            ordered_columns = [col for col in column_order if col in df.columns]
            # 순서에 없는 추가 컬럼도 포함
            extra_columns = [col for col in df.columns if col not in column_order]
            final_columns = ordered_columns + extra_columns
            
            df = df[final_columns]
            
            file_exists = os.path.exists(output_file)
            
            if mode == 'a' and file_exists:
                existing_df = pd.read_excel(output_file, sheet_name='전체_이메일')
                df = pd.concat([existing_df, df], ignore_index=True)
            
            with pd.ExcelWriter(output_file, engine='openpyxl', mode='w') as writer:
                df.to_excel(writer, sheet_name='전체_이메일', index=False)
                
                # 폴더별 통계
                if 'FolderPath' in df.columns:
                    folder_stats = df.groupby('FolderPath').size().reset_index(name='Count')
                    folder_stats = folder_stats.sort_values('Count', ascending=False)
                    folder_stats.to_excel(writer, sheet_name='폴더별_통계', index=False)
                
                # 발신자별 통계
                if 'SenderEmail' in df.columns:
                    sender_stats = df.groupby('SenderEmail').size().reset_index(name='Count')
                    sender_stats = sender_stats.sort_values('Count', ascending=False)
                    sender_stats.to_excel(writer, sheet_name='발신자별_통계', index=False)
            
            print(f"💾 배치 저장: {len(self.email_data)}개 → {output_file}")
            self.email_data = []
            
        except Exception as e:
            print(f"⚠️ 배치 저장 실패: {e}")
    
    def scan_folder_only(self, folder, folder_path, output_file):
        """단일 폴더만 스캔 (하위 폴더 제외)"""
        print(f"\n📁 스캔: {folder_path}")
        
        try:
            num_messages = folder.get_number_of_sub_messages()
            print(f"   📧 메시지 수: {num_messages}")
            
            for i in range(num_messages):
                try:
                    message = folder.get_sub_message(i)
                    self.total_scanned += 1
                    
                    data = self.extract_message_data(message)
                    
                    if data:
                        data['FolderPath'] = folder_path
                        self.email_data.append(data)
                        self.total_matched += 1
                        
                        if len(self.email_data) >= self.batch_size:
                            self.save_batch(output_file, mode='a')
                    else:
                        self.total_skipped += 1
                    
                    self.print_progress()
                        
                except Exception as e:
                    continue
                    
        except Exception as e:
            print(f"   ❌ 폴더 스캔 오류: {e}")
    
    def analyze_selected(self, pst_path, selected_indices, output_excel):
        """선택된 폴더만 분석"""
        print("\n" + "="*70)
        print("🔍 폴더 선택 PST 스캐너 v5.0")
        
        if self.start_date:
            print(f"   시작: {self.start_date.strftime('%Y-%m-%d')}")
        if self.end_date:
            print(f"   종료: {self.end_date.strftime('%Y-%m-%d')}")
        
        print(f"   선택 폴더: {len(selected_indices)}개")
        print("="*70)
        
        self.start_time = time.time()
        self.last_report_time = self.start_time
        
        try:
            # 선택된 폴더만 스캔
            for idx in selected_indices:
                folder_info = self.folder_list[idx]
                self.scan_folder_only(
                    folder_info['folder_obj'],
                    folder_info['path'],
                    output_excel
                )
            
            # 마지막 배치 저장
            if self.email_data:
                self.save_batch(output_excel, mode='a')
            
            # 최종 결과
            self.print_progress(force=True)
            
            print("\n" + "="*70)
            print(f"✅ 분석 완료")
            print(f"   총 스캔: {self.total_scanned:,}개")
            print(f"   날짜 매칭: {self.total_matched:,}개")
            print(f"   날짜 스킵: {self.total_skipped:,}개")
            
            elapsed = time.time() - self.start_time
            print(f"   소요 시간: {elapsed/60:.1f}분")
            print("="*70)
            
            print(f"\n📊 결과 파일: {output_excel}")
            
            return True
            
        except Exception as e:
            print(f"\n❌ 분석 중 오류: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def run(self, pst_path, output_excel, auto_folders=None, auto_confirm=False):
        """전체 실행 흐름
        
        Args:
            pst_path: PST 파일 경로
            output_excel: 출력 엑셀 파일명
            auto_folders: 자동 폴더 선택 ('all' 또는 None)
            auto_confirm: 자동 확인 (True/False)
        """
        # PST 열기
        if not self.open_pst_readonly(pst_path):
            return False
        
        try:
            # 폴더 목록 생성
            print("\n⏳ 폴더 목록 생성 중...")
            root = self.pst_file.get_root_folder()
            self.list_all_folders(root)
            
            # 폴더 선택
            if auto_folders == 'all':
                # 자동 모드: 모든 폴더 선택
                selected_indices = list(range(len(self.folder_list)))
                print(f"\n✅ 자동 모드: 전체 {len(selected_indices)}개 폴더 선택됨")
            else:
                # 대화형 모드
                self.display_folder_menu()
                selected_indices = self.select_folders()
            
            # 확인 프롬프트
            if not auto_confirm and len(selected_indices) > 0:
                folder_info = self.folder_list[selected_indices[0]]
                total_msgs = sum([self.folder_list[i]['messages'] for i in selected_indices])
                print(f"\n⚠️  {len(selected_indices)}개 폴더, 약 {total_msgs:,}개 메시지 스캔 예정")
                confirm = input("   계속하시겠습니까? (y/n): ").strip().lower()
                if confirm != 'y':
                    print("❌ 사용자가 취소했습니다")
                    return False
            
            # 선택된 폴더 분석
            return self.analyze_selected(pst_path, selected_indices, output_excel)
            
        finally:
            if self.pst_file:
                self.pst_file.close()
                print("\n✅ PST 파일 닫힘")


# 메인 실행
if __name__ == "__main__":
    # Windows 콘솔 인코딩 설정
    import sys
    if sys.platform == 'win32':
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except:
            pass
    
    print("""
================================================================
      폴더 선택 PST 스캐너 v5.0
      개별 폴더 선택 + 날짜 필터링 + 대용량 최적화
================================================================
    """)
    
    # argparse 설정
    parser = argparse.ArgumentParser(
        description='PST 파일 폴더 선택 분석',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
예시:
  # 대화형 모드
  python LIBPST_FOLDER_SELECT_v5.py
  
  # 자동 실행 모드
  python LIBPST_FOLDER_SELECT_v5.py --pst "경로" --start 2025-07-01 --end 2025-07-30 --folders all --auto
        """
    )
    parser.add_argument('--pst', help='PST 파일 경로')
    parser.add_argument('--start', help='시작 날짜 (YYYY-MM-DD)')
    parser.add_argument('--end', help='종료 날짜 (YYYY-MM-DD)')
    parser.add_argument('--folders', default=None, help='폴더 선택 (all 또는 번호)')
    parser.add_argument('--auto', action='store_true', help='확인 없이 자동 실행')
    
    args = parser.parse_args()
    
    # PST 파일 경로
    if args.pst:
        pst_path = args.pst.strip('"')
    else:
        pst_path = input("\n📁 PST 파일 경로: ").strip('"')
    
    if not pst_path:
        print("❌ 경로가 입력되지 않았습니다")
        sys.exit(1)
    
    # 날짜 범위
    start_date = None
    end_date = None
    
    if args.start:
        try:
            start_date = datetime.strptime(args.start, "%Y-%m-%d")
        except ValueError:
            print("⚠️  시작 날짜 형식 오류, 무시됨")
    elif not args.pst:  # 대화형 모드일 때만
        print("\n📅 날짜 범위 설정 (YYYY-MM-DD 형식)")
        print("   (Enter만 누르면 전체 날짜)")
        start_date_str = input("   시작 날짜: ").strip()
        if start_date_str:
            try:
                start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
            except ValueError:
                print("⚠️  시작 날짜 형식 오류, 무시됨")
    
    if args.end:
        try:
            end_date = datetime.strptime(args.end, "%Y-%m-%d")
        except ValueError:
            print("⚠️  종료 날짜 형식 오류, 무시됨")
    elif not args.pst:  # 대화형 모드일 때만
        end_date_str = input("   종료 날짜: ").strip()
        if end_date_str:
            try:
                end_date = datetime.strptime(end_date_str, "%Y-%m-%d")
            except ValueError:
                print("⚠️  종료 날짜 형식 오류, 무시됨")
    
    # 출력 파일명 (OUTLOOK_YYYYMM 형식)
    if start_date:
        year_month = start_date.strftime("%Y%m")  # "202505"
        base_name = f"OUTLOOK_{year_month}"
        
        # 충돌 방지: 기존 파일이 있으면 타임스탬프 추가
        output_file = f"{base_name}.xlsx"
        output_path = Path("results") / output_file
        if output_path.exists():
            timestamp = datetime.now().strftime("%Y%m%d")
            output_file = f"{base_name}_{timestamp}.xlsx"
    else:
        # 날짜 지정 안 된 경우 타임스탬프 사용
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"OUTLOOK_ALL_{timestamp}.xlsx"
    
    # 스캐너 실행
    scanner = FolderSelectPSTScanner(
        start_date=start_date,
        end_date=end_date,
        max_body_length=500,
        batch_size=1000
    )
    
    success = scanner.run(
        pst_path, 
        output_file,
        auto_folders=args.folders,
        auto_confirm=args.auto
    )
    
    if success:
        print("\n✅ 프로그램 정상 종료")
    else:
        print("\n❌ 분석 실패")
    
    if not args.auto:
        input("\n계속하려면 Enter를 누르세요...")
