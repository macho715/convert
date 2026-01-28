# 🔒 Outlook PST 안전 스캔 가이드

## ⚠️ PST 파일 손상 방지를 위한 필수 사항

Outlook의 PST 파일은 손상되기 쉬우며, 한 번 손상되면 복구가 어렵습니다.
이 스캐너는 **읽기 전용 모드**로 설계되어 PST 파일 손상을 방지합니다.

---

## 🛡️ 안전 장치

### 1. **읽기 전용 접근**
- ✅ 메일 제목, 발신자, 날짜, 첨부파일 기본 정보만 읽음
- ❌ 메일 본문은 절대 읽지 않음 (PST 부하 방지)
- ❌ 메일 이동/삭제/수정 금지
- ❌ 플래그, 카테고리 등 메타데이터 변경 금지

### 2. **배치 처리**
- 한 번에 50개씩 배치로 처리
- 배치마다 메모리 자동 정리
- 대용량 스캔에도 안정적

### 3. **첨부파일 안전 처리**
- 파일 이름, 크기, 유형만 추출
- 파일 내용은 읽지 않음
- 50MB 이상 큰 파일은 자동 스킵
- 최대 10개까지만 추출

### 4. **COM 객체 관리**
- 사용 후 즉시 릴리즈
- 메모리 누수 방지
- 가비지 컬렉션 자동 실행

---

## 📋 스캔 범위

### ✅ 안전하게 스캔하는 정보
```
📧 기본 정보
  - 제목 (Subject)
  - 발신자 이름 (SenderName)
  - 발신자 이메일 (SenderEmailAddress)
  - 수신 시간 (ReceivedTime)
  - 폴더 이름

📎 첨부파일 기본 정보
  - 파일 이름 (FileName)
  - 파일 크기 (Size)
  - 파일 유형 (확장자)
  - 첨부파일 개수
```

### ❌ 절대 읽지 않는 정보
```
🚫 위험한 정보 (PST 부하)
  - Body (텍스트 본문)
  - HTMLBody (HTML 본문)
  - RTFBody (RTF 본문)
  - Attachments.Item().PropertyAccessor
  - 첨부파일 내용
```

---

## 🚀 안전한 사용법

### 1. 기본 스캔 (권장)
```bash
# PST 안전 모드로 최근 30일 스캔
python run_scan.py --source outlook --date-range 30 --fallback
```

### 2. 제한적 스캔 (매우 안전)
```bash
# 최대 500개만 스캔
python run_scan.py --source outlook --max-emails 500 --fallback
```

### 3. 특정 폴더만 스캔
```bash
# Inbox만 스캔
python run_scan.py --source outlook --folders Inbox --fallback
```

### 4. 연결 테스트 (스캔 전 권장)
```bash
# Outlook 연결 및 PST 상태 확인
python run_scan.py --test-outlook
```

---

## ⚡ 권장 설정

### 소량 스캔 (안전)
```bash
python run_scan.py \
  --source outlook \
  --max-emails 1000 \
  --date-range 30 \
  --folders Inbox "Sent Items" \
  --fallback
```

### 대량 스캔 (주의)
```bash
# 10,000개 이상은 시간이 오래 걸립니다
python run_scan.py \
  --source outlook \
  --max-emails 10000 \
  --date-range 180 \
  --fallback
```

---

## 🔍 스캔 결과 예시

```
🔍 Outlook 메일 정보 스캔 시작...

✅ Outlook 연결 성공 (받은 편지함: 1234개 메일)
🔒 PST 안전 모드: 읽기 전용으로 동작합니다
📅 날짜 필터: 2024-09-26 이후
📁 기본 폴더만 스캔 (PST 안전): ['Inbox', 'Sent Items']

📧 폴더 'Inbox' 스캔 시작 (1234개 메일)
⏳ 진행 중... 100개 메일 처리됨
⏳ 진행 중... 200개 메일 처리됨
✅ 폴더 'Inbox' 완료: 234개 메일 처리

🎉 스캔 완료: 총 456개 메일

✅ 스캔 완료! 456개 메일

🎯 추출된 케이스: 23개

📋 케이스 목록 (처음 10개):
  1. HVDC-2024-001
  2. HVDC-2024-002
  ...

📎 첨부파일 통계:
  - 첨부파일 있는 메일: 123개
  - 총 첨부파일 수: 234개
  - 총 용량: 456.78 MB

  파일 유형별 통계:
    1. .pdf: 89개
    2. .xlsx: 45개
    3. .docx: 34개
    4. .jpg: 23개
    5. .msg: 12개

📧 발신자 통계 (상위 5명):
  1. John Doe: 45개
  2. Jane Smith: 32개
  ...
```

---

## 🚨 주의사항

### DO ✅
1. **스캔 전에 Outlook을 실행하세요**
2. **작은 범위부터 테스트하세요** (--max-emails 100)
3. **--fallback 옵션을 항상 사용하세요**
4. **--test-outlook으로 먼저 테스트하세요**
5. **네트워크가 안정적인지 확인하세요** (Exchange 사용 시)

### DON'T ❌
1. **Outlook이 동기화 중일 때 스캔하지 마세요**
2. **PST 파일이 손상된 상태로 스캔하지 마세요**
3. **동시에 여러 스캔을 실행하지 마세요**
4. **너무 큰 범위를 한 번에 스캔하지 마세요**
5. **안전 모드를 끄지 마세요**

---

## 🛠️ PST 파일 복구 (손상 시)

### 1. ScanPST.exe 사용 (Outlook 2021 LTS Pro)
```
⚠️ Outlook 2021 LTS Pro 경로:

64비트:
C:\Program Files\Microsoft Office\root\Office16\SCANPST.EXE

32비트 (드물지만):
C:\Program Files (x86)\Microsoft Office\root\Office16\SCANPST.EXE

실행 방법:
1. Outlook 완전 종료 (작업 관리자에서도 확인)
2. Windows 검색에서 "SCANPST" 입력
3. 또는 위 경로에서 직접 실행
4. PST 파일 선택
5. "시작" 클릭하여 스캔
6. 오류 발견 시 "복구" 클릭
7. 복구 완료 후 Outlook 재시작
```

### 2. PST 파일 위치 찾기 (Outlook 2021)
```
Outlook 2021 기본 위치:

데이터 파일:
C:\Users\[사용자명]\Documents\Outlook Files\*.pst

또는:
C:\Users\[사용자명]\AppData\Local\Microsoft\Outlook\*.pst

Outlook에서 확인:
파일 → 계정 설정 → 계정 설정 → 데이터 파일 탭
```

### 3. 백업 권장
```bash
# 스캔 전에 PST 파일 백업 (필수!)
복사: C:\Users\SAMSUNG\Documents\Outlook Files\*.pst
대상: D:\Backup\Outlook\[날짜]\

# 또는 Outlook 내보내기 사용
Outlook → 파일 → 열기/내보내기 → 가져오기/내보내기
```

---

## 📊 성능 가이드

### 예상 스캔 시간
- 100개 메일: ~10초
- 1,000개 메일: ~1-2분
- 10,000개 메일: ~10-20분
- 50,000개 메일: ~1-2시간

### 메모리 사용량
- 안전 모드: ~100-200MB
- 일반 모드: ~500MB-1GB

---

## 🔧 문제 해결

### Q: "COM 초기화 실패" 오류
**A:** Outlook을 완전히 종료하고 다시 실행하세요

### Q: "폴더 접근 실패" 오류
**A:** Outlook 프로필이 제대로 설정되었는지 확인하세요

### Q: "MAPI 네임스페이스 획득 실패" 오류
**A:** 관리자 권한으로 실행하거나 Outlook을 재설치하세요

### Q: 스캔이 너무 느림
**A:** 
1. --max-emails로 제한하세요
2. --date-range로 최근 메일만 스캔하세요
3. 특정 폴더만 지정하세요

### Q: PST 파일이 손상됨
**A:**
1. 즉시 스캔 중단
2. Outlook 완전 종료
3. ScanPST.exe로 복구 시도
4. 백업에서 복원

---

## 📞 지원

문제가 발생하면:
1. 로그 파일 확인: `hvdc_scripts_execution.log`
2. --test-outlook으로 진단
3. --fallback 옵션 사용
4. 작은 범위로 재시도

---

## 🎯 요약

✅ **안전한 사용**
```bash
python run_scan.py \
  --source outlook \
  --max-emails 1000 \
  --date-range 30 \
  --test-outlook    # 먼저 테스트
  --fallback        # 실패 시 자동 전환
```

이 설정으로 PST 파일을 안전하게 보호하면서 필요한 정보를 추출할 수 있습니다! 🚀
