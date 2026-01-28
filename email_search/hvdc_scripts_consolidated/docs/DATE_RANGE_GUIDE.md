# 📅 날짜 범위 지정 스캔 가이드

Outlook 스캔 시 **특정 날짜 범위**를 지정할 수 있습니다!

---

## 🎯 지원되는 날짜 옵션

### 1. **최근 N일** (`--date-range`)
```bash
# 최근 7일
python run_scan.py --source outlook --date-range 7 --fallback

# 최근 30일 (권장)
python run_scan.py --source outlook --date-range 30 --fallback

# 최근 90일
python run_scan.py --source outlook --date-range 90 --fallback
```

### 2. **시작 날짜 ~ 종료 날짜** (`--start-date`, `--end-date`) ✨ NEW!
```bash
# 2024년 전체
python run_scan.py --source outlook \
  --start-date 2024-01-01 \
  --end-date 2024-12-31 \
  --fallback

# 2024년 10월만
python run_scan.py --source outlook \
  --start-date 2024-10-01 \
  --end-date 2024-10-31 \
  --fallback

# 특정 분기 (Q3 2024)
python run_scan.py --source outlook \
  --start-date 2024-07-01 \
  --end-date 2024-09-30 \
  --fallback
```

### 3. **시작 날짜부터 현재까지**
```bash
# 2024년 1월 1일부터 지금까지
python run_scan.py --source outlook \
  --start-date 2024-01-01 \
  --fallback
```

### 4. **특정 날짜까지**
```bash
# 2024년 12월 31일까지
python run_scan.py --source outlook \
  --end-date 2024-12-31 \
  --fallback
```

---

## 📋 날짜 형식

**형식**: `YYYY-MM-DD` (필수)

✅ **올바른 예시:**
```
2024-01-01
2024-10-26
2023-12-31
```

❌ **잘못된 예시:**
```
2024/01/01    (슬래시 사용 금지)
01-01-2024    (순서 틀림)
2024-1-1      (두 자리 필수)
24-01-01      (네 자리 연도 필수)
```

---

## 🎨 실전 예제

### 프로젝트 진행 기간 스캔
```bash
# HVDC 프로젝트 기간 (2024년 상반기)
python run_scan.py --source outlook \
  --start-date 2024-01-01 \
  --end-date 2024-06-30 \
  --max-emails 5000 \
  --fallback
```

### 월별 스캔
```bash
# 2024년 10월 스캔
python run_scan.py --source outlook \
  --start-date 2024-10-01 \
  --end-date 2024-10-31 \
  --folders Inbox "Sent Items" \
  --fallback
```

### 분기별 스캔
```bash
# Q4 2024 (10월~12월)
python run_scan.py --source outlook \
  --start-date 2024-10-01 \
  --end-date 2024-12-31 \
  --fallback
```

### 특정 계약 기간 스캔
```bash
# 2024년 3월 15일 ~ 2024년 9월 15일
python run_scan.py --source outlook \
  --start-date 2024-03-15 \
  --end-date 2024-09-15 \
  --fallback
```

---

## ⚠️ 주의사항

### 1. **날짜 옵션 충돌**
```bash
# ❌ 잘못됨: --date-range와 --start-date 동시 사용
python run_scan.py --source outlook \
  --date-range 30 \
  --start-date 2024-01-01  # 에러!

# ✅ 올바름: 둘 중 하나만 사용
python run_scan.py --source outlook \
  --start-date 2024-01-01 \
  --end-date 2024-12-31
```

**메시지:**
```
⚠️ --date-range와 --start-date/--end-date를 동시에 사용할 수 없습니다.
   --start-date/--end-date를 사용합니다.
```

### 2. **날짜 순서**
```bash
# ❌ 잘못됨: 시작이 종료보다 늦음
python run_scan.py --source outlook \
  --start-date 2024-12-31 \
  --end-date 2024-01-01  # 에러!
```

**메시지:**
```
❌ 날짜 형식 오류: 시작 날짜가 종료 날짜보다 늦습니다
```

### 3. **미래 날짜**
```bash
# ⚠️ 경고: 종료 날짜가 미래
python run_scan.py --source outlook \
  --start-date 2024-01-01 \
  --end-date 2026-12-31  # 경고
```

**메시지:**
```
⚠️ 종료 날짜가 미래입니다. 오늘 날짜로 조정합니다.
📅 시작 날짜: 2024-01-01 (Monday)
📅 종료 날짜: 2025-10-26 (Sunday)
```

---

## 📊 예상 출력

```bash
python run_scan.py --source outlook \
  --start-date 2024-01-01 \
  --end-date 2024-03-31 \
  --max-emails 1000 \
  --fallback
```

**출력:**
```
🔍 Outlook 메일 정보 스캔 시작...

📅 시작 날짜: 2024-01-01 (Monday)
📅 종료 날짜: 2024-03-31 (Sunday)

✅ Outlook 연결 성공 (받은 편지함: 1234개 메일)
🔒 PST 안전 모드 활성화
📅 시작 날짜: 2024-01-01
📅 종료 날짜: 2024-03-31
📁 기본 폴더만 스캔: ['Inbox', 'Sent Items']

📧 폴더 'Inbox' 스캔 시작 (1234개 메일)
⏳ 진행 중... 100개 메일 처리됨
⏳ 진행 중... 200개 메일 처리됨
✅ 폴더 'Inbox' 완료: 234개 메일 처리

🎉 스캔 완료: 총 456개 메일 (2024-01-01 ~ 2024-03-31)

✅ 스캔 완료! 456개 메일

🎯 추출된 케이스: 23개
📋 케이스 목록:
  1. HVDC-2024-001
  2. HVDC-2024-002
  ...
```

---

## 🎯 권장 사용 패턴

### 일상 업무 (최근 30일)
```bash
python run_scan.py --source outlook --date-range 30 --fallback
```

### 월별 보고서 작성
```bash
# 이번 달
python run_scan.py --source outlook \
  --start-date 2024-10-01 \
  --end-date 2024-10-31 \
  --fallback

# 지난 달
python run_scan.py --source outlook \
  --start-date 2024-09-01 \
  --end-date 2024-09-30 \
  --fallback
```

### 프로젝트 기간별 분석
```bash
# 프로젝트 Phase 1
python run_scan.py --source outlook \
  --start-date 2024-01-01 \
  --end-date 2024-03-31 \
  --folders Inbox \
  --fallback

# 프로젝트 Phase 2
python run_scan.py --source outlook \
  --start-date 2024-04-01 \
  --end-date 2024-06-30 \
  --folders Inbox \
  --fallback
```

### 감사/규정 준수
```bash
# 연간 감사 (2024년 전체)
python run_scan.py --source outlook \
  --start-date 2024-01-01 \
  --end-date 2024-12-31 \
  --max-emails 10000 \
  --fallback
```

---

## 💡 팁 & 트릭

### 1. **대용량 날짜 범위는 max-emails로 제한**
```bash
# 1년치 메일 중 최대 5000개만
python run_scan.py --source outlook \
  --start-date 2024-01-01 \
  --end-date 2024-12-31 \
  --max-emails 5000 \
  --fallback
```

### 2. **특정 폴더만 스캔으로 속도 향상**
```bash
# Inbox만 스캔
python run_scan.py --source outlook \
  --start-date 2024-01-01 \
  --end-date 2024-12-31 \
  --folders Inbox \
  --fallback
```

### 3. **단계별 스캔 (큰 범위는 나누기)**
```bash
# 1분기
python run_scan.py --source outlook --start-date 2024-01-01 --end-date 2024-03-31 --fallback

# 2분기
python run_scan.py --source outlook --start-date 2024-04-01 --end-date 2024-06-30 --fallback

# 3분기
python run_scan.py --source outlook --start-date 2024-07-01 --end-date 2024-09-30 --fallback

# 4분기
python run_scan.py --source outlook --start-date 2024-10-01 --end-date 2024-12-31 --fallback
```

---

## 🔧 문제 해결

### Q: 날짜 형식 오류가 계속 발생
**A:** 형식을 정확히 확인하세요
```bash
# ❌ 잘못됨
--start-date 2024/10/26
--start-date 26-10-2024
--start-date 2024-10-26 00:00:00

# ✅ 올바름
--start-date 2024-10-26
```

### Q: 너무 오래된 메일은 안나옴
**A:** Outlook이 오래된 메일을 보관/삭제했을 수 있습니다
```
파일 → 옵션 → 고급 → 자동 보관 설정 확인
```

### Q: 날짜 지정했는데 다른 날짜 메일도 나옴
**A:** 메일의 ReceivedTime 필드가 없거나 잘못되었을 수 있습니다
```
이런 메일은 자동으로 스킵됩니다 (PST 안전)
```

---

## 📚 관련 문서

- [PST_SAFETY_GUIDE.md](PST_SAFETY_GUIDE.md) - PST 안전 가이드
- [OUTLOOK_2021_GUIDE.md](OUTLOOK_2021_GUIDE.md) - Outlook 2021 전용 가이드
- [OUTLOOK_SCANNER_README.md](OUTLOOK_SCANNER_README.md) - 기본 사용법

---

## 🎊 완료!

이제 **정확한 날짜 범위**를 지정해서 필요한 기간의 메일만 스캔할 수 있습니다! 🚀

```bash
# 시작하세요!
python run_scan.py --source outlook --start-date 2024-01-01 --end-date 2024-12-31 --fallback
```
