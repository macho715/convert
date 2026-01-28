# HVDC TR Transportation - Document Tracker VBA 사용 가이드

## 📋 개요
본 Excel 파일은 HVDC 변압기 운송 프로젝트의 서류 준비 현황을 추적하기 위한 도구입니다.
VBA 매크로를 통해 자동화된 상태 관리, 진행률 계산, 보고서 생성 기능을 제공합니다.

---

## 🔧 VBA 모듈 설치 방법

### 방법 1: .bas 파일 가져오기
1. Excel 파일을 `.xlsm` 형식으로 저장 (매크로 사용 통합 문서)
2. `Alt + F11` 키를 눌러 VBA 편집기 열기
3. 메뉴에서 `파일 → 가져오기` 선택
4. `TR_DocTracker_VBA_Module.bas` 파일 선택
5. `파일 → 닫고 Microsoft Excel로 돌아가기`
6. 파일 저장

### 방법 2: 직접 복사/붙여넣기
1. Excel 파일을 `.xlsm` 형식으로 저장
2. `Alt + F11` 키를 눌러 VBA 편집기 열기
3. 왼쪽 프로젝트 창에서 `VBAProject` 우클릭 → `삽입 → 모듈`
4. `.bas` 파일 내용을 새 모듈에 복사/붙여넣기
5. 저장

---

## ⚡ 매크로 기능 목록

### 1. TR_ApplyStatusFormatting (상태별 색상 적용)
- **기능**: 상태 값에 따라 자동으로 색상 서식 적용
- **실행**: `Alt + F8` → `TR_ApplyStatusFormatting` 선택 → 실행
- **색상 규칙**:
  - 🟢 Complete: 연두색
  - 🟡 In Progress: 연노랑
  - 🔴 Not Started: 연빨강
  - 🔵 Pending Review: 연파랑
  - ⚪ N/A: 회색

### 2. TR_CalculateProgress (진행률 계산)
- **기능**: 각 항차별 서류 준비 진행률 계산
- **결과**: Dashboard 시트의 Voyage Status 열 자동 업데이트

### 3. TR_CalculatePartyProgress (파티별 진행률)
- **기능**: 담당 파티별 서류 제출 현황 집계
- **결과**: Dashboard 시트의 Party-wise Progress 테이블 업데이트

### 4. TR_HighlightOverdue (마감일 초과 강조)
- **기능**: 마감일이 지났으나 완료되지 않은 서류 강조
- **표시**: 분홍색 배경으로 해당 행 강조

### 5. EXP_ExportToPDF (PDF 내보내기)
- **기능**: 현재 시트를 PDF로 저장
- **저장 위치**: Excel 파일과 동일한 폴더

### 6. EXP_SendEmailReminder (이메일 알림)
- **기능**: Outlook을 통해 미완료 서류 알림 이메일 생성
- **요구사항**: Microsoft Outlook 설치 필요

### 7. TR_AutoRefreshAll (전체 새로고침)
- **기능**: 모든 기능을 한 번에 실행
- **추천**: 주기적으로 실행하여 데이터 동기화

### 8. TR_CreateVoyageSummaryReport (요약 보고서 생성)
- **기능**: 항차별 서류 현황 요약 시트 자동 생성
- **결과**: `Summary_Report` 시트 생성

### 9. TR_QuickStatusUpdate (빠른 상태 변경)
- **기능**: 선택한 셀의 상태를 숫자 입력으로 빠르게 변경
- **사용법**: 상태 셀 선택 → 매크로 실행 → 숫자 입력

### 10. TR_FilterByParty (파티별 필터)
- **기능**: 특정 담당 파티의 서류만 표시
- **사용법**: 매크로 실행 → 파티 번호 입력

---

## ⌨️ 키보드 단축키 (선택 설정)

VBA 편집기에서 `ThisWorkbook` 모듈에 아래 코드를 추가하면 단축키 사용 가능:

```vba
Private Sub Workbook_Open()
    Application.OnKey "^+R", "TR_AutoRefreshAll"      ' Ctrl+Shift+R
    Application.OnKey "^+S", "TR_QuickStatusUpdate"   ' Ctrl+Shift+S
    Application.OnKey "^+P", "EXP_ExportToPDF"         ' Ctrl+Shift+P
End Sub
```

---

## 📊 시트 구조

### 1. Dashboard
- 프로젝트 개요 및 일정
- 항차별 진행 현황
- 파티별 진행률
- 주요 연락처

### 2. Document_Tracker (메인)
- 전체 서류 목록 (50개 항목)
- 항차별 상태 관리 (V1~V4)
- 날짜 및 비고 입력

### 3. Voyage_1 ~ Voyage_4
- 개별 항차 상세 정보
- 일정 및 서류 체크리스트

### 4. Gate_Pass_MZP
- Mina Zayed Port 게이트 패스
- 장비 및 인원 별도 관리

### 5. Instructions
- 사용법 및 범례

---

## 📅 주요 일정 (2026년)

| 항차 | TR Units | MZP 도착 | 서류 마감 | Land Permit 신청 |
|------|----------|----------|-----------|------------------|
| 1차 | TR 1-2 | 01-27 | 01-23 | ~01-22 |
| 2차 | TR 3-4 | 02-06 | 02-03 | ~02-02 |
| 3차 | TR 5-6 | 02-15 | 02-12 | ~02-11 |
| 4차 | TR 7 | 02-24 | 02-20 | ~02-19 |

---

## 📧 담당자 연락처

| Party | 담당자 | Email |
|-------|--------|-------|
| OFCO Agency | Nanda Kumar | nkk@ofco-int.com |
| OFCO Agency | Das Gopal | das@ofco-int.com |
| ADNOC L&S | Mahmoud Ouda | moda@adnoc.ae |
| Mammoet | Yulia Frolova | Yulia.Frolova@mammoet.com |
| DSV Solutions | Jay Manaloto | jay.manaloto@dsv.com |
| LCT Bushra | Vessel Ops | lct.bushra@khalidfarajshipping.com |

---

## ⚠️ 주의사항

1. **매크로 활성화**: 파일을 열 때 "매크로 사용"을 허용해야 VBA 기능 사용 가능
2. **파일 형식**: 반드시 `.xlsm` 형식으로 저장 (xlsx는 매크로 저장 불가)
3. **백업**: 중요 변경 전 파일 백업 권장
4. **Outlook**: 이메일 기능은 Outlook 설치/로그인 필요

---

## 🔄 업데이트 이력

| 버전 | 날짜 | 변경 내용 |
|------|------|-----------|
| 1.0 | 2026-01-19 | 초기 버전 생성 |

---

**작성**: Samsung C&T - HVDC Project Team  
**문의**: Project Manager (Cha)
