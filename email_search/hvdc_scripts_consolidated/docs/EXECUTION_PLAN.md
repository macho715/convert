# HVDC PST 스캐너 실행 계획

**최종 업데이트**: 2025-01-XX  
**프로젝트**: HVDC Email Processing System  
**버전**: v1.0

## 개요

이 문서는 HVDC PST 스캐너(`outlook_pst_scanner.py`)를 실행하기 위한 단계별 계획과 검증 절차를 제공합니다.

## 1. 사전 준비 검증

### 1.1 환경 요구사항 확인

#### Windows 버전
- 요구사항: Windows 10/11
- 확인 방법: `winver` 명령 실행

#### Python 버전
- 요구사항: Python 3.11 이상
- 확인 명령: `python --version`
- 확인 결과: Python 3.11.8 (✓ 충족)

#### 메모리 요구사항
- 최소: 4GB RAM
- 권장: 대용량 PST 처리 시 8GB RAM

#### 디스크 공간
- 최소: 1GB 여유 공간 (결과 파일 저장용)

### 1.2 의존성 설치 확인

#### 필수 패키지 설치
```bash
pip install -r requirements.txt
```

#### 필수 패키지 목록
- pandas >= 1.5.0
- numpy >= 1.21.0
- openpyxl >= 3.0.0
- python-dateutil >= 2.8.0

#### libpff-python 설치 (PST 스캔 필수)
```bash
pip install libpff-python
```

**검증 결과** (2025-01-XX):
- pandas: 2.3.3 (✓ 설치됨)
- numpy: 1.26.4 (✓ 설치됨)
- libpff-python: 20231205 (✓ 설치됨)

**확인 명령**:
```bash
pip list | findstr /i "pandas numpy libpff pypff"
python -c "import pypff; print('pypff module: OK')"
```

### 1.3 PST 파일 경로 확인

#### 하드코딩된 경로 (배치 파일)
두 배치 파일 모두 다음 경로를 사용합니다:
```
C:\Users\SAMSUNG\Documents\Outlook 파일\minkyu.cha@samsung.comswe - outlook.RECOVERED.20251002-092839.pst
```

**파일 존재 확인**:
```powershell
Test-Path "C:\Users\SAMSUNG\Documents\Outlook 파일\minkyu.cha@samsung.comswe - outlook.RECOVERED.20251002-092839.pst"
```

**검증 결과**: ✓ 파일 존재 확인됨

**경로 변경이 필요한 경우**:
1. `quick_run_2025_06.bat` 파일 열기
2. 21번 라인의 `set PST_PATH=` 값 수정
3. `run_scanner.bat` 파일 열기
4. 22번 라인의 `set PST_PATH=` 값 수정

### 1.4 디렉토리 구조 확인

#### 필수 디렉토리
- `results/`: 스캔 결과 Excel 파일 저장 위치
- `output/logs/`: 로그 파일 저장 위치
- `output/data/`: 중간 데이터 파일 저장 위치
- `output/reports/`: 보고서 파일 저장 위치

**디렉토리 생성 (없는 경우)**:
```bash
mkdir results
mkdir output\logs
mkdir output\data
mkdir output\reports
```

**검증 결과**:
- ✓ `results/` 폴더 존재
- ✓ `output/` 폴더 및 하위 폴더 존재

## 2. 실행 옵션별 절차

### 2.1 옵션 A: 빠른 실행 (2025년 6월 데이터)

**파일**: `quick_run_2025_06.bat`

#### 실행 전 확인사항
1. Outlook 실행 여부 확인 (자동 종료되지만 확인 권장)
2. PST 파일 경로 확인
3. 디스크 여유 공간 확인

#### 실행 단계

1. **Outlook 자동 종료**
   - 배치 파일이 자동으로 Outlook을 종료합니다
   - 명령: `taskkill /F /IM outlook.exe /T`

2. **설정 확인**
   - 시작 날짜: `2025-06-01`
   - 종료 날짜: `2025-06-30`
   - 배치 크기: `5,000` (Standard)
   - 폴더: `all` (프로그램에서 선택 가능)

3. **사용자 확인**
   - 배치 파일 실행 시 확인 프롬프트 표시
   - `y` 입력 시 실행, `n` 입력 시 취소

4. **Python 스크립트 실행**
   ```bash
   python outlook_pst_scanner.py --pst "%PST_PATH%" --start 2025-06-01 --end 2025-06-30 --folders all --batch-size 5000
   ```

5. **실행 완료 대기**
   - 스캔 진행 상황이 콘솔에 표시됩니다
   - 완료 메시지 표시 후 `pause`로 대기

#### 예상 결과 파일

`results/` 폴더에 다음 파일이 생성됩니다:

1. **기본 스캔 결과**
   - `OUTLOOK_202506.xlsx`
   - 또는 타임스탬프 포함: `OUTLOOK_202506_YYYYMMDD.xlsx`
   - 시트:
     - `전체_이메일`: 스캔된 모든 이메일 데이터
     - `폴더별_통계`: 폴더별 이메일 통계
     - `발신자별_통계`: 발신자별 이메일 통계

2. **HVDC 온톨로지 분석** (자동 실행되는 경우)
   - `OUTLOOK_HVDC_ONTOLOGY_202506.xlsx`

3. **HVDC 보고서** (자동 실행되는 경우)
   - `OUTLOOK_HVDC_REPORT_202506.xlsx`

#### 실행 예시
```batch
===============================================================
       PST Scanner v6.0 - Quick Run (2025 June)
===============================================================

PST File: C:\Users\SAMSUNG\Documents\Outlook 파일\minkyu.cha@samsung.comswe - outlook.RECOVERED.20251002-092839.pst

Settings:
   - Start date: 2025-06-01
   - End date: 2025-06-30
   - Batch size: 5,000 (Standard)
   - Folders: Select in program

Continue with these settings? (y/n): y

[Running PST Scanner...]
```

### 2.2 옵션 B: 자동 모드 실행

**파일**: `run_scanner.bat`

#### 실행 단계

1. **Outlook 자동 종료**
   - 배치 파일이 자동으로 Outlook을 종료합니다

2. **PST 파일 경로 확인**
   - 하드코딩된 경로 사용

3. **자동 모드 실행**
   ```bash
   python outlook_pst_scanner.py --pst "%PST_PATH%" --auto
   ```

4. **프로그램 내 설정**
   - 프로그램이 실행되면 날짜 범위 입력 프롬프트 표시
   - 폴더 선택 옵션 제공
   - 설정 후 스캔 시작

#### 실행 예시
```batch
===============================================================
       PST Scanner v6.0 - Simple Run
===============================================================

PST File: C:\Users\SAMSUNG\Documents\Outlook 파일\minkyu.cha@samsung.comswe - outlook.RECOVERED.20251002-092839.pst

Starting PST Scanner...
```

## 3. 실행 후 확인 사항

### 3.1 결과 파일 확인

#### Excel 파일 생성 확인
1. `results/` 폴더 열기
2. 다음 파일명 패턴 확인:
   - `OUTLOOK_YYYYMM.xlsx`
   - `OUTLOOK_HVDC_ONTOLOGY_YYYYMM.xlsx`
   - `OUTLOOK_HVDC_REPORT_YYYYMM.xlsx`

#### 파일 내용 확인
1. 파일 크기 확인 (빈 파일 여부)
2. Excel 파일 열어서 시트 확인
3. 데이터 행 수 확인

**확인 명령** (PowerShell):
```powershell
Get-ChildItem results\OUTLOOK_202506*.xlsx | Select-Object Name, Length, LastWriteTime
```

### 3.2 로그 파일 확인

#### 로그 파일 위치
- `output/logs/email_scan_YYYYMMDD_HHMMSS.log`

#### 확인 내용
1. 스캔 시작/종료 시간
2. 처리된 이메일 수
3. 오류 메시지 (있는 경우)
4. 경고 메시지 (있는 경우)

**확인 명령**:
```powershell
Get-Content output\logs\email_scan_*.log | Select-Object -Last 50
```

### 3.3 HVDC 분석 결과 확인

#### 온톨로지 분석 파일
- `results/OUTLOOK_HVDC_ONTOLOGY_YYYYMM.xlsx`
- 추출된 케이스 번호 확인
- 사이트 매핑 확인

#### 보고서 파일
- `results/OUTLOOK_HVDC_REPORT_YYYYMM.xlsx`
- 요약 통계 확인

## 4. 문제 해결 체크리스트

### 4.1 Outlook 종료 실패

#### 증상
- PST 파일이 잠겨 있음
- 파일 접근 오류

#### 해결 방법

1. **수동 Outlook 종료**
   ```bash
   taskkill /F /IM outlook.exe /T
   ```

2. **프로세스 확인**
   ```bash
   tasklist | findstr outlook
   ```

3. **PST 파일 사용 중인 프로세스 확인**
   - 작업 관리자에서 확인
   - 다른 프로그램이 파일을 열고 있는지 확인

### 4.2 PST 파일 접근 오류

#### 증상
- `FileNotFoundError`
- `PermissionError`

#### 해결 방법

1. **파일 경로 확인**
   - 배치 파일의 `PST_PATH` 값 확인
   - 파일 실제 존재 여부 확인

2. **파일 권한 확인**
   - 읽기 권한이 있는지 확인
   - 파일 속성에서 읽기 전용 해제

3. **파일 손상 확인**
   - 파일 크기 확인 (0KB가 아닌지)
   - Outlook에서 직접 열어보기

### 4.3 Python 모듈 오류

#### 증상
```
ModuleNotFoundError: No module named 'pypff'
```

#### 해결 방법

1. **libpff-python 설치**
   ```bash
   pip install libpff-python
   ```

2. **다른 패키지 설치**
   ```bash
   pip install -r requirements.txt
   ```

3. **Python 환경 확인**
   ```bash
   python --version
   pip --version
   ```

### 4.4 메모리 부족

#### 증상
- 스캔 중 프로그램 종료
- 느린 처리 속도

#### 해결 방법

1. **배치 크기 조정**
   - `--batch-size` 옵션 값을 낮춤 (예: 1000, 500)
   - 기본값: 5000

2. **다른 프로그램 종료**
   - 메모리 사용량이 높은 프로그램 종료
   - 브라우저 탭 줄이기

3. **스캔 범위 축소**
   - 날짜 범위를 줄임
   - 특정 폴더만 선택

### 4.5 인코딩 오류

#### 증상
- 한글 파일명 오류
- 문자 깨짐

#### 해결 방법

1. **배치 파일 인코딩 확인**
   - UTF-8 (BOM 없음)로 저장
   - `chcp 65001` 명령 포함 확인

2. **PowerShell 인코딩 설정**
   ```powershell
   [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
   ```

## 5. 실행 우선순위

### 권장 실행 순서

1. **1단계: 환경 검증**
   - [ ] Python 버전 확인
   - [ ] 필수 패키지 설치 확인
   - [ ] PST 파일 경로 확인
   - [ ] 디렉토리 구조 확인

2. **2단계: 빠른 실행 테스트**
   - [ ] `quick_run_2025_06.bat` 실행
   - [ ] 실행 결과 확인

3. **3단계: 결과 검증**
   - [ ] Excel 파일 생성 확인
   - [ ] 로그 파일 확인
   - [ ] 데이터 정확성 확인

4. **4단계: 필요시 추가 실행**
   - [ ] 다른 날짜 범위 스캔
   - [ ] 자동 모드 실행 (`run_scanner.bat`)

## 6. 성능 최적화 팁

### 배치 크기 조정
- 메모리가 충분한 경우: `--batch-size 10000`
- 메모리가 부족한 경우: `--batch-size 1000`

### 날짜 범위 최적화
- 전체 스캔보다 특정 월/기간 스캔 권장
- 여러 번 나누어 스캔 후 결과 병합

### 폴더 선택
- 필요한 폴더만 선택하여 처리 시간 단축
- `--folders` 옵션으로 특정 폴더만 스캔

## 7. 추가 리소스

### 관련 문서
- [PST 안전 가이드](PST_SAFETY_GUIDE.md)
- [날짜 범위 가이드](DATE_RANGE_GUIDE.md)
- [README Lite](README_LITE.md)
- [README](../README.md)

### 지원
문제가 발생하면 다음을 확인하세요:
1. 로그 파일 (`output/logs/`)
2. 에러 메시지
3. 시스템 환경
4. 재현 단계

---

**문서 버전**: 1.0  
**작성일**: 2025-01-XX  
**최종 검증**: 2025-01-XX

## 8. 환경 검증 스크립트

자동 환경 검증을 위해 `verify_execution_env.py` 스크립트를 제공합니다.

### 실행 방법
```bash
python verify_execution_env.py
```

### 검증 항목
1. Python 버전 (3.11+)
2. 필수 모듈 (pandas, numpy, openpyxl, pypff)
3. 디렉토리 구조 (results/, output/ 등)
4. PST 파일 경로
5. Outlook 실행 상태

### 검증 결과
2025-01-XX 검증 결과: 모든 항목 통과 (12/12)

- ✓ Python 3.11.8
- ✓ 모든 필수 모듈 설치됨
- ✓ 디렉토리 구조 확인됨
- ✓ PST 파일 존재 확인 (60.6 GB)
- ✓ Outlook 실행 중 아님

