## 🚀 빠른 시작

```bash
# 0. 환경 검증 (권장)
python verify_execution_env.py

# 1. 의존성 설치
pip install -r requirements.txt

# 2. 파일 시스템 스캔 (기본)
python run_scan.py

# 3. PST 이메일 스캔 (libpst 기반)
.\quick_run_2025_06.bat

# 4. 전체 파이프라인 실행
python run_all_scripts.py
```

## 📁 디렉토리 구조

```
hvdc_scripts_consolidated/
├── hvdc/                           # 메인 모듈 (완성)
│   ├── core/                      # 핵심 기능
│   ├── extractors/                # 데이터 추출
│   ├── scanner/                   # 파일 스캔 (fs_scanner.py)
│   ├── parser/                    # 파싱
│   ├── cli/                       # CLI 명령
│   └── report/                    # 보고서 생성
├── results/                       # PST 스캔 결과 및 HVDC 분석
├── docs/                          # 문서 (PST 가이드 포함)
├── tests/                         # 유닛 테스트
├── tools/                         # 유틸리티
├── legacy/                        # 레거시 스크립트
├── output/                        # 출력 파일
├── outlook_pst_scanner.py         # PST 스캐너 (libpst 기반)
├── outlook_hvdc_analyzer.py       # HVDC 온톨로지 분석
├── quick_run_2025_06.bat          # PST 빠른 실행
└── run_scanner.bat                # PST 스캐너 실행
```

## 📖 문서

- [시스템 아키텍처](docs/SYSTEM_ARCHITECTURE.md)
- [API 참조](docs/API_REFERENCE.md)
- [사용자 가이드](docs/USER_GUIDE.md)
- [실행 계획](docs/EXECUTION_PLAN.md) ⭐ **실행 전 필독**
- [PST 날짜 범위 가이드](docs/DATE_RANGE_GUIDE.md)
- [PST 안전 가이드](docs/PST_SAFETY_GUIDE.md)
- [README Lite](docs/README_LITE.md)

## 📧 PST 이메일 스캐너 (libpst 기반)

### 빠른 실행
```bash
# 2025년 6월 데이터 스캔 (권장 설정)
.\quick_run_2025_06.bat

# 또는 사용자 정의 스캔
python outlook_pst_scanner.py
```

### 주요 기능
- **자동 Outlook 종료**: 파일 잠금 방지
- **날짜 범위 필터링**: YYYY-MM-DD 형식
- **폴더 선택 스캔**: all, 번호, 범위 선택
- **안전한 읽기 전용**: libpst 기반 (PST 파일 절대 수정 안 함)
- **Excel 결과 출력**: 전체/폴더별/발신자별 3개 시트

### 스캔 결과
모든 결과는 `results/` 폴더에 저장됩니다:
- `OUTLOOK_YYYYMM.xlsx` - PST 스캔 결과 (월별)
- `OUTLOOK_HVDC_ONTOLOGY_YYYYMM.xlsx` - HVDC 온톨로지 분석
- `OUTLOOK_HVDC_REPORT_YYYYMM.xlsx` - HVDC 요약 보고서

**파일명 예시:**
- `OUTLOOK_202505.xlsx` - 2025년 5월 스캔
- `OUTLOOK_202506_20251029.xlsx` - 2025년 6월 재스캔 (충돌 방지)
- `OUTLOOK_HVDC_ONTOLOGY_202508.xlsx` - 8월 온톨로지

### 기술적 장점
- ✅ 완전한 읽기 전용 (PST 손상 위험 0%)
- ✅ Outlook 프로세스 불필요
- ✅ 대용량 PST 처리 가능 (60GB+)
- ✅ 안정적이고 빠른 성능

## 🔧 주요 기능

- **이메일 스캔**: EMAIL 폴더 전체 스캔 + PST 파일 직접 스캔
- **케이스 추출**: HVDC-ADOPT, PRL, JPTW/GRM 추출
- **사이트 매핑**: AGI, DAS, ZAK, MIR, SHU
- **보고서 생성**: Excel, JSON, MD 형식

## 📊 성능

- 처리 속도: 0.02ms/텍스트
- 메모리 효율: 0.55 KB/텍스트
- 테스트 커버리지: 100% (core 모듈)

## 🧪 테스트

```bash
# 유닛 테스트
pytest tests/unit/ -v

# 스모크 테스트
python tools/smoke_extract.py
```

## 📝 레거시 스크립트

이전 독립 스크립트들은 `legacy/scripts/`에 보관되어 있습니다.
새로운 시스템은 `hvdc/` 모듈을 사용하세요.

## 🔍 HVDC 코드 형식

### 케이스 번호
- **HVDC-ADOPT**: `HVDC-ADOPT-HE-0476`
- **일반 HVDC**: `HVDC-DSV-HE-MOSB-187`
- **PRL**: `PRL-O-046-O4(HE-0486)`
- **JPTW/GRM**: `JPTW-71 / GRM-123`

### 사이트 코드
- **AGI**: Abu Dhabi Global Island
- **DAS**: Das Island
- **ZAK**: Zakum
- **MIR**: Mirfa
- **SHU**: Shuweihat

## 📈 최근 성과

- **총 파일 스캔**: 435개 파일
- **케이스 추출**: 41개 케이스 (412.5% 증가)
- **사이트 식별**: 5개 주요 사이트
- **LPO 추출**: 8개 LPO 번호

## 🛠️ 기술 스택

- **Python**: 3.11+
- **주요 라이브러리**: pandas, pathlib, re, json
- **테스트**: pytest
- **문서화**: Markdown
- **PST 스캔**: libpff-python (pypff) - 읽기 전용 libpst 바인딩
- **Outlook 자동화**: subprocess (taskkill)

## 📋 시스템 요구사항

- **운영체제**: Windows 10/11
- **Python**: 3.11 이상
- **메모리**: 최소 4GB RAM (대용량 PST 처리 시 8GB 권장)
- **디스크**: 최소 1GB 여유 공간

## 🔧 설정

### EMAIL 폴더 경로
```python
# hvdc/core/config.py
EMAIL_ROOT = Path("C:/Users/SAMSUNG/Documents/EMAIL")
```

### PST 파일 경로
```batch
# quick_run_2025_06.bat 또는 run_scanner.bat에서 설정
set PST_PATH=C:\Users\SAMSUNG\Documents\Outlook 파일\[파일명].pst
```

### 출력 폴더
```python
EXCEL_OUTDIR = Path("output/reports")
```

## 📊 출력 파일

### 로그 파일
- `output/logs/email_scan_*.log`: 스캔 과정 로그

### 데이터 파일
- `output/data/email_scan_results_*.json`: 스캔 결과 데이터
- `output/data/email_folder_stats_*.csv`: 폴더별 통계

### 보고서 파일
- `output/reports/email_scan_report_*.md`: 마크다운 보고서
- `output/reports/hvdc_email_report_*.xlsx`: Excel 보고서

## 🚨 문제 해결

### 모듈 임포트 오류
```python
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
```

### 파일 인코딩 오류
```python
from hvdc.core.config import ENCODING_FALLBACKS
# 자동 인코딩 폴백 사용
```

### PST 파일 접근 오류
```bash
# Outlook 종료 확인 (자동 종료되지만, 수동으로도 가능)
taskkill /F /IM outlook.exe /T

# libpff-python 설치 확인
pip install libpff-python
```

### 메모리 부족
```python
# 샘플 제한 설정
MAX_FILES = 100
OUTLOOK_MAX_EMAILS = 1000
```

## 📚 추가 리소스

- [시스템 아키텍처](docs/SYSTEM_ARCHITECTURE.md)
- [API 참조](docs/API_REFERENCE.md)
- [사용자 가이드](docs/USER_GUIDE.md)
- [레거시 스크립트](legacy/scripts/)

## 🆘 지원

문제가 발생하면 다음을 확인하세요:
1. 로그 파일 (`output/logs/`)
2. 에러 메시지
3. 시스템 환경
4. 재현 단계

---

**프로젝트**: HVDC Email Processing System  
**버전**: v1.0  
**최종 업데이트**: 2025-10-26  
**상태**: ✅ **완료**

## 📋 포함된 스크립트 목록

### 1. 이메일 처리 스크립트
- **`email_folder_scanner.py`** (17,562 bytes, 434 lines)
  - EMAIL 폴더의 모든 하위 폴더 스캔
  - Outlook 파일 제외하고 이메일 데이터 추출
  - 12개 함수, 4개 try-except 블록

- **`email_ontology_mapper.py`** (19,276 bytes, 440 lines)
  - 이메일 데이터를 HVDC 온톨로지 시스템에 매핑
  - 10개 함수, 2개 try-except 블록

- **`create_complete_email_excel.py`** (15,074 bytes, 336 lines)
  - 매핑된 이메일 데이터를 종합 엑셀 보고서로 생성
  - 5개 함수, 에러 처리 없음

### 2. 폴더 분석 스크립트
- **`folder_title_mapper.py`** (21,568 bytes, 524 lines)
  - 폴더 제목 기반 케이스 번호, 날짜, 사이트 매핑
  - 10개 함수, 1개 try-except 블록

- **`simple_folder_analyzer.py`** (11,866 bytes, 296 lines)
  - 폴더 제목 간단 분석
  - 9개 함수, 1개 try-except 블록

### 3. 종합 매핑 스크립트
- **`comprehensive_email_mapper.py`** (23,027 bytes, 556 lines)
  - 종합 이메일 매핑 및 네트워크 시각화
  - 11개 함수, 2개 try-except 블록

### 4. 화물 추적 스크립트
- **`hvdc_cargo_tracking_system.py`** (20,074 bytes, 497 lines)
  - HVDC 화물 추적 시스템
  - 11개 함수, 2개 try-except 블록

### 5. 패턴 업데이트 스크립트
- **`update_email_pattern_rules.py`** (9,357 bytes, 254 lines)
  - 이메일 패턴 규칙 업데이트
  - 5개 함수, 에러 처리 없음

### 6. 모니터링 스크립트
- **`monitor_scan_progress.py`** (2,235 bytes, 71 lines)
  - 스캔 진행상황 모니터링
  - 1개 함수, 에러 처리 없음

### 7. 분석 보고서 스크립트
- **`analysis_summary_report.py`** (8,123 bytes, 222 lines)
  - 분석 로직 종합 보고서 생성
  - 2개 함수, 1개 try-except 블록

## 🔧 사용 방법

### 기본 실행 순서
1. **폴더 스캔**: `python email_folder_scanner.py`
2. **진행상황 모니터링**: `python monitor_scan_progress.py`
3. **이메일 매핑**: `python email_ontology_mapper.py`
4. **엑셀 보고서 생성**: `python create_complete_email_excel.py`
5. **종합 분석**: `python comprehensive_email_mapper.py`

### 개별 실행
각 스크립트는 독립적으로 실행 가능합니다.

## 📊 스크립트 통계

| 스크립트명 | 크기 | 라인 수 | 함수 수 | 에러 처리 | 복잡도 |
|------------|------|---------|---------|-----------|--------|
| comprehensive_email_mapper.py | 23KB | 556 | 11 | 2 | 🔴 복잡 |
| folder_title_mapper.py | 21KB | 524 | 10 | 1 | 🔴 복잡 |
| hvdc_cargo_tracking_system.py | 20KB | 497 | 11 | 2 | 🟡 중간 |
| email_ontology_mapper.py | 19KB | 440 | 10 | 2 | 🟡 중간 |
| email_folder_scanner.py | 17KB | 434 | 12 | 4 | 🟡 중간 |
| create_complete_email_excel.py | 15KB | 336 | 5 | 0 | 🟡 중간 |
| simple_folder_analyzer.py | 11KB | 296 | 9 | 1 | 🟡 중간 |
| update_email_pattern_rules.py | 9KB | 254 | 5 | 0 | 🟡 중간 |
| analysis_summary_report.py | 8KB | 222 | 2 | 1 | 🟡 중간 |
| monitor_scan_progress.py | 2KB | 71 | 1 | 0 | 🟢 단순 |

## 🚨 주요 문제점

### 1. 코드 품질 문제
- **복잡한 스크립트**: 2개 스크립트가 500라인 이상
- **에러 처리 부족**: 3개 스크립트에 try-except 블록 없음
- **에러 처리 비율 낮음**: 평균 0.5% (권장: 5% 이상)

### 2. 구조적 문제
- **단일 책임 원칙 위반**: 하나의 스크립트가 너무 많은 기능 담당
- **코드 중복**: 유사한 로직이 여러 스크립트에 반복
- **하드코딩**: 설정값이 코드에 직접 포함

## 💡 개선 방안

### Phase 1: 기반 구조 개선
1. **공통 모듈 분리**
   - `utils/email_parser.py`: 이메일 파싱 공통 로직
   - `utils/pattern_matcher.py`: 패턴 매칭 공통 로직
   - `utils/file_handler.py`: 파일 처리 공통 로직

2. **설정 관리 통합**
   - `config/settings.py`: 모든 설정값 중앙 관리
   - `config/patterns.py`: 정규식 패턴 정의

### Phase 2: 코드 품질 개선
1. **리팩토링 우선순위**
   - `comprehensive_email_mapper.py` (556 lines)
   - `folder_title_mapper.py` (524 lines)
   - `hvdc_cargo_tracking_system.py` (497 lines)

2. **에러 처리 표준화**
   - 모든 스크립트에 try-except 블록 추가
   - 커스텀 예외 클래스 정의

### Phase 3: 성능 및 안정성 개선
1. **비동기 처리 도입**
2. **메모리 최적화**
3. **테스트 코드 추가**

## 📋 실행 체크리스트

### 즉시 실행 가능
- [ ] 공통 모듈 디렉토리 생성
- [ ] 설정 파일 통합
- [ ] 로깅 표준화

### 단기 개선 (1-2주)
- [ ] 우선순위 스크립트 리팩토링
- [ ] 에러 처리 표준화
- [ ] 단위 테스트 추가

### 중기 개선 (1-2개월)
- [ ] 비동기 처리 도입
- [ ] 성능 최적화
- [ ] 통합 테스트 구축

## 🎯 성공 지표

### 코드 품질
- **코드 복잡도**: 평균 라인 수 < 300
- **에러 처리**: try-except 비율 > 5%
- **테스트 커버리지**: > 80%

### 성능
- **실행 시간**: 기존 대비 50% 단축
- **메모리 사용량**: 기존 대비 30% 감소
- **에러율**: < 1%

### 유지보수성
- **코드 중복**: < 10%
- **의존성**: 최소화 및 명확화
- **문서화**: 모든 함수에 docstring

---

**생성일시**: 2025-10-03 14:46:00  
**총 스크립트 수**: 10개  
**총 코드 라인**: 3,530 lines  
**총 파일 크기**: 158,158 bytes
