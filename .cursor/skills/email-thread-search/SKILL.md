---
name: email-thread-search
description: email_search 모듈에서 Outlook Excel export 기반 검색/스레드 추적을 표준화한다. "outlook export", "thread", "메일 검색" 요청에 사용.
---

# email-thread-search

## 핵심 원칙(PII)
- 운영 메일/전화/주소 등 PII는 커밋/공유 금지
- 샘플 데이터는 익명화된 최소 컬럼만 사용

## 입력 카드
- Excel/CSV 경로(Outlook export)
- 검색 조건: subject/from/to/date range/keyword
- 출력: 결과 CSV/리포트 경로(out/)

## 절차
1) 엔트리포인트 확인
- email_search 폴더의 README, streamlit app, CLI 스크립트(--help) 우선

2) 검색 1회(샘플 우선)
- 샘플 데이터로 "검색 1건 + 스레드 빌드 1회" 재현

3) 결과 정리
- 결과를 out/ 아래에 저장
- 리포트: | Query | Hits | Threaded? | Output Path | Notes |

## Ask first
- 대용량 원본(운영) export 전체를 로드/분석
- 추가 라이브러리 설치
