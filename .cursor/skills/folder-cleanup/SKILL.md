---
name: folder-cleanup
description: CONVERT 폴더에서 임시 파일 삭제, 중복 파일 식별, 폴더 구조 재구성, 아카이브 생성을 표준화한다. "cleanup", "정리", "폴더 정리", "임시 파일", "중복 파일" 요청에 사용.
---

# folder-cleanup

## 언제 사용
- CONVERT 폴더의 임시 파일 정리가 필요할 때
- 중복 파일을 식별하고 정리해야 할 때
- 폴더 구조를 재구성하거나 아카이브를 생성해야 할 때
- 프로젝트 정리 작업 전 분석이 필요할 때

## 핵심 원칙 (안전 우선)

**AGENTS.md Section 2 (안전/권한) 규칙 준수**

### 기본 원칙
- **모든 작업은 dry-run 모드로 시작** (실제 변경 없음, 리포트만 생성)
- **3단계 확인 프로세스**: Analysis → Review → Execution

### Ask first (필수 승인)
- 임시 파일 삭제도 10개 이상이면 Ask first
- 폴더 이동/재구성 (레거시 경로 파손 위험)
- 아카이브 생성 (파일 이동 포함)
- Git 추적 파일 삭제/이동
- 운영 데이터(PII 포함) 폴더 작업
- Excel 파일(.xlsx/.xlsm) 삭제 (서식 손실 위험)

### Never (절대 금지)
- Git 추적 파일 자동 삭제 (명시적 요청 없이)
- 운영 스크립트 핵심 파일 삭제
- 백업 없이 대량 삭제
- PII 포함 데이터 자동 정리

### 보호 대상
- `.git/` 디렉토리 및 모든 Git 추적 파일
- `AGENTS.md`, `README.md` 등 핵심 문서
- `requirements.txt`, `pyproject.toml` 등 설정 파일
- `*.py` 소스 파일 (임시/중복 확인 후에만)
- 운영 중인 Excel 템플릿/매크로 파일

## 표준 실행 (3단계 프로세스)

### 1) Analysis Phase (읽기 전용, 자동 실행 가능)
- 스캔: 임시 파일, 중복 파일, 폴더 구조 분석
- 리포트 생성: JSON/Markdown 형식
- 카테고리별 분류: 안전 삭제 / 검토 필요 / 보호 대상

**실행:**
```bash
python .cursor/skills/folder-cleanup/scripts/cleanup_analyzer.py --root . --dry-run --out out/cleanup_report.json
```

### 2) Review Phase (사용자 확인 필수)
- 표 형식으로 제시: | 파일/폴더 | 타입 | 크기 | 위험도 | 권장 조치 |
- 위험도 분류: LOW (임시 파일), MEDIUM (중복), HIGH (소스/설정)
- 예상 영향 범위 요약
- 사용자 승인 대기

### 3) Execution Phase (명시적 승인 후)
- 승인된 항목만 처리
- 각 작업 전 최종 확인 (대량 작업 시)
- 실행 로그 생성 (롤백 정보 포함)
- 실행 후 검증: `/verifier`로 영향 확인

**실행 (승인 후):**
```bash
python .cursor/skills/folder-cleanup/scripts/cleanup_analyzer.py --root . --execute --confirm
```

## 입력 카드
- 대상 폴더 경로 (기본: 프로젝트 루트)
- 정리 타입: 임시 파일 / 중복 파일 / 폴더 구조 / 아카이브
- 출력 경로: 리포트 저장 위치 (기본: `out/`)

## 절차 (보수적)

1) 엔트리포인트 확인
- `cleanup_analyzer.py` 스크립트의 `--help` 확인
- 사용 가능한 옵션 및 안전 기능 확인

2) Dry-run 분석
- 기본적으로 dry-run 모드로 실행
- 리포트를 검토하여 영향 범위 파악

3) 승인 및 실행
- 사용자 명시적 승인 후에만 실제 작업 수행
- 대량 작업 시 단계별 확인

4) 검증
- 실행 후 `/verifier`로 영향 확인
- 삭제된 파일 목록 로그 확인 (롤백 가능하도록)

## 리포트 포맷 (권장)
- Evidence Table: | 파일/폴더 | 타입 | 크기 | 위험도 | 권장 조치 | 상태 |
- 위험도별 그룹화: LOW / MEDIUM / HIGH
- 실행 로그: 삭제/이동된 파일 목록 (JSON 형식)
- FAIL이면: 원인 1줄 + 최소 수정안 + 재실행 커맨드

## 통합
- `/convert-scoper`: 정리 전 인벤토리 생성
- `/verifier`: 정리 후 영향 검증
- `convert-toolbox`: 스모크 테스트로 정리 후 검증
