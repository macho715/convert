# Cursor Subagents & Skills 사용 가이드

**CONVERT 프로젝트 전용**

이 가이드는 `.cursor/agents/`와 `.cursor/skills/`를 실제로 사용하는 방법을 설명합니다.

---

## 📋 목차

1. [Subagents 사용법](#subagents-사용법)
2. [Skills 사용법](#skills-사용법)
3. [실전 워크플로우 예시](#실전-워크플로우-예시)
4. [문제 해결](#문제-해결)

---

## Subagents 사용법

Subagents는 Cursor 채팅에서 **슬래시 명령어**로 호출합니다.

### 1. convert-scoper (인벤토리 스코퍼)

**언제 사용**: 프로젝트 구조를 빠르게 파악해야 할 때

**사용법**:
```
/convert-scoper
```

또는 자연어로:
```
프로젝트 구조 파악해줘
엔트리포인트 찾아줘
```

**결과**: 
- 모듈별 엔트리포인트 목록
- 입출력 계약 요약
- 스모크 커맨드 후보

**예시 출력**:
```
| Module | Entry Points | Inputs | Outputs | Risks |
| --- | --- | --- | --- | --- |
| mrconvert_v1 | mrconvert --help | PDF/DOCX/XLSX | out/*.txt | 대용량 파일 |
```

---

### 2. verifier (검증자)

**언제 사용**: 작업 완료 후 검증이 필요할 때

**사용법**:
```
/verifier
```

또는 자연어로:
```
검증해줘
테스트 돌려줘
작업이 제대로 됐는지 확인해줘
```

**결과**:
- PASS/FAIL 판정
- Evidence Table (테스트 결과)
- 실패 시 수정안 제시

**예시 출력**:
```
VERDICT: PASS

| Check | Result | Command | Notes |
| --- | --- | --- | --- |
| compileall | PASS | python -m compileall -q . | - |
| pytest | PASS | pytest -q | 15 tests passed |
```

---

### 3. excel-style-guardian (Excel 서식 가디언)

**언제 사용**: Excel 파일 서식 유지가 중요할 때 (CIPL, 간트 차트 등)

**사용법**:
```
/excel-style-guardian
```

또는 자연어로:
```
서식 유지해줘
CIPL 포맷 확인해줘
Excel 서식 회귀 체크해줘
```

**결과**:
- Visual Regression Checklist
- 서식 변경 위험도 평가
- 안전한 수정 방향 제시

**예시 출력**:
```
| Item | Same? | Evidence | Risk |
| --- | --- | --- | --- |
| 시트명 | ✅ | Sheet1 유지 | LOW |
| 헤더 라인 | ✅ | Row 1 유지 | LOW |
| 병합셀 | ⚠️ | A1:B1 변경됨 | MEDIUM |
```

---

### 4. agi-schedule-updater (AGI TR Schedule 업데이트)

**언제 사용**: AGI TR Unit 1 Schedule HTML의 공지란·Weather & Marine Risk 블록을 매일 갱신할 때

**사용법**:
```
/agi-schedule-updater
```

또는 자연어로:
```
AGI TR Schedule 공지 업데이트해줘
날씨 블록 갱신해줘
Mina Zayed weather 반영해줘
```

**결과**:
- 공지란: 사용자 제공 날짜·텍스트로 블록 교체
- Weather & Marine Risk: 웹 검색 후 포맷에 맞춰 블록 교체, Last Updated 갱신

**관련 스킬**:
- `agi-schedule-daily-update`: 공지·날씨 블록 갱신 (트리거: AGI schedule 공지, 날씨 블록 업데이트, Mina Zayed weather)
- `agi-schedule-shift`: pivot date 이후 전체 일정 delta일 시프트 (트리거: 일정 시프트, schedule shift, 일정 연기)

---

## Skills 사용법

Skills는 **자동으로 트리거**되거나, **명령어로 직접 실행**할 수 있습니다.

### 1. convert-toolbox (도구 상자)

**트리거 키워드**: `inventory`, `smoke`, `verify`, `package`

**직접 실행**:

```bash
# 인벤토리 생성
python .cursor/skills/convert-toolbox/scripts/convert_inventory.py --root . --out out/inventory.json

# 스모크 테스트
python .cursor/skills/convert-toolbox/scripts/run_smoke.py --root .

# 구조 검증
python .cursor/skills/convert-toolbox/scripts/validate_agent_assets.py --root .
```

**자동 트리거 예시**:
```
프로젝트 인벤토리 만들어줘  # → convert-toolbox 자동 사용
스모크 테스트 돌려줘        # → convert-toolbox 자동 사용
```

---

### 2. mrconvert-run (문서 변환)

**트리거 키워드**: `mrconvert`, `convert pdf`, `OCR`, `table extract`

**사용 예시**:
```
PDF를 텍스트로 변환해줘
OCR로 이미지에서 텍스트 추출해줘
테이블 추출해줘
```

**절차**:
1. 엔트리포인트 확인 (`mrconvert_v1/README` 또는 `--help`)
2. Dry-run 샘플 변환
3. 출력 규칙 확인 (`out/` 또는 `output/`)

---

### 3. email-thread-search (이메일 검색)

**트리거 키워드**: `outlook export`, `thread`, `메일 검색`

**사용 예시**:
```
Outlook export로 이메일 검색해줘
스레드 추적해줘
메일 검색해줘
```

**주의사항**:
- PII(개인정보) 포함 데이터는 익명화 필수
- 샘플 데이터로 먼저 테스트

---

### 4. cipl-excel-build (CIPL Excel 생성)

**트리거 키워드**: `CIPL`, `invoice packing list`, `xlsx template`

**사용 예시**:
```
CIPL Excel 만들어줘
Invoice packing list 생성해줘
```

**절차**:
1. 템플릿 기준 고정 (서식 SSOT)
2. 생성 스크립트 엔트리포인트 확인
3. 서식 회귀 체크 (`/excel-style-guardian` 사용)

---

### 5. folder-cleanup (폴더 정리)

**트리거 키워드**: `cleanup`, `정리`, `폴더 정리`, `임시 파일`, `중복 파일`

**직접 실행**:

```bash
# 분석 (dry-run, 기본)
python .cursor/skills/folder-cleanup/scripts/cleanup_analyzer.py --root . --out out/cleanup_report.json
```

**사용 예시**:
```
임시 파일 정리해줘
중복 파일 찾아줘
폴더 정리해줘
```

**안전 기능**:
- 기본적으로 dry-run 모드 (실제 변경 없음)
- Git 추적 파일 자동 보호
- 3단계 확인 프로세스 (Analysis → Review → Execution)

**절차**:
1. **Analysis Phase**: 스캔 및 리포트 생성 (읽기 전용)
2. **Review Phase**: 사용자 확인 및 승인 대기
3. **Execution Phase**: 명시적 승인 후에만 실행

---

## 실전 워크플로우 예시

### 예시 1: 새 모듈 추가 전 구조 파악

```
1. /convert-scoper
   → 프로젝트 구조 파악

2. convert-toolbox (인벤토리)
   → python .cursor/skills/convert-toolbox/scripts/convert_inventory.py --root . --out out/inventory.json

3. 작업 수행
   → 새 모듈 구현

4. /verifier
   → 검증 및 테스트
```

---

### 예시 2: Excel 파일 생성 (CIPL)

```
1. cipl-excel-build 스킬 트리거
   → "CIPL Excel 만들어줘"

2. 템플릿 확인
   → 서식 기준 고정

3. Excel 생성
   → make_cipl_set.py 실행

4. /excel-style-guardian
   → 서식 회귀 체크

5. /verifier
   → 최종 검증
```

---

### 예시 3: 프로젝트 정리

```
1. folder-cleanup 스킬 트리거
   → "임시 파일 정리해줘"

2. 분석 리포트 확인
   → out/cleanup_report.json

3. 검토 및 승인
   → 위험도별 분류 확인

4. 실행 (필요 시)
   → --execute --confirm (주의!)

5. /verifier
   → 정리 후 영향 검증
```

---

### 예시 4: 문서 변환 작업

```
1. mrconvert-run 스킬 트리거
   → "PDF를 텍스트로 변환해줘"

2. 엔트리포인트 확인
   → mrconvert_v1/README 또는 --help

3. 샘플 변환 (dry-run)
   → 익명 샘플로 테스트

4. 실제 변환
   → 출력은 out/ 또는 output/에 저장

5. /verifier
   → 변환 결과 검증
```

---

## 문제 해결

### Q: Subagent가 작동하지 않아요

**확인 사항**:
1. `.cursor/agents/<name>.md` 파일이 존재하는가?
2. YAML frontmatter가 올바른가?
3. Cursor가 프로젝트 루트를 인식하는가?

**해결**:
```bash
# 구조 검증
python .cursor/skills/convert-toolbox/scripts/validate_agent_assets.py --root .
```

---

### Q: Skill이 자동으로 트리거되지 않아요

**확인 사항**:
1. 트리거 키워드가 정확한가?
2. `.cursor/skills/<name>/SKILL.md`가 존재하는가?
3. frontmatter의 `name:`이 폴더명과 일치하는가?

**해결**:
- 명시적으로 스킬 이름을 언급: "convert-toolbox 사용해서..."
- 직접 스크립트 실행 (위의 "직접 실행" 섹션 참고)

---

### Q: Windows에서 스크립트 실행 오류

**문제**: 인코딩 오류, 경로 오류

**해결**:
```powershell
# UTF-8 인코딩 설정
$env:PYTHONIOENCODING='utf-8'
python .cursor\skills\folder-cleanup\scripts\cleanup_analyzer.py --root .
```

---

### Q: folder-cleanup이 실제로 파일을 삭제하지 않아요

**설명**: 이것은 **의도된 안전 기능**입니다.

- 기본적으로 dry-run 모드로 실행 (실제 변경 없음)
- 실제 삭제는 안전을 위해 구현되지 않음
- 리포트를 검토한 후 수동으로 삭제하거나, 필요 시 스크립트 확장

---

## 빠른 참조

### Subagents 슬래시 명령어

| 명령어 | 목적 | 권한 |
| --- | --- | --- |
| `/convert-scoper` | 프로젝트 구조 파악 | readonly |
| `/verifier` | 작업 검증 | 수정 가능 |
| `/excel-style-guardian` | Excel 서식 체크 | readonly |

### Skills 트리거 키워드

| 스킬 | 키워드 | 직접 실행 스크립트 |
| --- | --- | --- |
| `convert-toolbox` | inventory, smoke, verify, package | `convert_inventory.py`, `run_smoke.py`, `validate_agent_assets.py` |
| `mrconvert-run` | mrconvert, convert pdf, OCR | (문서형, 스크립트 없음) |
| `email-thread-search` | outlook export, thread, 메일 검색 | (문서형, 스크립트 없음) |
| `cipl-excel-build` | CIPL, invoice packing list, xlsx template | (문서형, 스크립트 없음) |
| `folder-cleanup` | cleanup, 정리, 폴더 정리, 임시 파일 | `cleanup_analyzer.py` |

### 자주 사용하는 스크립트

```bash
# 구조 검증
python .cursor/skills/convert-toolbox/scripts/validate_agent_assets.py --root .

# 인벤토리 생성
python .cursor/skills/convert-toolbox/scripts/convert_inventory.py --root . --out out/inventory.json

# 스모크 테스트
python .cursor/skills/convert-toolbox/scripts/run_smoke.py --root .

# 폴더 정리 분석 (dry-run)
python .cursor/skills/folder-cleanup/scripts/cleanup_analyzer.py --root . --out out/cleanup_report.json
```

---

## 통합 워크플로우

### AGENTS.md Section 4 (표준 작업 루틴)와의 통합

1. **Locate 단계**: `/convert-scoper`로 구조 파악
2. **Plan 단계**: 인벤토리 생성 (`convert-toolbox`)
3. **Implement 단계**: 해당 스킬 사용 (예: `mrconvert-run`, `cipl-excel-build`)
4. **Verify 단계**: `/verifier` 또는 `convert-toolbox` 스모크 실행
5. **Document 단계**: 변경 사항 문서화
6. **Package 단계**: `folder-cleanup`으로 정리 (선택)

---

## 추가 정보

- **상세 가이드**: `subagentandskillguide.md`
- **프로젝트 규칙**: `AGENTS.md` Section 12
- **각 스킬 상세**: `.cursor/skills/<name>/SKILL.md`
- **각 서브에이전트 상세**: `.cursor/agents/<name>.md`

---

## 안전 규칙 요약

모든 Subagents와 Skills는 **AGENTS.md Section 2 (안전/권한)** 규칙을 준수합니다:

- **Allowed without prompt**: 파일 읽기, 문서 업데이트, 단일 파일 스모크
- **Ask first**: 새 의존성 설치, 대량 삭제/이동, Excel 매크로 수정
- **Never**: 자격증명 커밋, 외부 데이터 전송, 운영 스크립트 핵심 로직 변경

특히 `folder-cleanup`은:
- 기본적으로 dry-run 모드
- Git 추적 파일 자동 보호
- 3단계 확인 프로세스 필수

---

**마지막 업데이트**: 2026-01-28
