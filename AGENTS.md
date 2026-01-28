AGENTS.md — CONVERT 업무자동화 프로그램 개발 규칙(문서변환·이메일·AGI·CIPL·Stability)
Last updated: 2026-01-28 (TZ=Asia/Dubai)
Version: 1.1

> 당신은 이 리포지토리(또는 CONVERT 폴더)에서 작업하는 **자율 코딩 에이전트**다.
> 목표: 사용자 개입 최소화로 “계획→구현→검증→문서→패키징”을 자동 수행한다.
> 충돌 시 우선순위: **사용자 프롬프트 > (가장 가까운) AGENTS.md > 상위 AGENTS.md**.
> 절대 규칙: 기존 운영 중 스크립트/엑셀 매크로/출력 포맷을 **깨지 말 것**(Backwards compatible).

---

## 0) 미션 / 불변 조건 (Non-negotiables)
- 이 폴더의 주 목적은 **업무 관련 프로그램 개발/개선**이다. (물류 문서·이메일·일정·정산 자동화)
- “이미 운영 중”인 모듈은 **동작/입력/출력 포맷을 유지**한 채로 개선한다.
- PII(이메일/전화/주소/계정) 또는 자격증명(API Key, Token) **커밋 금지**.
- 생성물(변환 결과)은 기본적으로 `out/` 또는 `output/`에 저장하고, Git 추적 제외가 원칙이다.

---

## 1) 프로젝트 개요 (Project Overview)
CONVERT는 아래 업무 자동화 모듈의 집합이다(폴더 기준):
- `mrconvert_v1/` : PDF/DOCX/XLSX → TXT/MD/JSON 변환(+OCR/테이블 추출)
  - 온톨로지 문서 변환 지원 (`ontology_machine_readable/`)
  - Cursor 통합 패키지 (`cursor_only_pack_v1/`)
- `email_search/` : Outlook Excel 데이터 기반 검색 + 스레드 추적 + Streamlit UI
  - 대시보드 (`dashboard/`)
  - 통합 스크립트 (`hvdc_scripts_consolidated/`)
- `AGI DOCS/` : AGI TR 문서 추적(빌더/VBA/룰테이블)
  - TR_DocHub_AGI_2026_Complete 패키지 포함
- `AGI TR 1-6 Transportation Master Gantt Chart/` : 간트/시나리오/히트맵/추적
  - AGI_TR6_READY_PACK_v1, AGI_TR7_Dynamic_Gantt 포함
- `CIPL/` : Commercial Invoice & Packing List 자동 생성(최신: `CIPL_PATCH_PACKAGE/`)
  - 레거시 코드 (`CIPL_LEGACY/`)
- `vessel_stability_python/` : Excel 안정성 계산의 Python 변환 + 테스트
- `JPT71/` : 콘텐츠 캘린더(Excel→Python)
- `scripts/` : 유틸(날씨/검증/변환/정리 등)
- `mammoet/` : Mammoet 관련 자료(보관/추출 스크립트 포함)
- `cursor_only_pack_v1/` : 루트 레벨 Cursor 통합 패키지
- `out/`, `output/` : 변환/생성 결과 저장소

---

## 2) 안전/권한 (Safety & Permissions)
### Allowed without prompt (자동 실행 허용)
- 파일 읽기/목록/검색(예: `rg`, `ls`)
- **단일 파일 수준** 스모크 실행(아래 “검증 루틴” 범위 내)
- 문서/README/AGENTS.md 업데이트(규칙/커맨드/경로 정정)

### Ask first (사용자 승인 필요)
- 새 의존성 설치/업그레이드(예: `pip install ...`, `apt-get ...`, `brew ...`)
- 대량 리네임/대량 삭제/폴더 이동(레거시 경로 파손 위험)
- Excel 매크로(VBA) 자동 수정(바이너리 손상/호환 리스크)
- 실제 운영 데이터(PII 포함)로 테스트 실행

### Never (금지)
- 자격증명/개인정보 커밋/출력
- 외부로 데이터 업로드/전송(로그 포함)
- “작동할 것”이라는 추측으로 운영 스크립트 핵심 로직 변경

---

## 3) 개발 환경/설치 (Setup)
> 이 레포는 모듈 혼합형(스크립트 다수)일 가능성이 높다. 따라서 **먼저 자동 탐색** 후, 가장 보수적인 방식으로 환경을 구성한다.

### 3.1 자동 탐색(필수)
- 루트/각 폴더에서 아래 파일을 먼저 찾는다:
  - 의존성 관리: `pyproject.toml`, `requirements.txt`, `environment.yml`, `Pipfile`, `setup.cfg`
  - 문서: `README.md`, `AGENTS.md`, `agent_*.md` (모듈별 가이드)
  - 엔트리포인트: `*_cli.py`, `main.py`, `app.py`, `__main__.py`, `run_*.py`
  - 설정: `*.json`, `*.yaml`, `*.yml` (설정 파일)

### 3.2 가상환경(권장)
- Linux/macOS:
  - `python -m venv .venv && source .venv/bin/activate`
- Windows PowerShell:
  - `python -m venv .venv; .\.venv\Scripts\Activate.ps1`

### 3.3 의존성 설치(조건부)
- `requirements.txt`가 있으면:
  - `python -m pip install -r requirements.txt`
- `pyproject.toml`만 있으면(패키지 형태일 때):
  - `python -m pip install -e .`
- 어떤 것도 없으면:
  - **Ask first** 후 최소 의존성만 추가한다(예: `pytest`, `openpyxl` 등).

---

## 4) 표준 작업 루틴(에이전트 실행 프로토콜)
에이전트는 모든 작업을 아래 순서로 수행한다:
1) **Locate**: 관련 폴더/엔트리포인트/입출력 포맷 파악(README 우선)
2) **Plan**: 변경 범위 최소화(작은 diff)
3) **Implement**: 기능 추가/버그 수정
4) **Verify**: 스모크/테스트/샘플 실행(아래 5절)
5) **Document**: 변경 사항을 해당 폴더 README/AGENTS.md에 반영(SSOT)
6) **Package**: 필요 시 실행 스크립트/예제/샘플 데이터(익명) 포함

---

## 5) 검증(테스트/스모크) 루틴 (Testing / Smoke)
> “전체 빌드”가 아니라 **파일/모듈 단위 검증**을 우선한다.

### 5.1 공통 스모크(항상 가능)
- 파이썬 구문/임포트 최소 검증:
  - `python -m compileall -q .`

### 5.2 pytest가 있을 때(조건부)
- `pytest.ini` 또는 `tests/` 또는 `pyproject.toml`에 pytest 설정이 있으면:
  - `pytest -q`

### 5.3 모듈별 스모크(엔트리포인트 확정 후 AGENTS.md에 “고정 커맨드”로 기록)
- `mrconvert_v1/` : CLI `--help` + 샘플 1건 변환
  - `python -m mrconvert --help` 또는 `mrconvert --help`
- `email_search/` : 샘플 Excel(익명)로 검색 1건 + 스레드 빌드 1회
  - Streamlit 대시보드 실행 확인: `streamlit run dashboard/app.py` (가능 시)
- `CIPL/` : `CIPL_PATCH_PACKAGE/make_cipl_set.py` 실행(익명 입력) + 출력 포맷 체크
- `vessel_stability_python/` : `tests/` 통과(가능 시)
  - `pytest tests/` 또는 모듈별 테스트 실행
- `AGI DOCS/` : 빌더 실행은 “Ask first”(엑셀 출력/서식 중요)

### 5.4 고정 커맨드(스모크 후보 3~8)
| # | 용도 | 커맨드 |
|---|------|--------|
| 1 | 공통 구문 검증 | `python -m compileall -q .` |
| 2 | pytest(조건부) | `pytest -q` |
| 3 | mrconvert_v1 | `python -m mrconvert --help` · `mrconvert sample.pdf --out out --format md json --tables csv` |
| 4 | email_search | `streamlit run email_search/dashboard/app.py` · `python email_search/scripts/run_full_export.py --excel <익명> --sheet "전체_데이터" --out email_search/outputs/threads_smoke --query "LPO-1599"` |
| 5 | CIPL | `python CIPL/CIPL_PATCH_PACKAGE/make_cipl_set.py --in CIPL/CIPL_PATCH_PACKAGE/voyage_input_sample_full.json --out out/CIPL_smoke.xlsx` |
| 6 | vessel_stability_python | `pytest vessel_stability_python/tests/ -q` 또는 `python vessel_stability_python/example_usage.py` |
| 7 | 패키지 정합성 | `python .cursor/skills/convert-toolbox/scripts/validate_agent_assets.py --root .` |
| 8 | 인벤토리 생성 | `python .cursor/skills/convert-toolbox/scripts/convert_inventory.py --root . --out out/convert_inventory.json` |

---

## 6) 코딩 스타일/규칙 (Code Style & Conventions)
- Python:
  - 함수/모듈 단위로 쪼개기(거대 스크립트 방지)
  - I/O(파일 읽기/쓰기)와 로직(변환/계산)을 분리
  - 로그는 `logging` 사용(프린트 남발 금지)
  - 신규 코드는 가능한 한 타입힌트 추가(과도한 강제는 금지)
- Excel/VBA:
  - 바이너리 파일(xlsm)은 자동 수정을 피하고, 필요 시 “Ask first”
  - 서식/병합셀/테두리 등 시각 요소는 **회귀 테스트(샘플 비교)**를 설계

---

## 7) 데이터/보안 (Security / Data Handling)
- 이메일/문서 데이터에는 PII가 포함될 수 있다.
  - 리포에 커밋하는 샘플은 **익명화(마스킹)** 된 데이터만 허용
- `.env`/토큰/키:
  - `.env.example`만 커밋하고, 실제 `.env`는 절대 커밋하지 않는다.
- 출력물:
  - 기본 출력 경로는 `out/`, `output/` 사용
  - 대용량 결과물은 Git 제외(필요 시 압축/샘플만)

---

## 8) Git 워크플로우 (Git Workflow & PR)
- 브랜치: `feat/<topic>`, `fix/<topic>`, `refactor/<topic>`
- 커밋 메시지: `feat: ...`, `fix: ...`, `refactor: ...`, `docs: ...`
- PR 전 체크:
  - 관련 스모크/테스트 실행 로그를 PR 본문에 요약
  - 변경된 동작/출력 포맷이 기존과 호환되는지 명시

---

## 9) 폴더별 전문 AGENTS.md(권장: Option B)
아래 폴더에는 별도 `AGENTS.md`를 두고, 각자의 “고정 커맨드/출력 규칙”을 적어라:
- `mrconvert_v1/agent_mrconvert.md` ✅ (존재)
- `email_search/AGENTS.md` (권장)
- `CIPL/AGENTS.md` (권장)
- `vessel_stability_python/AGENTS.md` (권장)
- `scripts/AGENTS.md` (위험 스크립트: dry-run 규칙 포함)
- `AGI DOCS/AGENTS.md` (권장)
- `JPT71/AGENTS.md` (권장)

> 참고: 파일명은 `AGENTS.md` 또는 `agent_<module>.md` 형식을 사용할 수 있다.

---

## 10) 도구 호환(심볼릭 링크) (Interoperability)
도구가 다른 파일명을 요구하면 루트에서 링크를 만든다(필요 시):
- Cursor: `.cursorrules` → `AGENTS.md`
- Claude: `CLAUDE.md` → `AGENTS.md`
- Gemini: `GEMINI.md` → `AGENTS.md`
- 구형: `agents.md` 또는 `AGENT.md` → `AGENTS.md`

### Windows PowerShell (현재 환경)
```powershell
# 심볼릭 링크 생성 (관리자 권한 필요할 수 있음)
New-Item -ItemType SymbolicLink -Path ".cursorrules" -Target "AGENTS.md"
New-Item -ItemType SymbolicLink -Path "CLAUDE.md" -Target "AGENTS.md"
New-Item -ItemType SymbolicLink -Path "GEMINI.md" -Target "AGENTS.md"
New-Item -ItemType SymbolicLink -Path "agents.md" -Target "AGENTS.md"
```

### Linux/macOS
```bash
ln -s AGENTS.md .cursorrules
ln -s AGENTS.md CLAUDE.md
ln -s AGENTS.md GEMINI.md
ln -s AGENTS.md agents.md
```

> 참고: Windows에서 심볼릭 링크 생성이 실패하면 관리자 권한으로 PowerShell을 실행하거나, 개발자 모드를 활성화하세요.

---

## 11) “막히면” 체크리스트 (When Stuck)
- 1) 해당 폴더 README/엔트리포인트(`*_cli.py`, `--help`)부터 확인했는가?
- 2) 입력/출력 포맷(엑셀 서식 포함)을 **샘플로 재현**했는가?
- 3) 변경 범위를 최소화했는가(레거시 경로/이름 유지)?
- 4) 스모크/테스트 로그를 남겼는가?
- 5) 문서(README/AGENTS.md)를 SSOT로 업데이트했는가?
- 6) 모듈별 `agent_*.md` 또는 `AGENTS.md`를 확인했는가?
- 7) 의존성 파일(`requirements.txt`, `pyproject.toml`)을 확인했는가?
- 8) `out/` 또는 `output/` 폴더의 기존 출력 포맷을 참고했는가?

---

## 12) Subagents and Skills (서브에이전트 및 스킬)

CONVERT 프로젝트는 Cursor Subagents와 Agent Skills를 사용하여 작업을 분리하고 표준화한다.

**사용 가이드**: `.cursor/USAGE_GUIDE.md` - 실전 사용법 및 예시  
**상세 가이드**: `subagentandskillguide.md` - 구조 및 설치 가이드

### 12.1 Subagents (서브에이전트)

Subagents는 `.cursor/agents/` 디렉토리에 위치하며, 특정 작업을 격리하여 수행한다:

| Subagent | 목적 | 사용 시기 | 권한 |
| --- | --- | --- | --- |
| `convert-scoper` | 인벤토리/엔트리포인트/스모크 커맨드 후보 생성 | 대규모 탐색 필요 시 | readonly |
| `verifier` | 완료된 작업 검증(테스트/스모크) | 작업 완료 후 확인 | 수정 가능 |
| `excel-style-guardian` | Excel 서식 회귀 방지 | CIPL/간트 등 Excel 산출물 작업 시 | readonly |
| `agi-schedule-updater` | AGI TR Schedule HTML 공지란·Weather & Marine Risk 블록 매일 갱신 | 공지/날씨 업데이트 필요 시 | 수정 가능 |

**사용 예시:**
- `/convert-scoper` - 프로젝트 구조 파악이 필요할 때
- `/verifier` - 구현 완료 후 검증이 필요할 때
- `/excel-style-guardian` - Excel 파일 서식 유지가 중요할 때
- `/agi-schedule-updater` - AGI TR Unit 1 Schedule 공지·날씨 블록 갱신이 필요할 때

### 12.2 Skills (스킬)

Skills는 `.cursor/skills/<name>/SKILL.md`에 위치하며, 반복 작업을 표준화한다:

| Skill | 목적 | 트리거 키워드 |
| --- | --- | --- |
| `convert-toolbox` | 인벤토리/스모크/패키지 검증 자동화 | inventory, smoke, verify, package |
| `mrconvert-run` | PDF/DOCX/XLSX 변환 실행 표준화 | mrconvert, convert pdf, OCR |
| `email-thread-search` | 이메일 검색/스레드 추적 표준화 | outlook export, thread, 메일 검색 |
| `cipl-excel-build` | CIPL Excel 생성(서식 유지) | CIPL, invoice packing list, xlsx template |
| `folder-cleanup` | 폴더 정리/임시 파일 삭제/중복 파일 식별/아카이브 생성 표준화 | cleanup, 정리, 폴더 정리, 임시 파일, 중복 파일 |
| `agi-schedule-daily-update` | AGI TR Schedule HTML 공지란·Weather & Marine Risk 블록 매일 갱신 | AGI schedule 공지, 날씨 블록 업데이트, Mina Zayed weather |
| `agi-schedule-shift` | AGI TR 일정(JSON/HTML) pivot 이후 전체 일정 delta일 시프트 | 일정 시프트, schedule shift, 일정 연기, AGI schedule delay |
| `weather-go-nogo` | SEA TRANSIT Go/No-Go 의사결정(3단 Gate: 임계값·Squall/피크파 버퍼·연속 window) | sea transit Go/No-Go, weather window, Hs/Hmax, squall buffer, marine weather decision |

### 12.3 검증 및 실행

**구조 검증:**
```bash
python .cursor/skills/convert-toolbox/scripts/validate_agent_assets.py --root .
```

**인벤토리 생성:**
```bash
python .cursor/skills/convert-toolbox/scripts/convert_inventory.py --root . --out out/convert_inventory.json
```

**스모크 테스트:**
```bash
python .cursor/skills/convert-toolbox/scripts/run_smoke.py --root .
```

**폴더 정리 분석 (dry-run):**
```bash
python .cursor/skills/folder-cleanup/scripts/cleanup_analyzer.py --root . --out out/cleanup_report.json
```

### 12.4 Codex 호환성

Codex를 사용하는 경우, `.codex/skills/`에 심볼릭 링크를 생성하거나 폴더를 복사할 수 있다:

**Windows PowerShell (관리자 권한 필요할 수 있음):**
```powershell
New-Item -ItemType SymbolicLink -Path ".codex\skills\convert-toolbox" -Target "..\..\.cursor\skills\convert-toolbox"
```

**Linux/macOS:**
```bash
ln -s ../../.cursor/skills/convert-toolbox .codex/skills/convert-toolbox
```

> 참고: Windows에서 심볼릭 링크 생성이 실패하면 폴더 복사로 대체할 수 있다.

### 12.5 워크플로우 통합

Subagents와 Skills는 기존 작업 루틴(섹션 4)과 통합된다:

1. **Locate 단계**: `/convert-scoper`로 구조 파악
2. **Verify 단계**: `/verifier` 또는 `convert-toolbox` 스모크 실행
3. **Excel 작업**: `/excel-style-guardian`으로 서식 회귀 방지
4. **모듈별 작업**: 해당 스킬 사용 (예: `mrconvert-run`, `email-thread-search`)
5. **AGI TR Schedule**: `/agi-schedule-updater` 또는 `agi-schedule-daily-update`(공지·날씨), `agi-schedule-shift`(일정 시프트)

---

## ZERO log

* 본 건은 “문서 작성” 작업이며 UAE 규정/통관/요율/ETA 등 실시간 근거 요구 항목이 없어 ZERO 게이트 비적용.

(내부 참고 문서)   

[1]: https://github.blog/changelog/2025-08-28-copilot-coding-agent-now-supports-agents-md-custom-instructions/ "Copilot coding agent now supports AGENTS.md custom instructions - GitHub Changelog"
[2]: https://developers.openai.com/codex/guides/agents-md/?utm_source=chatgpt.com "Custom instructions with AGENTS.md"
[3]: https://www.builder.io/blog/agents-md?utm_source=chatgpt.com "Improve your AI code output with AGENTS.md (+ my best ..."
[4]: https://www.anthropic.com/engineering/equipping-agents-for-the-real-world-with-agent-skills?utm_source=chatgpt.com "Equipping agents for the real world with Agent Skills"
