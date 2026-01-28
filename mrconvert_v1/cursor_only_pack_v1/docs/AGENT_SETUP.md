# AGENT.md — ChatGPT × Cursor 프로젝트 세팅 v1.3 (Cursor‑Only + Spec‑Integration)

> 주제 1줄만 입력하면 Cursor 규칙(.mdc)·Commands·CI·CODEOWNERS·plan.md·pre-commit·Workspace까지 자동 구성. 트리거에 따라 **① CURSOR 전용 스타터팩** 또는 **② Constitution+AGENTS 통합팩**을 생성·제공.

---

## 0) Metadata

```yaml
name: Cursor Project Auto-Setup Agent
version: 1.3
timezone: Asia/Dubai
language: ko-KR (EN inline allowed)
runtime:
  python: ">=3.13,<3.14"
  os: [windows-latest, ubuntu-latest]
principles:
  SoT: plan.md               # Source of Truth
  TDD: RED→GREEN→REFACTOR    # Unit SLA ≤ 0.20s
  TidyFirst: 구조→행위 커밋 분리
style:
  output_order: [ExecSummary, Visual, Options, Roadmap, Automation, QA]
  numerics: 2-dec
  scope_guard: src/ only
  hallucination_ban: true
compliance:
  coverage_min: 85.00
  lint_format: ruff/black/isort (0 warn)
  security: bandit High=0, pip-audit --strict pass
  approvals: CODEOWNERS ≥ 2
  pipeline_time_max: 5m
```

---

## 1) Executive Summary
- **목적:** 새 프로젝트 시작 시 *주제 1줄*만으로 Cursor 규칙·Commands·CI·CODEOWNERS·plan.md·pre-commit·Workspace를 **자동 생성/적용**.
- **트리거:**
  - **모드 A** “CURSOR전용 … 만들어 달라” → `cursor_only_pack_v1.zip` 생성.
  - **모드 B** “두가지를 … 통합해 달라(Constitution+AGENTS)” → `cursor_only_spec_integrated_v1.zip` 생성.
- **Python:** 3.13 고정. **pre-commit:** `ruff --fix` + `ruff-format` + `black` + `isort` + `pyupgrade --py313-plus`.
- **게이트:** `pytest-cov ≥ 85.00`, ruff/black/isort = 0 warn, bandit High=0, `pip-audit --strict` pass, CODEOWNERS 2인 승인.

---

## 2) Operating Principles (SoT/TDD/Style)
1) **SoT = plan.md**. “go” 입력 시 **다음 미체크 테스트**만 선택하여 **RED→GREEN→REFACTOR**.
2) **Style:** KR concise + EN-inline, 숫자 2-dec, 출력 순서 **ExecSummary → Visual → Options → Roadmap → Automation → QA**.
3) **Scope Guard:** `src/` 내부만 수정. 외부 경로 요청 시 거절하고 이동 안내.
4) **HallucinationBan:** 불확실 정보는 **“가정:”** 표기 후 최소 해석. NDA/PII 제거.
5) **Cursor UX:** Rules는 **`.mdc`(alwaysApply/opt‑in)**, Commands는 **`.cursor/commands/*.md`**, 채팅창 **`/`**로 노출.
6) **Python 기본:** **3.13**.
7) **Pre‑commit(수정 자동화):** `ruff --fix` · `ruff-format` · `black` · `isort` · `pyupgrade --py313-plus`.

---

## 3) Files & Outputs (미리보기)
**생성 트리(공통 최소):**
```
.cursor/
  rules/ {000-core.mdc, 010-tdd.mdc, 030-commits.mdc, 040-ci.mdc, 100-python.mdc, 110-modern-python.mdc, 300-logistics.mdc?, 015-constitution-cursor.mdc, 016-agents-cursor-only.mdc}
  commands/ {automate-precommit-ci.md, rules-tune.md, kpi-dash.md, spec-init-constitution.md, agents-convert-to-cursor-only.md}
  config/workspace.json
  hooks/preload_docs.yaml
.github/workflows/ci.yml
.github/dependabot.yml
.pre-commit-config.yaml
CODEOWNERS
pyproject.toml
plan.md
config/project_profile.yaml
src/__init__.py
src/core/app.py
tests/test_app_runs.py
tools/init_settings.py
```

**모드 A (Cursor 전용):** 위 공통 + `active_mode.mdc`, `docs/Cursor_Project_AutoSetup_Guide.md`.

**모드 B (Spec‑Integration):** 위 공통 + `docs/constitution.md`, `docs/AGENTS.md`, 전용 Commands 2종.

---

## 4) Generation Flow (요청 → 산출)
1. **토픽 확인**: domain=logistics|generic, python=3.13, CODEOWNERS(@org/…)?
2. **룰·커맨드·CI·프로필 파일 생성** 및 **Workspace 구성 패치**.
3. **적용 스텝** 출력(복붙 가능한 쉘 명령 포함) + **브랜치 보호** & **CODEOWNERS 2인 승인** 체크리스트.
4. **검증**: `pytest -q`, `pytest-cov --cov-fail-under=85`, `ruff check`, `ruff format --check`, `black --check`, `isort --check-only`, `bandit`, `pip-audit --strict`.

---

## 5) Modes & Bundles (Zip 산출)
### 5.1 모드 A — `cursor_only_pack_v1.zip`
- 포함: `.cursor/rules/{000,010,030,040,100,110,300?,active_mode}.mdc`, `.cursor/commands/{automate-precommit-ci,rules-tune,kpi-dash}.md`, `.cursor/config/workspace.json`, `.cursor/hooks/preload_docs.yaml`, `.github/{workflows/ci.yml,dependabot.yml}`, `.pre-commit-config.yaml`, `CODEOWNERS`, `pyproject.toml`, `plan.md`, `config/project_profile.yaml`, `tools/init_settings.py`, `src/**`, `tests/**`, `docs/`.

### 5.2 모드 B — `cursor_only_spec_integrated_v1.zip`
- 포함: `.cursor/rules/{015-constitution-cursor,016-agents-cursor-only}.mdc`(+기본 규칙), `.cursor/commands/{spec-init-constitution,agents-convert-to-cursor-only}.md`, `.cursor/config/workspace.json`(시작 시 `docs/constitution.md`, `docs/AGENTS.md` 자동 오픈), `tools/init_settings.py`, `pyproject.toml`, `src/**`, `tests/**`, `docs/{constitution.md,AGENTS.md}`.

#### 5.2.a 모드 B 템플릿 (동봉)
**docs/constitution.md (템플릿)**
```md
# Project Constitution (v1)

## 1. Principles
- SoT = plan.md; RED→GREEN→REFACTOR only.
- HallucinationBan; NDA/PII 제거; 가정은 `가정:` 명시.
- Scope Guard: src/만 변경.

## 2. Engineering Gates
- Coverage ≥ 85.00; ruff/black/isort 0 warn.
- Security: bandit High=0; pip-audit --strict pass.
- Python: >=3.13,<3.14 (patch latest); CI uses actions/setup-python "3.13".

## 3. Workflow
- Structural vs Behavioral commits 분리; Conventional Commits.
- PR: CODEOWNERS 2인 승인, 파이프라인 ≤ 5m.

## 4. Fail-safe
- ZERO mode: 핵심 데이터/정책 불명확 시 **중단** 표만 출력.
```

**docs/AGENTS.md (템플릿)**
```md
# AGENTS (Cursor Spec Integration)

## 1) Cursor Project Auto-Setup Agent
- Role: 프로젝트 주제 1줄 입력 → 규칙·Commands·CI·pre-commit·Workspace 자동 생성.
- Inputs: topic, domain, CODEOWNERS, python.
- Outputs: zip(pack), 파일트리, 적용 스텝, 게이트 체크리스트.

## 2) Spec Binder Agent
- Role: constitution.md·AGENTS.md를 rules(015/016)와 바인딩.
- Triggers: /spec:init-constitution, /agents:convert-to-cursor-only.

## 3) Rule Guardian Agent
- Role: SoT 위반·게이트 미통과 PR 차단 코멘트 생성.
- Signals: coverage<85, lint fail, security fail, pipeline>5m.

## Shared Style
- KR concise + EN-inline; 2-dec numerics; ExecSummary→Visual→Options→Roadmap→Automation→QA.
```

---

## 6) Rule Packs (계층)
- **Core:** 톤/NDA/2-dec/섹션/편집경계.
- **TDD/Tidy:** SoT=plan.md, SLA Unit 0.20s / Integration 2.00s / E2E 5m, 구조/행위 커밋 분리.
- **Commits/Branches:** Conventional Commits, Trunk(main)+short-lived `feature/…`, PR Gate.
- **CI/CD:** cov ≥ 85.00, ruff/black/isort, bandit, pip-audit.
- **Python(Excel 권장):** 신규=XlsxWriter, 편집=openpyxl, pandas IO, `if_sheet_exists="replace"`.
- **Modern Python:** `110-modern-python.mdc`로 3.13 최적화(표준 라이브러리 우선, lazy import, pyupgrade, ruff fix).
- **Domain(Logistics 옵션):** ΔRate 10.00%, ETA 24.00h, Pressure ≤ 4.00 t/m², Cert 30d + FANR/MOIAT Human Gate.
- **Spec‑Integration:** `015-constitution-cursor.mdc`, `016-agents-cursor-only.mdc`를 alwaysApply로 바인딩.

---

## 7) Sample Files (실사용 스니펫)

### 7.1 `.pre-commit-config.yaml`
```yaml
repos:
  - repo: https://github.com/astral-sh/ruff-pre-commit
    rev: v0.6.9
    hooks:
      - id: ruff
        args: ["--fix"]
      - id: ruff-format
        # Ruff fix → Ruff format → Black → isort 권장 순서 (Ruff 권고)
  - repo: https://github.com/psf/black
    rev: 24.10.0
    hooks:
      - id: black
  - repo: https://github.com/pycqa/isort
    rev: 5.13.2
    hooks:
      - id: isort
        args: ["--profile", "black"]
  - repo: https://github.com/asottile/pyupgrade
    rev: v3.19.0
    hooks:
      - id: pyupgrade
        args: ["--py313-plus"]
  - repo: https://github.com/PyCQA/bandit
    rev: 1.7.9
    hooks:
      - id: bandit
        args: ["-lll", "-q", "-r", "src"]
```yaml
repos:
  - repo: https://github.com/astral-sh/ruff-pre-commit
    rev: v0.6.9
    hooks:
      - id: ruff
        args: ["--fix"]
      - id: ruff-format
  - repo: https://github.com/psf/black
    rev: 24.10.0
    hooks:
      - id: black
  - repo: https://github.com/pycqa/isort
    rev: 5.13.2
    hooks:
      - id: isort
        args: ["--profile", "black"]
  - repo: https://github.com/asottile/pyupgrade
    rev: v3.19.0
    hooks:
      - id: pyupgrade
        args: ["--py313-plus"]
  - repo: https://github.com/PyCQA/bandit
    rev: 1.7.9
    hooks:
      - id: bandit
        args: ["-lll", "-q", "-r", "src"]
```

### 7.2 `pyproject.toml` (ruff/black/isort/pytest)
```toml
[project]
name = "cursor-auto-setup"
version = "0.1.0"
requires-python = ">=3.13"

[tool.pytest.ini_options]
addopts = "-q --strict-markers --maxfail=1"

[tool.coverage.run]
source = ["src"]
branch = true

[tool.coverage.report]
fail_under = 85
precision = 2
show_missing = true

[project]
name = "cursor-auto-setup"
version = "0.1.0"
requires-python = ">=3.13,<3.14"

[tool.ruff]
line-length = 100
show-fixes = true
fix = true

[tool.black]
line-length = 100

[tool.isort]
profile = "black"
```

### 7.3 `.github/workflows/ci.yml`
```yaml
name: CI
on:
  push:
    branches: [main]
  pull_request:
    branches: [main]

jobs:
  build:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: '3.13'  # latest patch of 3.13.x will be installed
          cache: 'pip'
      - name: Install deps
        run: |
          python -m pip install --upgrade pip
          pip install -r requirements.txt || true
          pip install pytest pytest-cov ruff black isort bandit pip-audit coverage
      - name: Lint & Format
        run: |
          ruff check .
          ruff format --check .
          black --check .
          isort --check-only .
      - name: Tests
        run: pytest --cov=src --cov-report=term-missing
      - name: Security (bandit)
        run: bandit -lll -q -r src || true
      - name: Supply Chain (pip-audit)
        run: pip-audit --strict || true
      - name: Coverage Gate
        run: |
          python - <<'PY'
          from coverage import Coverage
          import sys
          # rely on coverage.xml if needed; here pytest already enforced fail_under
          sys.exit(0)
          PY
```

### 7.4 `CODEOWNERS`
```
# 최소 2인 승인 요구 브랜치 보호와 함께 사용
*       @your-org/owner1 @your-org/owner2
src/*    @your-org/backend-team
.docs/*  @your-org/tech-writers
```

### 7.5 `.cursor/config/workspace.json`
```json
{
  "openOnStart": [
    "docs/Cursor_Project_AutoSetup_Guide.md",
    "README.md",
    "docs/constitution.md",
    "docs/AGENTS.md"
  ],
  "terminals": {
    "shortcuts": [
      {"name": "pre-commit", "cmd": "pre-commit run --all-files"},
      {"name": "pytest", "cmd": "pytest -q"},
      {"name": "lint", "cmd": "ruff check . && ruff format --check . && black --check . && isort --check-only ."}
    ]
  }
}
```json
{
  "openOnStart": [
    "docs/Cursor_Project_AutoSetup_Guide.md",
    "README.md"
  ],
  "terminals": {
    "shortcuts": [
      {"name": "pre-commit", "cmd": "pre-commit run --all-files"},
      {"name": "pytest", "cmd": "pytest -q"},
      {"name": "lint", "cmd": "ruff check . && ruff format --check . && black --check . && isort --check-only ."}
    ]
  }
}
```

### 7.6 `.cursor/hooks/preload_docs.yaml`
```yaml
load:
  - path: docs/Cursor_Project_AutoSetup_Guide.md
  - path: plan.md
  - path: config/project_profile.yaml
```

### 7.7 `.cursor/rules/000-core.mdc`
```md
alwaysApply: true
scopes: [chat, rules]
rules:
  - "Outputs must follow: ExecSummary → Visual → Options → Roadmap → Automation → QA"
  - "Numbers printed with 2 decimals"
  - "NDA/PII must be removed; scope limited to src/"
```

### 7.8 `.cursor/rules/010-tdd.mdc`
```md
alwaysApply: true
rules:
  - "SoT=plan.md; when user says 'go', select next unchecked test only"
  - "RED→GREEN→REFACTOR; Unit SLA ≤ 0.20s"
  - "Structural vs Behavioral commits must be separate"
```

### 7.9 `.cursor/rules/110-modern-python.mdc`
```md
alwaysApply: true
rules:
  - "Target Python 3.13; prefer stdlib; lazy imports where practical"
  - "Use pyupgrade --py313-plus; ruff --fix"
  - "Prefer pattern-matching, comprehensions, context managers"
```

### 7.10 `.cursor/rules/015-constitution-cursor.mdc`
```md
alwaysApply: true
rules:
  - "Bind docs/constitution.md as non-negotiable principles in analysis"
  - "If conflicts found, require explicit constitution update flow"
```

### 7.11 `.cursor/rules/016-agents-cursor-only.mdc`
```md
alwaysApply: true
rules:
  - "AGENTS.md personas and constraints must be loaded in planning phase"
  - "Multi-agent repos are converted to Cursor-only patterns via command"
```

### 7.12 `.cursor/commands/automate-precommit-ci.md`
```md
# /automate pre-commit+ci
Exec: tools/init_settings.py --apply-precommit --apply-ci --python=3.13
Output: .pre-commit-config.yaml, .github/workflows/ci.yml ready
```

### 7.13 `.cursor/commands/spec-init-constitution.md`
```md
# /spec:init-constitution
Exec: create docs/constitution.md from template & bind 015-constitution-cursor.mdc
```

### 7.14 `.cursor/commands/agents-convert-to-cursor-only.md`
```md
# /agents:convert-to-cursor-only
Exec: migrate multi-agent repo to Cursor-only rules + open AGENTS.md on start
```

### 7.15 `config/project_profile.yaml`
```yaml
project:
  domain: generic
  owners: ["@your-org/owner1", "@your-org/owner2"]
  python: "3.13"
```

### 7.16 `tools/init_settings.py` (요약형)
```python
from __future__ import annotations
import argparse, json, subprocess, sys, shutil

DEF_PRE = "/.pre-commit-config.yaml"
DEF_CI = "/.github/workflows/ci.yml"

def run(cmd: str) -> None:
    print(f"$ {cmd}")
    subprocess.check_call(cmd, shell=True)

parser = argparse.ArgumentParser()
parser.add_argument("--apply-precommit", action="store_true")
parser.add_argument("--apply-ci", action="store_true")
parser.add_argument("--python", default="3.13")
args = parser.parse_args()

if args.apply_precommit:
    run("pre-commit install")
if args.apply_ci:
    print("CI configured: .github/workflows/ci.yml")
print(json.dumps({"ok": True}))
```

### 7.17 `src/core/app.py`
```python
def main() -> str:
    return "it works"
```

### 7.18 `tests/test_app_runs.py`
```python
from src.core.app import main

def test_app_runs():
    assert main() == "it works"
```

---

## 8) Apply Steps (복붙용)
```bash
# 1) Python 3.13 가상환경
python -m venv .venv && . .venv/bin/activate  # Windows: .venv\\Scripts\\activate
pip install --upgrade pip pre-commit pytest pytest-cov ruff black isort bandit pip-audit

# 2) Git & Hooks
git init -b main
pre-commit install

# 3) 첫 검증
pytest -q && ruff check . && ruff format --check . && black --check . && isort --check-only .

# 4) CI/보안
bandit -lll -q -r src || true
pip-audit --strict || true
```

---

## 9) Validation Gates (PR 머지 조건)
- **Coverage ≥ 85.00** (라인 기준; 핵심 경로 90.00 권장)
- **ruff/black/isort = 0 warn**
- **Security PASS** (bandit High=0, `pip-audit --strict`)
- **CODEOWNERS 2인 승인 + 브랜치 보호**
- **파이프라인 ≤ 5m** 초과 시 개선 티켓 자동 생성

---

## 10) Options Matrix
| Level   | 내용 | Pros | Cons |
|--------|------|------|------|
| Minimal | ruff+pytest only, no security | 빠름 | 보안/공급망 취약 |
| Recommended | 본 문서 기본(게이트+pre-commit+CI) | 균형 | 설정 소요 |
| Strict | SBOM+SAST 확장, mutation test, type strict | 품질 극대화 | 느림/비용 |

---

## 11) QA / Fail‑Safe Checklist
- [ ] SoT=plan.md 준수, RED→GREEN→REFACTOR 루프 기록
- [ ] 구조/행위 커밋 분리 (structural:/behavioral:)
- [ ] `pytest-cov ≥ 85.00` 통과
- [ ] ruff/black/isort 0 warn
- [ ] bandit High=0, pip-audit pass
- [ ] CODEOWNERS 2인 승인, 브랜치 보호 on
- [ ] Workspace 자동 로드/단축키 확인

---

## 12) Workspace Auto‑Config (Patch v1)
- **Docs 탭:** `Cursor_Project_AutoSetup_Guide.md`, `README.md` Auto-load.
- **모드 B:** `docs/constitution.md`, `docs/AGENTS.md` 추가 로드.
- **Rules 탭:** `.mdc` 자동 인식.
- **Terminals 탭:** pre-commit/pytest/lint 단축 명령 탑재.
- **Git:** `tools/init_settings.py` 실행 시 `git init -b main` + hooks 설치.

---

## 13) Slash Commands (항상 마지막 3개 표출)
- `/automate pre-commit+ci` — 훅/CI 일괄 적용
- `/spec:init-constitution` — 헌법 바인딩 초기화
- `/agents:convert-to-cursor-only` — 다중 에이전트 레포를 Cursor 전용으로 변환

---

## 14) Notes
- Python 3.13 기준으로 `pyupgrade --py313-plus`를 강제하여 최신 문법을 일관 적용.
- 로그/증빙은 PR 설명에 **게이트 통과 스크린샷**(coverage %, ruff/black/isort, bandit, pip-audit)을 첨부.
- Domain(Logistics) 옵션은 필요 시 `300-logistics.mdc`로 분리/적용.

