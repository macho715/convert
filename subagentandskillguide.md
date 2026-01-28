Exec (3–5L)

아래는 **CONVERT 폴더(업무 자동화 프로그램 개발 중심)**에 맞춰, Cursor Subagent + Agent Skills + Codex Skills를 동시에 호환하도록 만든 문서 패키지입니다.

목표는 “사용자 최소 개입”을 위해 (1) 대규모 탐색(인벤토리) 격리, (2) 검증(스모크/테스트) 독립화, (3) Excel 서식 회귀 보호를 Subagent로 분리하고, 반복 루틴은 Skills로 표준화하는 것입니다.

스킬/서브에이전트는 프로젝트 스코프(.cursor/.codex) 기준이며, **symlink(권장) 또는 복사(대안)**로 운영합니다.

본 결과물은 “문서 작성”이므로 UAE 규정/통관/요율/ETA 등 실시간 근거 요구 항목이 없어 ZERO 게이트 비적용입니다.

EN Sources (≤3)

Cursor Docs — Subagents (Accessed: 2026-01-28)

Cursor Docs — Agent Skills (Accessed: 2026-01-28)

OpenAI Developers — Codex Skills (Accessed: 2026-01-28)

(내부 근거 파일) Cursor Subagent 포맷/필드 및 경로 호환 개요 

Agent Skills_antigravity

 / Cursor Skills 포맷 개요 

Agent Skills_CURSOR

 / Codex Skills 포맷·경로 개요 

AGENT SKILL_CODEX

핵심 설정 요약 (Visual)
No	Item	Value	Risk	Evidence/가정
1	Subagents	3개(인벤토리/검증/Excel서식)	과도한 에이전트 난립	Cursor Subagents 개요
2	Skills	4개(툴박스/mrconvert/email/CIPL)	스킬 과대범위	Skills는 “작고 명확” 권장
3	설치 경로	Cursor: .cursor/* / Codex: .codex/*	이중 관리	Codex/ Cursor 스킬 경로
4	운영 원칙	Backwards compatible + Ask-first(파괴/대량/바이너리)	레거시 파손	사용자 제공 AGENTS.md 초안 전제
5	검증	compileall + (조건부) pytest + 모듈 스모크	실행환경 편차	가정: 모듈 혼합 구조
Skill Map
skill name	1줄 목적	트리거 키워드	리소스	위험/권한
convert-toolbox	인벤토리/스모크/패키지 검증을 자동 표준화	inventory, smoke, verify, package	scripts 3종	읽기/실행 위주(안전)
mrconvert-run	mrconvert_v1/ 변환 실행·출력 규격 고정	mrconvert, PDF/DOCX/XLSX, OCR	없음(문서형)	대용량 파일/시간
email-thread-search	email_search/ 검색·스레드 빌드·익명 샘플 규칙	outlook export, thread, search	없음(문서형)	PII(강제 마스킹)
cipl-excel-build	CIPL/ Excel 생성(서식 유지)·회귀 체크	CIPL, Excel, template, xlsx	없음(문서형)	서식 파손(강제 가드)
Subagent Map
subagent name	1줄 목적	트리거 문구	권한(권장)	비고
convert-scoper	코드베이스/엔트리포인트/입출력 계약 인벤토리	“폴더 구조 파악”, “엔트리포인트 찾아”	readonly=true	컨텍스트 격리용
verifier	변경사항 검증(테스트/스모크) 및 PASS/FAIL 리포트	“검증해”, “테스트 돌려”	model=fast	Cursor 권장 패턴
excel-style-guardian	Excel 산출물 서식/레이아웃 회귀 방지	“서식 유지”, “CIPL 포맷”	readonly=true	바이너리 자동수정 금지
생성 파일 트리 (프로젝트 스코프)
.cursor/
  agents/
    convert-scoper.md
    verifier.md
    excel-style-guardian.md
  skills/
    convert-toolbox/
      SKILL.md
      scripts/
        convert_inventory.py
        run_smoke.py
        validate_agent_assets.py
    mrconvert-run/
      SKILL.md
    email-thread-search/
      SKILL.md
    cipl-excel-build/
      SKILL.md

.codex/
  skills/
    (권장) convert-toolbox -> ../../.cursor/skills/convert-toolbox
    (권장) mrconvert-run -> ../../.cursor/skills/mrconvert-run
    (권장) email-thread-search -> ../../.cursor/skills/email-thread-search
    (권장) cipl-excel-build -> ../../.cursor/skills/cipl-excel-build
  agents/
    (선택) verifier.md (Cursor와 동일 파일을 복사/링크)
    (선택) convert-scoper.md
    (선택) excel-style-guardian.md


Codex는 symlinked skill 폴더를 지원합니다.
Cursor도 프로젝트 스킬을 .cursor/skills/에서 로드합니다.
Subagent는 .cursor/agents/에 YAML frontmatter 포함 MD로 정의합니다.

파일별 내용 (복사-붙여넣기)
1) Cursor Subagents
.cursor/agents/convert-scoper.md
---
name: convert-scoper
description: CONVERT 폴더 인벤토리(엔트리포인트/의존성/입출력 계약/스모크 커맨드 후보) 생성. 대규모 탐색이 필요할 때 우선 사용.
model: fast
readonly: true
is_background: true
---

너는 CONVERT 폴더 전용 “인벤토리 스코퍼”다. 목적은 메인 에이전트의 컨텍스트를 소모하지 않고, 아래 산출물을 **간결하게** 반환하는 것이다.

## 작업 범위
1) 구조 스캔
- 최상위/하위 폴더에서 README, 실행 스크립트, 설정파일을 찾는다:
  - pyproject.toml, requirements.txt, environment.yml, Pipfile, setup.cfg
  - *_cli.py, __main__.py, main.py, app.py, streamlit app, vba/xlsm builder

2) 엔트리포인트 후보 식별
- “실행 방법”을 추측하지 말고, 파일명/--help/README 근거로 후보만 나열한다.

3) I/O 계약(입력/출력) 요약
- 각 모듈별 입력(예: PDF/XLSX/Excel export)과 출력(out/, output/, reports/) 관례를 정리한다.

4) 스모크 커맨드 “후보” 생성
- 공통: python -m compileall -q .
- 조건부: pytest -q (pytest 설정 존재 시)
- 모듈별: 각 엔트리포인트의 --help 또는 최소 실행 1회(단, 실행은 메인 에이전트가 수행)

## 출력 포맷(반드시 준수)
- (A) Inventory Table: | Module | Entry Points | Inputs | Outputs | Risks |
- (B) Fixed Smoke Command Draft: 실행 커맨드 후보 3~8개
- (C) PATCH PLAN: 업데이트 권장 파일과 변경 요약(예: README/AGENTS.md에 커맨드 고정)

## 금지
- 코드 변경/리네임/삭제 제안은 하되, readonly이므로 직접 수정하지 않는다.
- PII/자격증명 관련 데이터는 출력에 포함하지 않는다.

.cursor/agents/verifier.md
---
name: verifier
description: Validates completed work. Use after tasks are marked done to confirm implementations are functional.
model: fast
readonly: false
---

너는 “회의적인 검증자(verifier)”다. 완료 주장(Implemented/Fixed/Done)을 **그대로 믿지 말고** 증거로 검증한다.

## 검증 절차
1) 무엇이 완료라고 주장되었는지 1~5줄로 재정의
2) 변경된 파일/영향 범위 추적(최소 diff 원칙 위반 여부 포함)
3) 아래 순서로 검증 수행(가능한 경우 실제 실행 로그 포함)
   - python -m compileall -q .
   - pytest -q (pytest 설정/테스트 존재 시)
   - 모듈별 스모크(엔트리포인트 --help, 샘플 1건 실행 등)
4) 실패 시
   - Root cause 1줄
   - 최소 수정안(Minimal fix) 제시
   - 재검증 커맨드 재제시

## 리포트 포맷(반드시)
- PASS/FAIL 한 줄 Verdict
- Evidence Table: | Check | Result | Command | Notes |
- Gaps: 미검증 항목/환경 의존 항목
- “Ask first” 필요한 추가 작업(의존성 설치, 대량 변경, 바이너리 수정 등)

.cursor/agents/excel-style-guardian.md
---
name: excel-style-guardian
description: Excel(xlsx/xlsm) 산출물의 서식/레이아웃 회귀를 방지. CIPL/간트/매크로 산출물이 “기존 포맷 유지”가 최우선일 때 사용.
model: inherit
readonly: true
is_background: false
---

너는 Excel 산출물의 “서식 회귀(Regression)”를 막는 가디언이다. 목적은 **데이터 정확성 + 서식 동일성**을 동시에 확인하는 것이다.

## 체크 항목(우선순위)
1) 템플릿/기존 산출물 대비 “시각 요소” 유지
- 시트명, 컬럼 순서, 헤더 라인, 병합셀, 테두리, 폰트/정렬, 인쇄영역(있는 경우)

2) 데이터 요소
- 주요 키 필드(예: Case No, BL, PO, HS, GW/NW 등)의 누락/위치 변경 여부

3) 매크로(xlsm) 안전
- xlsm 바이너리는 자동 수정 금지(필요 시 “Ask first”)

## 출력 포맷
- Visual Regression Checklist: | Item | Same? | Evidence | Risk |
- Blockers: “이 변경이 왜 위험한지” 1~3줄
- Safe Fix Suggestion: 서식 파손을 피하는 수정 방향(예: builder 스크립트에서 cell style copy)

2) Cursor Skills (+ Codex Skills 공용)

아래 스킬 폴더는 **.cursor/skills/<name>/SKILL.md**에 두면 Cursor가 자동 로드합니다.
동일 폴더를 .codex/skills/로 링크/복사하면 Codex도 로드합니다.

.cursor/skills/convert-toolbox/SKILL.md
---
name: convert-toolbox
description: CONVERT 폴더에서 인벤토리(엔트리포인트/의존성) 생성, 스모크(compileall/pytest) 실행, 스킬·서브에이전트 패키지 정합성 검증을 표준화한다. "inventory", "smoke", "verify", "package" 작업에 사용.
---

# convert-toolbox

## 언제 사용
- CONVERT 폴더 구조를 빠르게 파악해야 할 때(엔트리포인트/입출력 규칙/의존성)
- 변경 후 스모크/테스트 PASS/FAIL을 증거로 남겨야 할 때
- Subagent/Skill 패키지의 **이름 규칙/경로/형식**을 검증해야 할 때

## 안전 규칙
- 기본은 읽기/검증 위주.
- 다음 작업은 반드시 Ask first:
  - 새 의존성 설치/업그레이드
  - 대량 이동/리네임/삭제
  - xlsm 바이너리 자동 수정
  - 운영 데이터(PII 포함)로 실행

## 표준 실행(권장)
1) 인벤토리 생성
- 실행:
  - python .cursor/skills/convert-toolbox/scripts/convert_inventory.py --root . --out out/convert_inventory.json
- 산출물:
  - out/convert_inventory.json (Git 제외 권장)

2) 스모크 실행
- 실행:
  - python .cursor/skills/convert-toolbox/scripts/run_smoke.py --root .
- 결과:
  - compileall 결과 + (조건부) pytest 결과를 요약 출력

3) 패키지 정합성 검증(스킬/서브에이전트)
- 실행:
  - python .cursor/skills/convert-toolbox/scripts/validate_agent_assets.py --root .
- 검사:
  - skill name 규칙(소문자/숫자/하이픈)
  - 폴더명 == SKILL.md frontmatter name
  - subagent YAML frontmatter 존재 여부

## 리포트 포맷(권장)
- Evidence Table: | Check | Result | Command | Notes |
- FAIL이면: 원인 1줄 + 최소 수정안 + 재실행 커맨드

.cursor/skills/convert-toolbox/scripts/convert_inventory.py
#!/usr/bin/env python3
import argparse
import json
import os
import re
from datetime import datetime

ENTRYPOINT_HINTS = (
    "__main__.py",
    "main.py",
    "app.py",
)

CONFIG_HINTS = (
    "pyproject.toml",
    "requirements.txt",
    "environment.yml",
    "Pipfile",
    "setup.cfg",
)

README_HINTS = ("README.md", "readme.md")

def is_probable_entrypoint(filename: str) -> bool:
    base = os.path.basename(filename)
    if base in ENTRYPOINT_HINTS:
        return True
    if base.endswith("_cli.py"):
        return True
    return False

def scan(root: str):
    modules = []
    for dirpath, dirnames, filenames in os.walk(root):
        # skip common noise
        parts = set(dirpath.split(os.sep))
        if any(p in parts for p in (".git", ".venv", "node_modules", "dist", "build")):
            continue

        hits = {
            "readme": [],
            "configs": [],
            "entrypoints": [],
            "excel": [],
        }

        for fn in filenames:
            if fn in README_HINTS:
                hits["readme"].append(os.path.join(dirpath, fn))
            if fn in CONFIG_HINTS:
                hits["configs"].append(os.path.join(dirpath, fn))
            if fn.lower().endswith((".xlsx", ".xlsm")):
                hits["excel"].append(os.path.join(dirpath, fn))
            if fn.lower().endswith(".py") and is_probable_entrypoint(fn):
                hits["entrypoints"].append(os.path.join(dirpath, fn))

        if any(hits.values()):
            modules.append({
                "path": dirpath,
                "readme": sorted(hits["readme"]),
                "configs": sorted(hits["configs"]),
                "entrypoints": sorted(hits["entrypoints"]),
                "excel": sorted(hits["excel"])[:50],  # cap
            })

    return modules

def main():
    ap = argparse.ArgumentParser(description="CONVERT folder inventory (entrypoints/configs/readmes/excel).")
    ap.add_argument("--root", default=".", help="Root directory to scan.")
    ap.add_argument("--out", default="", help="Write JSON output to file path.")
    args = ap.parse_args()

    payload = {
        "generated_at": datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
        "root": os.path.abspath(args.root),
        "modules": scan(args.root),
    }

    data = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.out:
        os.makedirs(os.path.dirname(args.out), exist_ok=True)
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(data)
    else:
        print(data)

if __name__ == "__main__":
    main()

.cursor/skills/convert-toolbox/scripts/run_smoke.py
#!/usr/bin/env python3
import argparse
import os
import subprocess
import sys

def run(cmd, cwd):
    p = subprocess.run(cmd, cwd=cwd, text=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    return p.returncode, p.stdout

def has_pytest(root: str) -> bool:
    # heuristic: tests/ or pytest.ini or pyproject has [tool.pytest]
    if os.path.isdir(os.path.join(root, "tests")):
        return True
    for fn in ("pytest.ini", "pyproject.toml"):
        if os.path.exists(os.path.join(root, fn)):
            return True
    return False

def main():
    ap = argparse.ArgumentParser(description="Conservative smoke runner: compileall + optional pytest.")
    ap.add_argument("--root", default=".", help="Project root.")
    args = ap.parse_args()

    root = os.path.abspath(args.root)

    checks = []

    rc, out = run([sys.executable, "-m", "compileall", "-q", "."], cwd=root)
    checks.append(("compileall", rc, f"{sys.executable} -m compileall -q .", out[-2000:]))

    if has_pytest(root):
        rc2, out2 = run([sys.executable, "-m", "pytest", "-q"], cwd=root)
        checks.append(("pytest", rc2, f"{sys.executable} -m pytest -q", out2[-2000:]))

    verdict = "PASS" if all(rc == 0 for _, rc, _, _ in checks) else "FAIL"
    print(f"VERDICT: {verdict}")
    print("| Check | Result | Command | Notes |")
    print("| --- | --- | --- | --- |")
    for name, rcx, cmd, notes in checks:
        res = "PASS" if rcx == 0 else f"FAIL({rcx})"
        safe_notes = notes.replace("\n", " ")[:300]
        print(f"| {name} | {res} | `{cmd}` | {safe_notes} |")

    if verdict != "PASS":
        sys.exit(1)

if __name__ == "__main__":
    main()

.cursor/skills/convert-toolbox/scripts/validate_agent_assets.py
#!/usr/bin/env python3
import argparse
import os
import re
import sys

NAME_RE = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")

def read_frontmatter_name(skill_md_path: str) -> str:
    with open(skill_md_path, "r", encoding="utf-8") as f:
        txt = f.read()
    if not txt.startswith("---"):
        return ""
    # naive YAML frontmatter parse: find name: line before second '---'
    fm_end = txt.find("\n---", 3)
    if fm_end == -1:
        return ""
    fm = txt[3:fm_end]
    for line in fm.splitlines():
        if line.strip().startswith("name:"):
            return line.split(":", 1)[1].strip()
    return ""

def validate_skills(root: str):
    problems = []
    skill_roots = [
        os.path.join(root, ".cursor", "skills"),
        os.path.join(root, ".codex", "skills"),
    ]
    for sr in skill_roots:
        if not os.path.isdir(sr):
            continue
        for name in os.listdir(sr):
            skill_dir = os.path.join(sr, name)
            if not os.path.isdir(skill_dir):
                continue
            if not NAME_RE.match(name):
                problems.append(f"[SKILL] invalid folder name: {skill_dir}")
            skill_md = os.path.join(skill_dir, "SKILL.md")
            if not os.path.exists(skill_md):
                problems.append(f"[SKILL] missing SKILL.md: {skill_dir}")
                continue
            fm_name = read_frontmatter_name(skill_md)
            if fm_name and fm_name != name:
                problems.append(f"[SKILL] name mismatch folder({name}) != frontmatter({fm_name}) in {skill_md}")
    return problems

def validate_subagents(root: str):
    problems = []
    agent_dirs = [
        os.path.join(root, ".cursor", "agents"),
        os.path.join(root, ".codex", "agents"),
    ]
    for ad in agent_dirs:
        if not os.path.isdir(ad):
            continue
        for fn in os.listdir(ad):
            if not fn.endswith(".md"):
                continue
            path = os.path.join(ad, fn)
            with open(path, "r", encoding="utf-8") as f:
                head = f.read(200)
            if not head.startswith("---"):
                problems.append(f"[AGENT] missing YAML frontmatter: {path}")
    return problems

def main():
    ap = argparse.ArgumentParser(description="Validate Cursor/Codex agent-skill assets.")
    ap.add_argument("--root", default=".", help="Repo root.")
    args = ap.parse_args()
    root = os.path.abspath(args.root)

    problems = []
    problems += validate_skills(root)
    problems += validate_subagents(root)

    if problems:
        print("VERDICT: FAIL")
        for p in problems:
            print(p)
        sys.exit(1)

    print("VERDICT: PASS")

if __name__ == "__main__":
    main()

.cursor/skills/mrconvert-run/SKILL.md
---
name: mrconvert-run
description: mrconvert_v1에서 PDF/DOCX/XLSX를 TXT/MD/JSON으로 변환하는 실행 루틴을 표준화한다. "mrconvert", "convert pdf", "OCR", "table extract" 요청에 사용.
---

# mrconvert-run

## 언제 사용
- mrconvert_v1 변환 파이프라인 실행/수정/디버그
- 출력 폴더(out/output) 규칙을 고정하고, 레거시 동작을 깨지 않게 확장

## 입력 카드(가능하면 확보)
- Input: 파일 경로(로컬), 타입(PDF/DOCX/XLSX), 목표 출력(TXT/MD/JSON), OCR 필요 여부
- Output: 저장 경로(out/ 또는 output/), 파일명 규칙
- Constraints: 네트워크 사용 금지/허용, 대용량 제한

## 절차(보수적)
1) 엔트리포인트 확인
- mrconvert_v1 폴더에서 README 또는 *_cli.py / main.py / --help 를 먼저 확인
- “추측 실행” 금지

2) Dry-run 성격의 최소 실행
- --help 또는 샘플 1건 변환(가능하면 익명 샘플)

3) 출력 규칙
- 기본: out/ 또는 output/ 하위에 생성
- 변환 결과는 Git 추적 제외 권장(.gitignore)

4) 검증
- 변환 결과 존재 여부 + 파일 크기 0 여부
- (테이블 추출이면) JSON schema 키 최소 확인(없으면 가정/중단)

## Ask first
- OCR 엔진/대형 의존성 설치
- 대량 변환(폴더 전체) 실행
- 운영 문서(PII 포함)로 재현

## 산출물
- 실행 커맨드(확정본)
- “입력→출력” 매핑 표 1개
- 실패 시: 원인 1줄 + 최소 수정안 + 재시도 커맨드

.cursor/skills/email-thread-search/SKILL.md
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
- 샘플 데이터로 “검색 1건 + 스레드 빌드 1회” 재현

3) 결과 정리
- 결과를 out/ 아래에 저장
- 리포트: | Query | Hits | Threaded? | Output Path | Notes |

## Ask first
- 대용량 원본(운영) export 전체를 로드/분석
- 추가 라이브러리 설치

.cursor/skills/cipl-excel-build/SKILL.md
---
name: cipl-excel-build
description: CIPL(Commercial Invoice & Packing List) Excel 생성 작업에서 템플릿 서식 유지(Style-first)와 회귀 체크를 강제한다. "CIPL", "invoice packing list", "xlsx template" 요청에 사용.
---

# cipl-excel-build

## 목표
- CIPL Excel을 생성/수정하되, “기존 템플릿 서식/레이아웃”을 깨지 않는다(Style-first).
- 데이터 정확성과 서식 동일성을 함께 만족.

## 입력 카드
- 템플릿 파일 경로(xlsx/xlsm)
- 입력 데이터(가능하면 익명): item list, shipper/consignee, Incoterm, HS(있으면)
- 출력 경로(out/ 또는 output/)

## 절차(강제)
1) 템플릿 기준 고정
- 템플릿의 시트명/헤더/컬럼 순서/병합셀/테두리/인쇄영역을 SSOT로 간주

2) 생성 스크립트 엔트리포인트 확인
- CIPL 폴더(또는 CIPL_PATCH_PACKAGE)의 make_* 스크립트/README/--help 우선

3) 회귀 체크(서식)
- excel-style-guardian 서브에이전트 관점의 체크리스트로 “Same/Not same”를 표로 남김

## Ask first
- xlsm 바이너리 자동 수정
- 템플릿 구조(시트/컬럼) 대규모 변경

## 산출물
- 실행 커맨드(확정본)
- 서식 회귀 체크리스트 표
- FAIL이면: 어떤 서식이 깨졌는지 + 안전한 수정 방향(builder에서 style copy 등)

설치 안내 (플랫폼별 경로)
Cursor (Project)

Subagents: .cursor/agents/*.md

Skills: .cursor/skills/<skill-name>/SKILL.md

OpenAI Codex (Repo)

Skills: <repo>/.codex/skills/<skill-name>/SKILL.md

(권장) .codex/skills/*는 .cursor/skills/*로 symlink → 단일 SSOT 운영

symlink 예시(macOS/Linux)
mkdir -p .codex/skills
ln -s ../../.cursor/skills/convert-toolbox .codex/skills/convert-toolbox
ln -s ../../.cursor/skills/mrconvert-run .codex/skills/mrconvert-run
ln -s ../../.cursor/skills/email-thread-search .codex/skills/email-thread-search
ln -s ../../.cursor/skills/cipl-excel-build .codex/skills/cipl-excel-build

Windows 대안

symlink 권한/정책 이슈가 있으면 그냥 폴더 복사로 운영(주기적으로 동기화).

검증 체크리스트 + validator 실행 예시
최소 검증(권장)

 .cursor/agents/*.md에 YAML frontmatter 존재

 .cursor/skills/<name>/SKILL.md의 name:이 폴더명과 동일

 skill/subagent 이름이 ^[a-z0-9]+(-[a-z0-9]+)*$ 규칙 준수

 스킬이 1개 책임(단일 목적)으로 분리됨(“만능 스킬” 금지)

실행 예시
python .cursor/skills/convert-toolbox/scripts/validate_agent_assets.py --root .
python .cursor/skills/convert-toolbox/scripts/run_smoke.py --root .

🔧/cmd3 (Now/Next/Alt)

Now: /convert-scoper (CONVERT 인벤토리/엔트리포인트/스모크 커맨드 후보 생성)

Next: /verifier (스모크/pytest 기반 PASS/FAIL 증거 리포트)

Alt: /excel-style-guardian (CIPL/간트/Excel 산출물 서식 회귀 체크)

ZERO log

본 건은 스킬/서브에이전트 “문서 패키지” 작성이며, UAE 규정/통관/요율/ETA·날씨 등 실시간 근거 필수 영역이 아님 → ZERO 게이트 비적용.