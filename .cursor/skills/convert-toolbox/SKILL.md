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

## 고정 스모크 후보(AGENTS.md §5.4 참조)
| # | 용도 | 커맨드 |
|---|------|--------|
| 1 | 공통 구문 검증 | `python -m compileall -q .` |
| 2 | pytest(조건부) | `pytest -q` |
| 7 | 패키지 정합성 | `python .cursor/skills/convert-toolbox/scripts/validate_agent_assets.py --root .` |
| 8 | 인벤토리 생성 | `python .cursor/skills/convert-toolbox/scripts/convert_inventory.py --root . --out out/convert_inventory.json` |

## 리포트 포맷(권장)
- Evidence Table: | Check | Result | Command | Notes |
- FAIL이면: 원인 1줄 + 최소 수정안 + 재실행 커맨드
