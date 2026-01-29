---
name: official-doc-drafter
description: |
  사용자 제공 DOCX/PDF 템플릿의 글자체·크기·단락·여백·A4 레이아웃을 최대한 유지하며 공문서(.docx)를 작성/생성한다. 템플릿이 없으면 ISO 216(A4)·ISO 8601(날짜) 및 통용 양식/커뮤니티 템플릿을 근거로 기본 양식을 구성한다. 생성 후 docx-style-verifier로 템플릿 일치도를 정량 검증한다.
compatibility:
  cursor: true
allowed-tools:
  - python
  - web
---

# official-doc-drafter

## 단일 파이프라인 (작성 → 검증)

**한 번의 트리거로 “작성 + 검증”이 끝나도록 구성한다.**

| 순서 | 담당 | 동작 |
|------|------|------|
| 1 | **/docstyle-researcher** | Discovery → content.json → 템플릿 채움 → 로고(선택) → `.docx` 저장 |
| 2 | **같은 런 내** | `compare_docx_style.py` 실행 (TEMPLATE_DOCX vs OUTPUT_DOCX) → `out/style-report.md`·`out/style-report.json` 생성 |
| 3 | **결과** | 생성 `.docx` + Verdict(PASS/WARN/FAIL) + 이슈 목록(있을 경우) |

- 사용자: “공문서 작성해줘” → **/docstyle-researcher** 한 번만 호출. 에이전트가 작성 후 반드시 검증 단계를 수행하고 Verdict를 보고한다.
- 검증만 따로 할 때: “이 문서 템플릿이랑 똑같아?” → **/docx-style-verifier** 단독 호출.

## 작업 폴더 (필수)

**모든 작업은 `OFFICIAL DOCS/` 폴더에서만 실행한다.**

- 스크립트 실행 시 **작업 디렉터리(cwd)를 `OFFICIAL DOCS/`** 로 두고 실행한다.
- content.json, 템플릿(.docx/.pdf), 로고, 산출 .docx는 **`OFFICIAL DOCS/` 또는 그 하위(`out/`, `output/`)** 에만 읽고 쓴다.
- **`OFFICIAL DOCS/` 밖**의 파일을 공문서 작업용으로 읽거나 쓰지 않는다.

## 언제 사용하나
- 사용자가 "이 포맷 그대로 공문서 작성해줘(Word/PDF)"를 요청할 때
- PDF 샘플만 있고 동일한 스타일의 Word가 필요할 때
- 템플릿이 없어서 "국제/관행/커뮤니티 양식 기반"으로 제작해야 할 때

## 핵심 원칙
- 템플릿 제공 시: **템플릿 우선** (폰트/여백/단락/헤더/서식 최대 보존)
- 템플릿 미제공 시: `references/template-sources.md`의 소스 우선순위로 템플릿을 채택/구성
- 작성 완료 후: **docx-style-verifier 또는 compare 스크립트로 정량 검증** (단순 육안 주장 금지)

## 템플릿이 없는 경우의 소스 우선순위(요약)
1) 관할/기관의 공식 서식(법령/정부 포털/공공기관 배포)
2) 벤더 제공 템플릿(예: Microsoft Create)
3) 신뢰 가능한 무료 템플릿 라이브러리
4) 커뮤니티/마켓(유료/회원제 포함) — 라이선스 확인 필수

상세 링크는 `references/template-sources.md` 참조.

## 검증(Verifier) 워크플로
### 빠른 검증(권장)
- `/docx-style-verifier`를 호출하여 템플릿 DOCX vs 결과 DOCX의
  - A4 페이지 크기
  - 여백
  - 지배적 폰트/폰트 크기 분포
  - 단락 정렬/줄간격/문단 간격
  - 헤더/푸터, placeholder 잔존 여부
  를 PASS/WARN/FAIL로 판정하게 한다.

### 수동 실행(스크립트)
- 프로파일:
  - `python scripts/docx_style_profile.py --docx TEMPLATE.docx --out-json template.profile.json`
  - `python scripts/docx_style_profile.py --docx OUTPUT.docx --out-json output.profile.json`
- 비교 리포트:
  - `python scripts/compare_docx_style.py --template TEMPLATE.docx --candidate OUTPUT.docx --out-md style-report.md --out-json style-report.json`

검증 기준(권장 허용 오차)은 `references/docx-style-verification.md` 참고.

## 로고/레터헤드
- 로고는 "공식 출처"를 우선으로 사용하고, 가능하면 SVG(벡터)로 보관 후 삽입한다.
- 공식 SVG가 없고 래스터만 존재하는 경우 벡터화는 품질 저하가 발생할 수 있으므로, verifier 결과에 WARN으로 기록한다.

## 템플릿 채우기: fill_template vs docxtpl

| 스크립트 | 문법 | 적합한 경우 |
|----------|------|-------------|
| `fill_template.py` | `{{KEY}}` 단순 치환 | 단순 플레이스홀더(문서번호, 날짜, 수신자 등) |
| `fill_template_docxtpl.py` | Jinja2 `{{ var }}`, `{% for %}`, `{% if %}` | 반복 블록·조건부 문단·복잡 템플릿 |

- docxtpl 사용 시 템플릿(.docx)에 `{{ company_name }}` 등 Jinja2 변수 삽입. context는 JSON으로 전달.
- 상세: `references/docxtpl-usage.md`. 의존성: `docxtpl` (requirements.txt에 포함).

## Skill Map (내부 스크립트)

| 이름 | 1줄 요약 | Trigger |
|------|----------|---------|
| `official-doc-drafter` | 템플릿(DOCX/PDF) 기반으로 스타일을 최대한 동일하게 유지하며 공문서 DOCX 생성 | 공문서, 템플릿 복제, PDF 스타일, 레터헤드 |
| `fill_template.py` | `{{KEY}}` 단순 치환으로 템플릿 채움 | content.json 기반 단순 치환 |
| `fill_template_docxtpl.py` | docxtpl(Jinja2)로 템플릿 렌더링 | `{% for %}`, `{% if %}` 등 필요 시 |
| `pdf_style_extract` (내부) | PDF에서 글자체/여백 추출 | PDF 스타일 파서, PDF 클론 |
| `logo_prepare` (내부) | 공식 로고 다운로드 및 SVG/PNG 준비 | 로고 검색, SVG 변환 |

## 관련 국제 기준
- **ISO 216 (A4 종이 규격)**: 210×297 mm. 공문서 기본 용지 규격.
- **ISO 8601 (날짜 표기)**: YYYY-MM-DD 형식. 공문서 표준 날짜 표현.

## 입력 파라미터

| 파라미터 | 설명 |
|----------|------|
| `--template` | `.docx` 또는 `.pdf` 사용자 샘플 |
| `--content-json` | 필드 기반 콘텐츠 JSON (문서번호, 수신자, 본문 등) |
| `--logo-url` | (선택) 공식 로고 URL |
| `--output` | 결과 `.docx` 경로 |

## Subagent 연동 (파이프라인)
- **/docstyle-researcher**: 파이프라인 1단계. 공문서 작성 후 **같은 런에서** 2단계 검증(compare_docx_style.py)을 반드시 실행. cwd = `OFFICIAL DOCS/`. 최종 산출: `.docx` + `out/style-report.md`(.json) + Verdict.
- **/docx-style-verifier**: 파이프라인 2단계(자동) 또는 단독. "검증해", "템플릿과 똑같아?" 트리거 시 단독 실행.

## 파일 위치
- Python 스크립트: `scripts/`
- 참조문서: `references/` (template-sources.md, docx-style-verification.md 포함)
- 템플릿/스키마: `assets/`
