# agent.md — **mrconvert_v1 · Document Conversion & Ontology Framework**

*v1.0 · 2025-11-07 · Asia/Dubai*

> Python-First · TDD + Tidy First v3.8 · Email→Ontology 파이프라인 · RDF Graph(TTL/JSON-LD/XML) · CI/보안 게이트 · 2× GitHub Cross-Check · 일일 자동화

---

## 0) 메타데이터

```yaml
name: mrconvert_v1
purpose: 문서(WhatsApp/Excel/Markdown/이미지) 변환 및 물류 온톨로지 통합
owner: Samsung C&T · HVDC Logistics (MACHO-GPT)
timezone: Asia/Dubai
language: ko-KR (tech docs EN 허용)
runtime:
  python: ">=3.13"
  os: [Windows, Linux, macOS]
entrypoints:
  cli: mrconvert
  packages: [src/mrconvert]
standards:
  test: pytest -q && pytest -q --cov=src --cov-report=term
  lint_format: ruff check && ruff format   # (또는 black 병행)
  typing: basedpyright(strict) | mypy --strict (선택)
  coverage: line>=0.85 (core path>=0.90)
  security: bandit -q, pip-audit/safety, gitleaks
  sbom: cyclonedx-py
build_lock: uv | pip-tools
docs_root: docs/
```

> ⚠️ `src/mrconvert.email`, `src/mrconvert.rdf`, RDF/SHACL 파이프라인 등은 아직 구현되지 않았습니다. 아래 로드맵 섹션을 참고하세요.

---

## 1) 목적 & 스코프

* **목적**: WhatsApp·Excel(이메일 데이터 포함)·Markdown·스캔이미지를 **기계 읽기 가능한 형식**(TXT/MD/JSON/CSV)으로 변환하고, **Email→Ontology** 파이프라인으로 **의도/프로세스/문서 참조**를 추출·연결하여 **RDF Graph**로 영속화한다.
* **스코프(In)**: 변환(Extract) · 온톨로지 강화(Enrich) · 문서링킹(Link) · 자동 라우팅(Route) · RDF 변환(Graph) · 일일 배치 자동화.
* **비스코프(Out)**: 외부 DMS/Foundry 배포, 사내 비밀키 운영, 대규모 데이터 마이그레이션(별도 런북).

### 1.1 운영 원칙 (TDD + Tidy First v3.8)

1. **RED→GREEN→REFACTOR** 루프 불변, 단일 테스트 SLA ≤200ms.
2. **Tidy First**: 구조 커밋(행위 불변) → 이후 행위 변경.
3. **Python-First 게이트**: ruff 0 · format OK · pytest 통과 · 커버리지 기준.
4. **Deterministic Build**: uv/pip-tools 잠금, CI=로컬 동형.
5. **2× Cross-Check**: 릴리즈/보고 전 최신 2레포 비교·근거 첨부.

---

## 2) 상태 요약 (Current vs Planned)

| 영역 | 상태 | 설명 |
| --- | --- | --- |
| PDF → TXT/MD/JSON 추출, 표 CSV 저장 | ✅ Running | `pdfplumber`, `rich` 기반 CLI 동작, OCR(auto/force) 선택 가능 |
| DOCX → 텍스트/마크다운 변환 | ✅ Running | `mammoth`, `python-docx` 이용 |
| PDF ↔ DOCX, DOCX → PDF 변환 | ✅ Running | `pdf2docx`, `docx2pdf`(soffice 폴백) |
| Markdown → DOCX/XLSX, DOCX/MD → MSG | ✅ Running | CLI `--to-docx/--to-xlsx/--to-msg` 지원 |
| WhatsApp/Excel/이미지 특화 파이프라인 | ⏳ Planned | 구조 설계만 명시, 구현 미착수 |
| Email→Ontology 추출/라우팅 | ⏳ Planned | JSON-LD 메타 파서를 제외한 분석·링커·라우터 미구현 |
| RDF(Graph)/SHACL 검증 | ⏳ Planned | `rdflib` 모듈 미구현, 그래프 스토어 없음 |
| 일일 배치 & 관측/알림 | ⏳ Planned | `scripts/daily_email_processing.py` 등 스크립트 부재 |
| 품질 게이트(ruff/bandit/pip-audit/SBOM) | ⏳ Planned | 테스트는 스모크 수준, CI·보안 게이트 미구축 |
| 2× GitHub Cross-Check 프로세스 | ⏳ Planned | 문서 템플릿만 존재, 자동화 없음 |

---

## 3) 시스템 개요(요약)

* **아키텍처 문서**:

  * `docs/m.md`(ASCII + Mermaid 5계층) — 초안
  * `docs/HVDC_System_Architecture.md`(Protégé 통합 풀스택 MVP) — 계획
* **온톨로지**: `docs/ontology/README.md`(구조만 정의, 실제 파이프라인 미연결)
* **가이드**: 프로젝트 인덱스/설정/사용자/시스템 아키텍처 가이드 일체
* **Email 파이프라인 세트**: 사용자/실전/구현/아키텍처/개발자참조/Graph Store 가이드 (※ 구현 예정 기능 기반)

---

## 4) 핵심 기능

### 4.1 현재 제공 기능

* PDF → TXT/MD/JSON 추출 (`--format`, `--tables`, `--ocr`, `--lang`)
* DOCX → 텍스트/마크다운 추출 (Markdown/표 포함)
* PDF ↔ DOCX, DOCX → PDF, DOCX/MD → MSG 변환
* Markdown → DOCX/XLSX 변환 (JSON-LD 메타 포함 시 시트 구성)

### 4.2 향후 계획 기능 (미구현)

* WhatsApp/Excel 이메일 파싱 및 의도/프로세스 추출
* Email→Ontology Multi-Key Linker 및 자동 라우팅
* RDF(Graph Store), SHACL 검증 및 규제 체크
* 일일 자동화 배치(`scripts/daily_email_processing.py`) 및 KPI 관측
* CI/보안/품질 게이트 고도화 (ruff/bandit/pip-audit/cyclonedx)

---

## 5) 디렉터리 구조

```
mrconvert_v1/
├── src/mrconvert/            # 현재 CLI/변환 구현
├── tests/                    # smoke + CLI 호환 테스트
├── converted_pdfs/           # 샘플 변환 산출물
├── ONTOLOGY/, ontology_*     # 온톨로지 참고 문서(정적)
├── a.md, README.md, ...      # 스펙/문서
├── docs/ (Planned)           # 아키텍처/가이드 정리 예정
├── data/, output/, archive/ (Planned)    # 파이프라인 데이터 루트
└── scripts/ (Planned)        # email_pipeline.py, daily_email_processing.py 등
```

---

## 6) 설치

```bash
# OCR 포함
pip install -e ".[ocr]"
# OCR 제외
pip install -e .
# RDF 옵션
pip install rdflib rdflib-jsonld
```

> Windows .doc → .docx: `soffice --headless --convert-to docx file.doc` (LibreOffice)

---

## 7) CLI 명령(요약)

### 7.1 Text Extraction (실행 중)

```bash
mrconvert INPUT_PATH --out OUT_DIR --format txt md json --tables csv --ocr auto --lang kor+eng
```

### 7.2 Bidirectional (실행 중)

```bash
mrconvert INPUT_PATH --to-docx    # PDF → DOCX
mrconvert INPUT_PATH --to-pdf     # DOCX → PDF
```

### 7.3 Email Conversion (Planned)

```bash
# TODO: CLI 확장 후 아래 옵션 공개 예정
# mrconvert email.xlsx --email-to-json --out output/excel --enrich-ontology --link --route
# mrconvert email.xlsx --email-to-json --to-rdf --rdf-format ttl --out output/excel
```

> 위 명령은 아직 CLI에 포함되어 있지 않습니다. 설계 로드맵 참고.

### 7.4 배치 & 일일 자동화 (Planned)

```bash
# python scripts/email_pipeline.py email.xlsx output/excel --all
# python scripts/daily_email_processing.py --full --file email.xlsx --graph-store output/excel/graph_store.ttl --report
```

> 배치 스크립트는 아직 제공되지 않습니다. 구현 시 상기 명령으로 실행 예정.

---

## 8) 출력 스키마 (요약)

### 8.1 Text JSON (실행 중)

```json
{
  "meta": {
    "source": "<path>", "type": "pdf|docx|image",
    "pages": 10, "parsed_at": "YYYY-MM-DDTHH:MM:SSZ",
    "ocr": {"used": true, "engine": "ocrmypdf|pytesseract|none", "lang": "kor+eng"}
  },
  "text": "...", "markdown": "...",
  "tables": [{"page":1,"index":0,"rows":[["A","B"],["1","2"]]}]
}
```

### 8.2 Email JSON(Planned)

```json
{
  "metadata": {
    "source_file":"email.xlsx",
    "total_emails":1999,
    "date_range":{"start":"2024-10-01T00:00:00+04:00","end":"2025-10-29T23:59:59+04:00"}
  },
  "emails_by_month": {"2024-10":[{ "date":"...", "intent":"request", "logistics_process":"Invoice", "project_tag":"HVDC-001", "document_refs":["Invoice:INV-2025-001","BL:ABC123"] }]},
  "statistics":{"total_emails":1999,"with_intent":500,"with_process":600,"with_project_tag":300,"with_document_refs":406}
}
```

> Email JSON/RDF 스키마는 문서화 단계이며, 실제 변환 로직은 아직 구현되지 않았습니다.

---

## 9) 테스트 & 품질 게이트

* **현재**: `pytest -q` 스모크 테스트 및 CLI 플래그 유효성 위주.
* **향후 목표**:
  * **TDD**: 유닛≫통합≫E2E(핵심 여정) 피라미드.
  * **커버리지**: 라인 ≥85%(핵심경로 90%).
  * **정적 분석**: `ruff check`=0, `ruff format`/`black` OK.
  * **타입(선택)**: `basedpyright --level strict` | `mypy --strict` 핵심 경로.
  * **성능 예산**: 변환 1문서(10p, 텍스트형) ≤1.5s, OCR 페이지 ≤2.5s(가이드).
  * **CI**: GH Actions(ubuntu-latest, windows-latest, py3.13), 캐시·병렬 테스트, SBOM/보안 스캔 단계 포함.

---

## 10) 보안·SBOM·컴플라이언스

* **현재**: 정적 보안 도구와 SBOM 생성 파이프라인 미구현.
* **향후 계획**:
  * `pip-audit`/`safety` 취약점 스캔, `gitleaks` 시크릿 검사
  * CycloneDX SBOM 생성, 배포 파이프라인 연동
  * Incoterms 2020, MOIAT/FANR 룰셋 + SHACL 검증
  * 변환 메타/해시/규칙 위반 JSONL 감사 로그 적재

---

## 11) 관측 가능성(Observability)

* **현재**: `rich` 프로그레스바 외 별도 메트릭·로그 미수집.
* **향후 계획**:
  * 처리속도(ms/op), OCR 사용률, 테이블 검출 성공률, 라우팅 카테고리 분포 수집
  * 일일 자동화 콘솔 프로그레스바 + `reports/` 요약 MD/CSV 생성
  * 실패율 상승/ SLA 초과 시 이메일·메신저 웹훅 알림

---

## 12) 한계·권고

* **스캔 품질 의존**: OCR 정확도는 해상도/언어팩/전처리에 좌우.
* **헤더/열 별칭**: 미규격 Excel은 매핑 테이블 보강 권장.
* **온톨로지**: SHACL/룰셋은 점진 확장. 미정의 개념은 `comm:Unknown`으로 격리.

---

## 13) 2× GitHub Cross-Check Gate (필수)

릴리즈/보고 전 아래 표를 채워 **최신 2개 이상 레포** 구현/인터페이스/엣지케이스를 비교한다.  
현재는 템플릿만 제공되며, 자동 수집 스크립트/프로세스는 구축되지 않았다.
(예: PDF 추출(PyMuPDF/pdfplumber/Camelot), Excel 파싱(openpyxl/pandas), RDFLib 활용 등)

| 항목            | Repo A | Repo B | 관찰 메모(서명/테스트/엣지/라이선스) | 채택/보류 사유 |
| ------------- | ------ | ------ | --------------------- | -------- |
| PDF 텍스트/표 추출  |        |        |                       |          |
| Excel→JSON 파싱 |        |        |                       |          |
| RDF Graph 변환  |        |        |                       |          |
| SHACL 검증 예시   |        |        |                       |          |

> 제출물에 링크·비교요약 첨부. 활동성(최근 커밋/이슈) 확인.

---

## 14) 수용 기준(Acceptance) & 준비 완료 체크리스트

### 현재 충족 항목

- [x] CLI: 텍스트 추출(`--format/--tables/--ocr`) 정상 동작
- [x] PDF ↔ DOCX, DOCX → PDF/MSG 변환
- [x] Markdown → DOCX/XLSX 변환

### 향후 달성 목표

- [ ] CLI: Email→Ontology/RDF/배치 시나리오(7.3~7.4) 구현 및 테스트
- [ ] Email JSON 스키마/통계 필드 일치, Asia/Dubai 타임존 정규화
- [ ] RDF(TTL/JSON-LD/RDF-XML) 동형성 테스트 통과
- [ ] SHACL 규칙 최소 1세트 적용 및 위반 리포트 생성
- [ ] 라우팅 규칙(COST-GUARD/LATTICE/PRIME/ORACLE) 분기 테스트
- [ ] 품질 게이트/보안/ SBOM/빌드 잠금 통과
- [ ] 2× Cross-Check 리포트 첨부

**권장 커밋 메시지**

```
docs(agent): add mrconvert_v1 agent spec (Email→Ontology, RDF, daily automation)
- CLI/outputs/schema defined, TDD+Tidy gates, CI/security/SBOM
- routing rules + SHACL checks + 2× GitHub cross-check gate
```

---

## 15) 부록 — CLI 예시 모음

```bash
# 1) PDF → MD+JSON (표는 CSV)
mrconvert sample.pdf --out out --format md json --tables csv

# 2) 폴더 일괄 + OCR 자동
mrconvert ./incoming --out ./out --format txt --ocr auto --lang kor+eng

# 3) 이미지 스캔 OCR
mrconvert invoice.png --out out --format txt md --ocr auto

# 4) PDF↔DOCX
mrconvert doc.pdf --to-docx --out ./converted
mrconvert doc.docx --to-pdf --out ./converted

# 5) Email 풀 파이프라인
# (Planned) CLI 옵션 확장 후 사용 가능
# mrconvert email.xlsx --email-to-json --email-to-md --enrich-ontology --to-rdf --link --route --out output/excel

# 6) 일일 자동화(권장)
# (Planned) 배치 스크립트 도입 시 사용
# python scripts/daily_email_processing.py --full --file email.xlsx --graph-store output/excel/graph_store.ttl --report
```

---

## 16) 연결 문서(필독)

* 프로젝트 인덱스: `docs/00_PROJECT_INDEX.md`
* 시스템 아키텍처: `docs/guides/SYSTEM_ARCHITECTURE_FINAL.md`
* 설정/사용자 가이드: `docs/guides/CONFIGURATION_GUIDE.md`, `docs/guides/USER_GUIDE.md`
* Email 파이프라인: `docs/guides/EMAIL_PIPELINE_*` 전 세트
* Graph Store 활용: `docs/guides/GRAPH_STORE_UTILITY_GUIDE.md`
* 개선 내역: `docs/guides/EMAIL_CONVERSION_IMPROVEMENTS.md`

---

## 17) 로드맵 & 우선순위

1. **CLI 확장**: `mrconvert email …` 서브커맨드, RDF/SHACL 옵션, 플러그인 아키텍처 도입.
2. **Email→Ontology 파이프라인**: JSON-LD 추출 + Intent/Process/DocRef 라우터 + Graph Store 업데이트.
3. **자동화 & Observability**: `scripts/email_pipeline.py`, KPI 로깅, Slack/Email 알림 연동.
4. **품질/보안 게이트**: `ruff`, `pytest --cov`, `bandit`, `pip-audit`, `cyclonedx-bom` CI 파이프라인 구축.
5. **Cross-Check & 문서화 자동화**: 외부 레포 비교 스크립트, `docs/` 갱신 자동 리포트, acceptance 체크 자동화.

> 이 문서는 **실행 사양**이다. 변경 시 테스트·CI·문서 동기화를 함께 수행한다.
