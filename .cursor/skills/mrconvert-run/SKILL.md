---
name: mrconvert-run
description: mrconvert_v1에서 PDF/DOCX/XLSX를 TXT/MD/JSON으로 변환하는 실행 루틴을 표준화한다. "mrconvert", "convert pdf", "OCR", "table extract" 요청에 사용.
---

# mrconvert-run

## 언제 사용
- mrconvert_v1 변환 파이프라인 실행/수정/디버그
- 출력 폴더(out/output) 규칙을 고정하고, 레거시 동작을 깨지 않게 확장
- WhatsApp/Excel/Markdown/이미지를 기계 읽기 가능한 형식으로 변환
- Email→Ontology 파이프라인 작업 (향후 계획)

## 프로젝트 메타데이터

**참조**: `mrconvert_v1/agent_mrconvert.md` (상세 사양)

- **Python 버전**: >=3.11
- **엔트리포인트**: `mrconvert` (CLI)
- **패키지 경로**: `src/mrconvert`
- **표준**: TDD + Tidy First v3.8, ruff/black, pytest, 커버리지 ≥85%

## 현재 제공 기능 (Running)

### 텍스트 추출
- PDF → TXT/MD/JSON 추출 (`--format`, `--tables`, `--ocr`, `--lang`)
- DOCX → 텍스트/마크다운 변환 (Markdown/표 포함)
- 이미지 OCR (자동/강제 선택 가능)

### 양방향 변환
- PDF ↔ DOCX 변환
- DOCX → PDF 변환 (soffice 폴백 지원)
- DOCX/MD → MSG 변환

### Markdown 변환
- Markdown → DOCX/XLSX 변환 (JSON-LD 메타 포함 시 시트 구성)

## 향후 계획 기능 (Planned)

- WhatsApp/Excel 이메일 파싱 및 의도/프로세스 추출
- Email→Ontology Multi-Key Linker 및 자동 라우팅
- RDF(Graph Store), SHACL 검증 및 규제 체크
- 일일 자동화 배치 및 KPI 관측

> ⚠️ Email→Ontology, RDF 파이프라인은 아직 구현되지 않았습니다. 현재는 텍스트 추출 및 기본 변환만 지원합니다.

## 입력 카드(가능하면 확보)
- Input: 파일 경로(로컬), 타입(PDF/DOCX/XLSX/이미지), 목표 출력(TXT/MD/JSON/CSV), OCR 필요 여부
- Output: 저장 경로(out/ 또는 output/), 파일명 규칙
- Constraints: 네트워크 사용 금지/허용, 대용량 제한
- 언어: `--lang kor+eng` (한국어+영어 OCR)

## 절차(보수적)

### 1) 엔트리포인트 확인
- mrconvert_v1 폴더에서 README 또는 `mrconvert --help` 먼저 확인
- "추측 실행" 금지
- 설치 상태 확인: `pip list | grep mrconvert` 또는 `mrconvert --version`

### 2) 설치 (필요 시)
```bash
# OCR 포함
pip install -e ".[ocr]"

# OCR 제외
pip install -e .

# RDF 옵션 (향후)
pip install rdflib rdflib-jsonld
```

> Windows .doc → .docx: `soffice --headless --convert-to docx file.doc` (LibreOffice 필요)

> **JPG/이미지 OCR**: Tesseract 사용 시 `eng.traineddata` 필요. CONVERT 프로젝트에서는 `out/tessdata/eng.traineddata`를 두고, `pdf_converter`가 `TESSDATA_PREFIX`를 `CONVERT/out/tessdata`로 자동 설정(미설정 시). 날씨 폴더 PDF+JPG 파싱 확인: `python scripts/weather_parse.py <weather/YYYYMMDD>`.

### 3) Dry-run 성격의 최소 실행
- `mrconvert --help` 확인
- 샘플 1건 변환(가능하면 익명 샘플)
- 출력 형식 확인

### 4) 출력 규칙
- 기본: `out/` 또는 `output/` 하위에 생성
- 변환 결과는 Git 추적 제외 권장(.gitignore)
- 파일명 규칙: 원본명_형식.확장자 (예: `sample_txt.txt`, `sample_md.md`)

### 5) 검증
- 변환 결과 존재 여부 + 파일 크기 0 여부
- (테이블 추출이면) JSON schema 키 최소 확인(없으면 가정/중단)
- OCR 사용 시 정확도 확인

## CLI 명령어 예시

### 텍스트 추출 (기본)
```bash
# PDF → MD+JSON (표는 CSV)
mrconvert sample.pdf --out out --format md json --tables csv

# 폴더 일괄 + OCR 자동
mrconvert ./incoming --out ./out --format txt --ocr auto --lang kor+eng

# 이미지 스캔 OCR
mrconvert invoice.png --out out --format txt md --ocr auto
```

### 양방향 변환
```bash
# PDF → DOCX
mrconvert doc.pdf --to-docx --out ./converted

# DOCX → PDF
mrconvert doc.docx --to-pdf --out ./converted
```

### Markdown 변환
```bash
# Markdown → DOCX/XLSX
mrconvert doc.md --to-docx --out ./converted
mrconvert doc.md --to-xlsx --out ./converted
```

### 향후 계획 (아직 미구현)
```bash
# Email 풀 파이프라인 (Planned)
# mrconvert email.xlsx --email-to-json --enrich-ontology --to-rdf --link --route --out output/excel

# 일일 자동화 (Planned)
# python scripts/daily_email_processing.py --full --file email.xlsx --graph-store output/excel/graph_store.ttl --report
```

## 출력 스키마

### Text JSON (현재 지원)
```json
{
  "meta": {
    "source": "<path>",
    "type": "pdf|docx|image",
    "pages": 10,
    "parsed_at": "YYYY-MM-DDTHH:MM:SSZ",
    "ocr": {
      "used": true,
      "engine": "ocrmypdf|pytesseract|none",
      "lang": "kor+eng"
    }
  },
  "text": "...",
  "markdown": "...",
  "tables": [
    {
      "page": 1,
      "index": 0,
      "rows": [["A", "B"], ["1", "2"]]
    }
  ]
}
```

### Email JSON (향후 계획)
```json
{
  "metadata": {
    "source_file": "email.xlsx",
    "total_emails": 1999,
    "date_range": {
      "start": "2024-10-01T00:00:00+04:00",
      "end": "2025-10-29T23:59:59+04:00"
    }
  },
  "emails_by_month": {
    "2024-10": [{
      "date": "...",
      "intent": "request",
      "logistics_process": "Invoice",
      "project_tag": "HVDC-001",
      "document_refs": ["Invoice:INV-2025-001", "BL:ABC123"]
    }]
  },
  "statistics": {
    "total_emails": 1999,
    "with_intent": 500,
    "with_process": 600,
    "with_project_tag": 300,
    "with_document_refs": 406
  }
}
```

## 디렉터리 구조

```
mrconvert_v1/
├── src/mrconvert/            # 현재 CLI/변환 구현
├── tests/                    # smoke + CLI 호환 테스트
├── converted_pdfs/           # 샘플 변환 산출물
├── ONTOLOGY/, ontology_*     # 온톨로지 참고 문서(정적)
├── docs/                     # 아키텍처/가이드
└── scripts/ (Planned)        # email_pipeline.py, daily_email_processing.py 등
```

## 테스트 & 품질 게이트

### 현재
- `pytest -q` 스모크 테스트 및 CLI 플래그 유효성 위주

### 향후 목표
- **TDD**: 유닛≫통합≫E2E(핵심 여정) 피라미드
- **커버리지**: 라인 ≥85% (핵심경로 90%)
- **정적 분석**: `ruff check`=0, `ruff format`/`black` OK
- **타입(선택)**: `basedpyright --level strict` | `mypy --strict`
- **성능 예산**: 변환 1문서(10p, 텍스트형) ≤1.5s, OCR 페이지 ≤2.5s

## Ask first
- OCR 엔진/대형 의존성 설치
- 대량 변환(폴더 전체) 실행
- 운영 문서(PII 포함)로 재현
- RDF/온톨로지 관련 기능 사용 (아직 미구현)

## 산출물
- 실행 커맨드(확정본)
- "입력→출력" 매핑 표 1개
- 변환 결과 파일 경로 및 메타데이터
- 실패 시: 원인 1줄 + 최소 수정안 + 재시도 커맨드

## 리포트 포맷(권장)
- Evidence Table: | 입력 파일 | 출력 형식 | OCR 사용 | 결과 | 경로 |
- 실패 시: 원인 1줄 + 최소 수정안 + 재시도 커맨드

## 통합
- `/convert-scoper`: mrconvert_v1 구조 파악
- `/verifier`: 변환 결과 검증
- `convert-toolbox`: 스모크 테스트로 변환 후 검증

## 참고 문서
- 상세 사양: `mrconvert_v1/agent_mrconvert.md`
- 프로젝트 인덱스: `mrconvert_v1/docs/00_PROJECT_INDEX.md` (Planned)
- 시스템 아키텍처: `mrconvert_v1/docs/guides/SYSTEM_ARCHITECTURE_FINAL.md` (Planned)
