# official-doc-drafter 워크플로우

## 1. Word 템플릿 제공 시

1. `make_default_template.py` 또는 사용자 `.docx` 템플릿 준비
2. `content.json` 작성 (DOC_NO, DATE, RECIPIENT, SUBJECT, BODY_PARAGRAPHS 등)
3. `fill_template.py --template T.docx --content-json content.json --out result.docx`
4. (선택) 로고: `logo_prepare.py` → `insert_logo.py`
5. (선택) `validate_output.py` 로 스타일 검증

## 2. PDF 템플릿만 제공 시

1. `pdf_style_extract.py --pdf sample.pdf --out-json style.json`
2. `docx_from_pdf_style.py --style-json style.json --content-json content.json --out base.docx`
3. 필요 시 `fill_template.py` 로 추가 치환 또는 `insert_logo.py` 로 로고 삽입
4. `validate_output.py --docx base.docx --expected-style style.json`

## 3. 템플릿 미제공 시

1. `make_default_template.py -o base_template.docx --org-name "기관명"`
2. `content.json` 작성
3. `fill_template.py -t base_template.docx -c content.json -o result.docx`
4. A4(ISO 216), 날짜 ISO 8601(YYYY-MM-DD) 권장

## 4. 로고 삽입

1. 공식 URL 확인 (Subagent `/docstyle-researcher` 로 조사 가능)
2. `logo_prepare.py -u https://example.org/logo.svg -d assets/logos`
3. `insert_logo.py -d result.docx -l assets/logos/logo.png -o final.docx -w 30`
