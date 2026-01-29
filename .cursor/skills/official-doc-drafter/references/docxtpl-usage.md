# python-docx-template (docxtpl) 사용 가이드

> **용도**: Word(.docx)에 **Jinja2 문법**을 넣어 변수·반복·조건으로 문서 생성.
> **패키지**: `pip install docxtpl` (의존: python-docx, Jinja2)

## fill_template.py vs docxtpl

| 항목 | fill_template.py (기본) | docxtpl |
|------|-------------------------|---------|
| 문법 | `{{KEY}}` 단순 치환 | Jinja2: `{{ var }}`, `{% for %}`, `{% if %}` |
| 범위 | 한 run 안에서만 치환 | 문단/테이블 행·열·run 단위 태그 지원 |
| 적합 | 단순 플레이스홀더 문서 | 반복 블록·조건부 블록·복잡 템플릿 |

템플릿에 **반복(목록)** 이나 **조건부 문단** 이 필요하면 docxtpl 사용을 검토한다.

## 기본 사용

```python
from docxtpl import DocxTemplate

doc = DocxTemplate("template.docx")
context = {
    "company_name": "Mammoet Malaysia Sdn Bhd",
    "date": "28 January 2026",
    "ref": "MMMY/HVDC/LS/CoC/Rev00",
}
doc.render(context)
doc.save("output.docx")
```

템플릿(.docx) 안에 `{{ company_name }}`, `{{ date }}`, `{{ ref }}` 를 넣어 두면 치환된다.

## Jinja2 문법 (템플릿 내)

- 변수: `{{ variable }}`
- 조건: `{% if condition %} ... {% endif %}`
- 반복: `{% for item in list %} ... {% endfor %}`

## docxtpl 전용 태그 (여러 문단/행에 걸칠 때)

- `{%p ... %}`: 문단 단위
- `{%tr ... %}`: 테이블 행 단위
- `{%tc ... %}`: 테이블 열 단위
- `{%r ... %}`: run 단위

표준 Jinja2 태그는 **한 run 안**에서만 동작한다. 문단/테이블을 넘나들면 위 태그 사용.

## 참고

- 문서: https://docxtpl.readthedocs.io/
- PyPI: https://pypi.org/project/docxtpl/
