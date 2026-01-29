# 콘텐츠 스키마 (content.json)

필드 기반 공문서 콘텐츠 예시:

| 필드 | 설명 |
|------|------|
| DOC_NO | 문서번호 |
| DATE | 일자 (ISO 8601 권장: YYYY-MM-DD) |
| RECIPIENT | 수신 |
| SUBJECT | 제목 |
| BODY_PARAGRAPHS | 본문 (문자열 또는 문자열 배열) |
| ATTACHMENTS | 첨부 |
| SIGNER_TITLE | 서명자 직함 |
| SIGNER_NAME | 서명자 성명 |
| ORG_NAME | 기관명 |
| ORG_ADDRESS | 기관 주소 |
| ORG_CONTACT | 기관 연락처 |

템플릿의 `{{PLACEHOLDER}}` 와 키 이름을 맞추면 `fill_template.py` 가 치환한다.
