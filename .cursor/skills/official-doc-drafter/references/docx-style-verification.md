# DOCX style verification 기준(권장)

## 목표
템플릿과 "동일한 출력(A4 프린트)"을 목표로, 정량 항목을 PASS/WARN/FAIL로 판정한다.

## 권장 허용 오차(기본값)
- Page size: ±1.0 mm
- Margins: ±2.0 mm
- Dominant font size: ±0.5 pt
- Paragraph spacing (before/after): ±1.0 pt (또는 동급 수준)
- Line spacing: 동일 rule(단일/배수/고정) + 값 근접

## 반드시 체크할 항목
1) 페이지/섹션
- A4(210×297mm) 여부
- 섹션별 여백(top/bottom/left/right)
- 헤더/푸터 거리(가능한 경우)

2) 타이포그래피
- 본문 지배적 폰트명/크기(템플릿과 일치하는지)
- 제목/헤더 스타일 차등

3) 단락
- 정렬(left/center/right/justify)
- 줄간격 rule(단일/배수/고정) 및 값
- 문단 전/후 간격
- 들여쓰기(첫줄/내어쓰기 포함)

4) 헤더/레터헤드
- 로고 위치/크기(대략)
- 기관명/주소 블록 정렬
- 1페이지/이후 페이지 헤더 차등(있는 경우)

5) Placeholder 제거
- `{{FIELD}}` 토큰 잔존 여부
- 누락된 필수 필드(문서번호/일자/수신/제목/서명 등)

## 복잡 템플릿 주의
- 텍스트박스/도형/앵커 이미지가 많은 경우 python-docx 기반 분석은 제한적일 수 있음 → WARN 처리 + 육안 QA 권고.
