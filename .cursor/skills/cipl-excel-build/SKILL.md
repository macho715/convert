---
name: cipl-excel-build
description: CIPL(Commercial Invoice & Packing List) Excel 생성 작업에서 템플릿 서식 유지(Style-first)와 회귀 체크를 강제한다. "CIPL", "invoice packing list", "xlsx template" 요청에 사용.
---

# cipl-excel-build

## 목표
- CIPL Excel을 생성/수정하되, "기존 템플릿 서식/레이아웃"을 깨지 않는다(Style-first).
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
- excel-style-guardian 서브에이전트 관점의 체크리스트로 "Same/Not same"를 표로 남김

## Ask first
- xlsm 바이너리 자동 수정
- 템플릿 구조(시트/컬럼) 대규모 변경

## 산출물
- 실행 커맨드(확정본)
- 서식 회귀 체크리스트 표
- FAIL이면: 어떤 서식이 깨졌는지 + 안전한 수정 방향(builder에서 style copy 등)
