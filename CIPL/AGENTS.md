# CIPL — AGENTS.md

Commercial Invoice & Packing List 자동 생성. 최신 코드는 **CIPL_PATCH_PACKAGE/** 사용. 루트 **AGENTS.md** 우선.

## 엔트리포인트

| 유형 | 경로 | 용도 |
|------|------|------|
| 빌더 | `CIPL/CIPL_PATCH_PACKAGE/make_cipl_set.py` | JSON → 4-sheet CIPL xlsx |

실행(문서 기준):

- `python CIPL/CIPL_PATCH_PACKAGE/make_cipl_set.py --in <voyage_input.json> [--out <output.xlsx>]`

## 입출력

- **입력**: `voyage_input*.json` (LDG/Commons 스키마). 샘플: `CIPL_PATCH_PACKAGE/voyage_input_sample_full.json`.
- **출력**: 4-sheet CIPL xlsx (기본/지정 out 경로). 생성물은 `out/` 또는 지정 경로.

## 고정 스모크(AGENTS.md §5.4 #5)

- `python CIPL/CIPL_PATCH_PACKAGE/make_cipl_set.py --in CIPL/CIPL_PATCH_PACKAGE/voyage_input_sample_full.json --out out/CIPL_smoke.xlsx` (출력 포맷/서식 확인)

## 주의

- 서식/템플릿 유지. 레거시는 `CIPL_LEGACY/` 참고만.
