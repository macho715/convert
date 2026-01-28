# vessel_stability_python — AGENTS.md

Excel 안정성 계산(Stability Booklet)의 Python 변환 + 테스트/검증. 루트 **AGENTS.md** 우선.

## 엔트리포인트

| 유형 | 경로 | 용도 |
|------|------|------|
| 예제 | `vessel_stability_python/example_usage.py` | 사용 예시 실행 |
| 테스트 | `vessel_stability_python/tests/` | 단위/검증 테스트 |
| 검증 | `vessel_stability_python/validation/` | validate_stability_calculations.py 등 |

실행(README/폴더 구조 기준):

- `python vessel_stability_python/example_usage.py`
- `pytest vessel_stability_python/tests/ -q`

## 입출력

- **입력**: Stability Booklet Excel(예: `data/1.Vessel Stability Booklet.xls`).
- **출력**: 계산 결과, 검증 리포트(문서/validation 출력).

## 고정 스모크(AGENTS.md §5.4 #6)

- `pytest vessel_stability_python/tests/ -q`
- 또는 `python vessel_stability_python/example_usage.py` (실행 가능 시)

## 주의

- Excel 데이터 의존. 테스트/검증은 모듈 내 paths 기준.
