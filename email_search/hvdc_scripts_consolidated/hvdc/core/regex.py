import re

# 단일 카탈로그 — 기존 스크립트에서 쓰던 변형을 포괄
# 필요시 tests로 스냅샷 고정
PATTERNS = {
    # 예) HVDC-ADOPT-HE-0476, HVDC-ADOPT-SCT-0136 등
    "HVDC_ADOPT": r"HVDC-ADOPT-[A-Z]+-[0-9]{3,5}",
    # 보다 일반화된 HVDC 코드(세그먼트 3~4개)
    "HVDC_GENERIC": r"HVDC-[A-Z]+-[A-Z0-9]+-[A-Z0-9\-]+",
    # PRL 코드(예시): PRL-O-046-O4(HE-0486)류
    "PRL": r"PRL-[A-Z]-\d{2,4}-[A-Z0-9\-]*(?:\([A-Z]{2}-\d{3,5}\))?",
    # JPTW/GRM 페어 추출: ... JPTW-71 / GRM-123 ...
    "JPTW_GRM": r"JPTW-(\d+)\s*/\s*GRM-(\d+)",
    # 괄호 벤더/부가 정보
    "PAREN_ANY": r"\(([^\)]+)\)",
    # 뒤따르는 트레일러 식별자(완충)
    "TRAILING": r":\s*([A-Z]+(?:-[A-Z0-9]+){2,})",
}

# 사전 컴파일(성능·일관)
COMPILED = {k: re.compile(v, re.IGNORECASE) for k, v in PATTERNS.items()}