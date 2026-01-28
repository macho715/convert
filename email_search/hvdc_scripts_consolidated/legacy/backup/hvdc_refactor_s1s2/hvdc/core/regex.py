
import re

PATTERNS = {
    "HVDC_ADOPT": r"HVDC-ADOPT-[A-Z]+-[0-9]{3,5}",
    "HVDC_GENERIC": r"HVDC-[A-Z]+-[A-Z]+-[0-9A-Z\-]+",
    "PRL": r"PRL-[A-Z]-\d{2,4}-[A-Z0-9]+(?:\([A-Z]{2}-\d{3,5}\))?",
    "JPTW_GRM": r"JPTW-(\d+)\s*/\s*GRM-(\d+)",
    "PAREN_ANY": r"\(([^\)]+)\)",
    "TRAILING": r":\s*([A-Z]+(?:-[A-Z0-9]+){2,})",
}
COMPILED = {k: re.compile(v, re.IGNORECASE) for k, v in PATTERNS.items()}
