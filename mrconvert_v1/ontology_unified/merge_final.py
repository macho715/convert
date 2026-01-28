# -*- coding: utf-8 -*-
import pathlib

base = pathlib.Path(".")

# Read files
w = (base / "temp_warehouse.md").read_text(encoding="utf-8")
b = (base / "temp_bulk.md").read_text(encoding="utf-8")
j = (base / "temp_jpt.md").read_text(encoding="utf-8")
c = pathlib.Path("../ontology_complete/cipl-bl-pre-arrival-guard.md").read_text(
    encoding="utf-8"
)
code = pathlib.Path("../ontology_complete/code.md").read_text(encoding="utf-8")
p = pathlib.Path("../ontology_complete/P.MD").read_text(encoding="utf-8")


# Remove YAML
def remove_yaml(text):
    if text.startswith("---"):
        parts = text.split("---", 2)
        if len(parts) >= 3:
            return parts[2].strip()
    return text


w2, b2, j2 = remove_yaml(w), remove_yaml(b), remove_yaml(j)
c2, code2, p2 = remove_yaml(c), remove_yaml(code), remove_yaml(p)

# Group 5
g5 = f"""---
title: "Operations Management Ontology"
type: "ontology-design"
domain: "operations-management"
sub-domains: ["warehouse", "bulk-cargo", "vessel-operations"]
version: "unified-1.0"
date: "2025-01-19"
tags: ["ontology", "warehouse", "bulk-cargo", "vessel", "jpt71", "operations", "hvdc"]
status: "active"
---

# Part 1: Warehouse Management

{w2}

---

# Part 2: Bulk Cargo Operations

{b2}

---

# Part 3: Vessel Operations (JPT 71)

{j2}
"""

# Group 6
g6 = f"""---
title: "Compliance & Customs Clearance Ontology"
type: "ontology-design"
domain: "customs-clearance"
version: "1.0"
date: "2025-01-19"
tags: ["ontology", "cipl", "bl", "pre-arrival", "customs", "hvdc", "compliance"]
status: "active"
---

{c2}
"""

# Group 7
g7 = f"""---
title: "Development & Tools Ontology"
type: "ontology-design"
domain: "software-development"
sub-domains: ["code-ontology", "document-conversion"]
version: "unified-1.0"
date: "2025-01-19"
tags: ["ontology", "code", "tdd", "development", "tools", "pdf", "docx"]
status: "active"
---

# Part 1: Code Writing Ontology

{code2}

---

# Part 2: PDF-DOCX Bidirectional Converter

{p2}
"""

# Write files
(base / "05-operations-management.md").write_text(g5, encoding="utf-8")
(base / "06-compliance-customs.md").write_text(g6, encoding="utf-8")
(base / "07-development-tools.md").write_text(g7, encoding="utf-8")

print("OK All 3 files created successfully!")
print(f"  - 05-operations-management.md ({len(g5)} bytes)")
print(f"  - 06-compliance-customs.md ({len(g6)} bytes)")
print(f"  - 07-development-tools.md ({len(g7)} bytes)")
