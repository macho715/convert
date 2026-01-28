# Builder Selection Guide

This guide helps choose the right builder based on your target workbook model.

## Quick Selection

- Normalized model (recommended): use `통합빌더.py`
- Legacy TR tracker: use `create_tr_document_tracker_v2.py`
- Legacy + DocGap integration: run `create_tr_document_tracker_v2.py` then `build_docgap_v3_1_operational.py`
- DocGap v2 upgrade: use `build_docgap_v3_fulloptions.py` (requires v2 source)
- Patch existing file: use `build_docgap_v3_1_operational.py`

## Builder Matrix

| Builder | Mode | Inputs | Outputs | Best For |
|---|---|---|---|---|
| `통합빌더.py` | CREATE | none | normalized template `.xlsx` | New unified model with rules table |
| `create_tr_document_tracker_v2.py` | CREATE + REFRESH | none | legacy TR template `.xlsx` | Legacy TR tracking with refresh |
| `build_docgap_v3_1_operational.py` | PATCH | existing `.xlsx` | patched `.xlsx` | Add Inputs + lead time mapping |
| `build_docgap_v3_fulloptions.py` | PATCH | DocGap v2 `.xlsx` | v3 fulloptions `.xlsx` + `.xlsm` | DocGap v2 migration |

## Recommended Paths

1) Normalized model (best long-term)
- Build with `통합빌더.py`
- Convert to `.xlsm`, import VBA modules

2) Legacy model + DocGap integration
- Build TR template with `create_tr_document_tracker_v2.py`
- Patch with `build_docgap_v3_1_operational.py`
- Convert to `.xlsm`, import VBA modules

3) DocGap-only upgrade
- Run `build_docgap_v3_fulloptions.py` using a v2 source file
- Import DocGap macros

## Notes

- Builder outputs are placed in `05_Templates`.
- VBA modules live in `02_VBA_Modules`.
- Sheet code snippets live in `03_Sheet_Codes`.
