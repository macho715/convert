# Sheet Mapping Guide

This guide maps sheet names between the normalized model and the legacy TR tracker.

## Normalized -> Legacy Mapping

| Normalized (통합빌더.py) | Legacy (create_tr_document_tracker_v2.py) | Notes |
|---|---|---|
| `C_Config` | `Config` | Key/value configuration |
| `S_Voyages` | `Voyage_Schedule` | Voyage schedule table |
| `M_DocCatalog` | `Doc_Matrix` | Document catalog |
| `M_Parties` | `Party_Contacts` | Parties/contacts |
| `R_DeadlineRules` | `Inputs` | Scenario + lead time rules |
| `T_Tracker` | `Document_Tracker` | Main tracker |
| `D_Dashboard` | `Dashboard` | KPI dashboard |
| `Calendar_View` | `Party_View` | Calendar vs party view |
| `Holidays` | (not present) | Optional in legacy |
| `VBA_Pasteboard` | `VBA_Pasteboard` | Unified VBA repo |
| `Instructions` | `Instructions` | Usage guide |

## DocGap Patch Sheets

| DocGap Patch | Description |
|---|---|
| `Inputs` | Scenario selector + lead time mapping |
| `OFCO_Req_1_15` | OFCO requirements with lead time columns |
| `NOC_Req_1_6` | NOC requirements with lead time columns |
| `Executive_Summary` | Summary stamp target |
| `VBA_Pasteboard` | DocGap macros appended |

## Notes

- In the legacy model, `Inputs` is optional unless DocGap patch is applied.
- `Voyage_Schedule` in legacy links Voyage 1 to `Inputs` when present.
