# Build Checklist

Use this checklist after generating a template.

## Build

- [ ] Run selected builder (or `run_all_builders.py`)
- [ ] Confirm output `.xlsx` appears in `05_Templates`

## Excel Packaging

- [ ] Open the `.xlsx` and save as `.xlsm`
- [ ] Import VBA modules from `02_VBA_Modules`
- [ ] Paste sheet code from `03_Sheet_Codes` into the correct sheet module
- [ ] Add shortcuts in `ThisWorkbook` (see `ThisWorkbook_Shortcuts.bas`)

## Validation

- [ ] Run `RefreshAll_ControlTower()`
- [ ] Verify KPI rows (D-7/D-3/D-1) update
- [ ] Verify Inputs -> Voyage 1 linkage (legacy + DocGap)
- [ ] Confirm VBA_Pasteboard contains TR + DocGap macro sections

## Packaging

- [ ] Save `.xlsm` and keep the original `.xlsx`
- [ ] Archive final outputs in `05_Templates`
