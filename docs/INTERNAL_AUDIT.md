# Internal Audit Workbooks

Simple workbooks for FY2025-26 internal audit programme based on existing audit plans in `docs/audit-plans/`.

## Workbooks

### IA Master (`ia_master.gs`)
Central coordination workbook with:
- Team allocation matrix
- Workpaper index
- Progress dashboard  
- Findings tracker

### Phase Workbooks (to be created as needed)
- H1 Revenue/OTC - 13 tests
- H1 P2P/Taxation - 26 tests  
- H1 Treasury/Systems - 20 tests
- Q3 Payroll/HR - 33 tests
- Q4 Fixed Assets/Close/IFC - 61 tests

## Usage

1. Build: `npm run build`
2. Copy `dist/ia_master_standalone.gs` to Google Apps Script
3. Run `createIAMasterWorkbook()`
4. Track progress and findings

## Reference

See detailed audit plans in `docs/audit-plans/`:
- `IA_FY2025-26_Master_Workbook.xlsx.md` - Complete plan
- `00_AUDIT_PLANS_INDEX.md` - Index of all plans
- Individual phase plans (H1, Q3, Q4)
