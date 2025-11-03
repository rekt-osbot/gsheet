# Ind AS Audit Builder

Automated audit workpaper generation for Indian Accounting Standards.

## Quick Start

1. Open Google Sheets
2. Extensions → Apps Script
3. Copy code from `dist/[workbook]_standalone.gs`
4. Save, refresh sheet
5. Use menu to create workbook

## Available Workbooks

| Workbook | Standard | Purpose |
|----------|----------|---------|
| Deferred Tax | AS 22 / Ind AS 12 | DTA/DTL calculations |
| Ind AS 109 | Financial Instruments | ECL, classification, measurement |
| Ind AS 115 | Revenue | Contract accounting, performance obligations |
| Ind AS 116 | Leases | ROU assets, lease liabilities |
| Fixed Assets | AS 10 / Ind AS 16 | Asset roll-forward |
| TDS Compliance | Income Tax Act | TDS tracking, 26AS reconciliation |
| ICFR P2P | SOX/COSO | Procure-to-pay controls |

## Features

- Auto-generated sheets with formulas
- Built-in validations and formatting
- Sample data for testing
- Audit trail and references
- No macros - pure formulas

## For Developers

```bash
npm install
npm run build        # Build all workbooks
npm run build:watch  # Auto-rebuild
```

Project structure:
- `src/common/` - Shared utilities
- `src/workbooks/` - Individual workbooks
- `dist/` - Built standalone files

## Documentation

- [Developer Guide](docs/README.md)
- [Code Improvements](docs/CODE_IMPROVEMENTS.md)
- Individual workbook READMEs in `docs/`

## License

MIT
