# Ind AS Audit Builder - Documentation

## Quick Start

1. Open Google Sheets
2. Extensions → Apps Script
3. Copy code from `dist/[workbook]_standalone.gs`
4. Save and refresh sheet
5. Use menu to create workbook
6. Use "Populate Sample Data" to see examples

## Available Workbooks

- **Deferred Tax** - AS 22 / Ind AS 12 compliance
- **Ind AS 109** - Financial instruments and ECL
- **Ind AS 115** - Revenue recognition
- **Ind AS 116** - Lease accounting
- **Fixed Assets** - Asset roll-forward
- **TDS Compliance** - Tax deduction tracking
- **ICFR P2P** - Procure-to-pay controls

## For Developers

### Build System
```bash
npm run build        # Build all workbooks
npm run build:watch  # Auto-rebuild on changes
```

### Project Structure
```
src/
  common/           # Shared utilities
  workbooks/        # Individual workbook code
dist/              # Built standalone files
```

### Key Functions

**Sheet Builders** (`sheetBuilders.gs`)
- `createStandardAuditSheet()` - Complete sheet with config
- `createDataTable()` - Headers + data
- `createInputSection()` - Label + input cells

**Testing** (`testing.gs`)
- `runAllTests()` - Run all tests
- Menu: 🧪 Testing → Run All Tests

**Sample Data** (`sampleData.gs`)
- `populateWorkbookSampleData()` - Auto-detects workbook type
- Menu: [Workbook] Tools → Populate Sample Data

### Code Standards

- Use `COLS` constants for column references
- Use sheet builders for new sheets
- Add sample data for new workbooks
- Test common functions

## User Guides

See individual workbook READMEs:
- [Deferred Tax](DEFERRED_TAX_README.md)
- [Ind AS 109](INDAS109_README.md)
- [Ind AS 115](INDAS115_README.md)
- [Ind AS 116](INDAS116_README.md)
- [Fixed Assets](FIXED_ASSETS_README.md)
- [TDS Compliance](TDS_COMPLIANCE_README.md)
- [ICFR P2P](ICFR_P2P_README.md)

## Recent Improvements

- ✅ Eliminated duplicate utility functions
- ✅ Added column constants (no magic numbers)
- ✅ PropertiesService-based workbook detection
- ✅ Sheet builder abstraction
- ✅ Automated testing framework
- ✅ Sample data for all workbooks

See [CODE_IMPROVEMENTS.md](CODE_IMPROVEMENTS.md) for details.
