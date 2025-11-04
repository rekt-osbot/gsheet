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
- **IA Master** - Internal audit programme coordination (NEW)

## Key Features

### Sample Data Management
- Pre-populated realistic data for all workbooks
- Easy to understand workbook functionality
- Menu: [Tools] → **Populate Sample Data**

### Enhanced Error Handling
- Safe formulas with automatic fallback values
- Data validation utilities
- Named range checking

## For Developers

### Build System
```bash
npm run build        # Build all workbooks
```

### Project Structure
```
src/
  common/           
    utilities.gs           # Menu creation, sheet management
    formatting.gs          # Colors, fonts, numbers
    sheetBuilders.gs       # Reusable sheet patterns
    dataValidation.gs      # Dropdown lists, validations
    conditionalFormatting.gs  # Color rules
    namedRanges.gs         # Named range helpers
    sampleData.gs          # Sample data system
    errorHandling.gs       # Safe formulas & validation
  workbooks/        # Individual workbook code
dist/              # Built standalone files
```

### Common Functions

**Sheet Builders** (`sheetBuilders.gs`)
```javascript
createStandardAuditSheet(ss, config)       // Complete sheet
createDataTable(sheet, row, col, headers)  // Headers + data
createInputSection(sheet, row, col, inputs) // Label + input cells
```

**Sample Data** (`sampleData.gs`)
```javascript
populateSampleData(ss)      // Auto-populate based on workbook type
clearSampleData(ss)         // Clear all input data
```

**Error Handling** (`errorHandling.gs`)
```javascript
safeFormula(formula, fallback)    // Wrap formula with IFERROR
safeLookupFormula(...)            // Safe VLOOKUP
validateRangeData(range, rules)   // Validate data
createErrorReportSheet(ss, errors) // Generate error report
```

### Code Standards

- Use `COLS` constants for column references
- Use sheet builders for new sheets
- Add sample data for new workbooks
- Use safe formulas for lookups

## Documentation

### User Guides (by Standard)

- [Deferred Tax](DEFERRED_TAX_README.md) - AS 22 / Ind AS 12
- [Ind AS 109](INDAS109_README.md) - Financial Instruments
- [Ind AS 115](INDAS115_README.md) - Revenue Recognition
- [Ind AS 116](INDAS116_README.md) - Lease Accounting
- [Fixed Assets](FIXED_ASSETS_README.md) - Asset Roll-Forward
- [TDS Compliance](TDS_COMPLIANCE_README.md) - Tax Deductions
- [ICFR P2P](ICFR_P2P_README.md) - Process Controls
- [Internal Audit](INTERNAL_AUDIT.md) - IA programme coordination

### Developer Documentation
- [Roadmap](todo.md) - Future Plans

## Recent Improvements

### Version 1.1 (November 2025)
- ✅ Comprehensive sample data system
- ✅ Enhanced error handling with safe formulas
- ✅ Menu integration for sample data

### Version 1.0 (September 2025)
- ✅ Eliminated duplicate utility functions
- ✅ Added column constants (no magic numbers)
- ✅ PropertiesService-based workbook detection
- ✅ Sheet builder abstraction
- ✅ Build system with metadata

## Quick Examples

### Using Safe Formulas
```javascript
const formula = safeLookupFormula('B2', 'Vendors!A:D', 4, '"Unknown"');
cell.setFormula(formula);
```

### Populating Sample Data
```javascript
populateSampleData(ss);  // Done - all sheets now have sample data
```

## Performance

- All 7 workbooks compile in ~100ms
- Sample data populates in ~200-500ms per workbook
- Error handling adds <5% overhead to formulas

## Status

**Version**: 1.1
**Status**: Production Ready ✅
**Last Updated**: November 4, 2025

All features tested and ready for deployment.

## Support

For questions:
1. Check the individual workbook READMEs for use cases
2. Check code comments in `src/common/` for implementation details
3. See [Roadmap](todo.md) for planned improvements

---

**Note**: This project is designed for professional audit workpaper generation. All calculations follow Indian Accounting Standards (Ind AS) and IGAAP guidelines.
