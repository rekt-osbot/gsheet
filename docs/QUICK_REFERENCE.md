# Quick Reference Card

## Common Tasks

### Build the Project
```bash
npm run build
```

### Use Column Constants
```javascript
// ❌ Don't
sheet.getRange(row, 7).setValue(value);

// ✅ Do
sheet.getRange(row, COLS.TEMP_DIFF.NATURE).setValue(value);
```

### Use Common Utilities
```javascript
// Available in all workbooks (from src/common/)
clearExistingSheets(ss);
setColumnWidths(sheet, [100, 200, 150]);
protectSheet(sheet, true);
formatHeader(sheet, row, startCol, endCol, text);
formatCurrency(range);
formatPercentage(range);
formatDate(range);
```

### Use Color Constants
```javascript
// Available in all workbooks (from src/common/formatting.gs)
COLORS.HEADER_BG        // "#1a237e" - Dark blue
COLORS.HEADER_TEXT      // "#ffffff" - White
COLORS.SUBHEADER_BG     // "#3949ab" - Medium blue
COLORS.INPUT_BG         // "#fff9c4" - Light yellow
COLORS.INPUT_ALT_BG     // "#b3e5fc" - Light blue
COLORS.CALC_BG          // "#e8eaf6" - Light purple-grey
COLORS.SECTION_BG       // "#c5cae9" - Light blue-grey
COLORS.TOTAL_BG         // "#ffccbc" - Light orange
COLORS.GRAND_TOTAL_BG   // "#ff8a65" - Orange
COLORS.WARNING_BG       // "#ffebee" - Light red
COLORS.SUCCESS_BG       // "#c8e6c9" - Light green
COLORS.INFO_BG          // "#e1f5fe" - Very light blue
COLORS.BORDER_COLOR     // "#757575" - Grey
```

### Set Workbook Type
```javascript
// In your main workbook creation function
function createMyWorkbook() {
  setWorkbookType('MY_WORKBOOK_TYPE');
  // ... rest of code
}
```

## File Structure

```
gsheet-audit-workpapers/
├── src/
│   ├── common/              # Shared utilities (DON'T duplicate these!)
│   │   ├── utilities.gs     # Sheet management, menu creation
│   │   ├── formatting.gs    # Colors, formatting functions
│   │   ├── dataValidation.gs
│   │   ├── conditionalFormatting.gs
│   │   ├── sheetBuilders.gs
│   │   └── namedRanges.gs
│   └── workbooks/           # Workbook-specific code
│       ├── deferredtax.gs
│       ├── indas109.gs
│       ├── indas115.gs
│       ├── indas116.gs
│       ├── far_wp.gs
│       ├── ifc_p2p.gs
│       └── tds_compliance.gs
├── dist/                    # Generated standalone files (DON'T edit!)
├── docs/                    # Documentation
└── build.js                 # Build script
```

## Workbook Types

| Type | Function Name | Menu Name |
|------|---------------|-----------|
| `DEFERRED_TAX` | `createDeferredTaxWorkbook()` | Deferred Tax Tools |
| `INDAS109` | `createIndAS109WorkingPapers()` | Ind AS 109 Tools |
| `INDAS115` | `buildIndAS115Workpaper()` | Ind AS 115 Tools |
| `INDAS116` | `createIndAS116Workbook()` | Ind AS 116 Tools |
| `FIXED_ASSETS` | `setupFixedAssetsWorkpaper()` | Fixed Assets Tools |
| `TDS_COMPLIANCE` | `createTDSComplianceWorkbook()` | TDS Tools |
| `ICFR_P2P` | `createICFRP2PWorkbook()` | ICFR Tools |

## Column Constants by Workbook

### Deferred Tax
```javascript
COLS.TEMP_DIFF.SR_NO
COLS.TEMP_DIFF.ITEM
COLS.TEMP_DIFF.CATEGORY
COLS.TEMP_DIFF.TAX_BASE
COLS.TEMP_DIFF.BOOK_BASE
COLS.TEMP_DIFF.TEMP_DIFF
COLS.TEMP_DIFF.NATURE
COLS.TEMP_DIFF.OPENING
COLS.TEMP_DIFF.ADDITIONS
COLS.TEMP_DIFF.REVERSALS
COLS.TEMP_DIFF.RATE_CHANGE
COLS.TEMP_DIFF.REMARKS
```

### Ind AS 109
```javascript
COLS.INSTRUMENTS_REGISTER.ID
COLS.INSTRUMENTS_REGISTER.NAME
COLS.INSTRUMENTS_REGISTER.TYPE
COLS.INSTRUMENTS_REGISTER.COUNTERPARTY
COLS.INSTRUMENTS_REGISTER.ISSUE_DATE
COLS.INSTRUMENTS_REGISTER.MATURITY_DATE
COLS.INSTRUMENTS_REGISTER.FACE_VALUE
COLS.INSTRUMENTS_REGISTER.COUPON_RATE
COLS.INSTRUMENTS_REGISTER.EIR
COLS.INSTRUMENTS_REGISTER.OPENING_BALANCE
// ... (see COLUMN_CONSTANTS_GUIDE.md for full list)
```

### Ind AS 115
```javascript
COLS.CONTRACT_REGISTER.SR_NO
COLS.CONTRACT_REGISTER.CONTRACT_ID
COLS.CONTRACT_REGISTER.CUSTOMER
COLS.CONTRACT_REGISTER.CONTRACT_DATE
COLS.CONTRACT_REGISTER.DESCRIPTION
// ... (see COLUMN_CONSTANTS_GUIDE.md for full list)

COLS.REVENUE_RECOGNITION.SR_NO
COLS.REVENUE_RECOGNITION.CONTRACT_ID
COLS.REVENUE_RECOGNITION.CUSTOMER
// ... (see COLUMN_CONSTANTS_GUIDE.md for full list)
```

### Ind AS 116
```javascript
COLS.LEASE_REGISTER.ID
COLS.LEASE_REGISTER.DESCRIPTION
COLS.LEASE_REGISTER.LESSOR
COLS.LEASE_REGISTER.COMMENCEMENT_DATE
// ... (see COLUMN_CONSTANTS_GUIDE.md for full list)

COLS.ROU_ASSET.LEASE_ID
COLS.ROU_ASSET.OPENING_BALANCE
COLS.ROU_ASSET.ADDITIONS
COLS.ROU_ASSET.DEPRECIATION
COLS.ROU_ASSET.CLOSING_BALANCE
```

### Fixed Assets
```javascript
COLS.ROLL_FORWARD.ASSET_CLASS
COLS.ROLL_FORWARD.OPENING_GROSS
COLS.ROLL_FORWARD.ADDITIONS
COLS.ROLL_FORWARD.DISPOSALS
COLS.ROLL_FORWARD.TRANSFERS
COLS.ROLL_FORWARD.CLOSING_GROSS
COLS.ROLL_FORWARD.OPENING_ACCUM_DEP
COLS.ROLL_FORWARD.DEPRECIATION
COLS.ROLL_FORWARD.DISPOSAL_DEP
COLS.ROLL_FORWARD.CLOSING_ACCUM_DEP
COLS.ROLL_FORWARD.OPENING_NBV
COLS.ROLL_FORWARD.CLOSING_NBV
```

### ICFR P2P
```javascript
COLS.RCM.CONTROL_ID
COLS.RCM.PROCESS
COLS.RCM.RISK
COLS.RCM.CONTROL_ACTIVITY
COLS.RCM.CONTROL_TYPE
COLS.RCM.FREQUENCY
COLS.RCM.OWNER
COLS.RCM.KEY_CONTROL

COLS.TEST_OF_DESIGN.CONTROL_ID
COLS.TEST_OF_DESIGN.CONTROL_DESC
COLS.TEST_OF_DESIGN.DESIGN_PROCEDURE
COLS.TEST_OF_DESIGN.EVIDENCE
COLS.TEST_OF_DESIGN.CONCLUSION
COLS.TEST_OF_DESIGN.TESTER
COLS.TEST_OF_DESIGN.DATE
```

## Common Patterns

### Creating a New Sheet
```javascript
function createMySheet(ss) {
  const sheet = ss.insertSheet('My_Sheet');
  
  // Set column widths
  setColumnWidths(sheet, [100, 200, 150, 120]);
  
  // Create header
  formatHeader(sheet, 1, 1, 4, 'MY SHEET TITLE', COLORS.HEADER_BG);
  
  // Create sub-header
  const headers = ['Column 1', 'Column 2', 'Column 3', 'Column 4'];
  formatSubHeader(sheet, 2, 1, headers, COLORS.SUBHEADER_BG);
  
  // Add data with formulas
  for (let row = 3; row <= 50; row++) {
    sheet.getRange(row, COLS.MY_SHEET.ID).setValue(row - 2);
    sheet.getRange(row, COLS.MY_SHEET.NAME).setBackground(COLORS.INPUT_BG);
    sheet.getRange(row, COLS.MY_SHEET.VALUE).setFormula(`=B${row}*2`);
  }
  
  // Freeze header rows
  sheet.setFrozenRows(2);
}
```

### Adding Data Validation
```javascript
const rule = SpreadsheetApp.newDataValidation()
  .requireValueInList(['Option 1', 'Option 2', 'Option 3'], true)
  .setAllowInvalid(false)
  .setHelpText('Select one option')
  .build();

sheet.getRange(3, COLS.MY_SHEET.STATUS, 50, 1).setDataValidation(rule);
```

### Formatting Ranges
```javascript
// Currency
formatCurrency(sheet.getRange(3, COLS.MY_SHEET.AMOUNT, 50, 1));

// Percentage
formatPercentage(sheet.getRange(3, COLS.MY_SHEET.RATE, 50, 1));

// Date
formatDate(sheet.getRange(3, COLS.MY_SHEET.DATE, 50, 1));

// Input cells
sheet.getRange(3, COLS.MY_SHEET.INPUT, 50, 1)
  .setBackground(COLORS.INPUT_BG);
```

## Dos and Don'ts

### ✅ Do
- Use column constants instead of numbers
- Use common utilities from `src/common/`
- Call `setWorkbookType()` in main function
- Run `npm run build` after changes
- Document complex formulas
- Test thoroughly before deploying

### ❌ Don't
- Edit files in `dist/` folder (they're auto-generated)
- Duplicate utility functions in workbook files
- Use magic numbers for columns/rows
- Forget to update column constants when adding columns
- Skip the build step
- Deploy without testing

## Troubleshooting

### Build Fails
```bash
# Check for syntax errors
npm run build

# If errors, check the error message for file and line number
```

### Menu Not Showing
```javascript
// Make sure you called setWorkbookType()
function createMyWorkbook() {
  setWorkbookType('MY_WORKBOOK_TYPE');  // Add this!
  // ...
}
```

### Column Constants Not Working
```javascript
// Make sure COLS is defined at top of workbook file
const COLS = {
  MY_SHEET: {
    ID: 1,
    NAME: 2,
    // ...
  }
};
```

### Function Not Found
```javascript
// Make sure you're using common utilities, not redefining them
// ❌ Don't redefine
function clearExistingSheets(ss) { ... }

// ✅ Just use it (it's in src/common/utilities.gs)
clearExistingSheets(ss);
```

## Getting Help

1. **Code Improvements**: See `docs/CODE_IMPROVEMENTS.md`
2. **Column Constants**: See `docs/COLUMN_CONSTANTS_GUIDE.md`
3. **Full Summary**: See `REFACTORING_SUMMARY.md`
4. **Build System**: See `docs/BUILD_SYSTEM.md`

## Version Info

- **Current Version**: 1.0.0
- **Last Updated**: November 2025
- **Build System**: Node.js + npm
- **Target Platform**: Google Apps Script
