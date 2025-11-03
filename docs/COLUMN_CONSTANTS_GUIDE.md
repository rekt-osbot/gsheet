# Column Constants Usage Guide

## Overview
This guide explains how to use the new column constants system to write more maintainable and readable code.

## Why Use Column Constants?

### ❌ Bad (Magic Numbers)
```javascript
// What does column 7 mean? Nobody knows without checking the sheet!
sheet.getRange(row, 7).setValue('DTL');
sheet.getRange(row, 4).setFormula('=E' + row + '-D' + row);
```

### ✅ Good (Named Constants)
```javascript
// Crystal clear what each column represents
sheet.getRange(row, COLS.TEMP_DIFF.NATURE).setValue('DTL');
sheet.getRange(row, COLS.TEMP_DIFF.TAX_BASE).setFormula(
  `=E${row}-D${row}`
);
```

## How to Use Column Constants

### 1. Find the Constants
Each workbook file has a `COLS` object at the top:

```javascript
// src/workbooks/deferredtax.gs
const COLS = {
  TEMP_DIFF: {
    SR_NO: 1,
    ITEM: 2,
    CATEGORY: 3,
    TAX_BASE: 4,
    BOOK_BASE: 5,
    TEMP_DIFF: 6,
    NATURE: 7,
    OPENING: 8,
    ADDITIONS: 9,
    REVERSALS: 10,
    RATE_CHANGE: 11,
    REMARKS: 12
  }
};
```

### 2. Use in Your Code

#### Setting Values
```javascript
// Old way
sheet.getRange(row, 2).setValue('Depreciation');

// New way
sheet.getRange(row, COLS.TEMP_DIFF.ITEM).setValue('Depreciation');
```

#### Setting Formulas
```javascript
// Old way
sheet.getRange(row, 6).setFormula('=E' + row + '-D' + row);

// New way
sheet.getRange(row, COLS.TEMP_DIFF.TEMP_DIFF).setFormula(
  `=E${row}-D${row}`
);
```

#### Formatting Ranges
```javascript
// Old way
sheet.getRange(7, 4, 50, 2).setBackground('#fff9c4');

// New way
const startCol = COLS.TEMP_DIFF.TAX_BASE;
const numCols = 2; // TAX_BASE and BOOK_BASE
sheet.getRange(7, startCol, 50, numCols).setBackground(COLORS.INPUT_BG);
```

#### Data Validation
```javascript
// Old way
sheet.getRange('C7:C50').setDataValidation(categoryRule);

// New way
const col = COLS.TEMP_DIFF.CATEGORY;
sheet.getRange(7, col, 44, 1).setDataValidation(categoryRule);
```

## Available Constants by Workbook

### Deferred Tax (`deferredtax.gs`)
```javascript
COLS.TEMP_DIFF.SR_NO          // Column 1
COLS.TEMP_DIFF.ITEM           // Column 2
COLS.TEMP_DIFF.CATEGORY       // Column 3
COLS.TEMP_DIFF.TAX_BASE       // Column 4
COLS.TEMP_DIFF.BOOK_BASE      // Column 5
COLS.TEMP_DIFF.TEMP_DIFF      // Column 6
COLS.TEMP_DIFF.NATURE         // Column 7
COLS.TEMP_DIFF.OPENING        // Column 8
COLS.TEMP_DIFF.ADDITIONS      // Column 9
COLS.TEMP_DIFF.REVERSALS      // Column 10
COLS.TEMP_DIFF.RATE_CHANGE    // Column 11
COLS.TEMP_DIFF.REMARKS        // Column 12

ROWS.ASSUMPTIONS.ENTITY_NAME       // Row 5
ROWS.ASSUMPTIONS.FY                // Row 6
ROWS.ASSUMPTIONS.FRAMEWORK         // Row 7
ROWS.ASSUMPTIONS.REPORTING_DATE    // Row 8
ROWS.ASSUMPTIONS.CURRENT_TAX_RATE  // Row 13
ROWS.ASSUMPTIONS.DT_RATE_CURRENT   // Row 14
```

### Ind AS 109 (`indas109.gs`)
```javascript
COLS.INSTRUMENTS_REGISTER.ID                // Column 1
COLS.INSTRUMENTS_REGISTER.NAME              // Column 2
COLS.INSTRUMENTS_REGISTER.TYPE              // Column 3
COLS.INSTRUMENTS_REGISTER.COUNTERPARTY      // Column 4
COLS.INSTRUMENTS_REGISTER.ISSUE_DATE        // Column 5
COLS.INSTRUMENTS_REGISTER.MATURITY_DATE     // Column 6
COLS.INSTRUMENTS_REGISTER.FACE_VALUE        // Column 7
COLS.INSTRUMENTS_REGISTER.COUPON_RATE       // Column 8
COLS.INSTRUMENTS_REGISTER.EIR               // Column 9
COLS.INSTRUMENTS_REGISTER.OPENING_BALANCE   // Column 10
COLS.INSTRUMENTS_REGISTER.CURRENCY          // Column 11
COLS.INSTRUMENTS_REGISTER.SECURITY_TYPE     // Column 12
COLS.INSTRUMENTS_REGISTER.CREDIT_RATING     // Column 13
COLS.INSTRUMENTS_REGISTER.DPD               // Column 14
COLS.INSTRUMENTS_REGISTER.SPPI_TEST         // Column 15
COLS.INSTRUMENTS_REGISTER.BUSINESS_MODEL    // Column 16
COLS.INSTRUMENTS_REGISTER.DESIGNATED_FVTPL  // Column 17
COLS.INSTRUMENTS_REGISTER.FVOCI_EQUITY      // Column 18
COLS.INSTRUMENTS_REGISTER.COUPON_FREQ       // Column 19
COLS.INSTRUMENTS_REGISTER.SIMPLIFIED_ECL    // Column 20

ROWS.INPUT_VARS.REPORTING_DATE       // Row 4
ROWS.INPUT_VARS.PREV_REPORTING_DATE  // Row 5
ROWS.INPUT_VARS.RISK_FREE_RATE       // Row 6
ROWS.INPUT_VARS.DAYS_IN_YEAR         // Row 7
ROWS.INPUT_VARS.DAYS_IN_PERIOD       // Row 8
```

### Ind AS 115 (`indas115.gs`)
```javascript
COLS.CONTRACT_REGISTER.SR_NO           // Column 1
COLS.CONTRACT_REGISTER.CONTRACT_ID     // Column 2
COLS.CONTRACT_REGISTER.CUSTOMER        // Column 3
COLS.CONTRACT_REGISTER.CONTRACT_DATE   // Column 4
COLS.CONTRACT_REGISTER.DESCRIPTION     // Column 5
COLS.CONTRACT_REGISTER.CONTRACT_VALUE  // Column 6
COLS.CONTRACT_REGISTER.GST_AMOUNT      // Column 7
COLS.CONTRACT_REGISTER.TOTAL_VALUE     // Column 8
COLS.CONTRACT_REGISTER.START_DATE      // Column 9
COLS.CONTRACT_REGISTER.END_DATE        // Column 10
COLS.CONTRACT_REGISTER.DURATION        // Column 11
COLS.CONTRACT_REGISTER.PATTERN         // Column 12
COLS.CONTRACT_REGISTER.NUM_PO          // Column 13
COLS.CONTRACT_REGISTER.STATUS          // Column 14
COLS.CONTRACT_REGISTER.NOTES           // Column 15

COLS.REVENUE_RECOGNITION.SR_NO            // Column 1
COLS.REVENUE_RECOGNITION.CONTRACT_ID      // Column 2
COLS.REVENUE_RECOGNITION.CUSTOMER         // Column 3
COLS.REVENUE_RECOGNITION.STEP1_IDENTIFIED // Column 4
COLS.REVENUE_RECOGNITION.STEP2_PO         // Column 5
COLS.REVENUE_RECOGNITION.STEP3_PRICE      // Column 6
COLS.REVENUE_RECOGNITION.STEP4_ALLOCATED  // Column 7
COLS.REVENUE_RECOGNITION.STEP5_RECOGNIZED // Column 8
COLS.REVENUE_RECOGNITION.CALC_BASIS       // Column 9
COLS.REVENUE_RECOGNITION.PROGRESS_PCT     // Column 10
```

### Ind AS 116 (`indas116.gs`)
```javascript
COLS.LEASE_REGISTER.ID                    // Column 1
COLS.LEASE_REGISTER.DESCRIPTION           // Column 2
COLS.LEASE_REGISTER.LESSOR                // Column 3
COLS.LEASE_REGISTER.COMMENCEMENT_DATE     // Column 4
COLS.LEASE_REGISTER.END_DATE              // Column 5
COLS.LEASE_REGISTER.TERM_MONTHS           // Column 6
COLS.LEASE_REGISTER.MONTHLY_PAYMENT       // Column 7
COLS.LEASE_REGISTER.IBR                   // Column 8
COLS.LEASE_REGISTER.INITIAL_DIRECT_COSTS  // Column 9
COLS.LEASE_REGISTER.LEASE_INCENTIVES      // Column 10

COLS.ROU_ASSET.LEASE_ID          // Column 1
COLS.ROU_ASSET.OPENING_BALANCE   // Column 2
COLS.ROU_ASSET.ADDITIONS         // Column 3
COLS.ROU_ASSET.DEPRECIATION      // Column 4
COLS.ROU_ASSET.CLOSING_BALANCE   // Column 5
```

### Fixed Assets (`far_wp.gs`)
```javascript
COLS.ROLL_FORWARD.ASSET_CLASS        // Column 1
COLS.ROLL_FORWARD.OPENING_GROSS      // Column 2
COLS.ROLL_FORWARD.ADDITIONS          // Column 3
COLS.ROLL_FORWARD.DISPOSALS          // Column 4
COLS.ROLL_FORWARD.TRANSFERS          // Column 5
COLS.ROLL_FORWARD.CLOSING_GROSS      // Column 6
COLS.ROLL_FORWARD.OPENING_ACCUM_DEP  // Column 7
COLS.ROLL_FORWARD.DEPRECIATION       // Column 8
COLS.ROLL_FORWARD.DISPOSAL_DEP       // Column 9
COLS.ROLL_FORWARD.CLOSING_ACCUM_DEP  // Column 10
COLS.ROLL_FORWARD.OPENING_NBV        // Column 11
COLS.ROLL_FORWARD.CLOSING_NBV        // Column 12
```

### ICFR P2P (`ifc_p2p.gs`)
```javascript
COLS.RCM.CONTROL_ID         // Column 1
COLS.RCM.PROCESS            // Column 2
COLS.RCM.RISK               // Column 3
COLS.RCM.CONTROL_ACTIVITY   // Column 4
COLS.RCM.CONTROL_TYPE       // Column 5
COLS.RCM.FREQUENCY          // Column 6
COLS.RCM.OWNER              // Column 7
COLS.RCM.KEY_CONTROL        // Column 8

COLS.TEST_OF_DESIGN.CONTROL_ID       // Column 1
COLS.TEST_OF_DESIGN.CONTROL_DESC     // Column 2
COLS.TEST_OF_DESIGN.DESIGN_PROCEDURE // Column 3
COLS.TEST_OF_DESIGN.EVIDENCE         // Column 4
COLS.TEST_OF_DESIGN.CONCLUSION       // Column 5
COLS.TEST_OF_DESIGN.TESTER           // Column 6
COLS.TEST_OF_DESIGN.DATE             // Column 7
```

## Best Practices

### 1. Always Use Constants for Column References
```javascript
// ❌ Don't do this
sheet.getRange(row, 7).setValue(value);

// ✅ Do this
sheet.getRange(row, COLS.TEMP_DIFF.NATURE).setValue(value);
```

### 2. Use Descriptive Variable Names
```javascript
// ❌ Unclear
const c = COLS.TEMP_DIFF.TAX_BASE;

// ✅ Clear
const taxBaseCol = COLS.TEMP_DIFF.TAX_BASE;
```

### 3. Calculate Column Ranges
```javascript
// When you need a range of columns
const startCol = COLS.TEMP_DIFF.TAX_BASE;
const endCol = COLS.TEMP_DIFF.BOOK_BASE;
const numCols = endCol - startCol + 1;

sheet.getRange(startRow, startCol, numRows, numCols)
  .setBackground(COLORS.INPUT_BG);
```

### 4. Document Complex Formulas
```javascript
// ✅ Good - explains what the formula does
// Calculate temporary difference: Book Base - Tax Base
sheet.getRange(row, COLS.TEMP_DIFF.TEMP_DIFF).setFormula(
  `=${columnToLetter(COLS.TEMP_DIFF.BOOK_BASE)}${row}` +
  `-${columnToLetter(COLS.TEMP_DIFF.TAX_BASE)}${row}`
);
```

### 5. Group Related Operations
```javascript
// ✅ Good - groups all operations on the same column
const natureCol = COLS.TEMP_DIFF.NATURE;

// Set formula
sheet.getRange(row, natureCol).setFormula(
  `=IF(F${row}>0,"DTL",IF(F${row}<0,"DTA","-"))`
);

// Add data validation
sheet.getRange(row, natureCol).setDataValidation(natureRule);

// Format
sheet.getRange(row, natureCol).setHorizontalAlignment('center');
```

## Adding New Constants

When adding new sheets or columns to a workbook:

### 1. Update the COLS Object
```javascript
const COLS = {
  // Existing sheets...
  
  // New sheet
  NEW_SHEET: {
    ID: 1,
    NAME: 2,
    VALUE: 3,
    STATUS: 4
  }
};
```

### 2. Use in Your Code
```javascript
function createNewSheet(ss) {
  const sheet = ss.insertSheet('New_Sheet');
  
  // Use the new constants
  sheet.getRange(1, COLS.NEW_SHEET.ID).setValue('ID');
  sheet.getRange(1, COLS.NEW_SHEET.NAME).setValue('Name');
  sheet.getRange(1, COLS.NEW_SHEET.VALUE).setValue('Value');
  sheet.getRange(1, COLS.NEW_SHEET.STATUS).setValue('Status');
}
```

### 3. Document in This Guide
Add the new constants to the "Available Constants by Workbook" section above.

## Helper Functions

### Convert Column Number to Letter
```javascript
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

// Usage
const col = COLS.TEMP_DIFF.NATURE;
const colLetter = columnToLetter(col); // Returns 'G'
```

### Get Column Range
```javascript
function getColumnRange(sheet, startRow, endRow, col) {
  return sheet.getRange(startRow, col, endRow - startRow + 1, 1);
}

// Usage
const natureRange = getColumnRange(
  sheet, 
  7, 
  50, 
  COLS.TEMP_DIFF.NATURE
);
natureRange.setHorizontalAlignment('center');
```

## Migration Checklist

When converting existing code to use constants:

- [ ] Identify all hardcoded column numbers
- [ ] Find or create appropriate constants
- [ ] Replace numbers with constants
- [ ] Test thoroughly
- [ ] Update documentation
- [ ] Commit changes

## Common Pitfalls

### 1. Off-by-One Errors
```javascript
// ❌ Wrong - columns are 1-indexed, not 0-indexed
const COLS = {
  TEMP_DIFF: {
    SR_NO: 0,  // Should be 1!
    ITEM: 1,   // Should be 2!
  }
};
```

### 2. Forgetting to Update Constants
```javascript
// If you add a column in the middle, update ALL subsequent columns!
const COLS = {
  TEMP_DIFF: {
    SR_NO: 1,
    ITEM: 2,
    NEW_COLUMN: 3,  // Added this
    CATEGORY: 4,    // Was 3, now 4
    TAX_BASE: 5,    // Was 4, now 5
    // ... update all following columns
  }
};
```

### 3. Using Wrong Sheet Constants
```javascript
// ❌ Wrong - using CONTRACT_REGISTER constants on REVENUE_RECOGNITION sheet
sheet.getRange(row, COLS.CONTRACT_REGISTER.CUSTOMER).setValue(name);

// ✅ Correct - use the right sheet's constants
sheet.getRange(row, COLS.REVENUE_RECOGNITION.CUSTOMER).setValue(name);
```

## Conclusion

Using column constants makes your code:
- **More readable**: Clear intent
- **Easier to maintain**: Change in one place
- **Less error-prone**: IDE autocomplete helps
- **Self-documenting**: No need to check sheet structure

Always use constants instead of magic numbers!
