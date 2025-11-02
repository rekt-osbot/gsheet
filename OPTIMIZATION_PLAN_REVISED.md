# Enhanced Optimization Plan for Google Apps Script Workbooks

## Analysis Summary
- **3 files to optimize**: indas109.gs (1909 lines), indas116.gs (1926 lines), deferredtax.gs (1968 lines)
- **18 for-loops identified** that use row-by-row operations
- **Primary bottleneck**: Individual API calls inside loops (setFormula, setBackground, formatInputCell, etc.)
- **Secondary issue**: Large ranges (:G1000) causing excessive recalculation overhead

## Optimization Strategy (Revised & Enhanced)

### Phase 1: Critical Fixes (Immediate Impact)
✅ **Already completed**: RANDBETWEEN removal (verified in indas109.gs:718)

### Phase 2: Batch Operations Refactoring (80-90% performance gain)

#### Core Pattern to Apply:
Instead of:
```javascript
for (let i = 0; i < 100; i++) {
  const row = 3 + i;
  const formulas = [/* ... */];
  formulas.forEach((formula, col) => {
    sheet.getRange(row, col + 1).setFormula(formula);  // 100 API calls per column!
  });
  formatInputCell(sheet.getRange(row, 10), '#ffebee');  // 100 more API calls!
}
```

Use:
```javascript
const numRows = 100;
const startRow = 3;
const numCols = 12;

// 1. Create 2D arrays for ALL attributes
const formulaArray = [];
const backgroundArray = Array(numRows).fill(null).map(() => Array(numCols).fill(null));

// 2. Populate arrays in memory (fast)
for (let i = 0; i < numRows; i++) {
  const row = startRow + i;
  formulaArray.push([
    `formula1 for row ${row}`,
    `formula2 for row ${row}`,
    // ... all formulas for this row
  ]);

  // Mark input cells in background array
  backgroundArray[i][9] = '#ffebee';  // Column J (index 9)
}

// 3. Write ONCE to sheet (1-2 API calls total!)
sheet.getRange(startRow, 1, numRows, numCols).setFormulas(formulaArray);
sheet.getRange(startRow, 1, numRows, numCols).setBackgrounds(backgroundArray);

// 4. Batch format entire column ranges
sheet.getRange(startRow, 4, numRows, 1).setNumberFormat('₹#,##0.00');
sheet.getRange(startRow, 6, numRows, 2).setNumberFormat('0.00%');
```

#### Benefits:
- **Before**: 100 rows × 13 cells × 2 calls (formula + format) = **2,600 API calls**
- **After**: 2 batch calls (formulas + backgrounds) + 3 format calls = **5 API calls**
- **Speed improvement**: ~520x faster for this section alone

### Phase 3: Advanced Optimizations

#### 3.1 Batch Number Formatting
Create format arrays for number formats, font weights, alignments:
```javascript
const numberFormatArray = Array(numRows).fill(null).map(() => Array(numCols).fill('@'));
for (let i = 0; i < numRows; i++) {
  numberFormatArray[i][3] = '₹#,##0.00';  // Column D
  numberFormatArray[i][5] = '0.00%';      // Column F
  // ... etc
}
sheet.getRange(startRow, 1, numRows, numCols).setNumberFormats(numberFormatArray);
```

#### 3.2 Minimize Range Lookups
```javascript
// Cache the main data range
const dataRange = sheet.getRange(startRow, 1, numRows, numCols);
dataRange.setFormulas(formulaArray);
dataRange.setBackgrounds(backgroundArray);
dataRange.setNumberFormats(numberFormatArray);
```

#### 3.3 Reduce Range Sizes
**Find and replace across all files:**
- `:G1000` → `:G250`
- `:E1000` → `:E250`
- `:P1000` → `:P250`
- etc.

Apply to:
- Formula ranges (SUMIF, COUNTIF, etc.)
- Conditional formatting rules
- Data validation ranges

### Phase 4: Function-Specific Optimizations

#### Functions requiring batch refactoring:

**indas109.gs:**
1. ✅ `createECLImpairmentSheet()` - lines 877-914 (100 row loop)
2. ✅ `createFairValueWorkingsSheet()` - lines 710-736 (100 row loop)
3. ✅ `createAmortizationScheduleSheet()` - needs review
4. ✅ `createInstrumentsRegisterSheet()` - needs review
5. ✅ `createClassificationMatrixSheet()` - needs review
6. ✅ Other create*Sheet functions with loops

**indas116.gs:** (5 loops to optimize)
**deferredtax.gs:** (6 loops to optimize)

## Implementation Order

1. **Start with indas109.gs** (most impactful, referenced in optimization doc)
   - createECLImpairmentSheet
   - createFairValueWorkingsSheet
   - createAmortizationScheduleSheet
   - Remaining functions

2. **Continue with indas116.gs**

3. **Finish with deferredtax.gs**

4. **Global range size reduction** (find/replace operation)

5. **Testing and validation**

## Expected Results

- **Script execution time**: Reduced by 80-90% (from ~60s to ~5-10s)
- **Sheet responsiveness**: Dramatically improved (no volatile functions, smaller ranges)
- **User experience**: Near-instant data entry and scrolling
- **Maintainability**: Cleaner, more understandable code structure

## Key Principles

1. **Never call sheet methods inside loops** unless absolutely necessary
2. **Build arrays in JavaScript, write to sheet once**
3. **Batch ALL operations**: formulas, values, formats, backgrounds, borders
4. **Cache range references** when used multiple times
5. **Use practical range sizes** (100-250 rows, not 1000)
