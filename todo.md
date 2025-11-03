# Issues Found in Scripts

## deferredtax.gs - Deferred Tax Workings

### 1. Flawed Movement Analysis (CONFIRMED)
**Location**: `createMovementAnalysisSheet()` function, lines 1000-1004
**Issue**: The Movement_Analysis sheet uses hardcoded, arbitrary percentages (0.2, 0.3, 0.25) to distribute the total opening DTA/DTL balance among different categories. This creates fabricated numbers that don't reflect actual opening balances per line item.
**Impact**: Entire Movement_Analysis sheet is unreliable and would fail audit review.
**Fix Required**: Invert the data model - enter opening temporary differences in Temp_Differences sheet, use SUMIFS to aggregate dynamically.

### 2. Arbitrary Distribution of Additions and Reversals
**Location**: Same as above, lines 1000-1004
**Issue**: Same arbitrary percentages applied to additions and reversals, assuming fixed proportions that don't match actual transaction patterns.
**Impact**: Misleading movement analysis that doesn't tie to underlying transactions.
**Fix Required**: Link additions and reversals directly from Temp_Differences sheet using SUMIF formulas.

## indas116.gs - Ind AS 116 Lease Accounting

### 3. Material Inaccuracies in Lease Calculations (CONFIRMED)
**Location**: `createLeaseLiabilityScheduleSheet()` function, lines 1141-1166
**Issue**: 
- **EIR Method Approximation**: Uses average balance method ((Opening + Closing pre-interest)/2 * IBR * Months/12) instead of true period-by-period EIR calculation required by Ind AS 116 Para 36.
- **Current Portion Calculation**: Uses approximation (MIN(Closing, 12×Payment - H×IBR×0.5)) instead of true forward-looking amortization schedule.
**Impact**: Misstated P&L (Interest Expense) and Balance Sheet (Current/Non-Current Liability).
**Fix Required**: Implement month-by-month amortization schedule for each lease, sum relevant periods for current portion.

### 4. Interest Calculation Method
**Location**: Lines 1147-1156
**Issue**: While improved with average balance, still not true EIR method. For credit-impaired assets, uses net carrying amount which is correct, but overall method is simplified.
**Impact**: Interest expense may be materially inaccurate over long lease terms.
**Fix Required**: Create detailed period-by-period calculation engine.

## indas109.gs - Ind AS 109 Financial Instruments

### 5. Omissions in Complex Calculations (CONFIRMED)
**Location**: ECL_Impairment sheet, line 928; Amortization_Schedule, line 1110
**Issue**:
- **ECL Discounting Omission**: ECL calculated as EAD × PD × LGD without present value discounting required by Ind AS 109.B5.5.29 for lifetime ECL.
- **Simplified EIR Interest**: Uses average balance method over days instead of true EIR compounding.
**Impact**: ECL provision and interest income inaccurate, especially for Stage 2/3 assets.
**Fix Required**: Implement PV function for ECL discounting and period-by-period EIR calculations.

### 6. ECL Stage Determination Logic
**Location**: ECL_Impairment sheet, lines 920-922
**Issue**: Simplified approach uses lifetime ECL from day 1 for trade receivables, but logic may not properly handle the 12-month vs lifetime distinction for other instruments.
**Impact**: Potential over-provisioning for non-trade receivables.
**Fix Required**: Refine stage determination based on instrument type and risk assessment.

## far_wp.gs - Fixed Assets Audit Workpaper

### 7. Template Nature - No Accounting Issues
**Assessment**: This script creates audit workpaper templates with proper cross-references. No accounting calculation flaws found - it's a documentation template.

## ifc_p2p.gs - ICFR P2P Workpaper

### 8. Template Nature - No Accounting Issues
**Assessment**: This script creates ICFR testing templates with predefined controls. No accounting calculation flaws - it's a compliance testing framework.

## indas115.gs - Ind AS 115 Revenue Recognition

### 9. Template Nature - No Accounting Issues
**Assessment**: This script creates comprehensive revenue recognition workpaper templates with proper 5-step model structure. No accounting calculation flaws found - it's a documentation template with automated cross-references.

## General Issues Across Scripts

### 9. Formula Dependencies and Error Handling
**Issue**: Scripts rely heavily on sheet references that could break if sheet names change or if data is not entered in expected format.
**Impact**: Runtime errors or incorrect calculations.
**Fix Required**: Add error handling and validation in formulas.

### 10. Hardcoded Values
**Issue**: Various hardcoded values (percentages, thresholds) that should be configurable.
**Impact**: Reduced flexibility for different entities.
**Fix Required**: Make parameters configurable through input sheets.

### 11. Performance Considerations
**Issue**: Some scripts create large ranges with formulas that may slow down large spreadsheets.
**Impact**: Performance degradation in large workbooks.
**Fix Required**: Optimize formula ranges and use batch operations where possible.

## Priority Action Items

1. **HIGH**: Fix Movement Analysis in deferredtax.gs - fundamental flaw affecting audit reliability
2. **HIGH**: Implement proper EIR calculations in indas116.gs - material impact on financial statements  
3. **HIGH**: Add ECL discounting in indas109.gs - required by standard
4. **MEDIUM**: Improve error handling and validation across all scripts
5. **LOW**: Make hardcoded values configurable for better flexibility