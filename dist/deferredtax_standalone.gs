/**
 * @name deferredtax
 * @version 1.0.1
 * @built 2025-11-03T12:27:00.894Z
 * @description Standalone script. Do not edit directly - edit source files in src/ folder.
 * 
 * This file is auto-generated from:
 * - Common utilities (src/common/*.gs)
 * - Workbook-specific code (src/workbooks/deferredtax.gs)
 * 
 * To make changes:
 * 1. Edit source files in src/ folder
 * 2. Run: npm run build
 * 3. Copy the generated file from dist/ folder to Google Apps Script
 */

/**
 * ════════════════════════════════════════════════════════════════════════════
 * COMMON UTILITY FUNCTIONS
 * ════════════════════════════════════════════════════════════════════════════
 * Shared utility functions used across all workbooks
 */

// ============================================================================
// SHEET MANAGEMENT
// ============================================================================

function clearExistingSheets(ss) {
  const sheets = ss.getSheets();
  
  // If there's only one sheet, just clear it instead of deleting
  if (sheets.length === 1) {
    sheets[0].clear();
    sheets[0].setName('_temp_sheet_');
    return;
  }
  
  // Keep at least one sheet - delete all except the last one
  for (let i = sheets.length - 1; i >= 0; i--) {
    if (sheets.length > 1) {  // Always keep at least one sheet
      ss.deleteSheet(sheets[i]);
    }
  }
  
  // Rename the remaining sheet to a temporary name
  if (ss.getSheets().length === 1) {
    ss.getSheets()[0].setName('_temp_sheet_');
  }
}

// ============================================================================
// MENU CREATION
// ============================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Try to get workbook type from script properties first (most reliable)
  const scriptProps = PropertiesService.getScriptProperties();
  let workbookType = scriptProps.getProperty('WORKBOOK_TYPE');
  
  // Fallback: detect from spreadsheet name if property not set
  if (!workbookType) {
    const sheetName = ss.getName();
    if (sheetName.includes('Deferred Tax') || sheetName.includes('DT')) {
      workbookType = 'DEFERRED_TAX';
    } else if (sheetName.includes('109')) {
      workbookType = 'INDAS109';
    } else if (sheetName.includes('115')) {
      workbookType = 'INDAS115';
    } else if (sheetName.includes('116')) {
      workbookType = 'INDAS116';
    } else if (sheetName.includes('Fixed Asset') || sheetName.includes('FAR')) {
      workbookType = 'FIXED_ASSETS';
    } else if (sheetName.includes('TDS')) {
      workbookType = 'TDS_COMPLIANCE';
    } else if (sheetName.includes('P2P') || sheetName.includes('ICFR')) {
      workbookType = 'ICFR_P2P';
    }
  }
  
  // Map workbook types to menu configurations
  const workbookConfig = {
    'DEFERRED_TAX': { menuName: 'Deferred Tax Tools', functionName: 'createDeferredTaxWorkbook' },
    'INDAS109': { menuName: 'Ind AS 109 Tools', functionName: 'createIndAS109WorkingPapers' },
    'INDAS115': { menuName: 'Ind AS 115 Tools', functionName: 'buildIndAS115Workpaper' },
    'INDAS116': { menuName: 'Ind AS 116 Tools', functionName: 'createIndAS116Workbook' },
    'FIXED_ASSETS': { menuName: 'Fixed Assets Tools', functionName: 'setupFixedAssetsWorkpaper' },
    'TDS_COMPLIANCE': { menuName: 'TDS Tools', functionName: 'createTDSWorkbook' },
    'ICFR_P2P': { menuName: 'ICFR Tools', functionName: 'createICFRP2PWorkbook' }
  };
  
  const config = workbookConfig[workbookType] || { menuName: 'Audit Tools', functionName: 'createWorkbook' };
  
  ui.createMenu(config.menuName)
    .addItem('Create/Refresh Workbook', config.functionName)
    .addSeparator()
    .addItem('About', 'showAbout')
    .addToUi();
}

/**
 * Set the workbook type property - call this from each workbook creation function
 */
function setWorkbookType(type) {
  PropertiesService.getScriptProperties().setProperty('WORKBOOK_TYPE', type);
}

function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'IGAAP-Ind AS Audit Builder',
    'Automated audit workpaper generation tool\n\n' +
    'Version: 1.0\n' +
    'Created: November 2025\n\n' +
    'This tool generates comprehensive audit workpapers compliant with ' +
    'Indian Accounting Standards (Ind AS) and IGAAP.',
    ui.ButtonSet.OK
  );
}


/**
 * ════════════════════════════════════════════════════════════════════════════
 * COMMON FORMATTING UTILITIES
 * ════════════════════════════════════════════════════════════════════════════
 * Shared color schemes and formatting functions used across all workbooks
 */

// ============================================================================
// COLOR SCHEME
// ============================================================================

const COLORS = {
  HEADER_BG: "#1a237e",           // Dark blue
  HEADER_TEXT: "#ffffff",          // White
  SUBHEADER_BG: "#3949ab",        // Medium blue
  INPUT_BG: "#fff9c4",            // Light yellow
  INPUT_ALT_BG: "#b3e5fc",        // Light blue
  CALC_BG: "#e8eaf6",             // Light purple-grey
  SECTION_BG: "#c5cae9",          // Light blue-grey
  TOTAL_BG: "#ffccbc",            // Light orange
  GRAND_TOTAL_BG: "#ff8a65",      // Orange
  WARNING_BG: "#ffebee",          // Light red
  SUCCESS_BG: "#c8e6c9",          // Light green
  INFO_BG: "#e1f5fe",             // Very light blue
  BORDER_COLOR: "#757575"         // Grey
};

const FONT_SIZES = {
  title: 14,
  header: 11,
  normal: 10,
  small: 9
};

// ============================================================================
// FORMATTING HELPER FUNCTIONS
// ============================================================================

function formatHeader(sheet, row, startCol, endCol, text, bgColor = '#1a237e') {
  const range = sheet.getRange(row, startCol, 1, endCol - startCol + 1);
  range.merge()
       .setValue(text)
       .setBackground(bgColor)
       .setFontColor('#ffffff')
       .setFontWeight('bold')
       .setFontSize(12)
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle');
  sheet.setRowHeight(row, 35);
}

function formatSubHeader(sheet, row, startCol, values, bgColor = '#283593') {
  values.forEach((value, index) => {
    sheet.getRange(row, startCol + index)
         .setValue(value)
         .setBackground(bgColor)
         .setFontColor('#ffffff')
         .setFontWeight('bold')
         .setHorizontalAlignment('center')
         .setVerticalAlignment('middle')
         .setWrap(true);
  });
  sheet.setRowHeight(row, 30);
}

function formatInputCell(range, bgColor = '#e3f2fd') {
  range.setBackground(bgColor)
       .setBorder(true, true, true, true, true, true, '#1976d2', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function formatCurrency(range) {
  range.setNumberFormat('₹#,##0.00');
}

function formatPercentage(range) {
  range.setNumberFormat('0.00%');
}

function formatDate(range) {
  range.setNumberFormat('dd-mmm-yyyy');
}

function setColumnWidths(sheet, widths) {
  widths.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });
}

function protectSheet(sheet, warningOnly = true) {
  const protection = sheet.protect();
  if (warningOnly) {
    protection.setWarningOnly(true);
  }
}


/**
 * ════════════════════════════════════════════════════════════════════════════
 * DATA VALIDATION HELPERS
 * ════════════════════════════════════════════════════════════════════════════
 * Common data validation builders to ensure consistency
 */

// ============================================================================
// COMMON VALIDATION LISTS
// ============================================================================

const VALIDATION_LISTS = {
  YES_NO: ['Yes', 'No'],
  YES_NO_NA: ['Yes', 'No', 'N/A'],
  PASS_FAIL: ['Pass', 'Fail'],
  PASS_FAIL_NOTE: ['Pass', 'Fail', 'Note'],
  PASS_FAIL_NA: ['Pass', 'Fail', 'N/A'],
  CHECK_MARKS: ['✓', '✗', 'N/A'],
  STATUS_ACTIVE: ['Active', 'Inactive', 'Pending'],
  STATUS_COMPLETE: ['Complete', 'In Progress', 'Not Started', 'N/A'],
  STATUS_OPEN: ['Open', 'Closed', 'Pending', 'Noted'],
  EFFECTIVENESS: ['Effective', 'Partially Effective', 'Ineffective', 'Not Tested'],
  OPERATING_EFFECTIVENESS: ['Operating Effectively', 'Operating Partially', 'Not Operating Effectively', 'Not Tested'],
  CONDITION_PHYSICAL: ['Good', 'Fair', 'Poor', 'N/A'],
  LOCATION_STATUS: ['✓ Yes', '✗ No', 'Unable to locate'],
  
  // Ind AS specific
  SPPI_TEST: ['Pass', 'Fail', 'Not Applicable'],
  BUSINESS_MODEL: ['Hold to Collect', 'Hold to Collect & Sell', 'Other (Trading)'],
  CREDIT_RATING: ['AAA', 'AA+', 'AA', 'AA-', 'A+', 'A', 'A-', 'BBB+', 'BBB', 'BBB-', 'BB', 'B', 'C', 'D', 'Not Rated'],
  SECURITY_TYPE: ['Secured', 'Unsecured', 'Equity', 'Sovereign', 'Units', 'Not Applicable'],
  INSTRUMENT_TYPE: ['Loan', 'Bond', 'Debenture', 'Equity', 'Mutual Fund', 'G-Sec', 'T-Bill', 'Receivable', 'Derivative', 'Other'],
  COUPON_FREQUENCY: ['Annual', 'Semi-Annual', 'Quarterly', 'Monthly', 'Not Applicable'],
  PAYMENT_FREQUENCY: ['Monthly', 'Quarterly', 'Half-Yearly', 'Annual'],
  REVENUE_PATTERN: ['Point in Time', 'Over Time', 'Mixed'],
  CONTRACT_STATUS: ['Active', 'Completed', 'Terminated', 'On Hold'],
  LEASE_CATEGORY: ['Property', 'Vehicles', 'Equipment', 'IT Equipment', 'Other'],
  EXEMPTION_TYPE: ['No', 'Low Value', 'Short-term', 'Both'],
  
  // TDS specific
  ENTITY_TYPE: ['Company', 'Individual', 'HUF', 'Firm', 'AOP/BOI', 'Trust', 'Government', 'Non-Resident'],
  TDS_QUARTER: ['Q1', 'Q2', 'Q3', 'Q4'],
  TDS_RETURN_TYPE: ['24Q (Salary)', '26Q (Non-Salary)'],
  PAYMENT_STATUS: ['Paid', 'Pending', 'Late Payment'],
  RECONCILIATION_STATUS: ['Matched', 'Variance', 'Pending'],
  
  // ICFR specific
  RISK_CATEGORY: ['High', 'Medium', 'Low'],
  CONTROL_TYPE: ['Preventive', 'Detective', 'Corrective'],
  CONTROL_FREQUENCY: ['Each Transaction', 'Daily', 'Weekly', 'Monthly', 'Quarterly', 'Annually'],
  TOD_TOE_STATUS: ['PASS', 'REVIEW', 'FAIL'],
  
  // Fixed Assets specific
  REPAIR_TYPE: ['Repair', 'Improvement', 'Betterment', 'Other'],
  CAPITALIZATION_DECISION: ['Capitalize', 'Expense', 'Review Required']
};

// ============================================================================
// VALIDATION BUILDERS
// ============================================================================

/**
 * Create a simple dropdown validation from a list
 * @param {Array<string>} values - List of valid values
 * @param {boolean} allowInvalid - Whether to allow invalid values (default: false)
 * @param {string} helpText - Optional help text
 * @returns {DataValidation}
 */
function createDropdownValidation(values, allowInvalid = false, helpText = '') {
  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(allowInvalid);
  
  if (helpText) {
    validation.setHelpText(helpText);
  }
  
  return validation.build();
}

/**
 * Create a Yes/No dropdown validation
 * @param {string} helpText - Optional help text
 * @returns {DataValidation}
 */
function createYesNoValidation(helpText = '') {
  return createDropdownValidation(VALIDATION_LISTS.YES_NO, false, helpText);
}

/**
 * Create a date validation
 * @param {string} helpText - Optional help text
 * @returns {DataValidation}
 */
function createDateValidation(helpText = 'Enter date in DD-MMM-YYYY format') {
  return SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText(helpText)
    .build();
}

/**
 * Create a number validation with optional range
 * @param {number} min - Minimum value (optional)
 * @param {number} max - Maximum value (optional)
 * @param {string} helpText - Optional help text
 * @returns {DataValidation}
 */
function createNumberValidation(min = null, max = null, helpText = '') {
  let validation = SpreadsheetApp.newDataValidation();
  
  if (min !== null && max !== null) {
    validation = validation.requireNumberBetween(min, max);
    helpText = helpText || `Enter a number between ${min} and ${max}`;
  } else if (min !== null) {
    validation = validation.requireNumberGreaterThanOrEqualTo(min);
    helpText = helpText || `Enter a number >= ${min}`;
  } else if (max !== null) {
    validation = validation.requireNumberLessThanOrEqualTo(max);
    helpText = helpText || `Enter a number <= ${max}`;
  } else {
    validation = validation.requireNumberBetween(-999999999, 999999999);
  }
  
  return validation
    .setAllowInvalid(false)
    .setHelpText(helpText)
    .build();
}

/**
 * Create a percentage validation (0-100%)
 * @param {string} helpText - Optional help text
 * @returns {DataValidation}
 */
function createPercentageValidation(helpText = 'Enter percentage (0-100%)') {
  return createNumberValidation(0, 1, helpText);
}

/**
 * Apply validation to a range using a predefined list
 * @param {Range} range - The range to apply validation to
 * @param {string} listName - Name of the validation list from VALIDATION_LISTS
 * @param {string} helpText - Optional help text
 */
function applyValidationList(range, listName, helpText = '') {
  const values = VALIDATION_LISTS[listName];
  if (!values) {
    Logger.log('Warning: Validation list "' + listName + '" not found');
    return;
  }
  
  const validation = createDropdownValidation(values, false, helpText);
  range.setDataValidation(validation);
}

/**
 * Apply multiple validations at once
 * @param {Sheet} sheet - The sheet to apply validations to
 * @param {Array<Object>} validations - Array of {range, type, options}
 * 
 * Example:
 * applyMultipleValidations(sheet, [
 *   {range: 'B5:B10', type: 'YES_NO'},
 *   {range: 'C5:C10', type: 'date'},
 *   {range: 'D5:D10', type: 'number', min: 0, max: 100}
 * ]);
 */
function applyMultipleValidations(sheet, validations) {
  validations.forEach(v => {
    const range = typeof v.range === 'string' ? sheet.getRange(v.range) : v.range;
    
    if (VALIDATION_LISTS[v.type]) {
      // Predefined list
      applyValidationList(range, v.type, v.helpText || '');
    } else if (v.type === 'date') {
      range.setDataValidation(createDateValidation(v.helpText || ''));
    } else if (v.type === 'number') {
      range.setDataValidation(createNumberValidation(v.min, v.max, v.helpText || ''));
    } else if (v.type === 'percentage') {
      range.setDataValidation(createPercentageValidation(v.helpText || ''));
    } else if (v.type === 'custom' && v.values) {
      range.setDataValidation(createDropdownValidation(v.values, v.allowInvalid || false, v.helpText || ''));
    }
  });
}

// ============================================================================
// COMMON VALIDATION PATTERNS
// ============================================================================

/**
 * Apply standard Ind AS 109 instrument validations
 * @param {Sheet} sheet - The sheet to apply validations to
 * @param {string} startRow - Starting row (e.g., '3')
 * @param {string} endRow - Ending row (e.g., '250')
 */
function applyIndAS109Validations(sheet, startRow, endRow) {
  applyMultipleValidations(sheet, [
    {range: `C${startRow}:C${endRow}`, type: 'INSTRUMENT_TYPE'},
    {range: `L${startRow}:L${endRow}`, type: 'SECURITY_TYPE'},
    {range: `M${startRow}:M${endRow}`, type: 'CREDIT_RATING'},
    {range: `O${startRow}:O${endRow}`, type: 'SPPI_TEST'},
    {range: `P${startRow}:P${endRow}`, type: 'BUSINESS_MODEL'},
    {range: `Q${startRow}:Q${endRow}`, type: 'YES_NO'},
    {range: `R${startRow}:R${endRow}`, type: 'YES_NO'},
    {range: `S${startRow}:S${endRow}`, type: 'COUPON_FREQUENCY'},
    {range: `T${startRow}:T${endRow}`, type: 'YES_NO_NA'}
  ]);
}

/**
 * Apply standard Ind AS 115 contract validations
 * @param {Sheet} sheet - The sheet to apply validations to
 * @param {string} startRow - Starting row
 * @param {string} endRow - Ending row
 */
function applyIndAS115Validations(sheet, startRow, endRow) {
  applyMultipleValidations(sheet, [
    {range: `L${startRow}:L${endRow}`, type: 'REVENUE_PATTERN'},
    {range: `N${startRow}:N${endRow}`, type: 'CONTRACT_STATUS'}
  ]);
}

/**
 * Apply standard Ind AS 116 lease validations
 * @param {Sheet} sheet - The sheet to apply validations to
 * @param {string} startRow - Starting row
 * @param {string} endRow - Ending row
 */
function applyIndAS116Validations(sheet, startRow, endRow) {
  applyMultipleValidations(sheet, [
    {range: `D${startRow}:D${endRow}`, type: 'LEASE_CATEGORY'},
    {range: `I${startRow}:I${endRow}`, type: 'PAYMENT_FREQUENCY'},
    {range: `J${startRow}:J${endRow}`, type: 'EXEMPTION_TYPE'}
  ]);
}

/**
 * Apply standard TDS validations
 * @param {Sheet} sheet - The sheet to apply validations to
 * @param {string} startRow - Starting row
 * @param {string} endRow - Ending row
 */
function applyTDSValidations(sheet, startRow, endRow) {
  applyMultipleValidations(sheet, [
    {range: `E${startRow}:E${endRow}`, type: 'ENTITY_TYPE'},
    {range: `K${startRow}:K${endRow}`, type: 'YES_NO'}
  ]);
}


/**
 * ════════════════════════════════════════════════════════════════════════════
 * CONDITIONAL FORMATTING HELPERS
 * ════════════════════════════════════════════════════════════════════════════
 * Common conditional formatting rules for consistency
 */

// ============================================================================
// COLOR DEFINITIONS FOR CONDITIONAL FORMATTING
// ============================================================================

const CF_COLORS = {
  GREEN: '#c8e6c9',      // Success/Pass/Positive
  LIGHT_GREEN: '#d9ead3',
  YELLOW: '#fff9c4',     // Warning/Review/Partial
  LIGHT_YELLOW: '#fff3cd',
  RED: '#ffcdd2',        // Error/Fail/Negative
  LIGHT_RED: '#f4cccc',
  DARK_RED: '#cc0000',
  BLUE: '#bbdefb',       // Info
  LIGHT_BLUE: '#e1f5fe',
  PURPLE: '#e1bee7',     // Special
  ORANGE: '#ffe0b2'      // Alert
};

// ============================================================================
// RULE BUILDERS
// ============================================================================

/**
 * Create a text equals rule
 * @param {string} text - Text to match
 * @param {string} backgroundColor - Background color
 * @param {string} fontColor - Font color (optional)
 * @returns {ConditionalFormatRule}
 */
function createTextEqualsRule(text, backgroundColor, fontColor = null) {
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(text)
    .setBackground(backgroundColor);
  
  if (fontColor) {
    rule.setFontColor(fontColor);
  }
  
  return rule.build();
}

/**
 * Create a text contains rule
 * @param {string} text - Text to search for
 * @param {string} backgroundColor - Background color
 * @param {string} fontColor - Font color (optional)
 * @returns {ConditionalFormatRule}
 */
function createTextContainsRule(text, backgroundColor, fontColor = null) {
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains(text)
    .setBackground(backgroundColor);
  
  if (fontColor) {
    rule.setFontColor(fontColor);
  }
  
  return rule.build();
}

/**
 * Create a number comparison rule
 * @param {string} comparison - 'greater', 'less', 'equal', 'notEqual', 'between', 'notBetween'
 * @param {number} value1 - First value
 * @param {number} value2 - Second value (for between/notBetween)
 * @param {string} backgroundColor - Background color
 * @param {string} fontColor - Font color (optional)
 * @returns {ConditionalFormatRule}
 */
function createNumberRule(comparison, value1, value2, backgroundColor, fontColor = null) {
  let rule = SpreadsheetApp.newConditionalFormatRule();
  
  switch(comparison) {
    case 'greater':
      rule = rule.whenNumberGreaterThan(value1);
      break;
    case 'less':
      rule = rule.whenNumberLessThan(value1);
      break;
    case 'equal':
      rule = rule.whenNumberEqualTo(value1);
      break;
    case 'notEqual':
      rule = rule.whenNumberNotEqualTo(value1);
      break;
    case 'between':
      rule = rule.whenNumberBetween(value1, value2);
      break;
    case 'notBetween':
      rule = rule.whenNumberNotBetween(value1, value2);
      break;
  }
  
  rule = rule.setBackground(backgroundColor);
  
  if (fontColor) {
    rule.setFontColor(fontColor);
  }
  
  return rule.build();
}

/**
 * Create a formula-based rule
 * @param {string} formula - Formula to evaluate (e.g., '=$A1>0')
 * @param {string} backgroundColor - Background color
 * @param {string} fontColor - Font color (optional)
 * @returns {ConditionalFormatRule}
 */
function createFormulaRule(formula, backgroundColor, fontColor = null) {
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setBackground(backgroundColor);
  
  if (fontColor) {
    rule.setFontColor(fontColor);
  }
  
  return rule.build();
}

// ============================================================================
// COMMON RULE SETS
// ============================================================================

/**
 * Apply Pass/Fail conditional formatting
 * @param {Range} range - Range to apply formatting to
 * @returns {Array<ConditionalFormatRule>}
 */
function createPassFailRules(range) {
  const passRule = createTextContainsRule('Pass', CF_COLORS.GREEN);
  const failRule = createTextContainsRule('FAIL', CF_COLORS.RED);
  
  passRule.setRanges([range]);
  failRule.setRanges([range]);
  
  return [passRule, failRule];
}

/**
 * Apply Pass/Review/Fail conditional formatting
 * @param {Range} range - Range to apply formatting to
 * @returns {Array<ConditionalFormatRule>}
 */
function createPassReviewFailRules(range) {
  const passRule = createTextEqualsRule('PASS', CF_COLORS.GREEN);
  const reviewRule = createTextEqualsRule('REVIEW', CF_COLORS.YELLOW);
  const failRule = createTextEqualsRule('FAIL', CF_COLORS.RED);
  
  passRule.setRanges([range]);
  reviewRule.setRanges([range]);
  failRule.setRanges([range]);
  
  return [passRule, reviewRule, failRule];
}

/**
 * Apply status conditional formatting (Active/Inactive/Pending)
 * @param {Range} range - Range to apply formatting to
 * @returns {Array<ConditionalFormatRule>}
 */
function createStatusRules(range) {
  const activeRule = createTextEqualsRule('Active', CF_COLORS.LIGHT_GREEN);
  const pendingRule = createTextEqualsRule('Pending', CF_COLORS.LIGHT_YELLOW);
  const inactiveRule = createTextEqualsRule('Inactive', CF_COLORS.LIGHT_RED);
  
  activeRule.setRanges([range]);
  pendingRule.setRanges([range]);
  inactiveRule.setRanges([range]);
  
  return [activeRule, pendingRule, inactiveRule];
}

/**
 * Apply effectiveness conditional formatting
 * @param {Range} range - Range to apply formatting to
 * @returns {Array<ConditionalFormatRule>}
 */
function createEffectivenessRules(range) {
  const effectiveRule = createTextEqualsRule('Effective', CF_COLORS.GREEN);
  const partialRule = createTextEqualsRule('Partially Effective', CF_COLORS.YELLOW);
  const ineffectiveRule = createTextEqualsRule('Ineffective', CF_COLORS.RED);
  
  effectiveRule.setRanges([range]);
  partialRule.setRanges([range]);
  ineffectiveRule.setRanges([range]);
  
  return [effectiveRule, partialRule, ineffectiveRule];
}

/**
 * Apply operating effectiveness conditional formatting
 * @param {Range} range - Range to apply formatting to
 * @returns {Array<ConditionalFormatRule>}
 */
function createOperatingEffectivenessRules(range) {
  const effectiveRule = createTextEqualsRule('Operating Effectively', CF_COLORS.GREEN);
  const partialRule = createTextEqualsRule('Operating Partially', CF_COLORS.YELLOW);
  const notEffectiveRule = createTextEqualsRule('Not Operating Effectively', CF_COLORS.RED);
  
  effectiveRule.setRanges([range]);
  partialRule.setRanges([range]);
  notEffectiveRule.setRanges([range]);
  
  return [effectiveRule, partialRule, notEffectiveRule];
}

/**
 * Apply positive/negative number formatting
 * @param {Range} range - Range to apply formatting to
 * @returns {Array<ConditionalFormatRule>}
 */
function createPositiveNegativeRules(range) {
  const positiveRule = createNumberRule('greater', 0, null, CF_COLORS.GREEN);
  const negativeRule = createNumberRule('less', 0, null, CF_COLORS.RED);
  
  positiveRule.setRanges([range]);
  negativeRule.setRanges([range]);
  
  return [positiveRule, negativeRule];
}

/**
 * Apply variance highlighting (non-zero values)
 * @param {Range} range - Range to apply formatting to
 * @param {number} tolerance - Tolerance for variance (default: 0)
 * @returns {Array<ConditionalFormatRule>}
 */
function createVarianceRules(range, tolerance = 0) {
  const varianceRule = createNumberRule('notBetween', -tolerance, tolerance, CF_COLORS.RED);
  varianceRule.setRanges([range]);
  
  return [varianceRule];
}

/**
 * Apply balance check formatting (should be zero)
 * @param {Range} range - Range to apply formatting to
 * @param {number} tolerance - Tolerance (default: 0.01)
 * @returns {Array<ConditionalFormatRule>}
 */
function createBalanceCheckRules(range, tolerance = 0.01) {
  const balancedRule = createNumberRule('between', -tolerance, tolerance, CF_COLORS.GREEN);
  const unbalancedRule = createNumberRule('notBetween', -tolerance, tolerance, CF_COLORS.DARK_RED, '#ffffff');
  
  balancedRule.setRanges([range]);
  unbalancedRule.setRanges([range]);
  
  return [balancedRule, unbalancedRule];
}

// ============================================================================
// IND AS SPECIFIC RULES
// ============================================================================

/**
 * Apply Ind AS 109 classification color coding
 * @param {Range} range - Range to apply formatting to
 * @returns {Array<ConditionalFormatRule>}
 */
function createIndAS109ClassificationRules(range) {
  const acRule = createTextEqualsRule('Amortized Cost', CF_COLORS.GREEN);
  const fvociRule = createTextEqualsRule('FVOCI', CF_COLORS.BLUE);
  const fvtplRule = createTextEqualsRule('FVTPL', CF_COLORS.ORANGE);
  
  acRule.setRanges([range]);
  fvociRule.setRanges([range]);
  fvtplRule.setRanges([range]);
  
  return [acRule, fvociRule, fvtplRule];
}

/**
 * Apply Ind AS 109 ECL stage color coding
 * @param {Range} range - Range to apply formatting to
 * @returns {Array<ConditionalFormatRule>}
 */
function createECLStageRules(range) {
  const stage1Rule = createTextEqualsRule('Stage 1', CF_COLORS.GREEN);
  const stage2Rule = createTextEqualsRule('Stage 2', CF_COLORS.YELLOW);
  const stage3Rule = createTextEqualsRule('Stage 3', CF_COLORS.RED);
  const simplifiedRule = createTextEqualsRule('Simplified (Lifetime)', CF_COLORS.PURPLE);
  
  stage1Rule.setRanges([range]);
  stage2Rule.setRanges([range]);
  stage3Rule.setRanges([range]);
  simplifiedRule.setRanges([range]);
  
  return [stage1Rule, stage2Rule, stage3Rule, simplifiedRule];
}

/**
 * Apply TDS payment status color coding
 * @param {Range} range - Range to apply formatting to
 * @returns {Array<ConditionalFormatRule>}
 */
function createTDSPaymentStatusRules(range) {
  const paidRule = createTextEqualsRule('Paid', CF_COLORS.LIGHT_GREEN);
  const pendingRule = createTextEqualsRule('Pending', CF_COLORS.LIGHT_YELLOW);
  const lateRule = createTextEqualsRule('Late Payment', CF_COLORS.LIGHT_RED);
  
  paidRule.setRanges([range]);
  pendingRule.setRanges([range]);
  lateRule.setRanges([range]);
  
  return [paidRule, pendingRule, lateRule];
}

/**
 * Apply TDS reconciliation status color coding
 * @param {Range} range - Range to apply formatting to
 * @returns {Array<ConditionalFormatRule>}
 */
function createTDSReconciliationRules(range) {
  const matchedRule = createTextEqualsRule('Matched', CF_COLORS.LIGHT_GREEN);
  const varianceRule = createTextEqualsRule('Variance', CF_COLORS.LIGHT_RED);
  
  matchedRule.setRanges([range]);
  varianceRule.setRanges([range]);
  
  return [matchedRule, varianceRule];
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Apply multiple conditional format rules to a sheet
 * @param {Sheet} sheet - The sheet to apply rules to
 * @param {Array<ConditionalFormatRule>} newRules - Array of rules to add
 * @param {boolean} clearExisting - Whether to clear existing rules (default: false)
 */
function applyConditionalFormatRules(sheet, newRules, clearExisting = false) {
  const existingRules = clearExisting ? [] : sheet.getConditionalFormatRules();
  sheet.setConditionalFormatRules(existingRules.concat(newRules));
}

/**
 * Apply standard audit workpaper conditional formatting
 * @param {Sheet} sheet - The sheet to apply formatting to
 * @param {Object} config - Configuration object with ranges
 * 
 * Example config:
 * {
 *   passFailRange: 'E5:E50',
 *   statusRange: 'F5:F50',
 *   varianceRange: 'G5:G50'
 * }
 */
function applyStandardAuditFormatting(sheet, config) {
  const rules = [];
  
  if (config.passFailRange) {
    const range = sheet.getRange(config.passFailRange);
    rules.push(...createPassFailRules(range));
  }
  
  if (config.statusRange) {
    const range = sheet.getRange(config.statusRange);
    rules.push(...createStatusRules(range));
  }
  
  if (config.varianceRange) {
    const range = sheet.getRange(config.varianceRange);
    rules.push(...createVarianceRules(range, config.varianceTolerance || 0));
  }
  
  if (config.balanceRange) {
    const range = sheet.getRange(config.balanceRange);
    rules.push(...createBalanceCheckRules(range, config.balanceTolerance || 0.01));
  }
  
  applyConditionalFormatRules(sheet, rules);
}


/**
 * ════════════════════════════════════════════════════════════════════════════
 * SHEET BUILDER HELPERS
 * ════════════════════════════════════════════════════════════════════════════
 * Common patterns for building sheets
 */

// ============================================================================
// SHEET CREATION HELPERS
// ============================================================================

/**
 * Get or create a sheet with a given name
 * @param {Spreadsheet} ss - The spreadsheet
 * @param {string} name - Sheet name
 * @param {number} index - Position index (optional)
 * @param {string} tabColor - Tab color (optional)
 * @param {boolean} clearIfExists - Clear existing sheet (default: true)
 * @returns {Sheet}
 */
function getOrCreateSheet(ss, name, index = null, tabColor = null, clearIfExists = true) {
  let sheet = ss.getSheetByName(name);
  
  if (sheet) {
    if (clearIfExists) {
      sheet.clear();
      sheet.clearConditionalFormatRules();
    }
  } else {
    if (index !== null) {
      sheet = ss.insertSheet(name, index);
    } else {
      sheet = ss.insertSheet(name);
    }
  }
  
  if (tabColor) {
    sheet.setTabColor(tabColor);
  }
  
  return sheet;
}

/**
 * Create a standard header section
 * @param {Sheet} sheet - The sheet
 * @param {string} title - Main title
 * @param {string} subtitle - Subtitle (optional)
 * @param {number} startCol - Starting column (default: 1)
 * @param {number} endCol - Ending column (default: 5)
 */
function createStandardHeader(sheet, title, subtitle = '', startCol = 1, endCol = 5) {
  // Main title
  formatHeader(sheet, 1, startCol, endCol, title, COLORS.HEADER_BG);
  
  // Subtitle if provided
  if (subtitle) {
    const range = sheet.getRange(2, startCol, 1, endCol - startCol + 1);
    range.merge()
         .setValue(subtitle)
         .setBackground(COLORS.SUBHEADER_BG)
         .setFontColor(COLORS.HEADER_TEXT)
         .setFontSize(10)
         .setFontStyle('italic')
         .setHorizontalAlignment('center')
         .setWrap(true);
    sheet.setRowHeight(2, 25);
  }
}

/**
 * Create a section header
 * @param {Sheet} sheet - The sheet
 * @param {number} row - Row number
 * @param {string} title - Section title
 * @param {number} startCol - Starting column (default: 1)
 * @param {number} endCol - Ending column (default: 5)
 */
function createSectionHeader(sheet, row, title, startCol = 1, endCol = 5) {
  const range = sheet.getRange(row, startCol, 1, endCol - startCol + 1);
  range.merge()
       .setValue(title)
       .setBackground(COLORS.SECTION_BG)
       .setFontWeight('bold')
       .setFontSize(11)
       .setHorizontalAlignment('center');
  sheet.setRowHeight(row, 25);
}

/**
 * Create a data table with headers
 * @param {Sheet} sheet - The sheet
 * @param {number} startRow - Starting row
 * @param {number} startCol - Starting column
 * @param {Array<string>} headers - Column headers
 * @param {Array<Array>} data - Data rows (optional)
 * @param {Object} options - Additional options
 * @returns {Object} - {headerRange, dataRange}
 */
function createDataTable(sheet, startRow, startCol, headers, data = [], options = {}) {
  const numCols = headers.length;
  
  // Create headers
  const headerRange = sheet.getRange(startRow, startCol, 1, numCols);
  headerRange.setValues([headers])
             .setBackground(options.headerBg || COLORS.HEADER_BG)
             .setFontColor(options.headerColor || COLORS.HEADER_TEXT)
             .setFontWeight('bold')
             .setHorizontalAlignment('center')
             .setWrap(true);
  
  if (options.headerHeight) {
    sheet.setRowHeight(startRow, options.headerHeight);
  }
  
  // Add data if provided
  let dataRange = null;
  if (data.length > 0) {
    dataRange = sheet.getRange(startRow + 1, startCol, data.length, numCols);
    dataRange.setValues(data);
  }
  
  // Apply borders if requested
  if (options.borders !== false) {
    const borderRange = sheet.getRange(
      startRow, 
      startCol, 
      (data.length || 1) + 1, 
      numCols
    );
    borderRange.setBorder(
      true, true, true, true, true, true,
      COLORS.BORDER_COLOR,
      SpreadsheetApp.BorderStyle.SOLID
    );
  }
  
  return {
    headerRange: headerRange,
    dataRange: dataRange
  };
}

/**
 * Create an input section with labels and input cells
 * @param {Sheet} sheet - The sheet
 * @param {number} startRow - Starting row
 * @param {number} labelCol - Label column
 * @param {number} inputCol - Input column
 * @param {Array<Object>} inputs - Array of {label, value, type, note}
 * @returns {number} - Next available row
 */
function createInputSection(sheet, startRow, labelCol, inputCol, inputs) {
  let row = startRow;
  
  inputs.forEach(input => {
    // Label
    sheet.getRange(row, labelCol)
         .setValue(input.label)
         .setFontWeight('bold');
    
    // Input cell
    const inputCell = sheet.getRange(row, inputCol);
    inputCell.setBackground(input.bgColor || COLORS.INPUT_BG);
    
    if (input.value !== undefined) {
      inputCell.setValue(input.value);
    }
    
    // Apply formatting based on type
    if (input.type === 'currency') {
      formatCurrency(inputCell);
    } else if (input.type === 'percentage') {
      formatPercentage(inputCell);
    } else if (input.type === 'date') {
      formatDate(inputCell);
    } else if (input.type === 'number' && input.format) {
      inputCell.setNumberFormat(input.format);
    }
    
    // Add note if provided
    if (input.note) {
      inputCell.setNote(input.note);
    }
    
    // Add validation if provided
    if (input.validation) {
      if (typeof input.validation === 'string') {
        // Use predefined validation
        applyValidationList(inputCell, input.validation);
      } else {
        // Custom validation object
        inputCell.setDataValidation(input.validation);
      }
    }
    
    row++;
  });
  
  return row;
}

/**
 * Create a summary/totals section
 * @param {Sheet} sheet - The sheet
 * @param {number} startRow - Starting row
 * @param {number} startCol - Starting column
 * @param {Array<Object>} totals - Array of {label, formula, format}
 * @param {string} title - Section title (optional)
 */
function createTotalsSection(sheet, startRow, startCol, totals, title = 'TOTALS') {
  let row = startRow;
  
  // Title if provided
  if (title) {
    const titleRange = sheet.getRange(row, startCol, 1, 2);
    titleRange.merge()
              .setValue(title)
              .setBackground(COLORS.SECTION_BG)
              .setFontWeight('bold')
              .setHorizontalAlignment('center');
    row++;
  }
  
  // Totals rows
  totals.forEach(total => {
    sheet.getRange(row, startCol)
         .setValue(total.label)
         .setFontWeight('bold');
    
    const valueCell = sheet.getRange(row, startCol + 1);
    
    if (total.formula) {
      valueCell.setFormula(total.formula);
    } else if (total.value !== undefined) {
      valueCell.setValue(total.value);
    }
    
    valueCell.setFontWeight('bold')
             .setBackground(total.bgColor || COLORS.TOTAL_BG);
    
    // Apply formatting
    if (total.format === 'currency') {
      formatCurrency(valueCell);
    } else if (total.format === 'percentage') {
      formatPercentage(valueCell);
    } else if (total.format) {
      valueCell.setNumberFormat(total.format);
    }
    
    row++;
  });
  
  return row;
}

/**
 * Create an instructions/notes section
 * @param {Sheet} sheet - The sheet
 * @param {number} row - Row number
 * @param {number} startCol - Starting column
 * @param {number} endCol - Ending column
 * @param {string} title - Title
 * @param {string} text - Instructions text
 */
function createInstructionsSection(sheet, row, startCol, endCol, title, text) {
  // Title
  sheet.getRange(row, startCol, 1, endCol - startCol + 1)
       .merge()
       .setValue(title)
       .setFontWeight('bold')
       .setBackground(COLORS.INFO_BG)
       .setHorizontalAlignment('center');
  
  // Text
  sheet.getRange(row + 1, startCol, 1, endCol - startCol + 1)
       .merge()
       .setValue(text)
       .setWrap(true)
       .setVerticalAlignment('top')
       .setBackground('#ffffff')
       .setBorder(true, true, true, true, false, false, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  
  return row + 2;
}

/**
 * Create a navigation/table of contents section
 * @param {Sheet} sheet - The sheet
 * @param {number} startRow - Starting row
 * @param {Array<Object>} items - Array of {sheet, description}
 * @returns {number} - Next available row
 */
function createNavigationSection(sheet, startRow, items) {
  const headers = ['Sheet Name', 'Description'];
  const data = items.map(item => [item.sheet, item.description]);
  
  createDataTable(sheet, startRow, 1, headers, data, {
    headerBg: COLORS.SUBHEADER_BG,
    headerHeight: 30
  });
  
  return startRow + items.length + 1;
}

/**
 * Create a sign-off section
 * @param {Sheet} sheet - The sheet
 * @param {number} row - Starting row
 * @param {number} startCol - Starting column
 */
function createSignOffSection(sheet, row, startCol = 1) {
  const signOffData = [
    ['Prepared By:', '', 'Date:', ''],
    ['Reviewed By:', '', 'Date:', ''],
    ['Approved By:', '', 'Date:', '']
  ];
  
  signOffData.forEach((rowData, index) => {
    sheet.getRange(row + index, startCol).setValue(rowData[0]).setFontWeight('bold');
    sheet.getRange(row + index, startCol + 1).setBackground(COLORS.INPUT_BG);
    sheet.getRange(row + index, startCol + 2).setValue(rowData[2]).setFontWeight('bold');
    sheet.getRange(row + index, startCol + 3).setBackground(COLORS.INPUT_BG);
    formatDate(sheet.getRange(row + index, startCol + 3));
  });
  
  return row + signOffData.length;
}

/**
 * Apply alternating row colors
 * @param {Sheet} sheet - The sheet
 * @param {number} startRow - Starting row
 * @param {number} endRow - Ending row
 * @param {number} startCol - Starting column
 * @param {number} endCol - Ending column
 * @param {string} color1 - First color (default: white)
 * @param {string} color2 - Second color (default: light grey)
 */
function applyAlternatingRows(sheet, startRow, endRow, startCol, endCol, color1 = '#ffffff', color2 = '#f2f2f2') {
  for (let row = startRow; row <= endRow; row++) {
    const color = (row - startRow) % 2 === 0 ? color1 : color2;
    sheet.getRange(row, startCol, 1, endCol - startCol + 1).setBackground(color);
  }
}

/**
 * Freeze header rows and columns
 * @param {Sheet} sheet - The sheet
 * @param {number} rows - Number of rows to freeze (default: 1)
 * @param {number} cols - Number of columns to freeze (default: 0)
 */
function freezeHeaders(sheet, rows = 1, cols = 0) {
  if (rows > 0) {
    sheet.setFrozenRows(rows);
  }
  if (cols > 0) {
    sheet.setFrozenColumns(cols);
  }
}

/**
 * Create a complete standard audit sheet
 * @param {Spreadsheet} ss - The spreadsheet
 * @param {Object} config - Configuration object
 * @returns {Sheet}
 * 
 * Example config:
 * {
 *   name: 'My Sheet',
 *   title: 'MY AUDIT SHEET',
 *   subtitle: 'Description',
 *   tabColor: '#1a237e',
 *   headers: ['Col1', 'Col2', 'Col3'],
 *   sampleData: [[1, 2, 3], [4, 5, 6]],
 *   instructions: 'Fill in the data...',
 *   freezeRows: 3
 * }
 */
function createStandardAuditSheet(ss, config) {
  const sheet = getOrCreateSheet(ss, config.name, config.index, config.tabColor);
  
  // Set column widths if provided
  if (config.columnWidths) {
    setColumnWidths(sheet, config.columnWidths);
  }
  
  let currentRow = 1;
  
  // Header
  createStandardHeader(sheet, config.title, config.subtitle, 1, config.headers.length);
  currentRow = config.subtitle ? 3 : 2;
  
  // Instructions if provided
  if (config.instructions) {
    currentRow = createInstructionsSection(
      sheet, currentRow, 1, config.headers.length,
      'INSTRUCTIONS', config.instructions
    );
    currentRow++;
  }
  
  // Data table
  const table = createDataTable(
    sheet, currentRow, 1,
    config.headers,
    config.sampleData || [],
    config.tableOptions || {}
  );
  
  // Freeze rows
  freezeHeaders(sheet, config.freezeRows || currentRow, config.freezeCols || 0);
  
  return sheet;
}


/**
 * ════════════════════════════════════════════════════════════════════════════
 * NAMED RANGES SETUP
 * ════════════════════════════════════════════════════════════════════════════
 * Common named range setup functions
 */

// ============================================================================
// NAMED RANGES
// ============================================================================

function setupNamedRanges(ss) {
  // This function can be overridden in specific workbooks
  // Default implementation does nothing
  Logger.log('Named ranges setup (default - no ranges created)');
}

function createNamedRange(ss, name, range) {
  try {
    // Remove existing named range if it exists
    const existingRange = ss.getNamedRanges().find(nr => nr.getName() === name);
    if (existingRange) {
      existingRange.remove();
    }
    
    // Create new named range
    ss.setNamedRange(name, range);
    Logger.log('Created named range: ' + name);
  } catch (error) {
    Logger.log('Error creating named range ' + name + ': ' + error.toString());
  }
}


/**
 * ════════════════════════════════════════════════════════════════════════════
 * IGAAP-IND AS AUDIT BUILDER: DEFERRED TAXATION WORKINGS
 * ════════════════════════════════════════════════════════════════════════════
 * 
 * Purpose: Generate comprehensive deferred tax audit workings compliant with
 *          IGAAP (AS 22) and Ind AS (Ind AS 12) frameworks
 * 
 * Features:
 * - Framework selection (IGAAP/Ind AS)
 * - Dynamic temporary difference tracking
 * - Automatic DTA/DTL classification
 * - Movement schedules with period closures
 * - P&L and Balance Sheet reconciliations
 * - Full audit trail with control totals
 * 
 * Author: IGAAP-Ind AS Audit Builder
 * Version: 1.0
 * Last Updated: November 2025
 * ════════════════════════════════════════════════════════════════════════════
 */

// ============================================================================
// MAIN EXECUTION FUNCTION
// ============================================================================

function createDeferredTaxWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Set workbook type for menu detection
  setWorkbookType('DEFERRED_TAX');
  
  // Clear existing sheets (optional - comment out if you want to preserve data)
  clearExistingSheets(ss);
  
  // Create sheets in order
  Logger.log("Creating Cover sheet...");
  createCoverSheet(ss);
  
  Logger.log("Creating Assumptions sheet...");
  createAssumptionsSheet(ss);
  
  Logger.log("Creating Input Variables sheet...");
  createInputVariablesSheet(ss);
  
  Logger.log("Creating Temporary Differences sheet...");
  createTempDifferencesSheet(ss);
  
  Logger.log("Creating Deferred Tax Schedule...");
  createDTScheduleSheet(ss);
  
  Logger.log("Creating Movement Analysis sheet...");
  createMovementAnalysisSheet(ss);
  
  Logger.log("Creating P&L Reconciliation sheet...");
  createPLReconciliationSheet(ss);
  
  Logger.log("Creating Balance Sheet Reconciliation sheet...");
  createBSReconciliationSheet(ss);
  
  Logger.log("Creating References sheet...");
  createReferencesSheet(ss);
  
  Logger.log("Creating Audit Notes sheet...");
  createAuditNotesSheet(ss);
  
  // Set up named ranges
  Logger.log("Setting up named ranges...");
  setupNamedRanges(ss);
  
  // Format all sheets
  Logger.log("Applying final formatting...");
  applyFinalFormatting(ss);
  
  // Activate Cover sheet
  ss.getSheetByName("Cover").activate();
  
  SpreadsheetApp.getUi().alert(
    "Deferred Tax Workbook Created Successfully!",
    "Your IGAAP/Ind AS Deferred Tax workbook is ready.\n\n" +
    "Next Steps:\n" +
    "1. Go to 'Assumptions' sheet and enter entity details\n" +
    "2. Select framework (IGAAP or Ind AS)\n" +
    "3. Enter tax rates and period information\n" +
    "4. Go to 'Temp_Differences' sheet and enter temporary differences\n" +
    "5. Review calculations in 'DT_Schedule' sheet\n\n" +
    "All formulas are dynamic and will update automatically.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================================================
// WORKBOOK-SPECIFIC CONFIGURATION
// ============================================================================

// Column mappings for Deferred Tax workbook
const COLS = {
  // Temp_Differences sheet columns
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
  },
  // Add other sheet column mappings as needed
};

// Key row numbers for reference
const ROWS = {
  ASSUMPTIONS: {
    ENTITY_NAME: 5,
    FY: 6,
    FRAMEWORK: 7,
    REPORTING_DATE: 8,
    CURRENT_TAX_RATE: 13,
    DT_RATE_CURRENT: 14
  }
};

// ============================================================================
// COVER SHEET
// ============================================================================

function createCoverSheet(ss) {
  let sheet = ss.getSheetByName("Cover");
  if (!sheet) {
    sheet = ss.insertSheet("Cover", 0);
  }
  
  sheet.clear();
  setColumnWidths(sheet, [50, 200, 250, 200, 150, 150]);
  
  // Title Section
  sheet.getRange("B2:F4").merge()
    .setValue("DEFERRED TAXATION WORKINGS")
    .setFontSize(24)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT);
  
  sheet.getRange("B5:F5").merge()
    .setValue("IGAAP (AS 22) & Ind AS (Ind AS 12) Compliant Audit Workings")
    .setFontSize(12)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.SUBHEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT);
  
  // Entity Information Section
  sheet.getRange("B7").setValue("Entity Name:")
    .setFontWeight("bold");
  sheet.getRange("C7").setFormula('=Assumptions!B3')
    .setFontSize(11);
  
  sheet.getRange("B8").setValue("Financial Year:")
    .setFontWeight("bold");
  sheet.getRange("C8").setFormula('=Assumptions!B4')
    .setFontSize(11);
  
  sheet.getRange("B9").setValue("Framework:")
    .setFontWeight("bold");
  sheet.getRange("C9").setFormula('=Assumptions!B5')
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor("#d32f2f");
  
  sheet.getRange("B10").setValue("Reporting Date:")
    .setFontWeight("bold");
  sheet.getRange("C10").setFormula('=TEXT(Assumptions!B6,"DD-MMM-YYYY")')
    .setFontSize(11);
  
  // Key Metrics Summary
  sheet.getRange("B12:F12").merge()
    .setValue("KEY METRICS SUMMARY")
    .setFontSize(14)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.SECTION_BG);
  
  const metrics = [
    ["Metric", "Current Year", "Prior Year", "Movement", "% Change"],
    ["Deferred Tax Assets (DTA)", "=Movement_Analysis!F65", "=Movement_Analysis!C65", "=Movement_Analysis!I65", "=IF(Movement_Analysis!C65<>0,Movement_Analysis!I65/Movement_Analysis!C65,\"-\")"],
    ["Deferred Tax Liabilities (DTL)", "=Movement_Analysis!F66", "=Movement_Analysis!C66", "=Movement_Analysis!I66", "=IF(Movement_Analysis!C66<>0,Movement_Analysis!I66/Movement_Analysis!C66,\"-\")"],
    ["Net DTA/(DTL)", "=Movement_Analysis!F67", "=Movement_Analysis!C67", "=Movement_Analysis!I67", "=IF(Movement_Analysis!C67<>0,Movement_Analysis!I67/Movement_Analysis!C67,\"-\")"],
    ["", "", "", "", ""],
    ["Deferred Tax Expense/(Income)", "=PL_Reconciliation!C8", "", "", ""],
    ["Effective Tax Rate", "=PL_Reconciliation!C15", "", "", ""]
  ];
  
  sheet.getRange(13, 2, metrics.length, 5).setValues(metrics);
  sheet.getRange("B13:F13").setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  
  sheet.getRange("B14:B17").setNumberFormat("#,##0");
  sheet.getRange("C14:E17").setNumberFormat("#,##0");
  sheet.getRange("F14:F17").setNumberFormat("0.00%");
  sheet.getRange("C19").setNumberFormat("#,##0");
  sheet.getRange("C20").setNumberFormat("0.00%");
  
  // Navigation Buttons Section
  sheet.getRange("B22:F22").merge()
    .setValue("QUICK NAVIGATION")
    .setFontSize(14)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.SECTION_BG);
  
  const navigationButtons = [
    ["Sheet", "Purpose"],
    ["Assumptions", "Entity details, framework selection, tax rates"],
    ["Temp_Differences", "Input temporary differences data"],
    ["DT_Schedule", "Main deferred tax calculation schedule"],
    ["Movement_Analysis", "Opening to closing reconciliation"],
    ["PL_Reconciliation", "Tax expense and effective rate analysis"],
    ["BS_Reconciliation", "Balance sheet presentation of DTA/DTL"],
    ["References", "Accounting standards guidance"],
    ["Audit_Notes", "Audit assertions and control checks"]
  ];
  
  sheet.getRange(23, 2, navigationButtons.length, 2).setValues(navigationButtons);
  sheet.getRange("B23:C23").setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  
  // Instructions
  sheet.getRange("B33:F33").merge()
    .setValue("INSTRUCTIONS FOR USE")
    .setFontSize(14)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.SECTION_BG);
  
  const instructions = [
    ["Step", "Action"],
    ["1", "Navigate to 'Assumptions' sheet and complete all yellow highlighted input cells"],
    ["2", "Select your accounting framework (IGAAP or Ind AS) - this will adjust calculations automatically"],
    ["3", "Enter your entity's tax rates (current and deferred tax rates)"],
    ["4", "Go to 'Temp_Differences' sheet and enter all temporary differences line by line"],
    ["5", "Review 'DT_Schedule' for detailed deferred tax calculations"],
    ["6", "Check 'Movement_Analysis' for period-wise movement"],
    ["7", "Verify 'PL_Reconciliation' for tax expense correctness"],
    ["8", "Review 'BS_Reconciliation' for balance sheet presentation"],
    ["9", "Use 'Audit_Notes' sheet to document review points and control checks"],
    ["", ""],
    ["NOTE", "All yellow/light blue cells are input cells. All other cells are formula-driven."]
  ];
  
  sheet.getRange(34, 2, instructions.length, 2).setValues(instructions);
  sheet.getRange("B34:C34").setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  
  sheet.getRange("B45").setValue("NOTE:")
    .setFontWeight("bold")
    .setFontColor("#d32f2f");
  sheet.getRange("C45").setValue("All yellow/light blue cells are input cells. All other cells are formula-driven.")
    .setFontStyle("italic");
  
  // Borders
  sheet.getRange("B7:C10").setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("B13:F20").setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("B23:C31").setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("B34:C45").setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  
  // Freeze header
  sheet.setFrozenRows(1);
}

// ============================================================================
// ASSUMPTIONS SHEET
// ============================================================================

function createAssumptionsSheet(ss) {
  let sheet = ss.getSheetByName("Assumptions");
  if (!sheet) {
    sheet = ss.insertSheet("Assumptions", 1);
  }
  
  sheet.clear();
  setColumnWidths(sheet, [30, 300, 200, 250, 150]);
  
  // Header
  sheet.getRange("A1:E1").merge()
    .setValue("ASSUMPTIONS & INPUT PARAMETERS")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT);
  
  // Entity Information Section
  sheet.getRange("A3:E3").merge()
    .setValue("ENTITY INFORMATION")
    .setFontSize(12)
    .setFontWeight("bold")
    .setBackground(COLORS.SECTION_BG);
  
  const entityInfo = [
    ["Parameter", "Value", "Instructions", ""],
    ["Entity Name", "", "Enter the legal name of the entity", ""],
    ["Financial Year", "FY 2024-25", "Enter financial year (e.g., FY 2024-25)", ""],
    ["Framework", "Ind AS", "Select: IGAAP or Ind AS", "DROPDOWN"],
    ["Reporting Date", "31-Mar-2025", "Enter balance sheet date", "DATE"],
    ["Prior Period Reporting Date", "31-Mar-2024", "Enter prior period balance sheet date", "DATE"]
  ];
  
  sheet.getRange(4, 1, entityInfo.length, 4).setValues(entityInfo);
  sheet.getRange("A4:D4").setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  
  // Input cells styling
  sheet.getRange("B5:B9").setBackground(COLORS.INPUT_BG);
  
  // Data validation for Framework
  const frameworkRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["IGAAP", "Ind AS"], true)
    .setAllowInvalid(false)
    .setHelpText("Select the accounting framework: IGAAP (AS 22) or Ind AS (Ind AS 12)")
    .build();
  sheet.getRange("B6").setDataValidation(frameworkRule);
  
  // Tax Rates Section
  sheet.getRange("A11:E11").merge()
    .setValue("TAX RATES")
    .setFontSize(12)
    .setFontWeight("bold")
    .setBackground(COLORS.SECTION_BG);
  
  const taxRates = [
    ["Tax Rate Parameter", "Rate (%)", "Instructions", ""],
    ["Current Tax Rate", "25.17%", "Enter applicable current tax rate (including surcharge & cess)", ""],
    ["Deferred Tax Rate - Current Year", "25.17%", "Enter DT rate for current year temporary differences", ""],
    ["Deferred Tax Rate - Prior Year", "25.17%", "Enter DT rate used in prior year", ""],
    ["Minimum Alternate Tax (MAT) Rate", "15.60%", "Enter MAT rate if applicable (including surcharge & cess)", ""],
    ["", "", "", ""],
    ["Note:", "Ind AS 12 requires use of substantively enacted rates", "", ""],
    ["", "IGAAP (AS 22) requires use of enacted rates", "", ""]
  ];
  
  sheet.getRange(12, 1, taxRates.length, 4).setValues(taxRates);
  sheet.getRange("A12:D12").setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  
  sheet.getRange("B13:B16").setBackground(COLORS.INPUT_BG)
    .setNumberFormat("0.00%");
  
  // Period Information Section
  sheet.getRange("A22:E22").merge()
    .setValue("PERIOD INFORMATION")
    .setFontSize(12)
    .setFontWeight("bold")
    .setBackground(COLORS.SECTION_BG);
  
  const periodInfo = [
    ["Period Parameter", "Amount (₹)", "Instructions", ""],
    ["Accounting Profit Before Tax (PBT)", "10,000,000", "Enter PBT from P&L statement", ""],
    ["Prior Year PBT", "8,500,000", "Enter prior year PBT for comparative analysis", ""],
    ["Current Tax Expense (Computed)", "2,517,000", "Enter current tax as per tax computation", ""],
    ["", "", "", ""],
    ["Opening Balance - DTA", "500,000", "Enter opening DTA from prior year balance sheet", ""],
    ["Opening Balance - DTL", "750,000", "Enter opening DTL from prior year balance sheet", ""],
  ];
  
  sheet.getRange(23, 1, periodInfo.length, 4).setValues(periodInfo);
  sheet.getRange("A23:D23").setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  
  sheet.getRange("B24:B29").setBackground(COLORS.INPUT_BG)
    .setNumberFormat("#,##0");
  
  // Recognition Thresholds (Ind AS specific)
  sheet.getRange("A32:E32").merge()
    .setValue("RECOGNITION CRITERIA (Ind AS 12)")
    .setFontSize(12)
    .setFontWeight("bold")
    .setBackground(COLORS.SECTION_BG);
  
  const recognition = [
    ["Recognition Parameter", "Policy", "Instructions", ""],
    ["Recognize DTA on Carry Forward Losses", "Yes", "Select Yes/No - Ind AS 12 allows recognition if probable", "DROPDOWN"],
    ["Recognize DTA on Unabsorbed Depreciation", "Yes", "Select Yes/No - based on future taxable profit availability", "DROPDOWN"],
    ["Apply Netting of DTA/DTL", "Yes", "Select Yes/No - allowed if legally enforceable right exists", "DROPDOWN"],
    ["", "", "", ""],
    ["Note:", "IGAAP (AS 22) requires virtual certainty for loss-related DTA", "", ""],
    ["", "Ind AS 12 requires reasonable certainty (probable future taxable profits)", "", ""]
  ];
  
  sheet.getRange(33, 1, recognition.length, 4).setValues(recognition);
  sheet.getRange("A33:D33").setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  
  sheet.getRange("B34:B36").setBackground(COLORS.INPUT_ALT_BG);
  
  // Data validation for Yes/No dropdowns
  const yesNoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Yes", "No"], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange("B34:B36").setDataValidation(yesNoRule);
  
  // Borders
  sheet.getRange("A4:D9").setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("A12:D19").setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("A23:D29").setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("A33:D39").setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  
  // Add cell notes
  sheet.getRange("B6").setNote("Critical: Framework selection affects recognition criteria and disclosure requirements.\n\nIGAAP (AS 22): More conservative, requires virtual certainty for DTA on losses.\n\nInd AS 12: Allows DTA recognition when probable (>50% likelihood) that taxable profits will be available.");
  
  sheet.getRange("B13").setNote("Include surcharge and cess.\nExample: Base rate 25% + Surcharge 10% + Cess 4% = 25 × 1.1 × 1.04 = 28.6% (rounded)");
  
  sheet.getRange("B14").setNote("Use the tax rate substantively enacted (Ind AS) or enacted (IGAAP) as at balance sheet date for recognizing deferred tax.");
  
  // Freeze rows only (columns removed to avoid splitting merged cells)
  sheet.setFrozenRows(1);
}

// ============================================================================
// INPUT VARIABLES SHEET
// ============================================================================

function createInputVariablesSheet(ss) {
  let sheet = ss.getSheetByName("Input_Variables");
  if (!sheet) {
    sheet = ss.insertSheet("Input_Variables", 2);
  }
  
  sheet.clear();
  setColumnWidths(sheet, [30, 200, 150, 250, 120, 150]);
  
  // Header
  sheet.getRange("A1:F1").merge()
    .setValue("INPUT VARIABLES CATALOG")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT);
  
  sheet.getRange("A2:F2").merge()
    .setValue("Complete list of all user input cells across the workbook")
    .setFontSize(10)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.SUBHEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT);
  
  // Column headers
  const headers = [["Sheet Name", "Cell Reference", "Parameter Name", "Purpose", "Data Type", "Sample Value"]];
  sheet.getRange(4, 1, 1, 6).setValues(headers);
  sheet.getRange("A4:F4").setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  
  // Input variables data
  const inputVars = [
    // Assumptions sheet
    ["Assumptions", "B5", "Entity Name", "Legal name of the entity", "Text", "ABC Private Limited"],
    ["Assumptions", "B6", "Financial Year", "Reporting financial year", "Text", "FY 2024-25"],
    ["Assumptions", "B7", "Framework", "Accounting framework selection", "Dropdown", "Ind AS"],
    ["Assumptions", "B8", "Reporting Date", "Current period balance sheet date", "Date", "31-Mar-2025"],
    ["Assumptions", "B9", "Prior Period Date", "Prior period balance sheet date", "Date", "31-Mar-2024"],
    ["Assumptions", "B13", "Current Tax Rate", "Applicable current tax rate with surcharge & cess", "Percentage", "25.17%"],
    ["Assumptions", "B14", "DT Rate - Current", "Deferred tax rate for current year", "Percentage", "25.17%"],
    ["Assumptions", "B15", "DT Rate - Prior", "Deferred tax rate for prior year", "Percentage", "25.17%"],
    ["Assumptions", "B16", "MAT Rate", "Minimum Alternate Tax rate", "Percentage", "15.60%"],
    ["Assumptions", "B24", "PBT - Current", "Accounting profit before tax", "Number", "10,000,000"],
    ["Assumptions", "B25", "PBT - Prior", "Prior year PBT", "Number", "8,500,000"],
    ["Assumptions", "B26", "Current Tax", "Current tax expense computed", "Number", "2,517,000"],
    ["Assumptions", "B28", "Opening DTA", "Opening balance of DTA", "Number", "500,000"],
    ["Assumptions", "B29", "Opening DTL", "Opening balance of DTL", "Number", "750,000"],
    ["Assumptions", "B34", "DTA on Losses", "Policy for recognizing DTA on carry forward losses", "Dropdown", "Yes"],
    ["Assumptions", "B35", "DTA on Depreciation", "Policy for DTA on unabsorbed depreciation", "Dropdown", "Yes"],
    ["Assumptions", "B36", "Apply Netting", "Whether to net DTA and DTL", "Dropdown", "Yes"],
    ["", "", "", "", "", ""],
    // Temp_Differences sheet
    ["Temp_Differences", "B7:B50", "Temporary Difference Items", "Description of each temporary difference", "Text", "Depreciation - Block 1"],
    ["Temp_Differences", "C7:C50", "Category", "Classification of temporary difference", "Dropdown", "Depreciation"],
    ["Temp_Differences", "D7:D50", "Tax Base", "Tax base of the asset/liability", "Number", "1,000,000"],
    ["Temp_Differences", "E7:E50", "Book Base", "Carrying amount per books", "Number", "850,000"],
    ["Temp_Differences", "G7:G50", "Nature", "Whether creates DTA or DTL", "Auto-calc", "DTL"],
    ["Temp_Differences", "H7:H50", "Opening Temp Difference", "Opening temporary difference balance for movement analysis", "Number", "250,000"],
    ["Temp_Differences", "I7:I50", "Additions", "New temporary differences in current year", "Number", "50,000"],
    ["Temp_Differences", "J7:J50", "Reversals", "Reversals of prior period differences", "Number", "30,000"],
    ["Temp_Differences", "K7:K50", "Rate Change Impact", "Impact of tax rate changes", "Number", "0"],
    ["", "", "", "", "", ""],
    ["NOTE", "", "Yellow cells = Primary inputs", "", "", ""],
    ["NOTE", "", "Light blue cells = Secondary inputs", "", "", ""],
    ["NOTE", "", "All other cells = Formula-driven calculations", "", "", ""]
  ];
  
  sheet.getRange(5, 1, inputVars.length, 6).setValues(inputVars);
  
  // Formatting
  sheet.getRange("A5:F" + (4 + inputVars.length))
    .setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  
  // Color code input types
  for (let i = 5; i < 5 + inputVars.length; i++) {
    const dataType = sheet.getRange(i, 5).getValue();
    if (dataType === "Number" || dataType === "Percentage" || dataType === "Date") {
      sheet.getRange(i, 1, 1, 6).setBackground("#fff9c4"); // Yellow for primary inputs
    } else if (dataType === "Dropdown") {
      sheet.getRange(i, 1, 1, 6).setBackground("#b3e5fc"); // Light blue for dropdowns
    }
  }

  // Freeze rows only (columns removed to avoid splitting merged cells)
  sheet.setFrozenRows(4);
}

// ============================================================================
// TEMPORARY DIFFERENCES SHEET
// ============================================================================

function createTempDifferencesSheet(ss) {
  let sheet = ss.getSheetByName("Temp_Differences");
  if (!sheet) {
    sheet = ss.insertSheet("Temp_Differences", 3);
  }
  
  sheet.clear();
  setColumnWidths(sheet, [40, 250, 150, 120, 120, 120, 80, 120, 120, 120, 120, 150]);
  
  // Header
  sheet.getRange("A1:L1").merge()
    .setValue("TEMPORARY DIFFERENCES INPUT SCHEDULE")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT);
  
  sheet.getRange("A2:L2").merge()
    .setValue("Enter all temporary differences between book and tax basis of assets and liabilities")
    .setFontSize(10)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.SUBHEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT);
  
  // Instructions
  sheet.getRange("A4:L4").merge()
    .setValue("INSTRUCTIONS: Enter temporary differences line by line. System will auto-calculate DTA/DTL. Yellow cells are inputs.")
    .setFontSize(10)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.INFO_BG);
  
  // Column headers
  const headers = [
    [
      "Sr.",
      "Temporary Difference Item",
      "Category",
      "Tax Base (A)",
      "Book Base (B)",
      "Temp Diff (C=B-A)",
      "Nature",
      "Opening Temp Diff (Prior)",
      "Additions (CY)",
      "Reversals (CY)",
      "Rate Change Impact",
      "Remarks"
    ]
  ];
  sheet.getRange(6, 1, 1, 12).setValues(headers);
  sheet.getRange("A6:L6").setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
  // Sample data and formulas (rows 7-50)
  const sampleData = [
    [
      1,
      "Depreciation - Plant & Machinery (Block 1)",
      "Depreciation",
      5000000,
      4500000,
      "=E7-D7",
      "=IF(F7>0,\"DTL\",IF(F7<0,\"DTA\",\"-\"))",
      250000,
      200000,
      150000,
      0,
      "Timing difference due to different depreciation rates"
    ],
    [
      2,
      "Provision for Doubtful Debts",
      "Provision",
      0,
      300000,
      "=E8-D8",
      "=IF(F8>0,\"DTL\",IF(F8<0,\"DTA\",\"-\"))",
      120000,
      50000,
      0,
      0,
      "Deductible when written off"
    ],
    [
      3,
      "Employee Benefits - Gratuity Provision",
      "Employee Benefits",
      0,
      450000,
      "=E9-D9",
      "=IF(F9>0,\"DTL\",IF(F9<0,\"DTA\",\"-\"))",
      60000,
      80000,
      30000,
      0,
      "Deductible on payment basis"
    ],
    [
      4,
      "Disallowance u/s 43B - Statutory Dues",
      "Section 43B",
      0,
      125000,
      "=E10-D10",
      "=IF(F10>0,\"DTL\",IF(F10<0,\"DTA\",\"-\"))",
      50000,
      125000,
      100000,
      0,
      "Timing difference - payment basis vs accrual"
    ],
    [
      5,
      "Carry Forward Business Losses",
      "Tax Losses",
      0,
      1500000,
      "=E11-D11",
      "=IF(F11>0,\"DTL\",IF(F11<0,\"DTA\",\"-\"))",
      900000,
      0,
      500000,
      0,
      "DTA subject to probability assessment"
    ]
  ];

  sheet.getRange(7, 1, sampleData.length, 12).setValues(sampleData);
  
  // Add empty rows with formulas for rows 12-50
  for (let row = 12; row <= 50; row++) {
    sheet.getRange(row, 1).setValue(row - 6);
    sheet.getRange(row, 6).setFormula(`=IF(OR(D${row}<>\"\",E${row}<>\"\"),E${row}-D${row},"")`);
    sheet.getRange(row, 7).setFormula(`=IF(F${row}>0,"DTL",IF(F${row}<0,"DTA",""))`);
  }
  
  // Category dropdown
  const categories = [
    "Depreciation",
    "Provision",
    "Employee Benefits",
    "Section 43B",
    "Tax Losses",
    "Unabsorbed Depreciation",
    "Revenue Recognition",
    "Financial Instruments",
    "Lease Accounting",
    "Other"
  ];
  
  const categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(categories, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange("C7:C50").setDataValidation(categoryRule);
  
  // Formatting
  sheet.getRange("B7:B50").setBackground(COLORS.INPUT_BG);
  sheet.getRange("C7:C50").setBackground(COLORS.INPUT_ALT_BG);
  sheet.getRange("D7:E50").setBackground(COLORS.INPUT_BG);
  sheet.getRange("H7:K50").setBackground(COLORS.INPUT_BG);
  sheet.getRange("L7:L50").setBackground("#ffffff");
  
  // Number formatting
  sheet.getRange("D7:F50").setNumberFormat("#,##0");
  sheet.getRange("H7:K50").setNumberFormat("#,##0");
  
  // Totals row
  const totalRow = 52;
  sheet.getRange(`A${totalRow}:B${totalRow}`).merge()
    .setValue("TOTAL TEMPORARY DIFFERENCES")
    .setFontWeight("bold")
    .setBackground(COLORS.TOTAL_BG);
  
  sheet.getRange(`F${totalRow}`).setFormula(`=SUM(F7:F51)`)
    .setFontWeight("bold")
    .setNumberFormat("#,##0")
    .setBackground(COLORS.TOTAL_BG);
  
  sheet.getRange(`H${totalRow}`).setFormula(`=SUM(H7:H51)`)
    .setFontWeight("bold")
    .setNumberFormat("#,##0")
    .setBackground(COLORS.TOTAL_BG);
  
  sheet.getRange(`I${totalRow}`).setFormula(`=SUM(I7:I51)`)
    .setFontWeight("bold")
    .setNumberFormat("#,##0")
    .setBackground(COLORS.TOTAL_BG);
  
  sheet.getRange(`J${totalRow}`).setFormula(`=SUM(J7:J51)`)
    .setFontWeight("bold")
    .setNumberFormat("#,##0")
    .setBackground(COLORS.TOTAL_BG);

  sheet.getRange(`K${totalRow}`).setFormula(`=SUM(K7:K51)`)
    .setFontWeight("bold")
    .setNumberFormat("#,##0")
    .setBackground(COLORS.TOTAL_BG);
  
  // Borders
  sheet.getRange("A6:L52").setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  
  // Add cell notes for guidance
  sheet.getRange("D7").setNote("Tax Base:\nThe amount attributed to this item for tax purposes.\n\nFor assets: Tax WDV or NIL if fully written off\nFor liabilities: Tax-deductible amount");
  
  sheet.getRange("E7").setNote("Book Base:\nCarrying amount as per books/financial statements.\n\nFor assets: NBV as per books\nFor liabilities: Provision amount in books");
  
  sheet.getRange("F7").setNote("Temporary Difference:\nDifference between book base and tax base.\n\nPositive = Taxable temporary difference (creates DTL)\nNegative = Deductible temporary difference (creates DTA)");
  
  sheet.getRange("H7").setNote("Opening Temp Difference:\nCarry-forward temporary difference balance as at the start of the current year. This should tie to prior year closing balance for the item.");

  sheet.getRange("I7").setNote("Additions:\nNew temporary differences arising in the current year that will reverse in future periods.");

  sheet.getRange("J7").setNote("Reversals:\nReversal of temporary differences from prior years that reversed in the current year.");

  sheet.getRange("K7").setNote("Rate Change Impact:\nImpact on deferred tax due to changes in enacted tax rates or laws affecting this temporary difference.");
  
  // Freeze rows only (columns removed to avoid splitting merged cells)
  sheet.setFrozenRows(6);
}

// ============================================================================
// DEFERRED TAX SCHEDULE SHEET
// ============================================================================

function createDTScheduleSheet(ss) {
  let sheet = ss.getSheetByName("DT_Schedule");
  if (!sheet) {
    sheet = ss.insertSheet("DT_Schedule", 4);
  }
  
  sheet.clear();
  setColumnWidths(sheet, [40, 250, 120, 120, 120, 120, 120, 120, 120, 150]);
  
  // Header
  sheet.getRange("A1:J1").merge()
    .setValue("DEFERRED TAX CALCULATION SCHEDULE")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT);
  
  sheet.getRange("A2:J2").merge()
    .setValue("Detailed computation of Deferred Tax Assets and Liabilities")
    .setFontSize(10)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.SUBHEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT);
  
  // Framework indicator
  sheet.getRange("A4").setValue("Framework:");
  sheet.getRange("B4").setFormula("=Assumptions!B7")
    .setFontWeight("bold")
    .setFontColor("#d32f2f");
  
  sheet.getRange("D4").setValue("Tax Rate:");
  sheet.getRange("E4").setFormula("=Assumptions!B14")
    .setFontWeight("bold")
    .setNumberFormat("0.00%");
  
  // Column headers
  const headers = [
    ["Sr.", "Temporary Difference Item", "Temp Diff Amount", "Applicable Rate", "Deferred Tax", "Classification", "DTA Amount", "DTL Amount", "Net DTA/(DTL)", "Remarks"]
  ];
  sheet.getRange(6, 1, 1, 10).setValues(headers);
  sheet.getRange("A6:J6").setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
  // Data rows with formulas (7-50)
  for (let row = 7; row <= 50; row++) {
    const srcRow = row;
    
    // Sr. No
    sheet.getRange(row, 1).setFormula(`=IF(Temp_Differences!B${srcRow}<>"",Temp_Differences!A${srcRow},"")`);
    
    // Item description
    sheet.getRange(row, 2).setFormula(`=Temp_Differences!B${srcRow}`);
    
    // Temp Diff Amount
    sheet.getRange(row, 3).setFormula(`=IF(Temp_Differences!F${srcRow}<>"",Temp_Differences!F${srcRow},"")`);
    
    // Applicable Rate
    sheet.getRange(row, 4).setFormula(`=IF(C${row}<>"",Assumptions!$B$14,"")`);
    
    // Deferred Tax
    sheet.getRange(row, 5).setFormula(`=IF(AND(C${row}<>"",D${row}<>""),C${row}*D${row},"")`);
    
    // Classification
    sheet.getRange(row, 6).setFormula(`=IF(E${row}<>"",IF(E${row}>0,"DTL",IF(E${row}<0,"DTA","")),"")`);
    
    // DTA Amount
    sheet.getRange(row, 7).setFormula(`=IF(F${row}="DTA",ABS(E${row}),"")`);
    
    // DTL Amount
    sheet.getRange(row, 8).setFormula(`=IF(F${row}="DTL",E${row},"")`);
    
    // Net DTA/(DTL)
    sheet.getRange(row, 9).setFormula(`=IF(E${row}<>"",E${row},"")`);
    
    // Remarks
    sheet.getRange(row, 10).setFormula(`=Temp_Differences!L${srcRow}`);
  }
  
  // Subtotals row
  const subtotalRow = 52;
  sheet.getRange(`A${subtotalRow}:B${subtotalRow}`).merge()
    .setValue("SUBTOTAL - Before Recognition Assessment")
    .setFontWeight("bold")
    .setBackground(COLORS.CALC_BG);
  
  sheet.getRange(`C${subtotalRow}`).setFormula(`=SUM(C7:C51)`)
    .setFontWeight("bold")
    .setNumberFormat("#,##0")
    .setBackground(COLORS.CALC_BG);
  
  sheet.getRange(`E${subtotalRow}`).setFormula(`=SUM(E7:E51)`)
    .setFontWeight("bold")
    .setNumberFormat("#,##0")
    .setBackground(COLORS.CALC_BG);
  
  sheet.getRange(`G${subtotalRow}`).setFormula(`=SUM(G7:G51)`)
    .setFontWeight("bold")
    .setNumberFormat("#,##0")
    .setBackground(COLORS.CALC_BG);
  
  sheet.getRange(`H${subtotalRow}`).setFormula(`=SUM(H7:H51)`)
    .setFontWeight("bold")
    .setNumberFormat("#,##0")
    .setBackground(COLORS.CALC_BG);
  
  sheet.getRange(`I${subtotalRow}`).setFormula(`=SUM(I7:I51)`)
    .setFontWeight("bold")
    .setNumberFormat("#,##0")
    .setBackground(COLORS.CALC_BG);
  
  // Recognition adjustments section
  sheet.getRange(`A${subtotalRow + 2}:J${subtotalRow + 2}`).merge()
    .setValue("RECOGNITION ADJUSTMENTS")
    .setFontWeight("bold")
    .setFontSize(12)
    .setHorizontalAlignment("center")
    .setBackground(COLORS.SECTION_BG);
  
  const adjustmentRow = subtotalRow + 3;
  sheet.getRange(`A${adjustmentRow}:F${adjustmentRow}`).merge()
    .setValue("Less: DTA not recognized (due to probability assessment)")
    .setFontStyle("italic");
  
  sheet.getRange(`G${adjustmentRow}`).setValue(0)
    .setBackground(COLORS.INPUT_ALT_BG)
    .setNumberFormat("#,##0");
  
  sheet.getRange(`J${adjustmentRow}`).setValue("Enter amount if DTA is not recognized")
    .setFontStyle("italic")
    .setFontSize(9);
  
  // Final recognized amounts
  const finalRow = subtotalRow + 5;
  sheet.getRange(`A${finalRow}:B${finalRow}`).merge()
    .setValue("TOTAL RECOGNIZED DTA")
    .setFontWeight("bold")
    .setBackground(COLORS.GRAND_TOTAL_BG);
  
  sheet.getRange(`G${finalRow}`).setFormula(`=G${subtotalRow}-G${adjustmentRow}`)
    .setFontWeight("bold")
    .setNumberFormat("#,##0")
    .setBackground(COLORS.GRAND_TOTAL_BG);
  
  sheet.getRange(`A${finalRow + 1}:B${finalRow + 1}`).merge()
    .setValue("TOTAL RECOGNIZED DTL")
    .setFontWeight("bold")
    .setBackground(COLORS.GRAND_TOTAL_BG);
  
  sheet.getRange(`H${finalRow + 1}`).setFormula(`=H${subtotalRow}`)
    .setFontWeight("bold")
    .setNumberFormat("#,##0")
    .setBackground(COLORS.GRAND_TOTAL_BG);
  
  sheet.getRange(`A${finalRow + 2}:B${finalRow + 2}`).merge()
    .setValue("NET DTA/(DTL)")
    .setFontWeight("bold")
    .setBackground(COLORS.GRAND_TOTAL_BG);
  
  sheet.getRange(`I${finalRow + 2}`).setFormula(`=G${finalRow}-H${finalRow + 1}`)
    .setFontWeight("bold")
    .setNumberFormat("#,##0")
    .setBackground(COLORS.GRAND_TOTAL_BG);
  
  // Framework-specific notes
  const notesRow = finalRow + 4;
  sheet.getRange(`A${notesRow}:J${notesRow}`).merge()
    .setValue("FRAMEWORK-SPECIFIC RECOGNITION CRITERIA")
    .setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center")
    .setBackground(COLORS.INFO_BG);
  
  const notes = [
    ["IGAAP (AS 22):", "DTA recognized only when there is virtual certainty (backed by convincing evidence) of sufficient future taxable income"],
    ["", "DTA on carry forward losses/unabsorbed depreciation recognized only if virtual certainty exists"],
    ["", ""],
    ["Ind AS 12:", "DTA recognized when it is probable (>50% likelihood) that taxable profit will be available"],
    ["", "More liberal than IGAAP - requires reasonable certainty rather than virtual certainty"],
    ["", "DTA on unused tax losses/credits recognized if probable that taxable profits will be available"],
    ["", ""],
    ["Current Selection:", "=Assumptions!B7"]
  ];
  
  sheet.getRange(notesRow + 1, 1, notes.length, 2).setValues(notes);
  sheet.getRange(`B${notesRow + 8}`).setFontWeight("bold").setFontColor("#d32f2f");
  
  // Number formatting
  sheet.getRange("C7:E50").setNumberFormat("#,##0");
  sheet.getRange("D7:D50").setNumberFormat("0.00%");
  sheet.getRange("G7:I50").setNumberFormat("#,##0");
  
  // Borders
  sheet.getRange("A6:J" + (finalRow + 2)).setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("A" + (notesRow) + ":C" + (notesRow + 8)).setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  
  // Freeze rows only (columns removed to avoid splitting merged cells)
  sheet.setFrozenRows(6);
}

// ============================================================================
// MOVEMENT ANALYSIS SHEET
// ============================================================================

function createMovementAnalysisSheet(ss) {
  let sheet = ss.getSheetByName("Movement_Analysis");
  if (!sheet) {
    sheet = ss.insertSheet("Movement_Analysis", 5);
  }

  sheet.clear();
  setColumnWidths(sheet, [40, 250, 120, 120, 120, 120, 120, 120, 120]);

  // Header
  sheet.getRange("A1:I1").merge()
    .setValue("DEFERRED TAX MOVEMENT ANALYSIS")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT);

  sheet.getRange("A2:I2").merge()
    .setValue("Reconciliation of opening to closing balances of Deferred Tax Assets and Liabilities")
    .setFontSize(10)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.SUBHEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT);

  // Column headers
  const headers = [
    ["Sr.", "Particular", "Opening Balance", "Additions (CY)", "Reversals (CY)", "Closing Balance", "Recognized in P&L", "Recognized in OCI/Equity", "Movement"]
  ];
  sheet.getRange(5, 1, 1, 9).setValues(headers);
  sheet.getRange("A5:I5").setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  const dtaHeaderRow = 7;
  const dtaDataStartRow = dtaHeaderRow + 1;
  const dtaDataRows = 20;
  const dtaSubtotalRow = dtaDataStartRow + dtaDataRows;
  const dtlHeaderRow = dtaSubtotalRow + 2;
  const dtlDataStartRow = dtlHeaderRow + 1;
  const dtlDataRows = 20;
  const dtlSubtotalRow = dtlDataStartRow + dtlDataRows;
  const netHeaderRow = dtlSubtotalRow + 2;
  const netDataRow = netHeaderRow + 1;
  const nettingHeaderRow = netDataRow + 2;
  const applyNettingRow = nettingHeaderRow + 1;
  const nettingTableStartRow = applyNettingRow + 1;
  const nettingRowsCount = 4;
  const finalHeaderRow = nettingTableStartRow + nettingRowsCount + 1;
  const finalDataStartRow = finalHeaderRow + 2;
  const finalRowsCount = 3;
  const journalHeaderRow = finalDataStartRow + finalRowsCount + 3;
  const journalTableHeaderRow = journalHeaderRow + 2;

  // DTA Section
  sheet.getRange(`A${dtaHeaderRow}:I${dtaHeaderRow}`).merge()
    .setValue("DEFERRED TAX ASSETS")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(COLORS.SECTION_BG);

  const dtaOffset = dtaDataStartRow - 1;
  sheet.getRange(dtaDataStartRow, 1, dtaDataRows, 1).setFormulaR1C1(`=IF(RC[1]="","",ROW()-${dtaOffset})`);
  sheet.getRange(dtaDataStartRow, 2, dtaDataRows, 1).setFormulaR1C1(`=IFERROR(INDEX(FILTER(Temp_Differences!R7C2:R1000C2,Temp_Differences!R7C2:R1000C2<>"",Temp_Differences!R7C7:R1000C7="DTA"),ROW()-${dtaOffset}),"")`);
  sheet.getRange(dtaDataStartRow, 3, dtaDataRows, 1).setFormulaR1C1(`=IF(RC[-1]="","",IFERROR(INDEX(FILTER(Temp_Differences!R7C8:R1000C8,Temp_Differences!R7C2:R1000C2<>"",Temp_Differences!R7C7:R1000C7="DTA"),ROW()-${dtaOffset}),0))`);
  sheet.getRange(dtaDataStartRow, 4, dtaDataRows, 1).setFormulaR1C1(`=IF(RC[-2]="","",IFERROR(INDEX(FILTER(Temp_Differences!R7C9:R1000C9,Temp_Differences!R7C2:R1000C2<>"",Temp_Differences!R7C7:R1000C7="DTA"),ROW()-${dtaOffset}),0))`);
  sheet.getRange(dtaDataStartRow, 5, dtaDataRows, 1).setFormulaR1C1(`=IF(RC[-3]="","",IFERROR(INDEX(FILTER(Temp_Differences!R7C10:R1000C10,Temp_Differences!R7C2:R1000C2<>"",Temp_Differences!R7C7:R1000C7="DTA"),ROW()-${dtaOffset}),0))`);
  sheet.getRange(dtaDataStartRow, 8, dtaDataRows, 1).setFormulaR1C1(`=IF(RC[-6]="","",IFERROR(INDEX(FILTER(Temp_Differences!R7C11:R1000C11,Temp_Differences!R7C2:R1000C2<>"",Temp_Differences!R7C7:R1000C7="DTA"),ROW()-${dtaOffset}),0))`);
  sheet.getRange(dtaDataStartRow, 7, dtaDataRows, 1).setFormulaR1C1('=IF(RC[-5]="","",RC[-3]-RC[-2])');
  sheet.getRange(dtaDataStartRow, 6, dtaDataRows, 1).setFormulaR1C1('=IF(RC[-4]="","",RC[-3]+RC[-2]-RC[-1]+IF(RC[2]="",0,RC[2]))');
  sheet.getRange(dtaDataStartRow, 9, dtaDataRows, 1).setFormulaR1C1('=IF(RC[-8]="","",RC[-2]+RC[-1])');

  // DTA Subtotal
  sheet.getRange(`A${dtaSubtotalRow}:B${dtaSubtotalRow}`).merge()
    .setValue("Subtotal - DTA")
    .setFontWeight("bold")
    .setBackground(COLORS.CALC_BG);
  for (let col = 3; col <= 9; col++) {
    const colLetter = String.fromCharCode(64 + col);
    sheet.getRange(dtaSubtotalRow, col).setFormula(`=SUM(${colLetter}${dtaDataStartRow}:${colLetter}${dtaSubtotalRow - 1})`)
      .setFontWeight("bold")
      .setBackground(COLORS.CALC_BG);
  }

  // DTL Section
  sheet.getRange(`A${dtlHeaderRow}:I${dtlHeaderRow}`).merge()
    .setValue("DEFERRED TAX LIABILITIES")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(COLORS.SECTION_BG);

  const dtlOffset = dtlDataStartRow - 1;
  sheet.getRange(dtlDataStartRow, 1, dtlDataRows, 1).setFormulaR1C1(`=IF(RC[1]="","",ROW()-${dtlOffset})`);
  sheet.getRange(dtlDataStartRow, 2, dtlDataRows, 1).setFormulaR1C1(`=IFERROR(INDEX(FILTER(Temp_Differences!R7C2:R1000C2,Temp_Differences!R7C2:R1000C2<>"",Temp_Differences!R7C7:R1000C7="DTL"),ROW()-${dtlOffset}),"")`);
  sheet.getRange(dtlDataStartRow, 3, dtlDataRows, 1).setFormulaR1C1(`=IF(RC[-1]="","",IFERROR(INDEX(FILTER(Temp_Differences!R7C8:R1000C8,Temp_Differences!R7C2:R1000C2<>"",Temp_Differences!R7C7:R1000C7="DTL"),ROW()-${dtlOffset}),0))`);
  sheet.getRange(dtlDataStartRow, 4, dtlDataRows, 1).setFormulaR1C1(`=IF(RC[-2]="","",IFERROR(INDEX(FILTER(Temp_Differences!R7C9:R1000C9,Temp_Differences!R7C2:R1000C2<>"",Temp_Differences!R7C7:R1000C7="DTL"),ROW()-${dtlOffset}),0))`);
  sheet.getRange(dtlDataStartRow, 5, dtlDataRows, 1).setFormulaR1C1(`=IF(RC[-3]="","",IFERROR(INDEX(FILTER(Temp_Differences!R7C10:R1000C10,Temp_Differences!R7C2:R1000C2<>"",Temp_Differences!R7C7:R1000C7="DTL"),ROW()-${dtlOffset}),0))`);
  sheet.getRange(dtlDataStartRow, 8, dtlDataRows, 1).setFormulaR1C1(`=IF(RC[-6]="","",IFERROR(INDEX(FILTER(Temp_Differences!R7C11:R1000C11,Temp_Differences!R7C2:R1000C2<>"",Temp_Differences!R7C7:R1000C7="DTL"),ROW()-${dtlOffset}),0))`);
  sheet.getRange(dtlDataStartRow, 7, dtlDataRows, 1).setFormulaR1C1('=IF(RC[-5]="","",RC[-3]-RC[-2])');
  sheet.getRange(dtlDataStartRow, 6, dtlDataRows, 1).setFormulaR1C1('=IF(RC[-4]="","",RC[-3]+RC[-2]-RC[-1]+IF(RC[2]="",0,RC[2]))');
  sheet.getRange(dtlDataStartRow, 9, dtlDataRows, 1).setFormulaR1C1('=IF(RC[-8]="","",RC[-2]+RC[-1])');

  sheet.getRange(`A${dtlSubtotalRow}:B${dtlSubtotalRow}`).merge()
    .setValue("Subtotal - DTL")
    .setFontWeight("bold")
    .setBackground(COLORS.CALC_BG);
  for (let col = 3; col <= 9; col++) {
    const colLetter = String.fromCharCode(64 + col);
    sheet.getRange(dtlSubtotalRow, col).setFormula(`=SUM(${colLetter}${dtlDataStartRow}:${colLetter}${dtlSubtotalRow - 1})`)
      .setFontWeight("bold")
      .setBackground(COLORS.CALC_BG);
  }

  // Net Position
  sheet.getRange(`A${netHeaderRow}:I${netHeaderRow}`).merge()
    .setValue("NET POSITION")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(COLORS.SECTION_BG);

  const netRows = [
    ["", "Net DTA/(DTL) - Before Netting", `=C${dtaSubtotalRow}-C${dtlSubtotalRow}`, `=D${dtaSubtotalRow}-D${dtlSubtotalRow}`, `=E${dtaSubtotalRow}-E${dtlSubtotalRow}`, `=F${dtaSubtotalRow}-F${dtlSubtotalRow}`, `=G${dtaSubtotalRow}-G${dtlSubtotalRow}`, `=H${dtaSubtotalRow}-H${dtlSubtotalRow}`, `=I${dtaSubtotalRow}-I${dtlSubtotalRow}`]
  ];

  sheet.getRange(netDataRow, 1, netRows.length, 9).setValues(netRows);
  sheet.getRange(`A${netDataRow}:B${netDataRow}`).merge().setFontWeight("bold");

  // Netting section (if applicable)
  sheet.getRange(`A${nettingHeaderRow}:I${nettingHeaderRow}`).merge()
    .setValue("NETTING OF DTA AND DTL (If Applicable per Framework)")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(COLORS.SECTION_BG);

  sheet.getRange(`A${applyNettingRow}:B${applyNettingRow}`).merge()
    .setValue("Apply Netting?")
    .setFontWeight("bold");
  sheet.getRange(`C${applyNettingRow}`).setFormula("=Assumptions!B36")
    .setFontWeight("bold")
    .setFontColor("#d32f2f");

  const nettingRows = [
    ["", "DTA (Gross)", `=C${dtaSubtotalRow}`, "", "", `=F${dtaSubtotalRow}`, "", "", ""],
    ["", "DTL (Gross)", `=C${dtlSubtotalRow}`, "", "", `=F${dtlSubtotalRow}`, "", "", ""],
    ["", "Amount Netted", `=IF($C${applyNettingRow}="Yes",MIN(C${nettingTableStartRow},C${nettingTableStartRow + 1}),0)`, "", "", `=IF($C${applyNettingRow}="Yes",MIN(F${nettingTableStartRow},F${nettingTableStartRow + 1}),0)`, "", "", ""],
    ["", "", "", "", "", "", "", "", ""]
  ];

  sheet.getRange(nettingTableStartRow, 1, nettingRows.length, 9).setValues(nettingRows);
  sheet.getRange(`A${nettingTableStartRow}:B${nettingTableStartRow + nettingRows.length - 1}`).mergeAcross().setFontWeight("bold");

  // Final Presentation
  sheet.getRange(`A${finalHeaderRow}:I${finalHeaderRow}`).merge()
    .setValue("FINAL PRESENTATION (Balance Sheet)")
    .setFontWeight("bold")
    .setFontSize(12)
    .setHorizontalAlignment("center")
    .setBackground(COLORS.GRAND_TOTAL_BG);

  const finalRows = [
    ["", "Deferred Tax Assets", `=C${dtaSubtotalRow}`, "", "", `=F${dtaSubtotalRow}`, "", "", `=I${dtaSubtotalRow}`],
    ["", "Deferred Tax Liabilities", `=C${dtlSubtotalRow}`, "", "", `=F${dtlSubtotalRow}`, "", "", `=I${dtlSubtotalRow}`],
    ["", "Net DTA/(DTL)", `=C${netDataRow}`, "", "", `=F${netDataRow}`, `=G${netDataRow}`, `=H${netDataRow}`, `=I${netDataRow}`]
  ];

  sheet.getRange(finalDataStartRow, 1, finalRows.length, 9).setValues(finalRows);
  sheet.getRange(`A${finalDataStartRow}:B${finalDataStartRow + finalRows.length - 1}`).merge().setFontWeight("bold").setBackground(COLORS.GRAND_TOTAL_BG);
  sheet.getRange(`C${finalDataStartRow}:I${finalDataStartRow + finalRows.length - 1}`).setFontWeight("bold").setBackground(COLORS.GRAND_TOTAL_BG);

  // Journal Entry section
  sheet.getRange(`A${journalHeaderRow}:I${journalHeaderRow}`).merge()
    .setValue("PERIOD CLOSURE JOURNAL ENTRY")
    .setFontSize(14)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT);

  const jeHeaders = [["Particulars", "Debit (₹)", "Credit (₹)"]];
  sheet.getRange(journalTableHeaderRow, 2, 1, 3).setValues(jeHeaders);
  sheet.getRange(`B${journalTableHeaderRow}:D${journalTableHeaderRow}`).setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");

  const debitCell = `I${finalDataStartRow}`;
  const creditCell = `I${finalDataStartRow + 1}`;
  const netCell = `I${finalDataStartRow + 2}`;

  const jeEntries = [
    ["Deferred Tax Asset A/c                              Dr.", `=IF(${netCell}>0,${netCell},"")`, ""],
    ["     To Deferred Tax Expense A/c", "", `=IF(${netCell}>0,${netCell},"")`],
    ["(Being DTA created/increased during the year)", "", ""],
    ["", "", ""],
    ["Deferred Tax Expense A/c                           Dr.", `=IF(${creditCell}>0,${creditCell},"")`, ""],
    ["     To Deferred Tax Liability A/c", "", `=IF(${creditCell}>0,${creditCell},"")`],
    ["(Being DTL created/increased during the year)", "", ""],
    ["", "", ""],
    ["TOTAL", "=SUM(C${journalTableHeaderRow + 1},C${journalTableHeaderRow + 5})", "=SUM(D${journalTableHeaderRow + 2},D${journalTableHeaderRow + 6})"]
  ];

  sheet.getRange(journalTableHeaderRow + 1, 2, jeEntries.length, 3).setValues(jeEntries);
  sheet.getRange(`B${journalTableHeaderRow + jeEntries.length}:D${journalTableHeaderRow + jeEntries.length}`).setBackground(COLORS.TOTAL_BG).setFontWeight("bold");

  // Number formatting
  sheet.getRange(`C${dtaDataStartRow}:I${finalDataStartRow + finalRows.length - 1}`).setNumberFormat("#,##0");
  sheet.getRange(`C${journalTableHeaderRow + 1}:D${journalTableHeaderRow + jeEntries.length}`).setNumberFormat("#,##0");

  // Borders
  sheet.getRange(`A5:I${finalDataStartRow + finalRows.length - 1}`).setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(`B${journalTableHeaderRow}:D${journalTableHeaderRow + jeEntries.length}`).setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);

  // Freeze rows only (columns removed to avoid splitting merged cells)
  sheet.setFrozenRows(5);
}


// ============================================================================
// P&L RECONCILIATION SHEET
// ============================================================================

function createPLReconciliationSheet(ss) {
  let sheet = ss.getSheetByName("PL_Reconciliation");
  if (!sheet) {
    sheet = ss.insertSheet("PL_Reconciliation", 6);
  }
  
  sheet.clear();
  setColumnWidths(sheet, [40, 350, 150, 200]);
  
  // Header
  sheet.getRange("A1:D1").merge()
    .setValue("P&L RECONCILIATION - TAX EXPENSE ANALYSIS")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT);
  
  sheet.getRange("A2:D2").merge()
    .setValue("Reconciliation of accounting profit to tax expense and effective tax rate calculation")
    .setFontSize(10)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.SUBHEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT);
  
  // Tax Expense Components
  sheet.getRange("A4:D4").merge()
    .setValue("TAX EXPENSE COMPONENTS")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORS.SECTION_BG);
  
  const taxComponents = [
    ["Particular", "Amount (₹)", "Notes"],
    ["Current Tax Expense", "=Assumptions!B26", "As per tax computation"],
    ["Deferred Tax Expense/(Income)", "=Movement_Analysis!G67", "Net change in deferred tax"],
    ["", "", ""],
    ["Total Tax Expense (Current + Deferred)", "=C6+C7", "Total charge to P&L"]
  ];
  
  sheet.getRange(5, 2, taxComponents.length, 3).setValues(taxComponents);
  sheet.getRange("B5:D5").setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  
  sheet.getRange("B9:C9").setBackground(COLORS.TOTAL_BG)
    .setFontWeight("bold");
  
  // Effective Tax Rate Calculation
  sheet.getRange("A12:D12").merge()
    .setValue("EFFECTIVE TAX RATE ANALYSIS")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORS.SECTION_BG);
  
  const etrCalc = [
    ["Particular", "Amount (₹)", "Notes"],
    ["Profit Before Tax (PBT)", "=Assumptions!B24", "As per P&L statement"],
    ["Total Tax Expense", "=C9", "Current + Deferred"],
    ["Effective Tax Rate (ETR)", "=IF(C14<>0,C15/C14,0)", "ETR = Tax Expense / PBT"],
    ["", "", ""],
    ["Statutory Tax Rate", "=Assumptions!B13", "Applicable statutory rate"],
    ["ETR vs Statutory Rate Variance", "=C16-C18", "Difference to be explained"]
  ];
  
  sheet.getRange(13, 2, etrCalc.length, 3).setValues(etrCalc);
  sheet.getRange("B13:D13").setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  
  sheet.getRange("C16").setNumberFormat("0.00%").setFontWeight("bold");
  sheet.getRange("C18:C19").setNumberFormat("0.00%");
  
  // Reconciliation of ETR to Statutory Rate
  sheet.getRange("A23:D23").merge()
    .setValue("RECONCILIATION: STATUTORY RATE TO EFFECTIVE RATE")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORS.SECTION_BG);
  
  const etrReconciliation = [
    ["Reconciliation Item", "Amount (₹)", "Rate Impact (%)"],
    ["Accounting Profit Before Tax (A)", "=C14", ""],
    ["", "", ""],
    ["Tax at Statutory Rate (B)", "=C25*Assumptions!B13", "=C26/C25"],
    ["", "", ""],
    ["Add/(Less): Permanent Differences", "", ""],
    ["   - Disallowances (e.g., CSR, penalties)", 100000, "=C29/C25"],
    ["   - Exempt income (e.g., dividend)", -50000, "=C30/C25"],
    ["   - Other permanent differences", 25000, "=C31/C25"],
    ["", "", ""],
    ["Tax Impact of Permanent Differences (C)", "=SUM(C29:C31)", "=C33/C25"],
    ["", "", ""],
    ["Tax at Effective Rate (D = B + C)", "=C27+C33", "=C35/C25"],
    ["", "", ""],
    ["Actual Tax Expense (per books) (E)", "=C9", "=C37/C25"],
    ["", "", ""],
    ["Difference (E - D)", "=C37-C35", "Should be minimal"],
    ["", "", ""],
    ["Effective Tax Rate (ETR)", "=C37/C25", "Final ETR"]
  ];
  
  sheet.getRange(24, 2, etrReconciliation.length, 3).setValues(etrReconciliation);
  sheet.getRange("B24:D24").setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  
  sheet.getRange("C29:C31").setBackground(COLORS.INPUT_BG);
  
  sheet.getRange("B27:D27").setBackground(COLORS.CALC_BG).setFontWeight("bold");
  sheet.getRange("B33:D33").setBackground(COLORS.CALC_BG).setFontWeight("bold");
  sheet.getRange("B35:D35").setBackground(COLORS.TOTAL_BG).setFontWeight("bold");
  sheet.getRange("B41:D41").setBackground(COLORS.GRAND_TOTAL_BG).setFontWeight("bold");
  
  sheet.getRange("D27:D41").setNumberFormat("0.00%");
  sheet.getRange("D41").setFontWeight("bold").setFontSize(11);
  
  // Variance Analysis
  sheet.getRange("A45:D45").merge()
    .setValue("VARIANCE ANALYSIS: CURRENT YEAR vs PRIOR YEAR")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORS.SECTION_BG);
  
  const varianceAnalysis = [
    ["Metric", "Current Year", "Prior Year", "Variance"],
    ["Profit Before Tax", "=Assumptions!B24", "=Assumptions!B25", "=C47-D47"],
    ["Current Tax", "=C6", "", ""],
    ["Deferred Tax", "=C7", "", ""],
    ["Total Tax Expense", "=C9", "", ""],
    ["Effective Tax Rate", "=C16", "", ""]
  ];
  
  sheet.getRange(46, 2, varianceAnalysis.length, 4).setValues(varianceAnalysis);
  sheet.getRange("B46:E46").setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  
  sheet.getRange("C51:E51").setNumberFormat("0.00%");
  
  // Number formatting
  sheet.getRange("C6:C9").setNumberFormat("#,##0");
  sheet.getRange("C14:C15").setNumberFormat("#,##0");
  sheet.getRange("C25:D41").setNumberFormat("#,##0");
  sheet.getRange("C47:E50").setNumberFormat("#,##0");
  
  // Borders
  sheet.getRange("B5:D9").setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("B13:D19").setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("B24:D41").setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("B46:E51").setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  
  // Freeze rows only (columns removed to avoid splitting merged cells)
  sheet.setFrozenRows(2);
}

// ============================================================================
// BALANCE SHEET RECONCILIATION SHEET
// ============================================================================

function createBSReconciliationSheet(ss) {
  let sheet = ss.getSheetByName("BS_Reconciliation");
  if (!sheet) {
    sheet = ss.insertSheet("BS_Reconciliation", 7);
  }
  
  sheet.clear();
  setColumnWidths(sheet, [40, 350, 150, 150, 200]);
  
  // Header
  sheet.getRange("A1:E1").merge()
    .setValue("BALANCE SHEET RECONCILIATION")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT);
  
  sheet.getRange("A2:E2").merge()
    .setValue("Presentation of Deferred Tax Assets and Liabilities in Balance Sheet")
    .setFontSize(10)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.SUBHEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT);
  
  // Balance Sheet Presentation
  sheet.getRange("A4:E4").merge()
    .setValue("BALANCE SHEET PRESENTATION")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORS.SECTION_BG);
  
  const bsPresentation = [
    ["Line Item", "Current Year (₹)", "Prior Year (₹)", "Note Reference"],
    ["NON-CURRENT ASSETS", "", "", ""],
    ["Deferred Tax Assets (Net)", "=Movement_Analysis!F65", "=Movement_Analysis!C65", "Note X - Deferred Taxation"],
    ["", "", "", ""],
    ["NON-CURRENT LIABILITIES", "", "", ""],
    ["Deferred Tax Liabilities (Net)", "=Movement_Analysis!F66", "=Movement_Analysis!C66", "Note X - Deferred Taxation"],
    ["", "", "", ""],
    ["NET DEFERRED TAX POSITION", "=C7-C10", "=D7-D10", ""]
  ];
  
  sheet.getRange(5, 2, bsPresentation.length, 4).setValues(bsPresentation);
  sheet.getRange("B5:E5").setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  
  sheet.getRange("B6:E6").setBackground(COLORS.CALC_BG).setFontWeight("bold");
  sheet.getRange("B9:E9").setBackground(COLORS.CALC_BG).setFontWeight("bold");
  sheet.getRange("B12:E12").setBackground(COLORS.TOTAL_BG).setFontWeight("bold");
  
  // Reconciliation with Schedule
  sheet.getRange("A15:E15").merge()
    .setValue("RECONCILIATION WITH DETAILED SCHEDULE")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORS.SECTION_BG);
  
  const scheduleRecon = [
    ["Source", "DTA (₹)", "DTL (₹)", "Net DTA/(DTL) (₹)"],
    ["Per Movement Analysis Schedule", "=Movement_Analysis!F65", "=Movement_Analysis!F66", "=Movement_Analysis!F67"],
    ["Per Balance Sheet (above)", "=C7", "=C10", "=C12"],
    ["", "", "", ""],
    ["Difference (Should be NIL)", "=C17-C18", "=D17-D18", "=E17-E18"]
  ];
  
  sheet.getRange(16, 2, scheduleRecon.length, 4).setValues(scheduleRecon);
  sheet.getRange("B16:E16").setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  
  sheet.getRange("B20:E20").setBackground(COLORS.WARNING_BG).setFontWeight("bold");
  
  // Conditional formatting for differences
  const diffRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberNotEqualTo(0)
    .setBackground("#ff0000")
    .setFontColor("#ffffff")
    .setRanges([sheet.getRange("C20:E20")])
    .build();
  const rules = sheet.getConditionalFormatRules();
  rules.push(diffRule);
  sheet.setConditionalFormatRules(rules);
  
  // Netting Disclosure
  sheet.getRange("A23:E23").merge()
    .setValue("NETTING DISCLOSURE")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORS.SECTION_BG);
  
  const nettingDisclosure = [
    ["Netting Status", "Amount (₹)", "Framework Compliance"],
    ["DTA (Gross before netting)", "=Movement_Analysis!C58", ""],
    ["DTL (Gross before netting)", "=Movement_Analysis!C59", ""],
    ["Amount Netted", "=Movement_Analysis!C60", ""],
    ["", "", ""],
    ["Netting Applied?", "=Assumptions!B36", ""],
    ["Framework", "=Assumptions!B7", ""],
    ["", "", ""],
    ["Netting Criteria:", "", ""],
    ["  - Legally enforceable right to set off", "Yes/No", "User to verify"],
    ["  - Intention to settle net or realize simultaneously", "Yes/No", "User to verify"],
    ["", "", ""],
    ["Note:", "Netting is permitted under both IGAAP and Ind AS if:", ""],
    ["", "1. Legally enforceable right to set off exists", ""],
    ["", "2. DTA and DTL relate to same taxing authority", ""],
    ["", "3. Entity intends to settle on net basis", ""]
  ];
  
  sheet.getRange(24, 2, nettingDisclosure.length, 3).setValues(nettingDisclosure);
  sheet.getRange("B24:D24").setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  
  sheet.getRange("C30:C31").setBackground(COLORS.INPUT_ALT_BG);
  
  // Balance Sheet Note Template
  sheet.getRange("A43:E43").merge()
    .setValue("NOTE DISCLOSURE TEMPLATE")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORS.SECTION_BG);
  
  const noteTemplate = [
    ["Note X: DEFERRED TAXATION", "", ""],
    ["", "", ""],
    ["The major components of deferred tax assets and liabilities are as follows:", "", ""],
    ["", "", ""],
    ["Particular", "As at " + "=TEXT(Assumptions!B8,\"DD-MMM-YYYY\")", "As at " + "=TEXT(Assumptions!B9,\"DD-MMM-YYYY\")"],
    ["Deferred Tax Assets:", "", ""],
    ["  - Provision for employee benefits", "=Movement_Analysis!F9", "=Movement_Analysis!C9"],
    ["  - Provision for doubtful debts", "=Movement_Analysis!F8", "=Movement_Analysis!C8"],
    ["  - Disallowances u/s 43B", "=Movement_Analysis!F10", "=Movement_Analysis!C10"],
    ["  - Carry forward losses", "=Movement_Analysis!F11", "=Movement_Analysis!C11"],
    ["", "", ""],
    ["Deferred Tax Liabilities:", "", ""],
    ["  - Depreciation differences", "=Movement_Analysis!F15", "=Movement_Analysis!C15"],
    ["  - Other timing differences", "=Movement_Analysis!F16", "=Movement_Analysis!C16"],
    ["", "", ""],
    ["Net Deferred Tax Asset/(Liability)", "=Movement_Analysis!F67", "=Movement_Analysis!C67"]
  ];
  
  sheet.getRange(44, 2, noteTemplate.length, 3).setValues(noteTemplate);
  sheet.getRange("B44:D44").setFontWeight("bold").setFontSize(11);
  sheet.getRange("B48:D48").setBackground(COLORS.SUBHEADER_BG)
    .setFontWeight("bold");
  sheet.getRange("B60:D60").setBackground(COLORS.TOTAL_BG)
    .setFontWeight("bold");
  
  // Number formatting
  sheet.getRange("C7:E12").setNumberFormat("#,##0");
  sheet.getRange("C17:E20").setNumberFormat("#,##0");
  sheet.getRange("C25:C27").setNumberFormat("#,##0");
  sheet.getRange("C49:D60").setNumberFormat("#,##0");
  
  // Borders
  sheet.getRange("B5:E12").setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("B16:E20").setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("B24:D40").setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("B48:D60").setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  
  // Freeze rows only (columns removed to avoid splitting merged cells)
  sheet.setFrozenRows(2);
}

// ============================================================================
// REFERENCES SHEET
// ============================================================================

function createReferencesSheet(ss) {
  let sheet = ss.getSheetByName("References");
  if (!sheet) {
    sheet = ss.insertSheet("References", 8);
  }
  
  sheet.clear();
  setColumnWidths(sheet, [50, 200, 600]);
  
  // Header
  sheet.getRange("A1:C1").merge()
    .setValue("ACCOUNTING STANDARDS REFERENCE")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT);
  
  sheet.getRange("A2:C2").merge()
    .setValue("IGAAP (AS 22) vs Ind AS (Ind AS 12) - Key Guidance and Differences")
    .setFontSize(10)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.SUBHEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT);
  
  // IGAAP Section
  sheet.getRange("A4:C4").merge()
    .setValue("IGAAP: AS 22 - ACCOUNTING FOR TAXES ON INCOME")
    .setFontWeight("bold")
    .setFontSize(13)
    .setBackground(COLORS.SECTION_BG);
  
  const igaapGuidance = [
    ["Topic", "AS 22 Requirement", "Key Points"],
    ["", "", ""],
    ["Recognition", "Timing differences result in deferred tax. DTA recognized only with reasonable certainty (virtual certainty for losses).", "• Conservative approach\n• High threshold for DTA on losses\n• Requires convincing evidence"],
    ["", "", ""],
    ["Measurement", "Deferred tax measured at tax rates enacted or substantively enacted by balance sheet date.", "• Use enacted rates\n• No discounting allowed\n• Tax rate as at reporting date"],
    ["", "", ""],
    ["Carry Forward Losses", "DTA on unabsorbed depreciation and carry forward losses recognized ONLY if virtual certainty backed by convincing evidence.", "• Very stringent criteria\n• Requires detailed business plans\n• Conservative recognition"],
    ["", "", ""],
    ["Presentation", "Deferred tax assets and liabilities shown separately. Netting allowed if legally enforceable right exists.", "• Separate line items\n• Netting with conditions\n• Classification as non-current"],
    ["", "", ""],
    ["Disclosure", "Break-up of DTA and DTL by major components. Reconciliation not mandatory.", "• Component-wise disclosure\n• Opening-closing reconciliation\n• Expiry date of losses"]
  ];
  
  sheet.getRange(5, 1, igaapGuidance.length, 3).setValues(igaapGuidance);
  sheet.getRange("A5:C5").setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  
  // Ind AS Section
  const indAsStartRow = 5 + igaapGuidance.length + 2;
  sheet.getRange(`A${indAsStartRow}:C${indAsStartRow}`).merge()
    .setValue("IND AS: IND AS 12 - INCOME TAXES")
    .setFontWeight("bold")
    .setFontSize(13)
    .setBackground(COLORS.SECTION_BG);
  
  const indAsGuidance = [
    ["Topic", "Ind AS 12 Requirement", "Key Points"],
    ["", "", ""],
    ["Recognition", "Temporary differences (not timing differences) result in deferred tax. DTA recognized when probable that taxable profit available.", "• Based on temporary differences\n• Probable = >50% likelihood\n• More liberal than AS 22"],
    ["", "", ""],
    ["Measurement", "Deferred tax measured at tax rates expected to apply when asset realized or liability settled (substantively enacted rates).", "• Substantively enacted rates\n• No discounting (except specific cases)\n• Future rate expectations"],
    ["", "", ""],
    ["Unused Tax Losses", "DTA recognized for unused tax losses and credits to extent probable that future taxable profit available.", "• Probable future profits required\n• Detailed assessment needed\n• Less stringent than AS 22"],
    ["", "", ""],
    ["Presentation", "Offset DTA and DTL if legally enforceable right and they relate to same taxation authority.", "• Offset is common\n• Single net line\n• Non-current classification"],
    ["", "", ""],
    ["Disclosure", "Extensive disclosures including reconciliation of accounting profit to tax expense, major temporary differences, unused losses.", "• Detailed reconciliation required\n• Tax rate reconciliation\n• Extensive note disclosures"]
  ];
  
  sheet.getRange(indAsStartRow + 1, 1, indAsGuidance.length, 3).setValues(indAsGuidance);
  sheet.getRange(`A${indAsStartRow + 1}:C${indAsStartRow + 1}`).setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  
  // Key Differences Section
  const diffStartRow = indAsStartRow + indAsGuidance.length + 2;
  sheet.getRange(`A${diffStartRow}:C${diffStartRow}`).merge()
    .setValue("KEY DIFFERENCES: IGAAP vs IND AS")
    .setFontWeight("bold")
    .setFontSize(13)
    .setBackground(COLORS.GRAND_TOTAL_BG);
  
  const keyDifferences = [
    ["Aspect", "IGAAP (AS 22)", "Ind AS (Ind AS 12)"],
    ["", "", ""],
    ["Basis", "Timing differences", "Temporary differences"],
    ["DTA Recognition Threshold", "Reasonable certainty (virtual certainty for losses)", "Probable (>50% likelihood)"],
    ["Tax Rate", "Enacted or substantively enacted", "Expected rate (substantively enacted)"],
    ["Unabsorbed Losses", "Virtual certainty required", "Probable future profits sufficient"],
    ["Initial Recognition Exemption", "Not applicable", "Exists for goodwill and certain transactions"],
    ["Investments in Subsidiaries", "Not specifically addressed", "Temporary differences recognized unless specific criteria met"],
    ["", "", ""],
    ["PRACTICAL IMPACT:", "", ""],
    ["DTA Recognition", "More conservative, lower DTA amounts", "More liberal, higher DTA possible"],
    ["Loss Recognition", "Rarely recognized without strong evidence", "More commonly recognized"],
    ["Transition Impact", "Lower DTA balances", "Typically increases DTA on transition"]
  ];
  
  sheet.getRange(diffStartRow + 1, 1, keyDifferences.length, 3).setValues(keyDifferences);
  sheet.getRange(`A${diffStartRow + 1}:C${diffStartRow + 1}`).setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  
  sheet.getRange(`A${diffStartRow + 9}:C${diffStartRow + 9}`).setBackground(COLORS.INFO_BG)
    .setFontWeight("bold");
  
  // Common Temporary Differences Section
  const commonStartRow = diffStartRow + keyDifferences.length + 2;
  sheet.getRange(`A${commonStartRow}:C${commonStartRow}`).merge()
    .setValue("COMMON TEMPORARY DIFFERENCES")
    .setFontWeight("bold")
    .setFontSize(13)
    .setBackground(COLORS.SECTION_BG);
  
  const commonDiffs = [
    ["Temporary Difference Type", "Creates", "Explanation"],
    ["", "", ""],
    ["Depreciation - Book vs Tax", "DTA or DTL", "Different depreciation rates/methods create timing difference"],
    ["Provisions (Doubtful Debts, Warranty)", "DTA", "Deductible when actually written off or paid"],
    ["Employee Benefits (Gratuity, Leave)", "DTA", "Deductible on payment basis, not accrual"],
    ["Section 43B Disallowances", "DTA", "Statutory dues deductible on payment"],
    ["Revenue Recognition Differences", "DTA or DTL", "Different recognition criteria for book vs tax"],
    ["Carry Forward Business Losses", "DTA", "Subject to probability assessment"],
    ["Unabsorbed Depreciation", "DTA", "Subject to probability assessment"],
    ["Fair Value Adjustments", "DTA or DTL", "Unrealized gains/losses create temporary differences"],
    ["Lease Accounting (Ind AS 116)", "DTA or DTL", "ROU asset vs tax depreciation timing"],
    ["Financial Instruments (Ind AS 109)", "DTA or DTL", "Fair value changes vs cost-based tax treatment"]
  ];
  
  sheet.getRange(commonStartRow + 1, 1, commonDiffs.length, 3).setValues(commonDiffs);
  sheet.getRange(`A${commonStartRow + 1}:C${commonStartRow + 1}`).setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  
  // Wrap text in column C
  sheet.getRange("C5:C" + (commonStartRow + commonDiffs.length)).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
  // Borders for all sections
  sheet.getRange("A5:C" + (5 + igaapGuidance.length - 1)).setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(`A${indAsStartRow + 1}:C${indAsStartRow + indAsGuidance.length}`).setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(`A${diffStartRow + 1}:C${diffStartRow + keyDifferences.length}`).setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(`A${commonStartRow + 1}:C${commonStartRow + commonDiffs.length}`).setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  
  // Freeze
  sheet.setFrozenRows(2);
}

// ============================================================================
// AUDIT NOTES SHEET
// ============================================================================

function createAuditNotesSheet(ss) {
  let sheet = ss.getSheetByName("Audit_Notes");
  if (!sheet) {
    sheet = ss.insertSheet("Audit_Notes", 9);
  }
  
  sheet.clear();
  setColumnWidths(sheet, [50, 300, 200, 150, 250]);
  
  // Header
  sheet.getRange("A1:E1").merge()
    .setValue("AUDIT NOTES & CONTROL CHECKS")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT);
  
  sheet.getRange("A2:E2").merge()
    .setValue("Audit Assertions, Control Totals, and Review Points")
    .setFontSize(10)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.SUBHEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT);
  
  // Control Totals Section
  sheet.getRange("A4:E4").merge()
    .setValue("CONTROL TOTALS")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORS.SECTION_BG);
  
  const controlTotals = [
    ["Control Check", "Amount/Status", "Expected", "Status", "Comments"],
    ["", "", "", "", ""],
    ["DTA per Movement Schedule", "=Movement_Analysis!F65", "=BS_Reconciliation!C7", "=IF(C7=D7,\"✓ OK\",\"⚠ MISMATCH\")", "Must match BS presentation"],
    ["DTL per Movement Schedule", "=Movement_Analysis!F66", "=BS_Reconciliation!C10", "=IF(C8=D8,\"✓ OK\",\"⚠ MISMATCH\")", "Must match BS presentation"],
    ["Net DTA/(DTL)", "=Movement_Analysis!F67", "=BS_Reconciliation!C12", "=IF(C9=D9,\"✓ OK\",\"⚠ MISMATCH\")", "Must match BS presentation"],
    ["", "", "", "", ""],
    ["Deferred Tax Expense per Movement", "=Movement_Analysis!G67", "=PL_Reconciliation!C7", "=IF(C11=D11,\"✓ OK\",\"⚠ MISMATCH\")", "Must match P&L reconciliation"],
    ["", "", "", "", ""],
    ["Total Tax Expense Check", "=PL_Reconciliation!C9", "=PL_Reconciliation!C6+PL_Reconciliation!C7", "=IF(C13=D13,\"✓ OK\",\"⚠ MISMATCH\")", "Current + Deferred"],
    ["", "", "", "", ""],
    ["Opening + Movement = Closing?", "=Movement_Analysis!F67", "=Movement_Analysis!C67+Movement_Analysis!I67", "=IF(C15=D15,\"✓ OK\",\"⚠ MISMATCH\")", "Movement reconciliation check"]
  ];
  
  sheet.getRange(5, 1, controlTotals.length, 5).setValues(controlTotals);
  sheet.getRange("A5:E5").setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  
  sheet.getRange("C7:C15").setNumberFormat("#,##0");
  sheet.getRange("D7:D15").setNumberFormat("#,##0");
  
  // Conditional formatting for status
  const statusRangeOK = sheet.getRange("E7:E15");
  const statusRuleOK = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("✓ OK")
    .setBackground(COLORS.SUCCESS_BG)
    .setFontColor("#2e7d32")
    .setRanges([statusRangeOK])
    .build();
  
  const statusRuleMismatch = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("⚠ MISMATCH")
    .setBackground(COLORS.WARNING_BG)
    .setFontColor("#c62828")
    .setRanges([statusRangeOK])
    .build();
  
  const rules = sheet.getConditionalFormatRules();
  rules.push(statusRuleOK);
  rules.push(statusRuleMismatch);
  sheet.setConditionalFormatRules(rules);
  
  // Audit Assertions Section
  const assertionStartRow = 21;
  sheet.getRange(`A${assertionStartRow}:E${assertionStartRow}`).merge()
    .setValue("AUDIT ASSERTIONS CHECKLIST")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORS.SECTION_BG);
  
  const assertions = [
    ["Assertion", "Check Performed", "Status", "Evidence", "Reviewer Notes"],
    ["", "", "", "", ""],
    ["EXISTENCE", "Verify that recognized DTA/DTL exist and pertain to the entity", "", "Tax computation, assessment orders", ""],
    ["", "", "", "", ""],
    ["COMPLETENESS", "All temporary differences identified and recorded", "", "Detailed trial balance review, tax computation", ""],
    ["", "", "", "", ""],
    ["ACCURACY", "DTA/DTL calculated correctly using appropriate tax rates", "", "Tax rate verification, formula checks", ""],
    ["", "", "", "", ""],
    ["VALUATION", "DTA recognized only when probable future taxable profits available", "", "Business forecasts, tax planning strategies", ""],
    ["", "", "", "", ""],
    ["CLASSIFICATION", "DTA and DTL properly classified as non-current", "", "Balance sheet presentation review", ""],
    ["", "", "", "", ""],
    ["PRESENTATION", "Adequate disclosure per AS 22/Ind AS 12 requirements", "", "Note disclosure review", ""],
    ["", "", "", "", ""],
    ["CUT-OFF", "Temporary differences recorded in correct period", "", "Period-end transaction review", ""]
  ];
  
  sheet.getRange(assertionStartRow + 1, 1, assertions.length, 5).setValues(assertions);
  sheet.getRange(`A${assertionStartRow + 1}:E${assertionStartRow + 1}`).setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  
  sheet.getRange(`C${assertionStartRow + 3}:C${assertionStartRow + assertions.length}`).setBackground(COLORS.INPUT_ALT_BG);
  sheet.getRange(`E${assertionStartRow + 3}:E${assertionStartRow + assertions.length}`).setBackground("#ffffff");
  
  // Review Points Section
  const reviewStartRow = assertionStartRow + assertions.length + 2;
  sheet.getRange(`A${reviewStartRow}:E${reviewStartRow}`).merge()
    .setValue("KEY REVIEW POINTS")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORS.SECTION_BG);
  
  const reviewPoints = [
    ["#", "Review Point", "Checked?", "Finding", "Action Required"],
    ["", "", "", "", ""],
    [1, "Are all temporary differences between book and tax basis identified?", "", "", ""],
    [2, "Is the tax rate used appropriate (enacted/substantively enacted)?", "", "", ""],
    [3, "For DTA on losses - is probability assessment documented?", "", "", ""],
    [4, "Are business forecasts/projections supporting DTA recognition available?", "", "", ""],
    [5, "Is framework (IGAAP vs Ind AS) consistently applied?", "", "", ""],
    [6, "Are permanent differences excluded from deferred tax calculation?", "", "", ""],
    [7, "Is netting of DTA/DTL appropriate and documented?", "", "", ""],
    [8, "Are disclosures complete per applicable standards?", "", "", ""],
    [9, "Is effective tax rate reconciliation prepared and explained?", "", "", ""],
    [10, "Are prior period adjustments, if any, properly disclosed?", "", "", ""],
    [11, "Have tax positions been discussed with tax consultants?", "", "", ""],
    [12, "Is there consistency with tax computation and return filing?", "", "", ""]
  ];
  
  sheet.getRange(reviewStartRow + 1, 1, reviewPoints.length, 5).setValues(reviewPoints);
  sheet.getRange(`A${reviewStartRow + 1}:E${reviewStartRow + 1}`).setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  
  sheet.getRange(`C${reviewStartRow + 3}:C${reviewStartRow + reviewPoints.length}`).setBackground(COLORS.INPUT_ALT_BG);
  sheet.getRange(`D${reviewStartRow + 3}:E${reviewStartRow + reviewPoints.length}`).setBackground("#ffffff");
  
  // Documentation Section
  const docStartRow = reviewStartRow + reviewPoints.length + 2;
  sheet.getRange(`A${docStartRow}:E${docStartRow}`).merge()
    .setValue("DOCUMENTATION & EVIDENCE")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORS.SECTION_BG);
  
  const documentation = [
    ["Document Type", "Description", "Obtained?", "File Reference", "Notes"],
    ["", "", "", "", ""],
    ["Tax Computation", "Income tax computation for the year", "", "", ""],
    ["Tax Assessment", "Prior year assessment orders", "", "", ""],
    ["Business Forecasts", "Supporting DTA recognition (minimum 3-5 years)", "", "", ""],
    ["Board Minutes", "Approval of tax strategies/planning", "", "", ""],
    ["Tax Consultant Opinion", "Opinion on deferred tax positions", "", "", ""],
    ["Management Representation", "Letter confirming temporary differences", "", "", ""],
    ["Fixed Asset Register", "For depreciation differences", "", "", ""],
    ["Provision Schedules", "Employee benefits, doubtful debts, etc.", "", "", ""]
  ];
  
  sheet.getRange(docStartRow + 1, 1, documentation.length, 5).setValues(documentation);
  sheet.getRange(`A${docStartRow + 1}:E${docStartRow + 1}`).setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  
  sheet.getRange(`C${docStartRow + 3}:C${docStartRow + documentation.length}`).setBackground(COLORS.INPUT_ALT_BG);
  sheet.getRange(`D${docStartRow + 3}:E${docStartRow + documentation.length}`).setBackground("#ffffff");
  
  // Borders
  sheet.getRange("A5:E" + (5 + controlTotals.length - 1)).setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(`A${assertionStartRow + 1}:E${assertionStartRow + assertions.length}`).setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(`A${reviewStartRow + 1}:E${reviewStartRow + reviewPoints.length}`).setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(`A${docStartRow + 1}:E${docStartRow + documentation.length}`).setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  
  // Freeze rows only (columns removed to avoid splitting merged cells)
  sheet.setFrozenRows(2);
}

// ============================================================================
// NAMED RANGES SETUP
// ============================================================================

function setupNamedRanges(ss) {
  // Key input ranges
  try {
    ss.setNamedRange("Framework", ss.getSheetByName("Assumptions").getRange("B7"));
    ss.setNamedRange("CurrentTaxRate", ss.getSheetByName("Assumptions").getRange("B13"));
    ss.setNamedRange("DeferredTaxRate", ss.getSheetByName("Assumptions").getRange("B14"));
    ss.setNamedRange("PBT_Current", ss.getSheetByName("Assumptions").getRange("B24"));
    ss.setNamedRange("Opening_DTA", ss.getSheetByName("Assumptions").getRange("B28"));
    ss.setNamedRange("Opening_DTL", ss.getSheetByName("Assumptions").getRange("B29"));
    
    // Output ranges
    ss.setNamedRange("Closing_DTA", ss.getSheetByName("Movement_Analysis").getRange("F28"));
    ss.setNamedRange("Closing_DTL", ss.getSheetByName("Movement_Analysis").getRange("F51"));
    ss.setNamedRange("Net_DTA_DTL", ss.getSheetByName("Movement_Analysis").getRange("F54"));
    ss.setNamedRange("DeferredTaxExpense", ss.getSheetByName("PL_Reconciliation").getRange("C7"));
    ss.setNamedRange("EffectiveTaxRate", ss.getSheetByName("PL_Reconciliation").getRange("C16"));
    
    Logger.log("Named ranges set up successfully");
  } catch (error) {
    Logger.log("Error setting up named ranges: " + error.toString());
  }
}

// ============================================================================
// FINAL FORMATTING
// ============================================================================

function applyFinalFormatting(ss) {
  const sheets = ss.getSheets();

  sheets.forEach(sheet => {
    // Set default font per 109 guide professional standards
    sheet.getDataRange().setFontFamily("Arial").setFontSize(10);

    // Hide gridlines for clean, professional sleek appearance
    sheet.setHiddenGridlines(true);
  });

  Logger.log("Final formatting applied - gridlines hidden for professional appearance");
}

// ============================================================================
// MENU FUNCTIONS
// ============================================================================

// onOpen() is handled by common/utilities.gs - auto-detects workbook type
    .addSeparator()
    .addItem('ℹ️ Help & Instructions', 'showHelp')
    .addToUi();
}

function refreshAllFormulas() {
  SpreadsheetApp.flush();
  SpreadsheetApp.getActiveSpreadsheet().toast("All formulas refreshed successfully!", "✓ Complete", 3);
}

function recalculateControls() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet = ss.getSheetByName("Audit_Notes");
  
  if (auditSheet) {
    // Force recalculation by getting values
    auditSheet.getRange("C7:C15").getValues();
    SpreadsheetApp.getActiveSpreadsheet().toast("Control totals recalculated!", "✓ Complete", 3);
  } else {
    SpreadsheetApp.getUi().alert("Audit_Notes sheet not found!");
  }
}

function addTempDiffRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Temp_Differences");
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Temp_Differences sheet not found!");
    return;
  }
  
  // Find the last row with data
  const lastRow = sheet.getLastRow();
  
  if (lastRow >= 50) {
    SpreadsheetApp.getUi().alert("Maximum rows reached. Please extend the range manually.");
    return;
  }
  
  // Insert a new row after the last data row
  sheet.insertRowAfter(lastRow);
  
  // Copy formatting and formulas from previous row
  const sourceRow = sheet.getRange(lastRow, 1, 1, 11);
  const targetRow = sheet.getRange(lastRow + 1, 1, 1, 11);
  sourceRow.copyTo(targetRow);
  
  // Clear input values but keep formulas
  sheet.getRange(lastRow + 1, 2).clearContent(); // Item name
  sheet.getRange(lastRow + 1, 3).clearContent(); // Category
  sheet.getRange(lastRow + 1, 4).clearContent(); // Tax Base
  sheet.getRange(lastRow + 1, 5).clearContent(); // Book Base
  sheet.getRange(lastRow + 1, 8).clearContent(); // Additions
  sheet.getRange(lastRow + 1, 9).clearContent(); // Reversals
  sheet.getRange(lastRow + 1, 10).clearContent(); // Rate change
  sheet.getRange(lastRow + 1, 11).clearContent(); // Remarks
  
  // Update serial number
  sheet.getRange(lastRow + 1, 1).setValue(lastRow - 5);
  
  SpreadsheetApp.getActiveSpreadsheet().toast("New temporary difference row added at row " + (lastRow + 1), "✓ Complete", 3);
  sheet.setActiveRange(sheet.getRange(lastRow + 1, 2));
}

function showHelp() {
  const ui = SpreadsheetApp.getUi();
  const helpText = 
    "DEFERRED TAXATION WORKINGS - HELP\n\n" +
    "This workbook helps you prepare comprehensive deferred tax schedules compliant with IGAAP (AS 22) and Ind AS (Ind AS 12).\n\n" +
    "QUICK START:\n" +
    "1. Go to 'Assumptions' sheet\n" +
    "2. Enter entity name, period, and framework\n" +
    "3. Enter tax rates\n" +
    "4. Go to 'Temp_Differences' sheet\n" +
    "5. Enter temporary differences\n" +
    "6. Review 'DT_Schedule' for calculations\n" +
    "7. Check 'Audit_Notes' for control checks\n\n" +
    "KEY FEATURES:\n" +
    "• Framework toggle (IGAAP/Ind AS)\n" +
    "• Automatic DTA/DTL calculation\n" +
    "• Movement analysis\n" +
    "• P&L and BS reconciliation\n" +
    "• Journal entries\n" +
    "• Control totals\n" +
    "• Audit assertions\n\n" +
    "SUPPORT:\n" +
    "For questions, refer to the 'References' sheet for accounting standards guidance.";
  
  ui.alert("Help & Instructions", helpText, ui.ButtonSet.OK);
}