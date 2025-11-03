/**
 * @name far_wp
 * @version 1.0.1
 * @built 2025-11-03T12:27:00.902Z
 * @description Standalone script. Do not edit directly - edit source files in src/ folder.
 * 
 * This file is auto-generated from:
 * - Common utilities (src/common/*.gs)
 * - Workbook-specific code (src/workbooks/far_wp.gs)
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
 * FIXED ASSETS AUDIT WORKPAPER AUTOMATION
 * One-click setup for comprehensive fixed assets audit documentation
 * 
 * INSTRUCTIONS:
 * 1. Open Google Sheets
 * 2. Go to Extensions > Apps Script
 * 3. Delete any existing code and paste this entire script
 * 4. Save the project (name it "Fixed Assets Audit Automation")
 * 5. Refresh your spreadsheet
 * 6. A new menu "Audit Tools" will appear
 * 7. Click "Audit Tools" > "Setup Fixed Assets Workpaper"
 * 8. Wait for completion message
 */

// ============================================================================
// WORKBOOK-SPECIFIC CONFIGURATION
// ============================================================================

// Column mappings for Fixed Assets workbook
const COLS = {
  ROLL_FORWARD: {
    ASSET_CLASS: 1,
    OPENING_GROSS: 2,
    ADDITIONS: 3,
    DISPOSALS: 4,
    TRANSFERS: 5,
    CLOSING_GROSS: 6,
    OPENING_ACCUM_DEP: 7,
    DEPRECIATION: 8,
    DISPOSAL_DEP: 9,
    CLOSING_ACCUM_DEP: 10,
    OPENING_NBV: 11,
    CLOSING_NBV: 12
  }
};

/**
 * Creates custom menu when spreadsheet opens
 */
// onOpen() is handled by common/utilities.gs - auto-detects workbook type
    .addToUi();
}

/**
 * Main function to setup the entire Fixed Assets audit workpaper
 */
function setupFixedAssetsWorkpaper() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Setup Fixed Assets Audit Workpaper',
    'This will create a complete fixed assets audit workpaper with multiple tabs.\n\n' +
    'Any existing sheets may be overwritten. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) {
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Set workbook type for menu detection
  setWorkbookType('FIXED_ASSETS');
  
  // Show progress
  ui.alert('Setting up workpaper...', 'This may take 15-30 seconds. Please wait.', ui.ButtonSet.OK);
  
  // Delete existing sheets except the first one (we'll recreate everything)
  const sheets = ss.getSheets();
  const keepFirstSheet = sheets[0];
  
  try {
    // Create all sheets
    createIndexSheet(ss);
    createSummarySheet(ss);
    createRollForwardSheet(ss);
    createDepreciationScheduleSheet(ss);
    createAdditionsTestingSheet(ss);
    createDisposalsTestingSheet(ss);
    createExistenceTestingSheet(ss);
    createCompletenessTestingSheet(ss);
    createDisclosureSheet(ss);
    createConclusionSheet(ss);
    
    // Delete the original blank sheet if it exists and is named "Sheet1"
    if (keepFirstSheet.getName() === "Sheet1") {
      ss.deleteSheet(keepFirstSheet);
    }
    
    // Set the Index as the active sheet
    ss.setActiveSheet(ss.getSheetByName("FA-Index"));
    
    // Success message
    ui.alert(
      'Success!',
      'Fixed Assets Audit Workpaper has been created successfully!\n\n' +
      'Navigate through the tabs and enter client data as needed.\n' +
      'All formulas and cross-references are set up automatically.',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert('Error', 'An error occurred: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Creates the Index/Table of Contents sheet
 */
function createIndexSheet(ss) {
  let sheet = ss.getSheetByName("FA-Index");
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet("FA-Index", 0);
  sheet.setTabColor('#1f4e78');
  
  // Set column widths
  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 350);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 250);
  
  // Main header
  sheet.getRange("A1:D1").merge().setValue("FIXED ASSETS AUDIT WORKPAPER")
    .setFontSize(FONT_SIZES.title).setFontWeight("bold")
    .setBackground(COLORS.header).setFontColor("#ffffff")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(1, 40);
  
  // Client information section
  const clientInfo = [
    ["Client Name:", "", "Period End:", ""],
    ["Engagement:", "", "Prepared By:", ""],
    ["Date:", "", "Reviewed By:", ""]
  ];
  
  sheet.getRange("A3:D5").setValues(clientInfo)
    .setFontWeight("bold")
    .setBackground(COLORS.preparer);
  
  sheet.getRange("B3:B5").setFontWeight("normal").setBackground("#ffffff");
  sheet.getRange("D3:D5").setFontWeight("normal").setBackground("#ffffff");
  
  // Table of Contents Header
  sheet.getRange("A7:D7").merge().setValue("TABLE OF CONTENTS")
    .setFontSize(FONT_SIZES.header).setFontWeight("bold")
    .setBackground(COLORS.subheader).setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(7, 30);
  
  // Column headers
  const headers = [["Ref", "Workpaper Description", "Preparer", "Reviewer"]];
  sheet.getRange("A8:D8").setValues(headers)
    .setFontWeight("bold")
    .setBackground(COLORS.sectionHeader)
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  
  // Index data with hyperlinks
  const indexData = [
    ["FA-1", "Summary & Conclusion", "", ""],
    ["FA-2", "Fixed Assets Roll Forward", "", ""],
    ["FA-3", "Depreciation Schedule", "", ""],
    ["FA-4", "Additions Testing", "", ""],
    ["FA-5", "Disposals Testing", "", ""],
    ["FA-6", "Existence Testing", "", ""],
    ["FA-7", "Completeness Testing", "", ""],
    ["FA-8", "Presentation & Disclosure", "", ""],
    ["FA-9", "Conclusion & Sign-off", "", ""]
  ];
  
  sheet.getRange(9, 1, indexData.length, 4).setValues(indexData);
  
  // Create hyperlinks to each sheet
  const sheetNames = [
    "FA-1 Summary",
    "FA-2 Roll Forward",
    "FA-3 Depreciation",
    "FA-4 Additions",
    "FA-5 Disposals",
    "FA-6 Existence",
    "FA-7 Completeness",
    "FA-8 Disclosure",
    "FA-9 Conclusion"
  ];
  
  for (let i = 0; i < sheetNames.length; i++) {
    const cell = sheet.getRange(9 + i, 2);
    const formula = `=HYPERLINK("#gid=" & INDIRECT("FA-Index!A1"), "${indexData[i][1]}")`;
    // We'll set actual hyperlinks after sheets are created
  }
  
  // Format data rows
  sheet.getRange(9, 1, indexData.length, 4)
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment("middle");
  
  sheet.getRange(9, 1, indexData.length, 1).setBackground(COLORS.referenceCell).setFontWeight("bold");
  
  // Freeze header rows
  sheet.setFrozenRows(8);
  
  applyStandardBorders(sheet, 9, 1, indexData.length, 4);
}

/**
 * Creates the Summary & Conclusion sheet (FA-1)
 */
function createSummarySheet(ss) {
  let sheet = ss.getSheetByName("FA-1 Summary");
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet("FA-1 Summary");
  sheet.setTabColor('#4472c4');
  
  // Set column widths
  sheet.setColumnWidths(1, 1, 100);
  sheet.setColumnWidths(2, 1, 300);
  sheet.setColumnWidths(3, 1, 150);
  sheet.setColumnWidths(4, 1, 150);
  sheet.setColumnWidths(5, 1, 150);
  
  // Header
  createWorkpaperHeader(sheet, "FA-1", "FIXED ASSETS - SUMMARY & CONCLUSION");
  
  // Summary of Balances
  let currentRow = 5;
  sheet.getRange(currentRow, 1, 1, 5).merge()
    .setValue("SUMMARY OF FIXED ASSETS BALANCES")
    .setFontWeight("bold")
    .setBackground(COLORS.sectionHeader)
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  
  currentRow++;
  const summaryHeaders = [["Description", "Beginning Balance", "Additions", "Disposals", "Ending Balance"]];
  sheet.getRange(currentRow, 1, 1, 5).setValues(summaryHeaders)
    .setFontWeight("bold")
    .setBackground(COLORS.subheader)
    .setFontColor("#ffffff");
  
  currentRow++;
  // Link to Roll Forward sheet
  sheet.getRange(currentRow, 1).setValue("Gross Fixed Assets");
  sheet.getRange(currentRow, 2).setFormula("='FA-2 Roll Forward'!E7");
  sheet.getRange(currentRow, 3).setFormula("='FA-2 Roll Forward'!F7");
  sheet.getRange(currentRow, 4).setFormula("='FA-2 Roll Forward'!G7");
  sheet.getRange(currentRow, 5).setFormula("='FA-2 Roll Forward'!H7");
  
  currentRow++;
  sheet.getRange(currentRow, 1).setValue("Accumulated Depreciation");
  sheet.getRange(currentRow, 2).setFormula("='FA-3 Depreciation'!E15");
  sheet.getRange(currentRow, 3).setFormula("='FA-3 Depreciation'!F15");
  sheet.getRange(currentRow, 4).setFormula("='FA-3 Depreciation'!G15");
  sheet.getRange(currentRow, 5).setFormula("='FA-3 Depreciation'!H15");
  
  currentRow++;
  sheet.getRange(currentRow, 1).setValue("Net Fixed Assets")
    .setFontWeight("bold");
  sheet.getRange(currentRow, 2).setFormula("=B7-B8")
    .setFontWeight("bold");
  sheet.getRange(currentRow, 3).setFormula("=C7-C8")
    .setFontWeight("bold");
  sheet.getRange(currentRow, 4).setFormula("=D7-D8")
    .setFontWeight("bold");
  sheet.getRange(currentRow, 5).setFormula("=E7-E8")
    .setFontWeight("bold");
  
  // Format numbers
  sheet.getRange(7, 2, 3, 4).setNumberFormat("#,##0.00");
  sheet.getRange(9, 2, 1, 4).setBackground(COLORS.totalRow);
  
  // Audit Procedures Summary
  currentRow = 12;
  sheet.getRange(currentRow, 1, 1, 5).merge()
    .setValue("AUDIT PROCEDURES PERFORMED")
    .setFontWeight("bold")
    .setBackground(COLORS.sectionHeader)
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  
  currentRow++;
  const procedures = [
    ["1", "Obtained and reviewed fixed asset roll forward", "FA-2", ""],
    ["2", "Tested additions for proper authorization and capitalization", "FA-4", ""],
    ["3", "Tested disposals and verified removal from records", "FA-5", ""],
    ["4", "Performed physical verification of selected assets", "FA-6", ""],
    ["5", "Tested completeness of fixed asset recording", "FA-7", ""],
    ["6", "Recalculated depreciation expense", "FA-3", ""],
    ["7", "Reviewed financial statement presentation and disclosure", "FA-8", ""]
  ];
  
  sheet.getRange(currentRow, 1, 1, 4).setValues([["#", "Procedure", "Ref", "Conclusion"]])
    .setFontWeight("bold")
    .setBackground(COLORS.subheader)
    .setFontColor("#ffffff");
  
  currentRow++;
  sheet.getRange(currentRow, 1, procedures.length, 4).setValues(procedures);
  
  // Conclusion section
  currentRow = currentRow + procedures.length + 2;
  sheet.getRange(currentRow, 1, 1, 5).merge()
    .setValue("AUDIT CONCLUSION")
    .setFontWeight("bold")
    .setBackground(COLORS.sectionHeader)
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  
  currentRow++;
  sheet.getRange(currentRow, 1, 1, 5).merge()
    .setValue("Based on the audit procedures performed, we conclude that:")
    .setWrap(true);
  
  currentRow++;
  sheet.getRange(currentRow, 1, 4, 5).merge()
    .setValue("[Enter conclusion here - e.g., 'Fixed assets are fairly stated in all material respects...']")
    .setWrap(true)
    .setVerticalAlignment("top")
    .setBackground("#ffffcc");
  sheet.setRowHeights(currentRow, 4, 25);
  
  // Sign-off
  currentRow += 5;
  createSignOffSection(sheet, currentRow);
  
  applyStandardBorders(sheet, 6, 1, 20, 5);
}

/**
 * Creates the Roll Forward sheet (FA-2)
 */
function createRollForwardSheet(ss) {
  let sheet = ss.getSheetByName("FA-2 Roll Forward");
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet("FA-2 Roll Forward");
  sheet.setTabColor('#5b9bd5');
  
  // Set column widths
  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 120);
  sheet.setColumnWidth(8, 120);
  
  createWorkpaperHeader(sheet, "FA-2", "FIXED ASSETS ROLL FORWARD");
  
  // Column headers
  const headers = [
    ["Ref", "Asset Category", "Useful Life", "Description", "Beginning Balance", "Additions", "Disposals", "Ending Balance"]
  ];
  
  sheet.getRange(5, 1, 1, 8).setValues(headers)
    .setFontWeight("bold")
    .setBackground(COLORS.subheader)
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center")
    .setWrap(true);
  sheet.setRowHeight(5, 40);
  
  // Asset categories
  const categories = [
    ["", "Land", "N/A", "Land - not depreciated"],
    ["", "Buildings", "39 years", "Office buildings and improvements"],
    ["", "Machinery & Equipment", "5-10 years", "Manufacturing equipment"],
    ["", "Furniture & Fixtures", "7 years", "Office furniture and fixtures"],
    ["", "Vehicles", "5 years", "Company vehicles"],
    ["", "Computer Equipment", "3-5 years", "Computers, servers, IT equipment"],
    ["", "Leasehold Improvements", "Lease term", "Improvements to leased property"]
  ];
  
  sheet.getRange(6, 1, categories.length, 4).setValues(categories);
  
  // Add formulas for totals
  for (let i = 0; i < categories.length; i++) {
    const row = 6 + i;
    // Beginning Balance (enter data here)
    // Additions and Disposals will be entered
    // Ending Balance formula
    sheet.getRange(row, 8).setFormula(`=E${row}+F${row}-G${row}`);
  }
  
  // Total row
  const totalRow = 6 + categories.length;
  sheet.getRange(totalRow, 2).setValue("TOTAL GROSS FIXED ASSETS")
    .setFontWeight("bold");
  
  sheet.getRange(totalRow, 5).setFormula(`=SUM(E6:E${totalRow-1})`)
    .setFontWeight("bold");
  sheet.getRange(totalRow, 6).setFormula(`=SUM(F6:F${totalRow-1})`)
    .setFontWeight("bold");
  sheet.getRange(totalRow, 7).setFormula(`=SUM(G6:G${totalRow-1})`)
    .setFontWeight("bold");
  sheet.getRange(totalRow, 8).setFormula(`=SUM(H6:H${totalRow-1})`)
    .setFontWeight("bold");
  
  sheet.getRange(totalRow, 1, 1, 8)
    .setBackground(COLORS.totalRow);
  
  // Format numbers
  sheet.getRange(6, 5, categories.length + 1, 4).setNumberFormat("#,##0.00");
  
  // Cross-reference column
  sheet.getRange(6, 1, categories.length, 1).setBackground(COLORS.referenceCell);
  
  // Reconciliation section
  let reconRow = totalRow + 3;
  sheet.getRange(reconRow, 1, 1, 8).merge()
    .setValue("RECONCILIATION TO GENERAL LEDGER")
    .setFontWeight("bold")
    .setBackground(COLORS.sectionHeader)
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  
  reconRow++;
  sheet.getRange(reconRow, 2).setValue("Per General Ledger:");
  sheet.getRange(reconRow, 5).setValue(0).setNumberFormat("#,##0.00");
  sheet.getRange(reconRow, 1).setValue("TB").setBackground(COLORS.referenceCell);
  
  reconRow++;
  sheet.getRange(reconRow, 2).setValue("Per Audit (above):");
  sheet.getRange(reconRow, 5).setFormula(`=H${totalRow}`).setNumberFormat("#,##0.00");
  
  reconRow++;
  sheet.getRange(reconRow, 2).setValue("Difference:")
    .setFontWeight("bold");
  sheet.getRange(reconRow, 5).setFormula(`=E${reconRow-2}-E${reconRow-1}`)
    .setFontWeight("bold")
    .setNumberFormat("#,##0.00")
    .setBackground(COLORS.totalRow);
  
  // Notes section
  reconRow += 2;
  sheet.getRange(reconRow, 1, 1, 8).merge()
    .setValue("NOTES & EXPLANATIONS")
    .setFontWeight("bold")
    .setBackground(COLORS.sectionHeader)
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  
  reconRow++;
  sheet.getRange(reconRow, 1, 3, 8).merge()
    .setValue("[Document any significant additions, disposals, or unusual transactions]")
    .setWrap(true)
    .setVerticalAlignment("top")
    .setBackground("#ffffcc");
  
  createSignOffSection(sheet, reconRow + 4);

  applyStandardBorders(sheet, 5, 1, totalRow - 4, 8);
  sheet.setFrozenRows(5);
}

/**
 * Creates the Depreciation Schedule sheet (FA-3)
 */
function createDepreciationScheduleSheet(ss) {
  let sheet = ss.getSheetByName("FA-3 Depreciation");
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet("FA-3 Depreciation");
  sheet.setTabColor('#70ad47');
  
  // Set column widths
  sheet.setColumnWidths(1, 8, 120);
  
  createWorkpaperHeader(sheet, "FA-3", "ACCUMULATED DEPRECIATION & DEPRECIATION EXPENSE");
  
  // Column headers
  const headers = [
    ["Asset Category", "Method", "Beginning Balance", "Current Year Expense", "Disposals", "Ending Balance", "Recalc", "Variance"]
  ];
  
  sheet.getRange(5, 1, 1, 8).setValues(headers)
    .setFontWeight("bold")
    .setBackground(COLORS.subheader)
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center")
    .setWrap(true);
  sheet.setRowHeight(5, 40);
  
  // Asset categories (matching roll forward)
  const categories = [
    ["Land", "N/A"],
    ["Buildings", "Straight-Line"],
    ["Machinery & Equipment", "Straight-Line"],
    ["Furniture & Fixtures", "Straight-Line"],
    ["Vehicles", "Straight-Line"],
    ["Computer Equipment", "Straight-Line"],
    ["Leasehold Improvements", "Straight-Line"]
  ];
  
  sheet.getRange(6, 1, categories.length, 2).setValues(categories);
  
  // Add formulas
  for (let i = 0; i < categories.length; i++) {
    const row = 6 + i;
    // Ending Balance formula
    sheet.getRange(row, 6).setFormula(`=C${row}+D${row}-E${row}`);
    // Variance formula (Ending - Recalc)
    sheet.getRange(row, 8).setFormula(`=F${row}-G${row}`);
  }
  
  // Depreciation calculation section
  let calcRow = 6 + categories.length + 2;
  sheet.getRange(calcRow, 1, 1, 8).merge()
    .setValue("DEPRECIATION EXPENSE RECALCULATION")
    .setFontWeight("bold")
    .setBackground(COLORS.sectionHeader)
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  
  calcRow++;
  const calcHeaders = [["Asset Category", "Gross Assets", "Useful Life", "Method", "Calculated Expense", "Per Client", "Variance", "Notes"]];
  sheet.getRange(calcRow, 1, 1, 8).setValues(calcHeaders)
    .setFontWeight("bold")
    .setBackground(COLORS.subheader)
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  
  // Total row for accumulated depreciation
  const totalRow = 6 + categories.length - 1 + 1;
  sheet.getRange(totalRow, 1).setValue("TOTAL ACCUMULATED DEPRECIATION")
    .setFontWeight("bold");
  
  sheet.getRange(totalRow, 3).setFormula(`=SUM(C6:C${totalRow-1})`)
    .setFontWeight("bold");
  sheet.getRange(totalRow, 4).setFormula(`=SUM(D6:D${totalRow-1})`)
    .setFontWeight("bold");
  sheet.getRange(totalRow, 5).setFormula(`=SUM(E6:E${totalRow-1})`)
    .setFontWeight("bold");
  sheet.getRange(totalRow, 6).setFormula(`=SUM(F6:F${totalRow-1})`)
    .setFontWeight("bold");
  sheet.getRange(totalRow, 7).setFormula(`=SUM(G6:G${totalRow-1})`)
    .setFontWeight("bold");
  sheet.getRange(totalRow, 8).setFormula(`=SUM(H6:H${totalRow-1})`)
    .setFontWeight("bold");
  
  sheet.getRange(totalRow, 1, 1, 8).setBackground(COLORS.totalRow);
  
  // Format numbers
  sheet.getRange(6, 3, categories.length + 1, 6).setNumberFormat("#,##0.00");
  
  createSignOffSection(sheet, calcRow + 10);
  
  applyStandardBorders(sheet, 5, 1, totalRow - 4, 8);
  sheet.setFrozenRows(5);
}

/**
 * Creates the Additions Testing sheet (FA-4)
 */
function createAdditionsTestingSheet(ss) {
  let sheet = ss.getSheetByName("FA-4 Additions");
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet("FA-4 Additions");
  sheet.setTabColor('#ffc000');
  
  sheet.setColumnWidths(1, 10, 110);
  
  createWorkpaperHeader(sheet, "FA-4", "ADDITIONS TESTING");
  
  // Testing objective
  sheet.getRange(5, 1, 1, 10).merge()
    .setValue("OBJECTIVE: Test additions to verify proper authorization, occurrence, and capitalization")
    .setFontWeight("bold")
    .setBackground(COLORS.sectionHeader)
    .setFontColor("#ffffff")
    .setWrap(true);
  
  // Sample selection
  sheet.getRange(7, 1, 1, 3).merge()
    .setValue("SAMPLE SELECTION")
    .setFontWeight("bold")
    .setBackground(COLORS.subheader)
    .setFontColor("#ffffff");
  
  sheet.getRange(8, 1).setValue("Total Additions:");
  sheet.getRange(8, 2).setFormula("='FA-2 Roll Forward'!F13")
    .setNumberFormat("#,##0.00");
  
  sheet.getRange(9, 1).setValue("Sample Size:");
  sheet.getRange(9, 2).setValue(25);
  
  sheet.getRange(10, 1).setValue("Sample Coverage:");
  sheet.getRange(10, 2).setFormula("=SUM(G13:G37)/B8")
    .setNumberFormat("0.0%");
  
  // Testing columns
  const testHeaders = [
    ["Date", "Description", "Category", "Vendor", "Invoice #", "Amount", "Authorization", "Capitalization", "Classification", "Conclusion"]
  ];
  
  sheet.getRange(12, 1, 1, 10).setValues(testHeaders)
    .setFontWeight("bold")
    .setBackground(COLORS.subheader)
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center")
    .setWrap(true);
  sheet.setRowHeight(12, 40);
  
  // Sample rows (25 rows for testing)
  const sampleCount = 25;
  for (let i = 0; i < sampleCount; i++) {
    const row = 13 + i;
    // Add data validation for certain columns
    sheet.getRange(row, 7).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['✓', '✗', 'N/A'], true)
        .build()
    );
    sheet.getRange(row, 8).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['✓', '✗', 'N/A'], true)
        .build()
    );
    sheet.getRange(row, 9).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['✓', '✗', 'N/A'], true)
        .build()
    );
    sheet.getRange(row, 10).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['Pass', 'Fail', 'Note'], true)
        .build()
    );
  }
  
  // Format amount column
  sheet.getRange(13, 6, sampleCount, 1).setNumberFormat("#,##0.00");
  
  // Total row
  const totalRow = 13 + sampleCount;
  sheet.getRange(totalRow, 1, 1, 5).merge()
    .setValue("TOTAL TESTED")
    .setFontWeight("bold");
  sheet.getRange(totalRow, 6).setFormula(`=SUM(F13:F${totalRow-1})`)
    .setFontWeight("bold")
    .setNumberFormat("#,##0.00");
  sheet.getRange(totalRow, 1, 1, 10).setBackground(COLORS.totalRow);
  
  // Exceptions section
  let exceptRow = totalRow + 2;
  sheet.getRange(exceptRow, 1, 1, 10).merge()
    .setValue("EXCEPTIONS & NOTES")
    .setFontWeight("bold")
    .setBackground(COLORS.sectionHeader)
    .setFontColor("#ffffff");
  
  exceptRow++;
  sheet.getRange(exceptRow, 1, 3, 10).merge()
    .setValue("[Document any exceptions, unusual items, or additional notes]")
    .setWrap(true)
    .setVerticalAlignment("top")
    .setBackground("#ffffcc");
  
  createSignOffSection(sheet, exceptRow + 4);

  applyStandardBorders(sheet, 12, 1, sampleCount + 1, 10);
  sheet.setFrozenRows(12);
}

/**
 * Creates the Disposals Testing sheet (FA-5)
 */
function createDisposalsTestingSheet(ss) {
  let sheet = ss.getSheetByName("FA-5 Disposals");
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet("FA-5 Disposals");
  sheet.setTabColor('#f4b084');
  
  sheet.setColumnWidths(1, 10, 110);
  
  createWorkpaperHeader(sheet, "FA-5", "DISPOSALS & RETIREMENTS TESTING");
  
  // Testing objective
  sheet.getRange(5, 1, 1, 10).merge()
    .setValue("OBJECTIVE: Test disposals to verify proper authorization, removal from records, and gain/loss calculation")
    .setFontWeight("bold")
    .setBackground(COLORS.sectionHeader)
    .setFontColor("#ffffff")
    .setWrap(true);
  
  // Sample selection
  sheet.getRange(7, 1, 1, 3).merge()
    .setValue("SAMPLE SELECTION")
    .setFontWeight("bold")
    .setBackground(COLORS.subheader)
    .setFontColor("#ffffff");
  
  sheet.getRange(8, 1).setValue("Total Disposals:");
  sheet.getRange(8, 2).setFormula("='FA-2 Roll Forward'!G13")
    .setNumberFormat("#,##0.00");
  
  sheet.getRange(9, 1).setValue("Sample Size:");
  sheet.getRange(9, 2).setValue(15);
  
  // Testing columns
  const testHeaders = [
    ["Date", "Asset Description", "Category", "Original Cost", "Accum. Depr.", "Net Book Value", "Proceeds", "Gain/(Loss)", "Authorization", "Conclusion"]
  ];
  
  sheet.getRange(12, 1, 1, 10).setValues(testHeaders)
    .setFontWeight("bold")
    .setBackground(COLORS.subheader)
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center")
    .setWrap(true);
  sheet.setRowHeight(12, 40);
  
  // Sample rows
  const sampleCount = 15;
  for (let i = 0; i < sampleCount; i++) {
    const row = 13 + i;
    // Net Book Value formula
    sheet.getRange(row, 6).setFormula(`=D${row}-E${row}`);
    // Gain/Loss formula
    sheet.getRange(row, 8).setFormula(`=G${row}-F${row}`);
    
    // Data validation
    sheet.getRange(row, 9).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['✓', '✗', 'N/A'], true)
        .build()
    );
    sheet.getRange(row, 10).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['Pass', 'Fail', 'Note'], true)
        .build()
    );
  }
  
  // Format numbers
  sheet.getRange(13, 4, sampleCount, 5).setNumberFormat("#,##0.00");
  
  // Conditional formatting for gain/loss
  const gainLossRange = sheet.getRange(13, 8, sampleCount, 1);
  const rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground("#c6efce")
    .setRanges([gainLossRange])
    .build();
  const rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground("#ffc7ce")
    .setRanges([gainLossRange])
    .build();
  sheet.setConditionalFormatRules([rule1, rule2]);
  
  // Total row
  const totalRow = 13 + sampleCount;
  sheet.getRange(totalRow, 1, 1, 3).merge()
    .setValue("TOTAL TESTED")
    .setFontWeight("bold");
  sheet.getRange(totalRow, 4).setFormula(`=SUM(D13:D${totalRow-1})`)
    .setFontWeight("bold")
    .setNumberFormat("#,##0.00");
  sheet.getRange(totalRow, 5).setFormula(`=SUM(E13:E${totalRow-1})`)
    .setFontWeight("bold")
    .setNumberFormat("#,##0.00");
  sheet.getRange(totalRow, 6).setFormula(`=SUM(F13:F${totalRow-1})`)
    .setFontWeight("bold")
    .setNumberFormat("#,##0.00");
  sheet.getRange(totalRow, 7).setFormula(`=SUM(G13:G${totalRow-1})`)
    .setFontWeight("bold")
    .setNumberFormat("#,##0.00");
  sheet.getRange(totalRow, 8).setFormula(`=SUM(H13:H${totalRow-1})`)
    .setFontWeight("bold")
    .setNumberFormat("#,##0.00");
  sheet.getRange(totalRow, 1, 1, 10).setBackground(COLORS.totalRow);
  
  createSignOffSection(sheet, totalRow + 5);

  applyStandardBorders(sheet, 12, 1, sampleCount + 1, 10);
  sheet.setFrozenRows(12);
}

/**
 * Creates the Existence Testing sheet (FA-6)
 */
function createExistenceTestingSheet(ss) {
  let sheet = ss.getSheetByName("FA-6 Existence");
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet("FA-6 Existence");
  sheet.setTabColor('#9dc3e6');
  
  sheet.setColumnWidths(1, 9, 120);
  
  createWorkpaperHeader(sheet, "FA-6", "PHYSICAL EXISTENCE VERIFICATION");
  
  // Testing objective
  sheet.getRange(5, 1, 1, 9).merge()
    .setValue("OBJECTIVE: Verify physical existence of selected fixed assets")
    .setFontWeight("bold")
    .setBackground(COLORS.sectionHeader)
    .setFontColor("#ffffff");
  
  // Sample selection
  sheet.getRange(7, 1, 1, 3).merge()
    .setValue("SAMPLE SELECTION")
    .setFontWeight("bold")
    .setBackground(COLORS.subheader)
    .setFontColor("#ffffff");
  
  sheet.getRange(8, 1).setValue("Total Fixed Assets:");
  sheet.getRange(8, 2).setFormula("='FA-2 Roll Forward'!H13")
    .setNumberFormat("#,##0.00");
  
  sheet.getRange(9, 1).setValue("Items Selected:");
  sheet.getRange(9, 2).setValue(30);
  
  // Testing columns
  const testHeaders = [
    ["Asset ID", "Description", "Category", "Location", "Book Value", "Observed?", "Condition", "Tag #", "Notes"]
  ];
  
  sheet.getRange(12, 1, 1, 9).setValues(testHeaders)
    .setFontWeight("bold")
    .setBackground(COLORS.subheader)
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center")
    .setWrap(true);
  sheet.setRowHeight(12, 40);
  
  // Sample rows
  const sampleCount = 30;
  for (let i = 0; i < sampleCount; i++) {
    const row = 13 + i;
    
    // Data validation
    sheet.getRange(row, 6).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['✓ Yes', '✗ No', 'Unable to locate'], true)
        .build()
    );
    
    sheet.getRange(row, 7).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['Good', 'Fair', 'Poor', 'N/A'], true)
        .build()
    );
  }
  
  // Format numbers
  sheet.getRange(13, 5, sampleCount, 1).setNumberFormat("#,##0.00");
  
  // Summary section
  const summaryRow = 13 + sampleCount + 2;
  sheet.getRange(summaryRow, 1, 1, 9).merge()
    .setValue("VERIFICATION SUMMARY")
    .setFontWeight("bold")
    .setBackground(COLORS.sectionHeader)
    .setFontColor("#ffffff");
  
  sheet.getRange(summaryRow + 1, 1).setValue("Assets Physically Verified:");
  sheet.getRange(summaryRow + 1, 2).setFormula(`=COUNTIF(F13:F${13+sampleCount-1},"✓ Yes")`);
  
  sheet.getRange(summaryRow + 2, 1).setValue("Assets Not Located:");
  sheet.getRange(summaryRow + 2, 2).setFormula(`=COUNTIF(F13:F${13+sampleCount-1},"Unable to locate")`);
  
  sheet.getRange(summaryRow + 3, 1).setValue("Verification Rate:");
  sheet.getRange(summaryRow + 3, 2).setFormula(`=B${summaryRow+1}/${sampleCount}`)
    .setNumberFormat("0.0%");
  
  createSignOffSection(sheet, summaryRow + 6);

  applyStandardBorders(sheet, 12, 1, sampleCount + 1, 9);
  sheet.setFrozenRows(12);
}

/**
 * Creates the Completeness Testing sheet (FA-7)
 */
function createCompletenessTestingSheet(ss) {
  let sheet = ss.getSheetByName("FA-7 Completeness");
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet("FA-7 Completeness");
  sheet.setTabColor('#a9d08e');
  
  sheet.setColumnWidths(1, 8, 130);
  
  createWorkpaperHeader(sheet, "FA-7", "COMPLETENESS TESTING");
  
  // Testing objective
  sheet.getRange(5, 1, 1, 8).merge()
    .setValue("OBJECTIVE: Test that all qualifying expenditures have been properly capitalized")
    .setFontWeight("bold")
    .setBackground(COLORS.sectionHeader)
    .setFontColor("#ffffff");
  
  // Procedure 1: Repair & Maintenance Expense Review
  sheet.getRange(7, 1, 1, 8).merge()
    .setValue("PROCEDURE 1: REVIEW REPAIR & MAINTENANCE EXPENSES")
    .setFontWeight("bold")
    .setBackground(COLORS.subheader)
    .setFontColor("#ffffff");
  
  const repairHeaders = [
    ["Date", "Vendor", "Description", "Amount", "Nature", "Capitalize?", "Adjustment", "Notes"]
  ];
  
  sheet.getRange(8, 1, 1, 8).setValues(repairHeaders)
    .setFontWeight("bold")
    .setBackground(COLORS.subheader)
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  
  const repairRows = 15;
  for (let i = 0; i < repairRows; i++) {
    const row = 9 + i;
    sheet.getRange(row, 5).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['Repair', 'Improvement', 'Betterment', 'Other'], true)
        .build()
    );
    sheet.getRange(row, 6).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['Yes', 'No'], true)
        .build()
    );
  }
  
  sheet.getRange(9, 4, repairRows, 1).setNumberFormat("#,##0.00");
  sheet.getRange(9, 7, repairRows, 1).setNumberFormat("#,##0.00");
  
  // Procedure 2: Construction in Progress
  let cipRow = 9 + repairRows + 2;
  sheet.getRange(cipRow, 1, 1, 8).merge()
    .setValue("PROCEDURE 2: CONSTRUCTION IN PROGRESS REVIEW")
    .setFontWeight("bold")
    .setBackground(COLORS.subheader)
    .setFontColor("#ffffff");
  
  cipRow++;
  const cipHeaders = [["Project", "Start Date", "Status", "Costs to Date", "Ready for Use?", "Transfer to FA?", "Notes", ""]];
  sheet.getRange(cipRow, 1, 1, 8).setValues(cipHeaders)
    .setFontWeight("bold")
    .setBackground(COLORS.subheader)
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  
  const cipRows = 10;
  for (let i = 0; i < cipRows; i++) {
    const row = cipRow + 1 + i;
    sheet.getRange(row, 3).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['In Progress', 'Complete', 'On Hold'], true)
        .build()
    );
    sheet.getRange(row, 5).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['Yes', 'No', 'Partial'], true)
        .build()
    );
    sheet.getRange(row, 6).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['Yes', 'No', 'N/A'], true)
        .build()
    );
  }
  
  sheet.getRange(cipRow + 1, 4, cipRows, 1).setNumberFormat("#,##0.00");
  
  createSignOffSection(sheet, cipRow + cipRows + 4);
  
  applyStandardBorders(sheet, 8, 1, repairRows + 1, 8);
  applyStandardBorders(sheet, cipRow, 1, cipRows + 1, 8);
  sheet.setFrozenRows(8);
}

/**
 * Creates the Disclosure sheet (FA-8)
 */
function createDisclosureSheet(ss) {
  let sheet = ss.getSheetByName("FA-8 Disclosure");
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet("FA-8 Disclosure");
  sheet.setTabColor('#c5e0b4');
  
  sheet.setColumnWidths(1, 5, 200);
  
  createWorkpaperHeader(sheet, "FA-8", "PRESENTATION & DISCLOSURE CHECKLIST");
  
  // Checklist section
  sheet.getRange(5, 1, 1, 5).merge()
    .setValue("DISCLOSURE REQUIREMENTS CHECKLIST")
    .setFontWeight("bold")
    .setBackground(COLORS.sectionHeader)
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  
  const checklistHeaders = [["Requirement", "Yes/No/N/A", "Reference", "Notes", ""]];
  sheet.getRange(6, 1, 1, 5).setValues(checklistHeaders)
    .setFontWeight("bold")
    .setBackground(COLORS.subheader)
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  
  const disclosures = [
    ["Balance of major classes of depreciable assets"],
    ["Accumulated depreciation by major class or in total"],
    ["General description of depreciation methods"],
    ["Depreciation expense for the period"],
    ["Significant additions and disposals"],
    ["Carrying amount of assets not in service"],
    ["Capitalization policy disclosure"],
    ["Useful lives by asset class"],
    ["Impairment losses (if any)"],
    ["Assets pledged as collateral"],
    ["Construction commitments"],
    ["Leased assets (if applicable)"],
    ["Fair value disclosures (if required)"],
    ["Related party transactions"]
  ];
  
  sheet.getRange(7, 1, disclosures.length, 1).setValues(disclosures);
  
  for (let i = 0; i < disclosures.length; i++) {
    const row = 7 + i;
    sheet.getRange(row, 2).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['Yes', 'No', 'N/A'], true)
        .build()
    );
  }
  
  // Financial Statement Presentation
  let fsRow = 7 + disclosures.length + 2;
  sheet.getRange(fsRow, 1, 1, 5).merge()
    .setValue("FINANCIAL STATEMENT PRESENTATION")
    .setFontWeight("bold")
    .setBackground(COLORS.sectionHeader)
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  
  fsRow++;
  const fsData = [
    ["Fixed Assets (Gross)", "", "=", ""],
    ["Less: Accumulated Depreciation", "", "=", ""],
    ["Fixed Assets (Net)", "", "=", ""],
    ["", "", "", ""],
    ["Depreciation Expense (P&L)", "", "=", ""]
  ];
  
  sheet.getRange(fsRow, 1, fsData.length, 4).setValues(fsData);
  sheet.getRange(fsRow, 1, fsData.length, 1).setFontWeight("bold");
  
  // Link to other workpapers
  sheet.getRange(fsRow, 2).setFormula("='FA-2 Roll Forward'!H13");
  sheet.getRange(fsRow + 1, 2).setFormula("='FA-3 Depreciation'!F15");
  sheet.getRange(fsRow + 2, 2).setFormula("=B" + fsRow + "-B" + (fsRow + 1));
  sheet.getRange(fsRow + 4, 2).setFormula("='FA-3 Depreciation'!D15");
  
  sheet.getRange(fsRow, 2, fsData.length, 1).setNumberFormat("#,##0.00");
  sheet.getRange(fsRow, 4, fsData.length, 1).setNumberFormat("#,##0.00");
  
  createSignOffSection(sheet, fsRow + 8);
  
  applyStandardBorders(sheet, 6, 1, disclosures.length + 1, 5);
  sheet.setFrozenRows(6);
}

/**
 * Creates the Conclusion sheet (FA-9)
 */
function createConclusionSheet(ss) {
  let sheet = ss.getSheetByName("FA-9 Conclusion");
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet("FA-9 Conclusion");
  sheet.setTabColor('#70ad47');
  
  sheet.setColumnWidths(1, 6, 150);
  
  createWorkpaperHeader(sheet, "FA-9", "AUDIT CONCLUSION & SIGN-OFF");
  
  // Summary of procedures
  sheet.getRange(5, 1, 1, 6).merge()
    .setValue("SUMMARY OF AUDIT PROCEDURES")
    .setFontWeight("bold")
    .setBackground(COLORS.sectionHeader)
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  
  const procedures = [
    ["Obtained and agreed fixed asset roll forward to general ledger", "FA-2", "✓"],
    ["Tested additions for proper authorization and capitalization", "FA-4", ""],
    ["Tested disposals and verified gain/loss calculations", "FA-5", ""],
    ["Performed physical verification of selected assets", "FA-6", ""],
    ["Tested completeness of fixed asset recording", "FA-7", ""],
    ["Recalculated depreciation expense", "FA-3", ""],
    ["Reviewed financial statement presentation and disclosure", "FA-8", ""]
  ];
  
  const procHeaders = [["Procedure", "Reference", "Complete"]];
  sheet.getRange(6, 1, 1, 3).setValues(procHeaders)
    .setFontWeight("bold")
    .setBackground(COLORS.subheader)
    .setFontColor("#ffffff");
  
  sheet.getRange(7, 1, procedures.length, 3).setValues(procedures);
  
  // Exceptions and findings
  let exceptRow = 7 + procedures.length + 2;
  sheet.getRange(exceptRow, 1, 1, 6).merge()
    .setValue("EXCEPTIONS & FINDINGS")
    .setFontWeight("bold")
    .setBackground(COLORS.sectionHeader)
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  
  exceptRow++;
  sheet.getRange(exceptRow, 1, 4, 6).merge()
    .setValue("[Document any exceptions, findings, or matters requiring attention]")
    .setWrap(true)
    .setVerticalAlignment("top")
    .setBackground("#ffffcc");
  sheet.setRowHeights(exceptRow, 4, 25);
  
  // Final conclusion
  exceptRow += 5;
  sheet.getRange(exceptRow, 1, 1, 6).merge()
    .setValue("FINAL AUDIT CONCLUSION")
    .setFontWeight("bold")
    .setBackground(COLORS.sectionHeader)
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  
  exceptRow++;
  sheet.getRange(exceptRow, 1, 5, 6).merge()
    .setValue(
      "Based on the audit procedures performed and documented in this workpaper, " +
      "we conclude that fixed assets are fairly stated in all material respects as of [date]. " +
      "The balance agrees with the general ledger, and no material exceptions were noted.\n\n" +
      "[Modify conclusion as appropriate based on findings]"
    )
    .setWrap(true)
    .setVerticalAlignment("top")
    .setBackground("#d9ead3");
  sheet.setRowHeights(exceptRow, 5, 25);
  
  // Final sign-off
  exceptRow += 6;
  createSignOffSection(sheet, exceptRow);
  
  applyStandardBorders(sheet, 6, 1, procedures.length + 1, 3);
}

/**
 * Helper function to create standard workpaper header
 */
function createWorkpaperHeader(sheet, reference, title) {
  // Title row
  sheet.getRange("A1:H1").merge()
    .setValue(title)
    .setFontSize(FONT_SIZES.title)
    .setFontWeight("bold")
    .setBackground(COLORS.header)
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 35);
  
  // Reference and metadata
  sheet.getRange("A2").setValue("Reference:");
  sheet.getRange("B2").setValue(reference)
    .setFontWeight("bold")
    .setBackground(COLORS.referenceCell);
  
  sheet.getRange("D2").setValue("Prepared By:");
  sheet.getRange("E2").setBackground("#ffffff");
  
  sheet.getRange("F2").setValue("Date:");
  sheet.getRange("G2").setBackground("#ffffff");
  
  sheet.getRange("A3").setValue("Client:");
  sheet.getRange("B3:C3").merge().setBackground("#ffffff");
  
  sheet.getRange("D3").setValue("Reviewed By:");
  sheet.getRange("E3").setBackground("#ffffff");
  
  sheet.getRange("F3").setValue("Date:");
  sheet.getRange("G3").setBackground("#ffffff");
  
  sheet.getRange("A2:A3").setFontWeight("bold");
  sheet.getRange("D2:F3").setFontWeight("bold");
}

/**
 * Helper function to create sign-off section
 */
function createSignOffSection(sheet, startRow) {
  sheet.getRange(startRow, 1, 1, 8).merge()
    .setValue("PREPARER & REVIEWER SIGN-OFF")
    .setFontWeight("bold")
    .setBackground(COLORS.sectionHeader)
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  
  startRow++;
  
  sheet.getRange(startRow, 1).setValue("Prepared By:");
  sheet.getRange(startRow, 2).setBackground(COLORS.preparer);
  sheet.getRange(startRow, 3).setValue("Date:");
  sheet.getRange(startRow, 4).setBackground(COLORS.preparer);
  
  startRow++;
  sheet.getRange(startRow, 1).setValue("Reviewed By:");
  sheet.getRange(startRow, 2).setBackground(COLORS.reviewer);
  sheet.getRange(startRow, 3).setValue("Date:");
  sheet.getRange(startRow, 4).setBackground(COLORS.reviewer);
  
  sheet.getRange(startRow - 1, 1, 2, 1).setFontWeight("bold");
  sheet.getRange(startRow - 1, 3, 2, 1).setFontWeight("bold");
}

/**
 * Helper function to apply standard borders
 */
function applyStandardBorders(sheet, startRow, startCol, numRows, numCols) {
  const range = sheet.getRange(startRow, startCol, numRows, numCols);
  range.setBorder(
    true, true, true, true, true, true,
    "#000000",
    SpreadsheetApp.BorderStyle.SOLID
  );
}

/**
 * Adds sample data to the workpaper for demonstration
 */
function addSampleData() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Add Sample Data',
    'This will populate the workpaper with sample data for demonstration. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Add sample data to Roll Forward
  const rollForward = ss.getSheetByName("FA-2 Roll Forward");
  if (rollForward) {
    const sampleData = [
      [1250000, 0, 0],
      [3500000, 250000, 0],
      [2800000, 450000, 125000],
      [450000, 75000, 25000],
      [325000, 45000, 30000],
      [680000, 125000, 85000],
      [420000, 95000, 15000]
    ];
    rollForward.getRange(6, 5, 7, 3).setValues(sampleData);
  }
  
  // Add sample data to Depreciation
  const depreciation = ss.getSheetByName("FA-3 Depreciation");
  if (depreciation) {
    const depData = [
      [0, 0, 0],
      [1250000, 89750, 0],
      [980000, 350000, 45000],
      [185000, 64285, 12500],
      [145000, 65000, 15000],
      [425000, 170000, 65000],
      [210000, 95000, 7500]
    ];
    depreciation.getRange(6, 3, 7, 3).setValues(depData);
  }
  
  ui.alert('Success', 'Sample data has been added to the workpaper.', ui.ButtonSet.OK);
}

/**
 * Clears all data but keeps formatting
 */
function clearAllData() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Clear All Data',
    'This will clear all entered data but keep the formatting. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  sheets.forEach(sheet => {
    const lastRow = sheet.getMaxRows();
    const lastCol = sheet.getMaxColumns();
    
    // Clear content but keep formatting
    for (let row = 1; row <= lastRow; row++) {
      for (let col = 1; col <= lastCol; col++) {
        const cell = sheet.getRange(row, col);
        const background = cell.getBackground();
        
        // Only clear white cells (data entry cells)
        if (background === "#ffffff" || background === "#ffffcc") {
          cell.clearContent();
        }
      }
    }
  });
  
  ui.alert('Success', 'All data has been cleared. Formatting preserved.', ui.ButtonSet.OK);
}

/**
 * Export workpaper to PDF (placeholder - requires additional setup)
 */
function exportToPDF() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Export to PDF',
    'To export this workpaper to PDF:\n\n' +
    '1. Click File > Download > PDF\n' +
    '2. Select "Workbook" to export all sheets\n' +
    '3. Or select specific sheets to export',
    ui.ButtonSet.OK
  );
}