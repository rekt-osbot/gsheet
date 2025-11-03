/**
 * @name tds_compliance
 * @version 1.0.1
 * @built 2025-11-03T12:27:00.930Z
 * @description Standalone script. Do not edit directly - edit source files in src/ folder.
 * 
 * This file is auto-generated from:
 * - Common utilities (src/common/*.gs)
 * - Workbook-specific code (src/workbooks/tds_compliance.gs)
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
 * TDS COMPLIANCE TRACKER - GOOGLE SHEETS AUTOMATION
 * 
 * Purpose: Comprehensive TDS (Tax Deducted at Source) compliance workbook
 * Covers: Section-wise calculations, vendor master, 26AS reconciliation, quarterly returns
 * 
 * Main Function: createTDSComplianceWorkbook()
 * 
 * Standards: Income Tax Act, 1961 (India)
 * Last Updated: November 2024
 */

/**
 * MAIN FUNCTION - Creates complete TDS Compliance Workbook
 * Run this function to generate all sheets
 */
function createTDSComplianceWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Set workbook type for menu detection
  setWorkbookType('TDS_COMPLIANCE');
  
  // Show progress
  SpreadsheetApp.getActiveSpreadsheet().toast('Creating TDS Compliance Workbook...', 'Please Wait', -1);
  
  // Create all sheets in logical order
  createCoverSheet(ss);
  createAssumptionsSheet(ss);
  createVendorMasterSheet(ss);
  createTDSRegisterSheet(ss);
  createSectionRatesSheet(ss);
  createLowerDeductionCertSheet(ss);
  createTDSPayableLedgerSheet(ss);
  create26ASReconciliationSheet(ss);
  createQuarterlyReturnSheet(ss);
  createInterestCalculatorSheet(ss);
  createDashboardSheet(ss);
  createAuditNotesSheet(ss);
  
  // Set up named ranges for easy reference
  setupNamedRanges(ss);
  
  // Reorder sheets
  reorderSheets(ss);
  
  // Show completion message
  SpreadsheetApp.getActiveSpreadsheet().toast('TDS Compliance Workbook created successfully!', 'Complete', 5);
  ss.getSheetByName('Dashboard').activate();
}

/**
 * Creates Cover Sheet with workbook information
 */
function createCoverSheet(ss) {
  let sheet = ss.getSheetByName('Cover');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Cover');
  
  // Title section
  sheet.getRange('B2').setValue('TDS COMPLIANCE TRACKER').setFontSize(24).setFontWeight('bold').setFontColor('#1a237e');
  sheet.getRange('B3').setValue('Tax Deducted at Source - Income Tax Act, 1961').setFontSize(12).setFontColor('#666666');
  
  // Entity details
  sheet.getRange('B5').setValue('Entity Name:').setFontWeight('bold');
  sheet.getRange('C5').setBackground('#e3f2fd');
  
  sheet.getRange('B6').setValue('PAN:').setFontWeight('bold');
  sheet.getRange('C6').setBackground('#e3f2fd');
  
  sheet.getRange('B7').setValue('TAN:').setFontWeight('bold');
  sheet.getRange('C7').setBackground('#e3f2fd');
  
  sheet.getRange('B8').setValue('Financial Year:').setFontWeight('bold');
  sheet.getRange('C8').setValue('FY 2024-25').setBackground('#e3f2fd');
  
  sheet.getRange('B9').setValue('Assessment Year:').setFontWeight('bold');
  sheet.getRange('C9').setValue('AY 2025-26').setBackground('#e3f2fd');
  
  // Workbook contents
  sheet.getRange('B11').setValue('WORKBOOK CONTENTS').setFontSize(14).setFontWeight('bold').setFontColor('#1a237e');
  
  const contents = [
    ['Sheet Name', 'Purpose'],
    ['Dashboard', 'Summary view of TDS compliance status'],
    ['Assumptions', 'Entity details and configuration'],
    ['Vendor_Master', 'Vendor database with PAN details'],
    ['TDS_Register', 'Transaction-wise TDS deduction entries'],
    ['Section_Rates', 'TDS rates by section and entity type'],
    ['Lower_Deduction_Cert', 'Certificate tracker (Sec 197)'],
    ['TDS_Payable_Ledger', 'Month-wise TDS liability'],
    ['26AS_Reconciliation', 'Form 26AS vs Books reconciliation'],
    ['Quarterly_Return', 'Form 24Q/26Q preparation'],
    ['Interest_Calculator', 'Late payment interest (Sec 201)'],
    ['Audit_Notes', 'Documentation and references']
  ];
  
  sheet.getRange(12, 2, contents.length, 2).setValues(contents);
  sheet.getRange('B12:C12').setBackground('#1a237e').setFontColor('white').setFontWeight('bold');
  
  // Instructions
  sheet.getRange('B25').setValue('QUICK START GUIDE').setFontSize(14).setFontWeight('bold').setFontColor('#1a237e');
  sheet.getRange('B26').setValue('1. Fill entity details in Assumptions sheet');
  sheet.getRange('B27').setValue('2. Add vendors in Vendor_Master sheet');
  sheet.getRange('B28').setValue('3. Enter transactions in TDS_Register sheet');
  sheet.getRange('B29').setValue('4. Review Dashboard for compliance status');
  sheet.getRange('B30').setValue('5. Use 26AS_Reconciliation for quarterly verification');
  
  // Formatting
  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 400);
  
  sheet.getRange('B12:C23').setBorder(true, true, true, true, true, true);
}

/**
 * Creates Assumptions Sheet with entity configuration
 */
function createAssumptionsSheet(ss) {
  let sheet = ss.getSheetByName('Assumptions');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Assumptions');
  
  // Header
  sheet.getRange('A1:F1').merge().setValue('TDS COMPLIANCE - ASSUMPTIONS & CONFIGURATION')
    .setBackground('#1a237e').setFontColor('white').setFontWeight('bold').setFontSize(14);
  
  // Entity Details Section
  sheet.getRange('A3').setValue('ENTITY DETAILS').setFontWeight('bold').setFontSize(12).setBackground('#e3f2fd');
  sheet.getRange('A3:F3').merge();
  
  const entityDetails = [
    ['Entity Name', '', '', '', '', ''],
    ['PAN', '', 'TAN', '', '', ''],
    ['Address', '', '', '', '', ''],
    ['City', '', 'State', '', 'PIN', ''],
    ['Contact Person', '', 'Email', '', 'Phone', ''],
    ['', '', '', '', '', ''],
    ['Financial Year', 'FY 2024-25', 'Assessment Year', 'AY 2025-26', '', ''],
    ['Deductor Type', 'Company', '', '', '', ''],
    ['', '', '', '', '', '']
  ];
  
  sheet.getRange(4, 1, entityDetails.length, 6).setValues(entityDetails);
  sheet.getRange('B4:B8').setBackground('#fff3e0');
  sheet.getRange('D4:D8').setBackground('#fff3e0');
  sheet.getRange('F4:F5').setBackground('#fff3e0');
  sheet.getRange('B10:B11').setBackground('#fff3e0');
  sheet.getRange('D10:D11').setBackground('#fff3e0');
  
  // TDS Payment Configuration
  sheet.getRange('A13').setValue('TDS PAYMENT CONFIGURATION').setFontWeight('bold').setFontSize(12).setBackground('#e3f2fd');
  sheet.getRange('A13:F13').merge();
  
  const paymentConfig = [
    ['Due Date for Payment', '7th of next month', '', '', '', ''],
    ['Bank Name', '', '', '', '', ''],
    ['Bank Account No.', '', 'IFSC', '', '', ''],
    ['', '', '', '', '', ''],
    ['Interest Rate (Sec 201)', '1% per month', '', '', '', ''],
    ['Interest Rate (Sec 201(1A))', '1.5% per month', '', '', '', '']
  ];
  
  sheet.getRange(14, 1, paymentConfig.length, 6).setValues(paymentConfig);
  sheet.getRange('B15:B16').setBackground('#fff3e0');
  sheet.getRange('D16').setBackground('#fff3e0');
  
  // Quarterly Return Dates
  sheet.getRange('A20').setValue('QUARTERLY RETURN DUE DATES').setFontWeight('bold').setFontSize(12).setBackground('#e3f2fd');
  sheet.getRange('A20:F20').merge();
  
  const quarterDates = [
    ['Quarter', 'Period', 'Due Date (24Q)', 'Due Date (26Q)', '', ''],
    ['Q1', 'Apr-Jun', '31-Jul', '31-Jul', '', ''],
    ['Q2', 'Jul-Sep', '31-Oct', '31-Oct', '', ''],
    ['Q3', 'Oct-Dec', '31-Jan', '31-Jan', '', ''],
    ['Q4', 'Jan-Mar', '31-May', '31-May', '', '']
  ];
  
  sheet.getRange(21, 1, quarterDates.length, 6).setValues(quarterDates);
  sheet.getRange('A21:D21').setBackground('#1a237e').setFontColor('white').setFontWeight('bold');
  sheet.getRange('A22:D26').setBorder(true, true, true, true, true, true);
  
  // Formatting
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 100);
  
  sheet.setFrozenRows(1);
}

/**
 * Creates Vendor Master Sheet
 */
function createVendorMasterSheet(ss) {
  let sheet = ss.getSheetByName('Vendor_Master');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Vendor_Master');
  
  // Header - don't merge to avoid freeze conflict
  sheet.getRange('A1:L1').setValue('VENDOR MASTER - TDS DEDUCTEE DATABASE')
    .setBackground('#1a237e').setFontColor('white').setFontWeight('bold').setFontSize(14);
  
  // Instructions - don't merge to avoid freeze conflict
  sheet.getRange('A2:L2').setValue('Instructions: Enter all vendors/deductees. PAN is mandatory. Entity Type determines TDS rate.')
    .setBackground('#fff3e0').setFontStyle('italic');
  
  // Column headers
  const headers = [
    'Vendor Code',
    'Vendor Name',
    'PAN',
    'PAN Valid?',
    'Entity Type',
    'Address',
    'City',
    'State',
    'Email',
    'Phone',
    'Lower Deduction Cert?',
    'Remarks'
  ];
  
  sheet.getRange(3, 1, 1, headers.length).setValues([headers])
    .setBackground('#1a237e').setFontColor('white').setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Sample data
  const sampleData = [
    ['V001', 'ABC Contractors Pvt Ltd', 'AAAAA1234A', '=IF(LEN(C4)=10,IF(REGEXMATCH(C4,"^[A-Z]{5}[0-9]{4}[A-Z]$"),"✓","✗"),"✗")', 'Company', '123 Main St', 'Mumbai', 'Maharashtra', 'abc@example.com', '9876543210', 'No', ''],
    ['V002', 'XYZ Consultants', 'BBBBB5678B', '=IF(LEN(C5)=10,IF(REGEXMATCH(C5,"^[A-Z]{5}[0-9]{4}[A-Z]$"),"✓","✗"),"✗")', 'Individual', '456 Park Ave', 'Delhi', 'Delhi', 'xyz@example.com', '9876543211', 'Yes', 'Cert No: XYZ123'],
    ['', '', '', '=IF(LEN(C6)=10,IF(REGEXMATCH(C6,"^[A-Z]{5}[0-9]{4}[A-Z]$"),"✓","✗"),"✗")', '', '', '', '', '', '', '', '']
  ];
  
  sheet.getRange(4, 1, sampleData.length, headers.length).setValues(sampleData);
  
  // Data validation for Entity Type
  const entityTypes = ['Company', 'Individual', 'HUF', 'Firm', 'AOP/BOI', 'Trust', 'Government', 'Non-Resident'];
  const entityTypeRule = SpreadsheetApp.newDataValidation().requireValueInList(entityTypes).build();
  sheet.getRange('E4:E1000').setDataValidation(entityTypeRule);
  
  // Data validation for Lower Deduction Cert
  const yesNoRule = SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No']).build();
  sheet.getRange('K4:K1000').setDataValidation(yesNoRule);
  
  // Formatting
  sheet.getRange('C4:C1000').setBackground('#fff3e0'); // PAN input
  sheet.getRange('E4:E1000').setBackground('#fff3e0'); // Entity Type input
  sheet.getRange('D4:D1000').setHorizontalAlignment('center'); // PAN validation
  
  // Column widths
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 180);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 100);
  sheet.setColumnWidth(9, 180);
  sheet.setColumnWidth(10, 100);
  sheet.setColumnWidth(11, 150);
  sheet.setColumnWidth(12, 200);
  
  sheet.setFrozenRows(3);
  sheet.setFrozenColumns(2); // Only freeze first 2 columns to avoid merge conflict
}

/**
 * Creates TDS Register Sheet - Main transaction log
 */
function createTDSRegisterSheet(ss) {
  let sheet = ss.getSheetByName('TDS_Register');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('TDS_Register');
  
  // Header - don't merge to avoid freeze conflict
  sheet.getRange('A1:R1').setValue('TDS REGISTER - TRANSACTION-WISE DEDUCTIONS')
    .setBackground('#1a237e').setFontColor('white').setFontWeight('bold').setFontSize(14);
  
  // Instructions - don't merge to avoid freeze conflict
  sheet.getRange('A2:R2').setValue('Instructions: Enter each payment transaction. TDS will be calculated automatically based on section and vendor type.')
    .setBackground('#fff3e0').setFontStyle('italic');
  
  // Column headers
  const headers = [
    'Entry No.',
    'Date',
    'Vendor Code',
    'Vendor Name',
    'PAN',
    'Entity Type',
    'TDS Section',
    'Nature of Payment',
    'Gross Amount',
    'Threshold Limit',
    'Applicable Rate %',
    'TDS Amount',
    'Net Payment',
    'Payment Date',
    'Challan No.',
    'Challan Date',
    'Quarter',
    'Remarks'
  ];
  
  sheet.getRange(3, 1, 1, headers.length).setValues([headers])
    .setBackground('#1a237e').setFontColor('white').setFontWeight('bold')
    .setHorizontalAlignment('center').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
  // Sample formulas (row 4)
  const row = 4;
  sheet.getRange(`C${row}`).setBackground('#fff3e0'); // Vendor Code - input
  sheet.getRange(`D${row}`).setFormula(`=IFERROR(VLOOKUP(C${row},Vendor_Master!A:B,2,FALSE),"")`); // Vendor Name lookup
  sheet.getRange(`E${row}`).setFormula(`=IFERROR(VLOOKUP(C${row},Vendor_Master!A:C,3,FALSE),"")`); // PAN lookup
  sheet.getRange(`F${row}`).setFormula(`=IFERROR(VLOOKUP(C${row},Vendor_Master!A:E,5,FALSE),"")`); // Entity Type lookup
  sheet.getRange(`G${row}`).setBackground('#fff3e0'); // TDS Section - input
  sheet.getRange(`H${row}`).setBackground('#fff3e0'); // Nature - input
  sheet.getRange(`I${row}`).setBackground('#fff3e0').setNumberFormat('#,##0.00'); // Gross Amount - input
  sheet.getRange(`J${row}`).setFormula(`=IFERROR(VLOOKUP(G${row},Section_Rates!A:C,3,FALSE),0)`); // Threshold lookup
  sheet.getRange(`K${row}`).setFormula(`=IFERROR(VLOOKUP(G${row}&F${row},Section_Rates!D:E,2,FALSE),0)`); // Rate lookup
  sheet.getRange(`L${row}`).setFormula(`=IF(I${row}>J${row},ROUND(I${row}*K${row}/100,0),0)`).setNumberFormat('#,##0.00'); // TDS calculation
  sheet.getRange(`M${row}`).setFormula(`=I${row}-L${row}`).setNumberFormat('#,##0.00'); // Net payment
  sheet.getRange(`N${row}`).setBackground('#fff3e0'); // Payment Date - input
  sheet.getRange(`O${row}`).setBackground('#fff3e0'); // Challan No - input
  sheet.getRange(`P${row}`).setBackground('#fff3e0'); // Challan Date - input
  sheet.getRange(`Q${row}`).setFormula(`=IF(MONTH(B${row})<=3,"Q4",IF(MONTH(B${row})<=6,"Q1",IF(MONTH(B${row})<=9,"Q2","Q3")))`); // Quarter
  sheet.getRange(`R${row}`).setBackground('#fff3e0'); // Remarks - input
  
  // Copy formulas down for 100 rows
  sheet.getRange(`C${row}:R${row}`).copyTo(sheet.getRange(`C${row}:R${row+99}`), SpreadsheetApp.CopyPasteType.PASTE_FORMULA);
  
  // Add date to first column
  sheet.getRange(`B${row}`).setBackground('#fff3e0');
  sheet.getRange(`B${row}:B${row+99}`).setNumberFormat('dd-mmm-yyyy');
  
  // Auto-number Entry No
  for (let i = 0; i < 100; i++) {
    sheet.getRange(row + i, 1).setFormula(`=IF(B${row+i}<>"",ROW()-3,"")`);
  }
  
  // Column widths
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 150);
  sheet.setColumnWidth(9, 120);
  sheet.setColumnWidth(10, 120);
  sheet.setColumnWidth(11, 100);
  sheet.setColumnWidth(12, 120);
  sheet.setColumnWidth(13, 120);
  sheet.setColumnWidth(14, 100);
  sheet.setColumnWidth(15, 120);
  sheet.setColumnWidth(16, 100);
  sheet.setColumnWidth(17, 80);
  sheet.setColumnWidth(18, 150);
  
  sheet.setFrozenRows(3);
  sheet.setFrozenColumns(2);
}

/**
 * Creates Section Rates Sheet - TDS rates by section and entity type
 */
function createSectionRatesSheet(ss) {
  let sheet = ss.getSheetByName('Section_Rates');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Section_Rates');
  
  // Header
  sheet.getRange('A1:J1').merge().setValue('TDS SECTION RATES - FY 2024-25')
    .setBackground('#1a237e').setFontColor('white').setFontWeight('bold').setFontSize(14);
  
  // Instructions
  sheet.getRange('A2:J2').merge().setValue('Reference: Income Tax Act rates as of FY 2024-25. Update rates as per amendments.')
    .setBackground('#fff3e0').setFontStyle('italic');
  
  // Column headers
  const headers = [
    'Section',
    'Nature of Payment',
    'Threshold (₹)',
    'Lookup Key',
    'Company',
    'Individual',
    'HUF',
    'Firm',
    'Non-Resident',
    'Remarks'
  ];
  
  sheet.getRange(3, 1, 1, headers.length).setValues([headers])
    .setBackground('#1a237e').setFontColor('white').setFontWeight('bold')
    .setHorizontalAlignment('center').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
  // TDS Rates Data
  const ratesData = [
    ['192', 'Salary', 0, '192Company', 'As per slab', 'As per slab', 'As per slab', 'As per slab', 'As per slab', 'Refer IT slabs'],
    ['192A', 'Premature EPF withdrawal', 50000, '192ACompany', 10, 10, 10, 10, 10, 'If PAN not available: 20%'],
    ['193', 'Interest on securities', 10000, '193Company', 10, 10, 10, 10, 20, 'Listed debentures'],
    ['194', 'Dividend', 5000, '194Company', 10, 10, 10, 10, 20, 'Resident shareholders'],
    ['194A', 'Interest other than securities', 40000, '194ACompany', 10, 10, 10, 10, 20, 'Bank/Co-op: 50,000; Senior citizen: 50,000'],
    ['194B', 'Lottery/Crossword/Game show', 10000, '194BCompany', 30, 30, 30, 30, 30, 'No threshold for non-residents'],
    ['194C', 'Payment to contractors', 100000, '194CCompany', 1, 1, 2, 1, 2, 'Individual/HUF: 2%; Others: 1%'],
    ['194C', 'Payment to contractors', 100000, '194CIndividual', 2, 2, 2, 1, 2, 'Individual/HUF: 2%; Others: 1%'],
    ['194C', 'Payment to contractors', 100000, '194CHUF', 2, 1, 2, 1, 2, 'Individual/HUF: 2%; Others: 1%'],
    ['194D', 'Insurance commission', 15000, '194DCompany', 5, 5, 5, 5, 5, 'Includes life insurance'],
    ['194DA', 'Life insurance maturity', 100000, '194DACompany', 5, 5, 5, 5, 5, 'On amount exceeding ₹1 lakh'],
    ['194EE', 'NSS deposit', 2500, '194EECompany', 10, 10, 10, 10, 10, 'National Savings Scheme'],
    ['194F', 'Mutual fund repurchase', 0, '194FCompany', 20, 20, 20, 20, 20, 'No threshold'],
    ['194G', 'Commission on lottery tickets', 15000, '194GCompany', 5, 5, 5, 5, 5, ''],
    ['194H', 'Commission/Brokerage', 15000, '194HCompany', 5, 5, 5, 5, 5, 'Excludes insurance commission'],
    ['194I', 'Rent - Plant & Machinery', 240000, '194ICompany', 2, 2, 2, 2, 2, 'Threshold per year'],
    ['194I', 'Rent - Land/Building/Furniture', 240000, '194ICompany', 10, 10, 10, 10, 10, 'Threshold per year'],
    ['194IA', 'Transfer of immovable property', 5000000, '194IACompany', 1, 1, 1, 1, 1, 'Buyer deducts TDS'],
    ['194IB', 'Rent by individual/HUF', 600000, '194IBIndividual', 5, 5, 5, 5, 5, 'Only if not liable for tax audit'],
    ['194IC', 'Joint Development Agreement', 0, '194ICCompany', 10, 10, 10, 10, 10, 'Landowner receipt'],
    ['194J', 'Professional/Technical fees', 30000, '194JCompany', 10, 10, 10, 10, 20, 'Call center: 2%'],
    ['194J', 'Royalty', 30000, '194JCompany', 10, 10, 10, 10, 20, ''],
    ['194K', 'Income from units (Mutual Fund)', 5000, '194KCompany', 10, 10, 10, 10, 20, 'Effective from 01-Apr-2020'],
    ['194LA', 'Compensation on land acquisition', 250000, '194LACompany', 10, 10, 10, 10, 10, 'Compulsory acquisition'],
    ['194LB', 'Interest on infrastructure debt fund', 5000, '194LBCompany', 5, 5, 5, 5, 5, ''],
    ['194LBA', 'Business trust income', 0, '194LBACompany', 10, 10, 10, 10, 10, 'REIT/InvIT'],
    ['194LBB', 'Investment fund income', 0, '194LBBCompany', 10, 10, 10, 10, 10, ''],
    ['194LBC', 'Income from securitization trust', 0, '194LBCCompany', 25, 25, 25, 25, 30, 'Individuals: 30%'],
    ['194M', 'Payment by individuals/HUF', 5000000, '194MCompany', 5, 5, 5, 5, 5, 'Aggregate threshold per year'],
    ['194N', 'Cash withdrawal', 10000000, '194NCompany', 2, 2, 2, 2, 2, 'Banking company/co-op/post office'],
    ['194O', 'E-commerce participants', 500000, '194OCompany', 1, 1, 1, 1, 1, 'By e-commerce operator'],
    ['194Q', 'Purchase of goods', 5000000, '194QCompany', 0.1, 0.1, 0.1, 0.1, 0.1, 'Buyer deducts if turnover >10Cr'],
    ['195', 'Non-resident payments', 0, '195Non-Resident', 'Varies', 'Varies', 'Varies', 'Varies', 'Varies', 'As per DTAA/Act']
  ];
  
  sheet.getRange(4, 1, ratesData.length, headers.length).setValues(ratesData);
  
  // Formatting
  sheet.getRange('C4:C' + (3 + ratesData.length)).setNumberFormat('#,##0');
  sheet.getRange('E4:I' + (3 + ratesData.length)).setNumberFormat('0.00"%"');
  
  // Alternate row colors
  for (let i = 4; i <= 3 + ratesData.length; i++) {
    if (i % 2 === 0) {
      sheet.getRange(i, 1, 1, headers.length).setBackground('#f5f5f5');
    }
  }
  
  // Column widths
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 90);
  sheet.setColumnWidth(6, 90);
  sheet.setColumnWidth(7, 80);
  sheet.setColumnWidth(8, 80);
  sheet.setColumnWidth(9, 110);
  sheet.setColumnWidth(10, 250);
  
  sheet.setFrozenRows(3);
}

/**
 * Creates Lower Deduction Certificate Tracker
 */
function createLowerDeductionCertSheet(ss) {
  let sheet = ss.getSheetByName('Lower_Deduction_Cert');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Lower_Deduction_Cert');
  
  // Header
  sheet.getRange('A1:J1').merge().setValue('LOWER DEDUCTION CERTIFICATE TRACKER (Section 197)')
    .setBackground('#1a237e').setFontColor('white').setFontWeight('bold').setFontSize(14);
  
  // Instructions
  sheet.getRange('A2:J2').merge().setValue('Instructions: Track certificates issued by Income Tax Department for lower/nil TDS deduction.')
    .setBackground('#fff3e0').setFontStyle('italic');
  
  // Column headers
  const headers = [
    'Cert No.',
    'Vendor Code',
    'Vendor Name',
    'PAN',
    'Valid From',
    'Valid To',
    'TDS Section',
    'Reduced Rate %',
    'Max Amount (₹)',
    'Status'
  ];
  
  sheet.getRange(3, 1, 1, headers.length).setValues([headers])
    .setBackground('#1a237e').setFontColor('white').setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Sample data with formulas
  const row = 4;
  sheet.getRange(`A${row}`).setBackground('#fff3e0'); // Cert No - input
  sheet.getRange(`B${row}`).setBackground('#fff3e0'); // Vendor Code - input
  sheet.getRange(`C${row}`).setFormula(`=IFERROR(VLOOKUP(B${row},Vendor_Master!A:B,2,FALSE),"")`); // Vendor Name lookup
  sheet.getRange(`D${row}`).setFormula(`=IFERROR(VLOOKUP(B${row},Vendor_Master!A:C,3,FALSE),"")`); // PAN lookup
  sheet.getRange(`E${row}`).setBackground('#fff3e0').setNumberFormat('dd-mmm-yyyy'); // Valid From - input
  sheet.getRange(`F${row}`).setBackground('#fff3e0').setNumberFormat('dd-mmm-yyyy'); // Valid To - input
  sheet.getRange(`G${row}`).setBackground('#fff3e0'); // Section - input
  sheet.getRange(`H${row}`).setBackground('#fff3e0').setNumberFormat('0.00"%"'); // Rate - input
  sheet.getRange(`I${row}`).setBackground('#fff3e0').setNumberFormat('#,##0.00'); // Max Amount - input
  sheet.getRange(`J${row}`).setFormula(`=IF(F${row}<TODAY(),"Expired",IF(E${row}>TODAY(),"Not Yet Valid","Active"))`); // Status
  
  // Copy formulas down
  sheet.getRange(`A${row}:J${row}`).copyTo(sheet.getRange(`A${row}:J${row+49}`), SpreadsheetApp.CopyPasteType.PASTE_FORMULA);
  
  // Conditional formatting for status
  const activeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Active')
    .setBackground('#d4edda')
    .setRanges([sheet.getRange('J4:J53')])
    .build();
  
  const expiredRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Expired')
    .setBackground('#f8d7da')
    .setRanges([sheet.getRange('J4:J53')])
    .build();
  
  sheet.setConditionalFormatRules([activeRule, expiredRule]);
  
  // Column widths
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 110);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 120);
  sheet.setColumnWidth(9, 130);
  sheet.setColumnWidth(10, 120);
  
  sheet.setFrozenRows(3);
}

/**
 * Creates TDS Payable Ledger - Month-wise liability tracking
 */
function createTDSPayableLedgerSheet(ss) {
  let sheet = ss.getSheetByName('TDS_Payable_Ledger');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('TDS_Payable_Ledger');
  
  // Header
  sheet.getRange('A1:H1').merge().setValue('TDS PAYABLE LEDGER - MONTH-WISE LIABILITY')
    .setBackground('#1a237e').setFontColor('white').setFontWeight('bold').setFontSize(14);
  
  // Instructions
  sheet.getRange('A2:H2').merge().setValue('Instructions: Auto-calculated from TDS_Register. Track monthly TDS liability and payment status.')
    .setBackground('#fff3e0').setFontStyle('italic');
  
  // Column headers
  const headers = [
    'Month',
    'TDS Deducted (₹)',
    'Due Date',
    'Challan No.',
    'Payment Date',
    'Amount Paid (₹)',
    'Balance (₹)',
    'Status'
  ];
  
  sheet.getRange(3, 1, 1, headers.length).setValues([headers])
    .setBackground('#1a237e').setFontColor('white').setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Month list for FY 2024-25
  const months = [
    ['Apr-2024'],
    ['May-2024'],
    ['Jun-2024'],
    ['Jul-2024'],
    ['Aug-2024'],
    ['Sep-2024'],
    ['Oct-2024'],
    ['Nov-2024'],
    ['Dec-2024'],
    ['Jan-2025'],
    ['Feb-2025'],
    ['Mar-2025']
  ];
  
  sheet.getRange(4, 1, months.length, 1).setValues(months);
  
  // Formulas for each month
  for (let i = 0; i < months.length; i++) {
    const row = 4 + i;
    const monthYear = months[i][0];
    
    // TDS Deducted - SUMIFS from TDS_Register
    sheet.getRange(`B${row}`).setFormula(
      `=SUMIFS(TDS_Register!L:L,TDS_Register!B:B,">="&DATE(${monthYear.split('-')[1]},${getMonthNumber(monthYear.split('-')[0])},1),TDS_Register!B:B,"<"&DATE(${monthYear.split('-')[1]},${getMonthNumber(monthYear.split('-')[0])}+1,1))`
    ).setNumberFormat('#,##0.00');
    
    // Due Date - 7th of next month
    sheet.getRange(`C${row}`).setFormula(
      `=DATE(${monthYear.split('-')[1]},${getMonthNumber(monthYear.split('-')[0])}+1,7)`
    ).setNumberFormat('dd-mmm-yyyy');
    
    // Challan No - input
    sheet.getRange(`D${row}`).setBackground('#fff3e0');
    
    // Payment Date - input
    sheet.getRange(`E${row}`).setBackground('#fff3e0').setNumberFormat('dd-mmm-yyyy');
    
    // Amount Paid - input
    sheet.getRange(`F${row}`).setBackground('#fff3e0').setNumberFormat('#,##0.00');
    
    // Balance
    sheet.getRange(`G${row}`).setFormula(`=B${row}-F${row}`).setNumberFormat('#,##0.00');
    
    // Status
    sheet.getRange(`H${row}`).setFormula(
      `=IF(B${row}=0,"No TDS",IF(G${row}=0,"Paid",IF(E${row}>C${row},"Late Payment",IF(E${row}="","Pending","Paid"))))`
    );
  }
  
  // Conditional formatting for status
  const paidRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Paid')
    .setBackground('#d4edda')
    .setRanges([sheet.getRange('H4:H15')])
    .build();
  
  const pendingRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Pending')
    .setBackground('#fff3cd')
    .setRanges([sheet.getRange('H4:H15')])
    .build();
  
  const lateRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Late Payment')
    .setBackground('#f8d7da')
    .setRanges([sheet.getRange('H4:H15')])
    .build();
  
  sheet.setConditionalFormatRules([paidRule, pendingRule, lateRule]);
  
  // Total row
  sheet.getRange('A16').setValue('TOTAL').setFontWeight('bold');
  sheet.getRange('B16').setFormula('=SUM(B4:B15)').setNumberFormat('#,##0.00').setFontWeight('bold');
  sheet.getRange('F16').setFormula('=SUM(F4:F15)').setNumberFormat('#,##0.00').setFontWeight('bold');
  sheet.getRange('G16').setFormula('=SUM(G4:G15)').setNumberFormat('#,##0.00').setFontWeight('bold');
  sheet.getRange('A16:H16').setBackground('#e3f2fd');
  
  // Column widths
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 150);
  sheet.setColumnWidth(8, 120);
  
  sheet.setFrozenRows(3);
}

// Helper function for month number
function getMonthNumber(monthName) {
  const months = {
    'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
    'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12
  };
  return months[monthName];
}

/**
 * Creates 26AS Reconciliation Sheet
 */
function create26ASReconciliationSheet(ss) {
  let sheet = ss.getSheetByName('26AS_Reconciliation');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('26AS_Reconciliation');
  
  // Header
  sheet.getRange('A1:J1').merge().setValue('FORM 26AS RECONCILIATION')
    .setBackground('#1a237e').setFontColor('white').setFontWeight('bold').setFontSize(14);
  
  // Instructions
  sheet.getRange('A2:J2').merge().setValue('Instructions: Download Form 26AS from TRACES. Enter data here and compare with books.')
    .setBackground('#fff3e0').setFontStyle('italic');
  
  // Section 1: Books Data
  sheet.getRange('A4').setValue('AS PER BOOKS').setFontWeight('bold').setFontSize(12).setBackground('#e3f2fd');
  sheet.getRange('A4:J4').merge();
  
  const booksHeaders = [
    'Quarter',
    'Section',
    'No. of Entries',
    'Gross Amount (₹)',
    'TDS Amount (₹)',
    '',
    '',
    '',
    '',
    ''
  ];
  
  sheet.getRange(5, 1, 1, booksHeaders.length).setValues([booksHeaders])
    .setBackground('#1a237e').setFontColor('white').setFontWeight('bold');
  
  // Books data with formulas
  const quarters = ['Q1', 'Q2', 'Q3', 'Q4'];
  for (let i = 0; i < quarters.length; i++) {
    const row = 6 + i;
    sheet.getRange(`A${row}`).setValue(quarters[i]);
    sheet.getRange(`B${row}`).setValue('All Sections');
    sheet.getRange(`C${row}`).setFormula(`=COUNTIF(TDS_Register!Q:Q,"${quarters[i]}")`);
    sheet.getRange(`D${row}`).setFormula(`=SUMIF(TDS_Register!Q:Q,"${quarters[i]}",TDS_Register!I:I)`).setNumberFormat('#,##0.00');
    sheet.getRange(`E${row}`).setFormula(`=SUMIF(TDS_Register!Q:Q,"${quarters[i]}",TDS_Register!L:L)`).setNumberFormat('#,##0.00');
  }
  
  // Total
  sheet.getRange('A10').setValue('TOTAL').setFontWeight('bold');
  sheet.getRange('C10').setFormula('=SUM(C6:C9)').setFontWeight('bold');
  sheet.getRange('D10').setFormula('=SUM(D6:D9)').setNumberFormat('#,##0.00').setFontWeight('bold');
  sheet.getRange('E10').setFormula('=SUM(E6:E9)').setNumberFormat('#,##0.00').setFontWeight('bold');
  sheet.getRange('A10:E10').setBackground('#e3f2fd');
  
  // Section 2: Form 26AS Data
  sheet.getRange('A12').setValue('AS PER FORM 26AS').setFontWeight('bold').setFontSize(12).setBackground('#fff3cd');
  sheet.getRange('A12:J12').merge();
  
  const form26ASHeaders = [
    'Quarter',
    'Section',
    'No. of Entries',
    'Gross Amount (₹)',
    'TDS Amount (₹)',
    '',
    '',
    '',
    '',
    ''
  ];
  
  sheet.getRange(13, 1, 1, form26ASHeaders.length).setValues([form26ASHeaders])
    .setBackground('#1a237e').setFontColor('white').setFontWeight('bold');
  
  // 26AS data - input fields
  for (let i = 0; i < quarters.length; i++) {
    const row = 14 + i;
    sheet.getRange(`A${row}`).setValue(quarters[i]);
    sheet.getRange(`B${row}`).setValue('All Sections');
    sheet.getRange(`C${row}`).setBackground('#fff3e0'); // Input
    sheet.getRange(`D${row}`).setBackground('#fff3e0').setNumberFormat('#,##0.00'); // Input
    sheet.getRange(`E${row}`).setBackground('#fff3e0').setNumberFormat('#,##0.00'); // Input
  }
  
  // Total
  sheet.getRange('A18').setValue('TOTAL').setFontWeight('bold');
  sheet.getRange('C18').setFormula('=SUM(C14:C17)').setFontWeight('bold');
  sheet.getRange('D18').setFormula('=SUM(D14:D17)').setNumberFormat('#,##0.00').setFontWeight('bold');
  sheet.getRange('E18').setFormula('=SUM(E14:E17)').setNumberFormat('#,##0.00').setFontWeight('bold');
  sheet.getRange('A18:E18').setBackground('#fff3cd');
  
  // Section 3: Variance Analysis
  sheet.getRange('A20').setValue('VARIANCE ANALYSIS').setFontWeight('bold').setFontSize(12).setBackground('#f8d7da');
  sheet.getRange('A20:J20').merge();
  
  const varianceHeaders = [
    'Quarter',
    'Variance - Entries',
    'Variance - Gross (₹)',
    'Variance - TDS (₹)',
    'Status',
    '',
    '',
    '',
    '',
    ''
  ];
  
  sheet.getRange(21, 1, 1, varianceHeaders.length).setValues([varianceHeaders])
    .setBackground('#1a237e').setFontColor('white').setFontWeight('bold');
  
  // Variance calculations
  for (let i = 0; i < quarters.length; i++) {
    const booksRow = 6 + i;
    const form26ASRow = 14 + i;
    const varianceRow = 22 + i;
    
    sheet.getRange(`A${varianceRow}`).setValue(quarters[i]);
    sheet.getRange(`B${varianceRow}`).setFormula(`=C${booksRow}-C${form26ASRow}`);
    sheet.getRange(`C${varianceRow}`).setFormula(`=D${booksRow}-D${form26ASRow}`).setNumberFormat('#,##0.00');
    sheet.getRange(`D${varianceRow}`).setFormula(`=E${booksRow}-E${form26ASRow}`).setNumberFormat('#,##0.00');
    sheet.getRange(`E${varianceRow}`).setFormula(`=IF(D${varianceRow}=0,"Matched","Variance")`);
  }
  
  // Conditional formatting for variance status
  const matchedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Matched')
    .setBackground('#d4edda')
    .setRanges([sheet.getRange('E22:E25')])
    .build();
  
  const varianceRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Variance')
    .setBackground('#f8d7da')
    .setRanges([sheet.getRange('E22:E25')])
    .build();
  
  sheet.setConditionalFormatRules([matchedRule, varianceRule]);
  
  // Column widths
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 150);
  
  sheet.setFrozenRows(5);
}

/**
 * Creates Quarterly Return Sheet (24Q/26Q preparation)
 */
function createQuarterlyReturnSheet(ss) {
  let sheet = ss.getSheetByName('Quarterly_Return');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Quarterly_Return');
  
  // Header
  sheet.getRange('A1:M1').merge().setValue('QUARTERLY TDS RETURN PREPARATION (Form 24Q/26Q)')
    .setBackground('#1a237e').setFontColor('white').setFontWeight('bold').setFontSize(14);
  
  // Instructions
  sheet.getRange('A2:M2').merge().setValue('Instructions: Select quarter to generate return data. Use this for filing quarterly TDS returns.')
    .setBackground('#fff3e0').setFontStyle('italic');
  
  // Quarter selection
  sheet.getRange('A4').setValue('Select Quarter:').setFontWeight('bold');
  sheet.getRange('B4').setBackground('#fff3e0');
  const quarterRule = SpreadsheetApp.newDataValidation().requireValueInList(['Q1', 'Q2', 'Q3', 'Q4']).build();
  sheet.getRange('B4').setDataValidation(quarterRule);
  sheet.getRange('B4').setValue('Q1');
  
  sheet.getRange('D4').setValue('Return Type:').setFontWeight('bold');
  sheet.getRange('E4').setBackground('#fff3e0');
  const returnTypeRule = SpreadsheetApp.newDataValidation().requireValueInList(['24Q (Salary)', '26Q (Non-Salary)']).build();
  sheet.getRange('E4').setDataValidation(returnTypeRule);
  sheet.getRange('E4').setValue('26Q (Non-Salary)');
  
  // Summary section
  sheet.getRange('A6').setValue('RETURN SUMMARY').setFontWeight('bold').setFontSize(12).setBackground('#e3f2fd');
  sheet.getRange('A6:M6').merge();
  
  const summaryLabels = [
    ['Total Deductees:', '=COUNTIF(TDS_Register!Q:Q,B4)', '', 'Total Transactions:', '=COUNTIF(TDS_Register!Q:Q,B4)'],
    ['Total Gross Amount:', '=SUMIF(TDS_Register!Q:Q,B4,TDS_Register!I:I)', '', 'Total TDS Deducted:', '=SUMIF(TDS_Register!Q:Q,B4,TDS_Register!L:L)'],
    ['TDS Deposited:', '', '', 'Balance Payable:', '']
  ];
  
  for (let i = 0; i < summaryLabels.length; i++) {
    sheet.getRange(7 + i, 1).setValue(summaryLabels[i][0]).setFontWeight('bold');
    sheet.getRange(7 + i, 2).setFormula(summaryLabels[i][1]).setNumberFormat('#,##0.00');
    sheet.getRange(7 + i, 4).setValue(summaryLabels[i][3]).setFontWeight('bold');
    sheet.getRange(7 + i, 5).setFormula(summaryLabels[i][4]).setNumberFormat('#,##0.00');
  }
  
  // Deductee-wise details
  sheet.getRange('A11').setValue('DEDUCTEE-WISE DETAILS').setFontWeight('bold').setFontSize(12).setBackground('#e3f2fd');
  sheet.getRange('A11:M11').merge();
  
  const detailHeaders = [
    'Sr. No.',
    'Deductee Name',
    'PAN',
    'Section',
    'Payment Date',
    'Gross Amount',
    'TDS Rate %',
    'TDS Amount',
    'Challan No.',
    'Challan Date',
    'BSR Code',
    'Challan Serial No.',
    'Remarks'
  ];
  
  sheet.getRange(12, 1, 1, detailHeaders.length).setValues([detailHeaders])
    .setBackground('#1a237e').setFontColor('white').setFontWeight('bold')
    .setHorizontalAlignment('center').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
  // Formula to pull data for selected quarter
  const startRow = 13;
  for (let i = 0; i < 50; i++) {
    const row = startRow + i;
    
    // This is simplified - in production, you'd use FILTER or QUERY function
    sheet.getRange(`A${row}`).setFormula(`=IF(ROW()-12<=COUNTIF(TDS_Register!Q:Q,$B$4),ROW()-12,"")`);
    
    // Note: These formulas are placeholders. In actual implementation, 
    // you'd use FILTER or QUERY to pull matching records
    sheet.getRange(`B${row}`).setValue(''); // Would use FILTER
    sheet.getRange(`C${row}`).setValue('');
    sheet.getRange(`D${row}`).setValue('');
    sheet.getRange(`E${row}`).setValue('').setNumberFormat('dd-mmm-yyyy');
    sheet.getRange(`F${row}`).setValue('').setNumberFormat('#,##0.00');
    sheet.getRange(`G${row}`).setValue('').setNumberFormat('0.00"%"');
    sheet.getRange(`H${row}`).setValue('').setNumberFormat('#,##0.00');
    sheet.getRange(`I${row}`).setValue('');
    sheet.getRange(`J${row}`).setValue('').setNumberFormat('dd-mmm-yyyy');
    sheet.getRange(`K${row}`).setBackground('#fff3e0'); // BSR Code - input
    sheet.getRange(`L${row}`).setBackground('#fff3e0'); // Serial No - input
    sheet.getRange(`M${row}`).setValue('');
  }
  
  // Add note about FILTER function
  sheet.getRange('A65').setValue('Note: Use FILTER or QUERY function to automatically populate deductee details based on selected quarter.')
    .setFontStyle('italic').setFontColor('#666666');
  sheet.getRange('A65:M65').merge();
  
  // Column widths
  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 90);
  sheet.setColumnWidth(8, 120);
  sheet.setColumnWidth(9, 120);
  sheet.setColumnWidth(10, 100);
  sheet.setColumnWidth(11, 100);
  sheet.setColumnWidth(12, 120);
  sheet.setColumnWidth(13, 150);
  
  sheet.setFrozenRows(12);
}

/**
 * Creates Interest Calculator Sheet (Section 201)
 */
function createInterestCalculatorSheet(ss) {
  let sheet = ss.getSheetByName('Interest_Calculator');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Interest_Calculator');
  
  // Header
  sheet.getRange('A1:H1').merge().setValue('INTEREST CALCULATOR - LATE TDS PAYMENT (Section 201)')
    .setBackground('#1a237e').setFontColor('white').setFontWeight('bold').setFontSize(14);
  
  // Instructions
  sheet.getRange('A2:H2').merge().setValue('Instructions: Interest calculated automatically for late TDS payments. Section 201(1): 1% per month; Section 201(1A): 1.5% per month.')
    .setBackground('#fff3e0').setFontStyle('italic');
  
  // Interest rates reference
  sheet.getRange('A4').setValue('INTEREST RATES').setFontWeight('bold').setFontSize(12).setBackground('#e3f2fd');
  sheet.getRange('A4:H4').merge();
  
  const ratesInfo = [
    ['Section 201(1)', 'Late payment of TDS', '1.00% per month', 'From due date to payment date'],
    ['Section 201(1A)', 'Late filing of TDS return', '1.50% per month', 'From due date to filing date']
  ];
  
  sheet.getRange(5, 1, ratesInfo.length, 4).setValues(ratesInfo);
  sheet.getRange('A5:D6').setBorder(true, true, true, true, true, true);
  
  // Late payment analysis
  sheet.getRange('A8').setValue('LATE PAYMENT ANALYSIS').setFontWeight('bold').setFontSize(12).setBackground('#e3f2fd');
  sheet.getRange('A8:H8').merge();
  
  const headers = [
    'Month',
    'TDS Amount (₹)',
    'Due Date',
    'Payment Date',
    'Delay (Days)',
    'Interest Rate',
    'Interest Amount (₹)',
    'Total Payable (₹)'
  ];
  
  sheet.getRange(9, 1, 1, headers.length).setValues([headers])
    .setBackground('#1a237e').setFontColor('white').setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Pull data from TDS_Payable_Ledger
  const months = [
    'Apr-2024', 'May-2024', 'Jun-2024', 'Jul-2024', 'Aug-2024', 'Sep-2024',
    'Oct-2024', 'Nov-2024', 'Dec-2024', 'Jan-2025', 'Feb-2025', 'Mar-2025'
  ];
  
  for (let i = 0; i < months.length; i++) {
    const row = 10 + i;
    const ledgerRow = 4 + i; // Corresponding row in TDS_Payable_Ledger
    
    sheet.getRange(`A${row}`).setValue(months[i]);
    
    // TDS Amount from ledger
    sheet.getRange(`B${row}`).setFormula(`=TDS_Payable_Ledger!B${ledgerRow}`).setNumberFormat('#,##0.00');
    
    // Due Date from ledger
    sheet.getRange(`C${row}`).setFormula(`=TDS_Payable_Ledger!C${ledgerRow}`).setNumberFormat('dd-mmm-yyyy');
    
    // Payment Date from ledger
    sheet.getRange(`D${row}`).setFormula(`=TDS_Payable_Ledger!E${ledgerRow}`).setNumberFormat('dd-mmm-yyyy');
    
    // Delay in days
    sheet.getRange(`E${row}`).setFormula(`=IF(D${row}="",0,MAX(0,D${row}-C${row}))`);
    
    // Interest Rate (1% per month = 0.0333% per day approx)
    sheet.getRange(`F${row}`).setValue('1% p.m.').setNumberFormat('0.00"% p.m."');
    
    // Interest calculation: TDS Amount × 1% × (Delay Days / 30)
    sheet.getRange(`G${row}`).setFormula(`=IF(E${row}>0,ROUND(B${row}*0.01*(E${row}/30),0),0)`).setNumberFormat('#,##0.00');
    
    // Total Payable
    sheet.getRange(`H${row}`).setFormula(`=B${row}+G${row}`).setNumberFormat('#,##0.00');
  }
  
  // Total row
  sheet.getRange('A22').setValue('TOTAL').setFontWeight('bold');
  sheet.getRange('B22').setFormula('=SUM(B10:B21)').setNumberFormat('#,##0.00').setFontWeight('bold');
  sheet.getRange('G22').setFormula('=SUM(G10:G21)').setNumberFormat('#,##0.00').setFontWeight('bold');
  sheet.getRange('H22').setFormula('=SUM(H10:H21)').setNumberFormat('#,##0.00').setFontWeight('bold');
  sheet.getRange('A22:H22').setBackground('#f8d7da');
  
  // Highlight rows with interest
  const interestRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$G10>0')
    .setBackground('#fff3cd')
    .setRanges([sheet.getRange('A10:H21')])
    .build();
  
  sheet.setConditionalFormatRules([interestRule]);
  
  // Additional calculator section
  sheet.getRange('A24').setValue('MANUAL INTEREST CALCULATOR').setFontWeight('bold').setFontSize(12).setBackground('#e3f2fd');
  sheet.getRange('A24:H24').merge();
  
  sheet.getRange('A25').setValue('TDS Amount (₹):').setFontWeight('bold');
  sheet.getRange('B25').setBackground('#fff3e0').setNumberFormat('#,##0.00');
  
  sheet.getRange('A26').setValue('Due Date:').setFontWeight('bold');
  sheet.getRange('B26').setBackground('#fff3e0').setNumberFormat('dd-mmm-yyyy');
  
  sheet.getRange('A27').setValue('Payment Date:').setFontWeight('bold');
  sheet.getRange('B27').setBackground('#fff3e0').setNumberFormat('dd-mmm-yyyy');
  
  sheet.getRange('A28').setValue('Delay (Days):').setFontWeight('bold');
  sheet.getRange('B28').setFormula('=MAX(0,B27-B26)');
  
  sheet.getRange('A29').setValue('Interest @ 1% p.m.:').setFontWeight('bold');
  sheet.getRange('B29').setFormula('=ROUND(B25*0.01*(B28/30),0)').setNumberFormat('#,##0.00');
  
  sheet.getRange('A30').setValue('Total Payable:').setFontWeight('bold');
  sheet.getRange('B30').setFormula('=B25+B29').setNumberFormat('#,##0.00').setBackground('#fff3cd');
  
  // Column widths
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 110);
  sheet.setColumnWidth(7, 130);
  sheet.setColumnWidth(8, 130);
  
  sheet.setFrozenRows(9);
}

/**
 * Creates Dashboard Sheet with summary metrics
 */
function createDashboardSheet(ss) {
  let sheet = ss.getSheetByName('Dashboard');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Dashboard');
  
  // Header
  sheet.getRange('A1:H1').merge().setValue('TDS COMPLIANCE DASHBOARD')
    .setBackground('#1a237e').setFontColor('white').setFontWeight('bold').setFontSize(18);
  
  // Entity info
  sheet.getRange('A2').setValue('Entity:').setFontWeight('bold');
  sheet.getRange('B2').setFormula('=Assumptions!B4');
  sheet.getRange('E2').setValue('FY:').setFontWeight('bold');
  sheet.getRange('F2').setFormula('=Assumptions!B10');
  
  // Key Metrics Section
  sheet.getRange('A4').setValue('KEY METRICS').setFontWeight('bold').setFontSize(14).setBackground('#e3f2fd');
  sheet.getRange('A4:H4').merge();
  
  // Metric cards
  const metrics = [
    ['Total Vendors', '=COUNTA(Vendor_Master!A4:A1000)-COUNTBLANK(Vendor_Master!A4:A1000)', 'B6', '#e3f2fd'],
    ['Total Transactions', '=COUNTA(TDS_Register!A4:A1000)-COUNTBLANK(TDS_Register!B4:B1000)', 'D6', '#e3f2fd'],
    ['Total TDS Deducted', '=SUM(TDS_Register!L:L)', 'F6', '#d4edda'],
    ['TDS Payable', '=SUM(TDS_Payable_Ledger!G4:G15)', 'B9', '#fff3cd'],
    ['Interest on Late Payment', '=SUM(Interest_Calculator!G10:G21)', 'D9', '#f8d7da'],
    ['Active Lower Deduction Certs', '=COUNTIF(Lower_Deduction_Cert!J:J,"Active")', 'F9', '#e3f2fd']
  ];
  
  for (let i = 0; i < metrics.length; i++) {
    const metric = metrics[i];
    const range = sheet.getRange(metric[2]);
    const labelRange = sheet.getRange(metric[2]).offset(-1, 0, 1, 2);
    
    labelRange.merge().setValue(metric[0]).setFontWeight('bold').setHorizontalAlignment('center');
    range.setFormula(metric[1]).setFontSize(16).setFontWeight('bold')
      .setHorizontalAlignment('center').setBackground(metric[3]);
    range.offset(0, 0, 1, 2).merge();
    
    // Add border
    range.offset(-1, 0, 2, 2).setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
  
  // Format currency metrics
  sheet.getRange('F6').setNumberFormat('"₹"#,##0');
  sheet.getRange('B9').setNumberFormat('"₹"#,##0');
  sheet.getRange('D9').setNumberFormat('"₹"#,##0');
  
  // Compliance Status Section
  sheet.getRange('A12').setValue('COMPLIANCE STATUS').setFontWeight('bold').setFontSize(14).setBackground('#e3f2fd');
  sheet.getRange('A12:H12').merge();
  
  const statusHeaders = ['Month', 'TDS Deducted', 'Status', 'Due Date', 'Payment Date', 'Delay (Days)'];
  sheet.getRange(13, 1, 1, statusHeaders.length).setValues([statusHeaders])
    .setBackground('#1a237e').setFontColor('white').setFontWeight('bold');
  
  // Pull status from TDS_Payable_Ledger
  for (let i = 0; i < 12; i++) {
    const row = 14 + i;
    const ledgerRow = 4 + i;
    
    sheet.getRange(`A${row}`).setFormula(`=TDS_Payable_Ledger!A${ledgerRow}`);
    sheet.getRange(`B${row}`).setFormula(`=TDS_Payable_Ledger!B${ledgerRow}`).setNumberFormat('#,##0');
    sheet.getRange(`C${row}`).setFormula(`=TDS_Payable_Ledger!H${ledgerRow}`);
    sheet.getRange(`D${row}`).setFormula(`=TDS_Payable_Ledger!C${ledgerRow}`).setNumberFormat('dd-mmm-yyyy');
    sheet.getRange(`E${row}`).setFormula(`=TDS_Payable_Ledger!E${ledgerRow}`).setNumberFormat('dd-mmm-yyyy');
    sheet.getRange(`F${row}`).setFormula(`=IF(E${row}="",0,MAX(0,E${row}-D${row}))`);
  }
  
  // Conditional formatting for status
  const paidRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Paid')
    .setBackground('#d4edda')
    .setRanges([sheet.getRange('C14:C25')])
    .build();
  
  const pendingRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Pending')
    .setBackground('#fff3cd')
    .setRanges([sheet.getRange('C14:C25')])
    .build();
  
  const lateRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Late Payment')
    .setBackground('#f8d7da')
    .setRanges([sheet.getRange('C14:C25')])
    .build();
  
  sheet.setConditionalFormatRules([paidRule, pendingRule, lateRule]);
  
  // Section-wise Summary
  sheet.getRange('A27').setValue('SECTION-WISE SUMMARY').setFontWeight('bold').setFontSize(14).setBackground('#e3f2fd');
  sheet.getRange('A27:H27').merge();
  
  const sectionHeaders = ['TDS Section', 'No. of Transactions', 'Total Amount (₹)'];
  sheet.getRange(28, 1, 1, sectionHeaders.length).setValues([sectionHeaders])
    .setBackground('#1a237e').setFontColor('white').setFontWeight('bold');
  
  // Common sections
  const sections = ['194C', '194J', '194I', '194A', '194H', 'Others'];
  for (let i = 0; i < sections.length; i++) {
    const row = 29 + i;
    sheet.getRange(`A${row}`).setValue(sections[i]);
    
    if (sections[i] === 'Others') {
      sheet.getRange(`B${row}`).setFormula(`=COUNTA(TDS_Register!G:G)-COUNTBLANK(TDS_Register!G:G)-${sections.slice(0, -1).map(s => `COUNTIF(TDS_Register!G:G,"${s}")`).join('-')}`);
      sheet.getRange(`C${row}`).setFormula(`=SUM(TDS_Register!L:L)-${sections.slice(0, -1).map(s => `SUMIF(TDS_Register!G:G,"${s}",TDS_Register!L:L)`).join('-')}`).setNumberFormat('#,##0.00');
    } else {
      sheet.getRange(`B${row}`).setFormula(`=COUNTIF(TDS_Register!G:G,"${sections[i]}")`);
      sheet.getRange(`C${row}`).setFormula(`=SUMIF(TDS_Register!G:G,"${sections[i]}",TDS_Register!L:L)`).setNumberFormat('#,##0.00');
    }
  }
  
  // Column widths
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 100);
  
  sheet.setFrozenRows(1);
}

/**
 * Creates Audit Notes Sheet
 */
function createAuditNotesSheet(ss) {
  let sheet = ss.getSheetByName('Audit_Notes');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Audit_Notes');
  
  // Header
  sheet.getRange('A1').setValue('TDS COMPLIANCE - AUDIT NOTES & REFERENCES')
    .setFontSize(16).setFontWeight('bold').setFontColor('#1a237e');
  
  const content = [
    ['', ''],
    ['WORKBOOK PURPOSE', ''],
    ['This workbook helps maintain TDS compliance records as per Income Tax Act, 1961.', ''],
    ['', ''],
    ['KEY COMPLIANCE REQUIREMENTS', ''],
    ['1. TDS Deduction', 'Deduct TDS at prescribed rates based on nature of payment and deductee type'],
    ['2. TDS Payment', 'Deposit TDS by 7th of next month (or as prescribed)'],
    ['3. TDS Return Filing', 'File quarterly returns (24Q for salary, 26Q for non-salary) by due dates'],
    ['4. TDS Certificate', 'Issue Form 16/16A to deductees within prescribed time'],
    ['5. Form 26AS', 'Reconcile with Form 26AS quarterly'],
    ['', ''],
    ['IMPORTANT SECTIONS', ''],
    ['Section 192', 'TDS on Salary'],
    ['Section 194A', 'TDS on Interest (other than securities)'],
    ['Section 194C', 'TDS on Payments to Contractors'],
    ['Section 194H', 'TDS on Commission/Brokerage'],
    ['Section 194I', 'TDS on Rent'],
    ['Section 194J', 'TDS on Professional/Technical Fees'],
    ['Section 201', 'Consequences of failure to deduct/pay TDS'],
    ['Section 197', 'Certificate for lower/nil deduction'],
    ['', ''],
    ['INTEREST & PENALTIES', ''],
    ['Section 201(1)', 'Interest @ 1% per month for late payment of TDS'],
    ['Section 201(1A)', 'Interest @ 1.5% per month for late filing of return'],
    ['Section 271H', 'Penalty for failure to file TDS return: ₹200 per day (min ₹10,000, max TDS amount)'],
    ['Section 271C', 'Penalty for failure to deduct TDS: Amount equal to TDS not deducted'],
    ['', ''],
    ['QUARTERLY RETURN DUE DATES', ''],
    ['Q1 (Apr-Jun)', 'Due by 31st July'],
    ['Q2 (Jul-Sep)', 'Due by 31st October'],
    ['Q3 (Oct-Dec)', 'Due by 31st January'],
    ['Q4 (Jan-Mar)', 'Due by 31st May'],
    ['', ''],
    ['USEFUL LINKS', ''],
    ['TRACES Portal', 'https://www.tdscpc.gov.in/'],
    ['Income Tax Department', 'https://www.incometax.gov.in/'],
    ['TDS Rates', 'https://www.incometax.gov.in/iec/foportal/help/individual/return-applicable-1/tds-rates'],
    ['', ''],
    ['BEST PRACTICES', ''],
    ['1. Maintain accurate vendor master with valid PANs', ''],
    ['2. Verify TDS rates before deduction', ''],
    ['3. Track lower deduction certificates and apply correctly', ''],
    ['4. Deposit TDS before due date to avoid interest', ''],
    ['5. Reconcile with Form 26AS every quarter', ''],
    ['6. Keep challan copies and acknowledgments safely', ''],
    ['7. Issue TDS certificates timely', ''],
    ['8. Maintain proper documentation for audit', ''],
    ['', ''],
    ['AUDIT CHECKLIST', ''],
    ['☐ All vendors have valid PAN', ''],
    ['☐ TDS rates applied correctly', ''],
    ['☐ Lower deduction certificates tracked', ''],
    ['☐ TDS deposited within due dates', ''],
    ['☐ Quarterly returns filed on time', ''],
    ['☐ Form 26AS reconciled', ''],
    ['☐ TDS certificates issued', ''],
    ['☐ Interest calculated for late payments', ''],
    ['☐ Documentation complete', ''],
    ['', ''],
    ['VERSION HISTORY', ''],
    ['Version 1.0 - November 2024', 'Initial release'],
    ['', ''],
    ['DISCLAIMER', ''],
    ['This workbook is a tool for TDS compliance tracking. Users should verify rates and rules', ''],
    ['as per latest Income Tax Act amendments. Consult a tax professional for specific situations.', '']
  ];
  
  sheet.getRange(2, 1, content.length, 2).setValues(content);
  
  // Formatting
  sheet.setColumnWidth(1, 300);
  sheet.setColumnWidth(2, 500);
  
  // Bold section headers
  const headerRows = [3, 6, 14, 23, 29, 35, 40, 49, 60, 63];
  headerRows.forEach(row => {
    sheet.getRange(row, 1).setFontWeight('bold').setFontSize(12).setBackground('#e3f2fd');
  });
}

/**
 * Setup named ranges for easy reference
 */
function setupNamedRanges(ss) {
  // Vendor Master
  ss.setNamedRange('VendorCodes', ss.getSheetByName('Vendor_Master').getRange('A4:A1000'));
  ss.setNamedRange('VendorNames', ss.getSheetByName('Vendor_Master').getRange('B4:B1000'));
  ss.setNamedRange('VendorPANs', ss.getSheetByName('Vendor_Master').getRange('C4:C1000'));
  
  // TDS Register
  ss.setNamedRange('TDSTransactions', ss.getSheetByName('TDS_Register').getRange('A4:R1000'));
  
  // Section Rates
  ss.setNamedRange('SectionRates', ss.getSheetByName('Section_Rates').getRange('A4:J100'));
}

/**
 * Reorder sheets in logical sequence
 */
function reorderSheets(ss) {
  const sheetOrder = [
    'Dashboard',
    'Cover',
    'Assumptions',
    'Vendor_Master',
    'TDS_Register',
    'Section_Rates',
    'Lower_Deduction_Cert',
    'TDS_Payable_Ledger',
    '26AS_Reconciliation',
    'Quarterly_Return',
    'Interest_Calculator',
    'Audit_Notes'
  ];
  
  sheetOrder.forEach((sheetName, index) => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(index + 1);
    }
  });
}

/**
 * POPULATE SAMPLE DATA - For demonstration and testing
 * Run this after creating the workbook to add realistic sample data
 */
function populateSampleData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  SpreadsheetApp.getActiveSpreadsheet().toast('Populating sample data...', 'Please Wait', -1);
  
  // Populate Assumptions
  populateAssumptionsSample(ss);
  
  // Populate Vendor Master
  populateVendorMasterSample(ss);
  
  // Populate TDS Register
  populateTDSRegisterSample(ss);
  
  // Populate Lower Deduction Certificates
  populateLowerDeductionCertSample(ss);
  
  // Populate TDS Payable Ledger (payment details)
  populateTDSPayableSample(ss);
  
  // Populate 26AS data
  populate26ASSample(ss);
  
  SpreadsheetApp.getActiveSpreadsheet().toast('Sample data populated successfully!', 'Complete', 5);
  ss.getSheetByName('Dashboard').activate();
}

/**
 * Populate Assumptions with sample entity data
 */
function populateAssumptionsSample(ss) {
  const sheet = ss.getSheetByName('Assumptions');
  
  sheet.getRange('B4').setValue('ABC Manufacturing Pvt Ltd');
  sheet.getRange('B5').setValue('AAAPL1234C');
  sheet.getRange('D5').setValue('MUMM12345D');
  sheet.getRange('B6').setValue('Plot No. 123, Industrial Area, Phase-II');
  sheet.getRange('B7').setValue('Mumbai');
  sheet.getRange('D7').setValue('Maharashtra');
  sheet.getRange('F7').setValue('400001');
  sheet.getRange('B8').setValue('Rajesh Kumar');
  sheet.getRange('D8').setValue('rajesh.kumar@abcmfg.com');
  sheet.getRange('F8').setValue('9876543210');
  
  sheet.getRange('B15').setValue('HDFC Bank Ltd');
  sheet.getRange('B16').setValue('50100123456789');
  sheet.getRange('D16').setValue('HDFC0001234');
}

/**
 * Populate Vendor Master with realistic vendors
 */
function populateVendorMasterSample(ss) {
  const sheet = ss.getSheetByName('Vendor_Master');
  
  const vendors = [
    ['V001', 'XYZ Contractors Pvt Ltd', 'AABCX1234F', '', 'Company', '45 MG Road, Andheri', 'Mumbai', 'Maharashtra', 'contact@xyzcontractors.com', '9876543211', 'No', 'Civil works contractor'],
    ['V002', 'Ramesh Consultants', 'AEMPR5678K', '', 'Individual', '12 Park Street', 'Pune', 'Maharashtra', 'ramesh@consultants.com', '9876543212', 'Yes', 'Technical consultant - Lower cert valid'],
    ['V003', 'Global Tech Solutions Pvt Ltd', 'AADCG9012L', '', 'Company', '789 IT Park, Whitefield', 'Bangalore', 'Karnataka', 'info@globaltech.com', '9876543213', 'No', 'Software development'],
    ['V004', 'Priya Advertising Agency', 'AHJPP3456M', '', 'Firm', '23 Commercial Complex', 'Delhi', 'Delhi', 'priya@advertising.com', '9876543214', 'No', 'Marketing services'],
    ['V005', 'Sharma & Associates', 'AAKFS7890N', '', 'Firm', '56 Legal Chambers', 'Mumbai', 'Maharashtra', 'sharma@legal.com', '9876543215', 'No', 'Legal consultancy'],
    ['V006', 'Metro Property Rentals', 'AABCM2345P', '', 'Company', '101 Real Estate Plaza', 'Mumbai', 'Maharashtra', 'metro@property.com', '9876543216', 'No', 'Office space rental'],
    ['V007', 'Suresh Transport Services', 'ACDPS6789Q', '', 'Individual', '78 Transport Nagar', 'Pune', 'Maharashtra', 'suresh@transport.com', '9876543217', 'No', 'Logistics services'],
    ['V008', 'ICICI Bank Ltd', 'AAACI1234R', '', 'Company', 'ICICI Towers, BKC', 'Mumbai', 'Maharashtra', 'corporate@icici.com', '1800123456', 'No', 'Interest on FD'],
    ['V009', 'Anita Design Studio', 'AHJPA4567S', '', 'Individual', '34 Creative Hub', 'Bangalore', 'Karnataka', 'anita@design.com', '9876543218', 'No', 'Graphic design services'],
    ['V010', 'BuildRight Engineers', 'AABCB8901T', '', 'Company', '67 Engineering Complex', 'Chennai', 'Tamil Nadu', 'info@buildright.com', '9876543219', 'No', 'Construction services']
  ];
  
  sheet.getRange(4, 1, vendors.length, 12).setValues(vendors);
}

/**
 * Populate TDS Register with realistic transactions across FY
 */
function populateTDSRegisterSample(ss) {
  const sheet = ss.getSheetByName('TDS_Register');
  
  const transactions = [
    // Q1 Transactions (Apr-Jun 2024)
    ['', new Date(2024, 3, 15), 'V001', '', '', '', '194C', 'Civil construction work', 250000, '', '', '', '', new Date(2024, 3, 20), '', '', '', 'Q1 - Foundation work'],
    ['', new Date(2024, 3, 25), 'V008', '', '', '', '194A', 'Interest on Fixed Deposit', 45000, '', '', '', '', new Date(2024, 3, 30), '', '', '', 'Q1 FD interest'],
    ['', new Date(2024, 4, 10), 'V003', '', '', '', '194J', 'Software development fees', 180000, '', '', '', '', new Date(2024, 4, 15), '', '', '', 'Custom ERP module'],
    ['', new Date(2024, 4, 20), 'V002', '', '', '', '194J', 'Technical consultancy', 85000, '', '', '', '', new Date(2024, 4, 25), '', '', '', 'Process optimization - Lower cert applied'],
    ['', new Date(2024, 5, 5), 'V006', '', '', '', '194I', 'Office rent - May 2024', 150000, '', '', '', '', new Date(2024, 5, 10), '', '', '', 'Monthly office rent'],
    ['', new Date(2024, 5, 15), 'V004', '', '', '', '194H', 'Marketing commission', 75000, '', '', '', '', new Date(2024, 5, 20), '', '', '', 'Q1 marketing campaign'],
    ['', new Date(2024, 5, 25), 'V007', '', '', '', '194C', 'Transportation charges', 120000, '', '', '', '', new Date(2024, 5, 30), '', '', '', 'Logistics - June'],
    
    // Q2 Transactions (Jul-Sep 2024)
    ['', new Date(2024, 6, 8), 'V001', '', '', '', '194C', 'Structural work', 320000, '', '', '', '', new Date(2024, 6, 12), '', '', '', 'Q2 - Building structure'],
    ['', new Date(2024, 6, 18), 'V005', '', '', '', '194J', 'Legal consultancy fees', 95000, '', '', '', '', new Date(2024, 6, 22), '', '', '', 'Contract review'],
    ['', new Date(2024, 7, 5), 'V006', '', '', '', '194I', 'Office rent - Aug 2024', 150000, '', '', '', '', new Date(2024, 7, 10), '', '', '', 'Monthly office rent'],
    ['', new Date(2024, 7, 15), 'V009', '', '', '', '194J', 'Graphic design services', 42000, '', '', '', '', new Date(2024, 7, 20), '', '', '', 'Brand identity design'],
    ['', new Date(2024, 7, 25), 'V003', '', '', '', '194J', 'IT support & maintenance', 125000, '', '', '', '', new Date(2024, 7, 28), '', '', '', 'Annual maintenance'],
    ['', new Date(2024, 8, 10), 'V010', '', '', '', '194C', 'Electrical installation', 280000, '', '', '', '', new Date(2024, 8, 15), '', '', '', 'Electrical work'],
    ['', new Date(2024, 8, 20), 'V004', '', '', '', '194H', 'Sales commission', 68000, '', '', '', '', new Date(2024, 8, 25), '', '', '', 'Q2 sales incentive'],
    
    // Q3 Transactions (Oct-Dec 2024)
    ['', new Date(2024, 9, 5), 'V006', '', '', '', '194I', 'Office rent - Oct 2024', 150000, '', '', '', '', new Date(2024, 9, 10), '', '', '', 'Monthly office rent'],
    ['', new Date(2024, 9, 15), 'V001', '', '', '', '194C', 'Interior finishing work', 195000, '', '', '', '', new Date(2024, 9, 20), '', '', '', 'Interior work'],
    ['', new Date(2024, 9, 28), 'V008', '', '', '', '194A', 'Interest on Fixed Deposit', 48000, '', '', '', '', new Date(2024, 10, 5), '', '', '', 'Q3 FD interest - LATE PAYMENT'],
    ['', new Date(2024, 10, 8), 'V002', '', '', '', '194J', 'Process audit services', 92000, '', '', '', '', new Date(2024, 10, 12), '', '', '', 'Annual process audit'],
    ['', new Date(2024, 10, 18), 'V007', '', '', '', '194C', 'Transportation charges', 135000, '', '', '', '', new Date(2024, 10, 22), '', '', '', 'Logistics - Nov'],
    ['', new Date(2024, 11, 5), 'V006', '', '', '', '194I', 'Office rent - Dec 2024', 150000, '', '', '', '', new Date(2024, 11, 10), '', '', '', 'Monthly office rent'],
    ['', new Date(2024, 11, 15), 'V003', '', '', '', '194J', 'Cloud services', 78000, '', '', '', '', new Date(2024, 11, 20), '', '', '', 'Cloud hosting charges'],
    ['', new Date(2024, 11, 28), 'V009', '', '', '', '194J', 'Website redesign', 115000, '', '', '', '', new Date(2025, 0, 5), '', '', '', 'Website project - LATE PAYMENT'],
    
    // Q4 Transactions (Jan-Mar 2025)
    ['', new Date(2025, 0, 10), 'V006', '', '', '', '194I', 'Office rent - Jan 2025', 150000, '', '', '', '', new Date(2025, 0, 15), '', '', '', 'Monthly office rent'],
    ['', new Date(2025, 0, 20), 'V010', '', '', '', '194C', 'HVAC installation', 245000, '', '', '', '', new Date(2025, 0, 25), '', '', '', 'Air conditioning work'],
    ['', new Date(2025, 1, 5), 'V006', '', '', '', '194I', 'Office rent - Feb 2025', 150000, '', '', '', '', new Date(2025, 1, 10), '', '', '', 'Monthly office rent'],
    ['', new Date(2025, 1, 15), 'V005', '', '', '', '194J', 'Legal retainer fees', 105000, '', '', '', '', new Date(2025, 1, 20), '', '', '', 'Quarterly retainer'],
    ['', new Date(2025, 1, 25), 'V004', '', '', '', '194H', 'Marketing commission', 82000, '', '', '', '', new Date(2025, 2, 10), '', '', '', 'Q4 marketing - LATE PAYMENT'],
    ['', new Date(2025, 2, 5), 'V006', '', '', '', '194I', 'Office rent - Mar 2025', 150000, '', '', '', '', new Date(2025, 2, 10), '', '', '', 'Monthly office rent'],
    ['', new Date(2025, 2, 15), 'V001', '', '', '', '194C', 'Final finishing work', 175000, '', '', '', '', new Date(2025, 2, 20), '', '', '', 'Project completion'],
    ['', new Date(2025, 2, 25), 'V003', '', '', '', '194J', 'Year-end IT support', 98000, '', '', '', '', new Date(2025, 2, 28), '', '', '', 'FY closing support']
  ];
  
  sheet.getRange(4, 1, transactions.length, 18).setValues(transactions);
}

/**
 * Populate Lower Deduction Certificate sample
 */
function populateLowerDeductionCertSample(ss) {
  const sheet = ss.getSheetByName('Lower_Deduction_Cert');
  
  const certificates = [
    ['CERT/2024/12345', 'V002', '', '', new Date(2024, 3, 1), new Date(2025, 2, 31), '194J', 2, 1000000, ''],
    ['CERT/2024/67890', 'V003', '', '', new Date(2024, 0, 1), new Date(2024, 11, 31), '194J', 5, 2000000, '']
  ];
  
  sheet.getRange(4, 1, certificates.length, 10).setValues(certificates);
}

/**
 * Populate TDS Payable Ledger with payment details
 */
function populateTDSPayableSample(ss) {
  const sheet = ss.getSheetByName('TDS_Payable_Ledger');
  
  // Add challan numbers and payment dates for most months
  const payments = [
    ['BSR0001234', new Date(2024, 4, 7), ''], // Apr - paid on time
    ['BSR0001235', new Date(2024, 5, 7), ''], // May - paid on time
    ['BSR0001236', new Date(2024, 6, 7), ''], // Jun - paid on time
    ['BSR0001237', new Date(2024, 7, 7), ''], // Jul - paid on time
    ['BSR0001238', new Date(2024, 8, 7), ''], // Aug - paid on time
    ['BSR0001239', new Date(2024, 9, 7), ''], // Sep - paid on time
    ['BSR0001240', new Date(2024, 10, 12), ''], // Oct - LATE (due 7th, paid 12th)
    ['BSR0001241', new Date(2024, 11, 7), ''], // Nov - paid on time
    ['BSR0001242', new Date(2025, 0, 7), ''], // Dec - paid on time
    ['BSR0001243', new Date(2025, 1, 7), ''], // Jan - paid on time
    ['BSR0001244', new Date(2025, 2, 15), ''], // Feb - LATE (due 7th, paid 15th)
    ['', '', ''] // Mar - not yet paid (pending)
  ];
  
  for (let i = 0; i < payments.length; i++) {
    const row = 4 + i;
    if (payments[i][0]) {
      sheet.getRange(`D${row}`).setValue(payments[i][0]); // Challan No
      sheet.getRange(`E${row}`).setValue(payments[i][1]); // Payment Date
      
      // Amount Paid = TDS Deducted (copy from column B)
      const tdsAmount = sheet.getRange(`B${row}`).getValue();
      if (tdsAmount) {
        sheet.getRange(`F${row}`).setValue(tdsAmount);
      }
    }
  }
}

/**
 * Populate 26AS Reconciliation with sample data
 */
function populate26ASSample(ss) {
  const sheet = ss.getSheetByName('26AS_Reconciliation');
  
  // Populate Form 26AS data (slightly different from books to show variance)
  const form26ASData = [
    [8, 1825000, 36500], // Q1 - matches books
    [7, 1193000, 23860], // Q2 - variance in count and amount
    [7, 1213000, 24260], // Q3 - variance
    [5, 1055000, 21100]  // Q4 - variance
  ];
  
  for (let i = 0; i < form26ASData.length; i++) {
    const row = 14 + i;
    sheet.getRange(`C${row}`).setValue(form26ASData[i][0]); // No. of Entries
    sheet.getRange(`D${row}`).setValue(form26ASData[i][1]); // Gross Amount
    sheet.getRange(`E${row}`).setValue(form26ASData[i][2]); // TDS Amount
  }
}

// onOpen() is handled by common/utilities.gs

/**
 * Refresh Dashboard (recalculate all formulas)
 */
function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.flush();
  ss.getSheetByName('Dashboard').activate();
  SpreadsheetApp.getActiveSpreadsheet().toast('Dashboard refreshed!', 'Complete', 3);
}

/**
 * Export quarterly return data to CSV
 */
function exportForReturnFiling() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    'Export Return Data',
    'This will prepare data for quarterly return filing. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result == ui.Button.YES) {
    // In a real implementation, this would create a CSV or formatted export
    ui.alert('Export Feature', 'Export functionality would generate CSV/Excel file for return filing software.', ui.ButtonSet.OK);
  }
}
