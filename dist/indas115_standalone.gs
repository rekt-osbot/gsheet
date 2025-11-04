/**
 * @name indas115
 * @version 1.1.0
 * @built 2025-11-04T10:11:10.770Z
 * @description Standalone script. Do not edit directly - edit source files in src/ folder.
 * 
 * This file is auto-generated from:
 * - Common utilities (src/common/*.gs)
 * - Workbook-specific code (src/workbooks/indas115.gs)
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
    'INDAS109': { menuName: 'Ind AS 109 Tools', functionName: 'createIndAS109Workbook' },
    'INDAS115': { menuName: 'Ind AS 115 Tools', functionName: 'createIndAS115Workbook' },
    'INDAS116': { menuName: 'Ind AS 116 Tools', functionName: 'createIndAS116Workbook' },
    'FIXED_ASSETS': { menuName: 'Fixed Assets Tools', functionName: 'createFixedAssetsWorkbook' },
    'TDS_COMPLIANCE': { menuName: 'TDS Tools', functionName: 'createTDSComplianceWorkbook' },
    'ICFR_P2P': { menuName: 'ICFR Tools', functionName: 'createICFRP2PWorkbook' },
    'IA_MASTER': { menuName: 'Internal Audit Tools', functionName: 'createIAMasterWorkbook' }
  };
  
  const config = workbookConfig[workbookType] || { menuName: 'Audit Tools', functionName: 'createWorkbook' };
  
  ui.createMenu(config.menuName)
    .addItem('Create/Refresh Workbook', config.functionName)
    .addSeparator()
    .addItem('Populate Sample Data', 'populateSampleData')
    .addItem('Clear All Input Data', 'clearSampleData')
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
 * ERROR HANDLING UTILITIES
 * ════════════════════════════════════════════════════════════════════════════
 * Robust error handling for formulas and operations
 */

/**
 * Safely create a formula with error handling
 * Wraps the formula with IFERROR to provide fallback values
 * @param {string} formula - The base formula (without the = sign)
 * @param {string} fallbackValue - Value to show if formula errors (default: "N/A")
 * @returns {string} - Complete formula with error handling
 */
function safeFormula(formula, fallbackValue = '"N/A"') {
  // Ensure formula doesn't already start with =
  const cleanFormula = formula.startsWith('=') ? formula.substring(1) : formula;
  return `=IFERROR(${cleanFormula}, ${fallbackValue})`;
}

/**
 * Safely create a lookup formula
 * Provides better error messages for missing lookups
 * @param {string} lookupValue - Value to look up
 * @param {string} lookupRange - Range to search in
 * @param {number} returnColumn - Column to return (1-based)
 * @param {string|number} fallbackValue - Value to show if not found (default: "Not Found")
 * @returns {string} - VLOOKUP formula with error handling
 */
function safeLookupFormula(lookupValue, lookupRange, returnColumn, fallbackValue = '"Not Found"') {
  const vlookup = `VLOOKUP(${lookupValue}, ${lookupRange}, ${returnColumn}, FALSE)`;
  return `=IFERROR(${vlookup}, ${fallbackValue})`;
}

/**
 * Safely create a sum formula
 * @param {string} range - Range to sum
 * @param {number} fallbackValue - Value to show if error (default: 0)
 * @returns {string} - SUM formula with error handling
 */
function safeSumFormula(range, fallbackValue = 0) {
  return `=IFERROR(SUM(${range}), ${fallbackValue})`;
}

/**
 * Safely create an IF statement with error checking
 * @param {string} condition - The IF condition
 * @param {string} trueValue - Value if true
 * @param {string} falseValue - Value if false
 * @param {string} errorValue - Value if error
 * @returns {string} - IF formula with error handling
 */
function safeIfFormula(condition, trueValue, falseValue, errorValue = '"Error"') {
  return `=IFERROR(IF(${condition}, ${trueValue}, ${falseValue}), ${errorValue})`;
}

/**
 * Create a conditional formula that handles missing references
 * @param {string} reference - Cell or range reference
 * @param {string} condition - Condition to check
 * @param {string} trueValue - Value if true
 * @param {string} falseValue - Value if false
 * @returns {string} - Formula with reference safety
 */
function safeConditionalFormula(reference, condition, trueValue, falseValue) {
  // Check if reference exists before evaluating
  const formula = `IF(ISBLANK(${reference}), "${falseValue}", IF(${reference} ${condition}, "${trueValue}", "${falseValue}"))`;
  return `=IFERROR(${formula}, "Error")`;
}

/**
 * Wrap a cell operation with error handling
 * Executes operation and catches errors
 * @param {Function} operation - Function that performs the operation
 * @param {string} errorMessage - Message to log if operation fails
 * @returns {*} - Result of operation or null if error
 */
function safeOperation(operation, errorMessage = 'Operation failed') {
  try {
    return operation();
  } catch (error) {
    Logger.log(`${errorMessage}: ${error}`);
    return null;
  }
}

/**
 * Safely set range value with validation
 * @param {Range} range - The range to set
 * @param {*} value - The value to set
 * @param {*} defaultValue - Default value if setting fails
 * @returns {Range} - The range object for chaining
 */
function safeRangeSet(range, value, defaultValue = '') {
  try {
    if (value === null || value === undefined) {
      range.setValue(defaultValue);
    } else {
      range.setValue(value);
    }
    return range;
  } catch (error) {
    Logger.log(`Error setting range value: ${error}`);
    range.setValue(defaultValue);
    return range;
  }
}

/**
 * Safely apply formatting to a range
 * @param {Range} range - The range to format
 * @param {Object} formatOptions - Object with format options
 * @returns {Range} - The range object for chaining
 */
function safeRangeFormat(range, formatOptions = {}) {
  try {
    if (formatOptions.background) {
      range.setBackground(formatOptions.background);
    }
    if (formatOptions.fontColor) {
      range.setFontColor(formatOptions.fontColor);
    }
    if (formatOptions.fontSize) {
      range.setFontSize(formatOptions.fontSize);
    }
    if (formatOptions.fontWeight) {
      range.setFontWeight(formatOptions.fontWeight);
    }
    if (formatOptions.numberFormat) {
      range.setNumberFormat(formatOptions.numberFormat);
    }
    return range;
  } catch (error) {
    Logger.log(`Error formatting range: ${error}`);
    return range;
  }
}

/**
 * Validate that required sheets exist
 * @param {Spreadsheet} ss - The spreadsheet
 * @param {Array<string>} requiredSheets - Array of required sheet names
 * @returns {Object} - {valid: boolean, missingSheets: Array<string>}
 */
function validateRequiredSheets(ss, requiredSheets) {
  const missingSheets = [];

  requiredSheets.forEach(sheetName => {
    if (!ss.getSheetByName(sheetName)) {
      missingSheets.push(sheetName);
    }
  });

  return {
    valid: missingSheets.length === 0,
    missingSheets: missingSheets
  };
}

/**
 * Check if a named range exists
 * @param {Spreadsheet} ss - The spreadsheet
 * @param {string} namedRangeName - Name of the range
 * @returns {boolean} - True if named range exists
 */
function namedRangeExists(ss, namedRangeName) {
  try {
    const range = ss.getRangeByName(namedRangeName);
    return range !== null;
  } catch (error) {
    return false;
  }
}

/**
 * Safely get a named range
 * @param {Spreadsheet} ss - The spreadsheet
 * @param {string} namedRangeName - Name of the range
 * @param {*} defaultValue - Value to return if range doesn't exist
 * @returns {Range|*} - The named range or default value
 */
function safeGetNamedRange(ss, namedRangeName, defaultValue = null) {
  try {
    return ss.getRangeByName(namedRangeName);
  } catch (error) {
    Logger.log(`Named range "${namedRangeName}" not found: ${error}`);
    return defaultValue;
  }
}

/**
 * Validate data in a range
 * @param {Range} range - The range to validate
 * @param {Object} rules - Validation rules {type, minValue, maxValue, allowEmpty}
 * @returns {Object} - {valid: boolean, errors: Array<string>}
 */
function validateRangeData(range, rules) {
  const values = range.getValues();
  const errors = [];
  let errorCount = 0;

  values.forEach((row, rowIndex) => {
    row.forEach((cell, colIndex) => {
      const cellAddress = range.getSheet().getName() + '!' + range.getCell(rowIndex + 1, colIndex + 1).getA1Notation();

      // Check if empty when not allowed
      if (!rules.allowEmpty && (cell === '' || cell === null)) {
        errors.push(`${cellAddress}: Value cannot be empty`);
        errorCount++;
      }

      // Check type
      if (rules.type === 'number' && cell !== '' && isNaN(cell)) {
        errors.push(`${cellAddress}: Expected number but got "${cell}"`);
        errorCount++;
      }

      // Check range
      if (rules.minValue !== undefined && cell < rules.minValue) {
        errors.push(`${cellAddress}: Value ${cell} is below minimum ${rules.minValue}`);
        errorCount++;
      }
      if (rules.maxValue !== undefined && cell > rules.maxValue) {
        errors.push(`${cellAddress}: Value ${cell} exceeds maximum ${rules.maxValue}`);
        errorCount++;
      }
    });
  });

  return {
    valid: errors.length === 0,
    errors: errors.slice(0, 10) // Limit to first 10 errors
  };
}

/**
 * Create error report sheet
 * @param {Spreadsheet} ss - The spreadsheet
 * @param {Array<string>} errors - Array of error messages
 * @returns {Sheet} - The error report sheet
 */
function createErrorReportSheet(ss, errors) {
  const sheetName = 'Error Report';
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }

  // Header
  const headerRange = sheet.getRange('A1:B1');
  headerRange.setValues([['Timestamp', 'Error Message']]);
  headerRange.setBackground('#ffebee').setFontWeight('bold');

  // Errors
  const timestamp = new Date().toLocaleString('en-IN');
  const errorData = errors.map(error => [timestamp, error]);

  if (errorData.length > 0) {
    sheet.getRange(2, 1, errorData.length, 2).setValues(errorData);
  }

  sheet.autoResizeColumns(1, 2);
  return sheet;
}


/**
 * ════════════════════════════════════════════════════════════════════════════
 * SAMPLE DATA MANAGER
 * ════════════════════════════════════════════════════════════════════════════
 * Centralized sample data for demonstration and testing
 */

/**
 * Sample data registry - maps workbook types to their sample data
 */
const SAMPLE_DATA_REGISTRY = {
  'TDS_COMPLIANCE': getSampleDataTDSCompliance,
  'DEFERRED_TAX': getSampleDataDeferredTax,
  'INDAS109': getSampleDataIndAS109,
  'INDAS115': getSampleDataIndAS115,
  'INDAS116': getSampleDataIndAS116,
  'FIXED_ASSETS': getSampleDataFixedAssets,
  'ICFR_P2P': getSampleDataICFRP2P
};

/**
 * Populate sample data for current workbook type
 * @param {Spreadsheet} ss - The spreadsheet
 */
function populateSampleData(ss) {
  const scriptProps = PropertiesService.getScriptProperties();
  const workbookType = scriptProps.getProperty('WORKBOOK_TYPE');

  if (!workbookType) {
    SpreadsheetApp.getUi().alert(
      'Cannot populate sample data',
      'Workbook type not detected. Please regenerate the workbook first.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const dataFunction = SAMPLE_DATA_REGISTRY[workbookType];
  if (!dataFunction) {
    SpreadsheetApp.getUi().alert(
      'Sample data not available',
      `No sample data configuration found for ${workbookType}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  try {
    const sampleData = dataFunction();
    applySampleDataToWorkbook(ss, sampleData);

    SpreadsheetApp.getUi().alert(
      'Sample Data Populated',
      'Sample data has been successfully populated. You can now explore the workbook and modify the data as needed.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Error populating sample data',
      error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    Logger.log(`Error populating sample data: ${error}`);
  }
}

/**
 * Apply sample data to specific sheets
 * @param {Spreadsheet} ss - The spreadsheet
 * @param {Object} sampleData - Sample data object with sheet names as keys and data arrays as values
 */
function applySampleDataToWorkbook(ss, sampleData) {
  for (const sheetName in sampleData) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`Warning: Sheet "${sheetName}" not found when applying sample data`);
      continue;
    }

    const sheetData = sampleData[sheetName];
    if (Array.isArray(sheetData)) {
      // It's raw cell data - apply to sheet
      applyCellData(sheet, sheetData);
    } else if (typeof sheetData === 'object') {
      // It's structured data with named ranges or specific positions
      applyStructuredData(sheet, sheetData);
    }
  }
}

/**
 * Apply raw cell data to a sheet
 * @param {Sheet} sheet - The sheet
 * @param {Array<Array>} data - 2D array of data
 */
function applyCellData(sheet, data) {
  if (data.length === 0) return;

  const range = sheet.getRange(1, 1, data.length, data[0].length);
  range.setValues(data);
}

/**
 * Apply structured data to a sheet
 * @param {Sheet} sheet - The sheet
 * @param {Object} dataMap - Object with ranges/positions as keys
 */
function applyStructuredData(sheet, dataMap) {
  for (const rangeName in dataMap) {
    try {
      const range = sheet.getRange(rangeName);
      const value = dataMap[rangeName];

      if (Array.isArray(value)) {
        range.setValues(value);
      } else {
        range.setValue(value);
      }
    } catch (error) {
      Logger.log(`Could not set range ${rangeName}: ${error}`);
    }
  }
}

/**
 * Get sample data for TDS Compliance workbook
 */
function getSampleDataTDSCompliance() {
  return {
    'Assumptions': [
      ['Entity Name:', 'ABC Limited', '', '', ''],
      ['PAN:', 'AABCU1234K', '', '', ''],
      ['TAN:', '0110001234', '', '', ''],
      ['Financial Year:', 'FY 2024-25', '', '', ''],
      ['', '', '', '', ''],
      ['Vendor Master Sample', '', '', '', ''],
      ['Vendor Name', 'PAN', 'Vendor Type', 'Contact', ''],
      ['XYZ Consultants', 'ABCDE1234K', 'Professional', 'xyz@example.com', ''],
      ['Service Providers Inc', 'FGHIJ5678K', 'Service Provider', 'sp@example.com', '']
    ],
    'Vendor_Master': [
      ['Vendor ID', 'Vendor Name', 'PAN', 'Vendor Type', 'Contact', 'State'],
      ['V001', 'XYZ Consultants', 'ABCDE1234K', 'Professional', 'xyz@example.com', 'Delhi'],
      ['V002', 'Service Providers Inc', 'FGHIJ5678K', 'Service Provider', 'sp@example.com', 'Mumbai'],
      ['V003', 'ABC Services', 'KLMNO9101K', 'Transporter', 'abc@example.com', 'Bangalore'],
      ['V004', 'Tech Solutions', 'PQRST1121K', 'IT Services', 'tech@example.com', 'Pune']
    ],
    'TDS_Register': [
      ['Transaction Date', 'Vendor Name', 'Invoice No', 'Amount', 'TDS Section', 'TDS Rate (%)', 'TDS Amount', 'Deduction Date'],
      ['01-Nov-2024', 'XYZ Consultants', 'INV-001', 50000, '194J', 10, 5000, '10-Nov-2024'],
      ['05-Nov-2024', 'Service Providers Inc', 'INV-002', 100000, '194C', 2, 2000, '10-Nov-2024'],
      ['10-Nov-2024', 'ABC Services', 'INV-003', 75000, '194C', 2, 1500, '15-Nov-2024']
    ],
    'Section_Rates': [
      ['TDS Section', 'Nature of Payment', 'Rate (%)', 'Applicability'],
      ['194J', 'Professional Fees', 10, 'Professionals & Consultants'],
      ['194C', 'Contractor Payments', 2, 'Contractors & Transporters'],
      ['194H', 'Brokerage', 5, 'Stock Brokers & Agents']
    ]
  };
}

/**
 * Get sample data for Deferred Tax workbook
 */
function getSampleDataDeferredTax() {
  return {
    'Assumptions': [
      ['Company Name:', 'ABC Limited', '', ''],
      ['Reporting Date:', '31-Mar-2025', '', ''],
      ['Tax Rate (%):', 25.168, '', ''],
      ['', '', '', ''],
      ['Key Adjustments:', '', '', ''],
      ['Depreciation Difference (L)', 500000, '', ''],
      ['Provision for Doubtful Debts (L)', 200000, '', '']
    ],
    'Timing_Differences': [
      ['Adjustment Item', 'Book Value', 'Tax Value', 'Difference', 'Deferred Tax (Asset/Liability)'],
      ['Fixed Assets - Depreciation', 5000000, 4500000, 500000, -125840],
      ['Provision for Doubtful Debts', 1000000, 800000, 200000, -50336],
      ['Employee Benefits', 300000, 250000, 50000, -12584]
    ]
  };
}

/**
 * Get sample data for Ind AS 109 workbook
 */
function getSampleDataIndAS109() {
  return {
    'Assumptions': [
      ['Portfolio Name:', 'Trade Receivables', '', ''],
      ['Reporting Date:', '31-Mar-2025', '', ''],
      ['Probability of Default (%):', 2, '', ''],
      ['Loss Given Default (%):', 25, '', '']
    ],
    'Credit_Exposure': [
      ['Customer Name', 'Invoice Amount', 'Days Outstanding', 'Credit Rating', 'PD (%)'],
      ['Customer A', 100000, 30, 'AAA', 0.5],
      ['Customer B', 250000, 45, 'AA', 1.0],
      ['Customer C', 150000, 60, 'BBB', 5.0]
    ]
  };
}

/**
 * Get sample data for Ind AS 115 workbook
 */
function getSampleDataIndAS115() {
  return {
    'Assumptions': [
      ['Company Name:', 'ABC Limited', '', ''],
      ['Reporting Period:', 'Q1 FY 2024-25', '', ''],
      ['', '', '', ''],
      ['Revenue Recognition Method:', 'Over Time', '', '']
    ],
    'Contracts': [
      ['Contract ID', 'Customer', 'Total Contract Value', 'Performance Obligation', '% Complete', 'Revenue Recognized'],
      ['C001', 'Customer X', 1000000, 'Service Delivery', 50, 500000],
      ['C002', 'Customer Y', 2500000, 'Goods & Services', 75, 1875000],
      ['C003', 'Customer Z', 750000, 'Support Services', 25, 187500]
    ]
  };
}

/**
 * Get sample data for Ind AS 116 workbook
 */
function getSampleDataIndAS116() {
  return {
    'Assumptions': [
      ['Reporting Date:', '31-Mar-2025', '', ''],
      ['Incremental Borrowing Rate (%):', 8, '', ''],
      ['', '', '', ''],
      ['Lease Portfolio:', '', '', '']
    ],
    'Lease_Register': [
      ['Lease ID', 'Asset Description', 'Lease Term (Months)', 'Monthly Lease Payment', 'Start Date', 'End Date'],
      ['L001', 'Office Building', 60, 100000, '01-Apr-2023', '31-Mar-2028'],
      ['L002', 'Plant & Machinery', 36, 50000, '01-Jun-2023', '31-May-2026'],
      ['L003', 'Equipment', 24, 25000, '01-Sep-2023', '31-Aug-2025']
    ]
  };
}

/**
 * Get sample data for Fixed Assets workbook
 */
function getSampleDataFixedAssets() {
  return {
    'Assumptions': [
      ['Reporting Date:', '31-Mar-2025', '', ''],
      ['Currency:', 'INR', '', ''],
      ['', '', '', ''],
      ['Depreciation Method:', 'Straight Line', '', '']
    ],
    'Fixed_Asset_Register': [
      ['Asset ID', 'Description', 'Category', 'Cost', 'Accumulated Depreciation', 'NBV', 'Useful Life (Years)'],
      ['FA001', 'Office Building', 'Building', 5000000, 500000, 4500000, 30],
      ['FA002', 'Motor Vehicle', 'Vehicle', 1500000, 450000, 1050000, 5],
      ['FA003', 'Furniture', 'Furniture', 300000, 150000, 150000, 10],
      ['FA004', 'Computer Equipment', 'IT Assets', 500000, 400000, 100000, 3]
    ]
  };
}

/**
 * Get sample data for ICFR P2P workbook
 */
function getSampleDataICFRP2P() {
  return {
    'Assumptions': [
      ['Organization:', 'ABC Limited', '', ''],
      ['Process:', 'Procure to Pay', '', ''],
      ['Review Period:', 'Q1 FY 2024-25', '', '']
    ],
    'Control_Registry': [
      ['Control ID', 'Description', 'Frequency', 'Owner', 'Test Status'],
      ['P2P-001', 'Purchase Requisition Review', 'Monthly', 'Procurement Manager', 'Tested'],
      ['P2P-002', 'Invoice Matching (3-way)', 'Monthly', 'Accounts Payable', 'Tested'],
      ['P2P-003', 'Vendor Master Maintenance', 'Quarterly', 'Procurement', 'Not Tested'],
      ['P2P-004', 'Payment Authorization', 'Monthly', 'Finance Manager', 'Tested']
    ]
  };
}

/**
 * Clear all sample data from input sections
 * @param {Spreadsheet} ss - The spreadsheet
 */
function clearSampleData(ss) {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Clear All Data?',
    'This will clear all data from input cells and tables. Are you sure?',
    ui.ButtonSet.YES_NO
  );

  if (result === ui.Button.YES) {
    // Find and clear all cells with INPUT_BG color
    const sheets = ss.getSheets();
    let clearedCount = 0;

    sheets.forEach(sheet => {
      const range = sheet.getDataRange();
      const values = range.getValues();
      const backgrounds = range.getBackgrounds();

      // Clear cells with input background color
      for (let i = 0; i < backgrounds.length; i++) {
        for (let j = 0; j < backgrounds[i].length; j++) {
          if (backgrounds[i][j] === COLORS.INPUT_BG) {
            sheet.getRange(i + 1, j + 1).clearContent();
            clearedCount++;
          }
        }
      }
    });

    ui.alert(
      'Data Cleared',
      `Cleared ${clearedCount} input cells across all sheets.`,
      ui.ButtonSet.OK
    );
  }
}


/**
 * ============================================================================
 * IGAAP-Ind AS 115 AUDIT WORKPAPER BUILDER
 * ============================================================================
 * Purpose: Automatically generate comprehensive Ind AS 115 (Revenue from 
 *          Contracts with Customers) audit workings for period-end closure
 * 
 * Standards Covered: Ind AS 115 - Revenue from Contracts with Customers
 * Compliance: ICAI Auditing Standards, Professional Audit Documentation
 * 
 * Created: 2025
 * Version: 1.0
 * 
 * INSTRUCTIONS:
 * 1. Open Google Sheets
 * 2. Go to Extensions > Apps Script
 * 3. Delete any existing code
 * 4. Paste this entire script
 * 5. Save the project (name it "IndAS115 Audit Builder")
 * 6. Run the function: createIndAS115Workbook()
 * 7. Authorize the script when prompted
 * 8. Wait 30-60 seconds for complete workbook generation
 * 
 * ============================================================================
 */

/**
 * MAIN EXECUTION FUNCTION
 * Run this function to build the entire Ind AS 115 audit workpaper
 */
function createIndAS115Workbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Set workbook type for menu detection
  setWorkbookType('INDAS115');
  
  // Show progress to user
  SpreadsheetApp.getActiveSpreadsheet().toast('Building Ind AS 115 Workpaper...', 'Progress', -1);
  
  // Clear existing sheets (optional - comment out if you want to keep existing data)
  clearExistingSheets(ss);
  
  // Create all sheets in order
  createCoverSheet(ss);
  createAssumptionsSheet(ss);
  createContractRegisterSheet(ss);
  createRevenueRecognitionSheet(ss);
  createContractBalancesSheet(ss);
  createPerformanceObligationsSheet(ss);
  createVariableConsiderationSheet(ss);
  createPeriodEndAdjustmentsSheet(ss);
  createIGAAPReconciliationSheet(ss);
  createReferencesSheet(ss);
  createAuditNotesSheet(ss);
  
  // Setup named ranges for key inputs
  setupNamedRanges(ss);
  
  // Final formatting and protection
  finalFormatting(ss);
  
  // Set Cover sheet as active
  ss.getSheetByName('Cover').activate();
  
  SpreadsheetApp.getActiveSpreadsheet().toast('✓ Ind AS 115 Workpaper Complete!', 'Success', 5);
}

// ============================================================================
// WORKBOOK-SPECIFIC CONFIGURATION
// ============================================================================

// Column mappings for Ind AS 115 workbook
const COLS = {
  CONTRACT_REGISTER: {
    SR_NO: 1,
    CONTRACT_ID: 2,
    CUSTOMER: 3,
    CONTRACT_DATE: 4,
    DESCRIPTION: 5,
    CONTRACT_VALUE: 6,
    GST_AMOUNT: 7,
    TOTAL_VALUE: 8,
    START_DATE: 9,
    END_DATE: 10,
    DURATION: 11,
    PATTERN: 12,
    NUM_PO: 13,
    STATUS: 14,
    NOTES: 15
  },
  REVENUE_RECOGNITION: {
    SR_NO: 1,
    CONTRACT_ID: 2,
    CUSTOMER: 3,
    STEP1_IDENTIFIED: 4,
    STEP2_PO: 5,
    STEP3_PRICE: 6,
    STEP4_ALLOCATED: 7,
    STEP5_RECOGNIZED: 8,
    CALC_BASIS: 9,
    PROGRESS_PCT: 10
  }
};

/**
 * ============================================================================
 * SHEET 1: COVER / DASHBOARD
 * ============================================================================
 */
function createCoverSheet(ss) {
  const sheet = getOrCreateSheet(ss, 'Cover', 0);
  setColumnWidths(sheet, [150, 200, 200, 200, 200]);
  
  // Title Section - custom formatting for this specific header
  sheet.getRange('B2:E2').merge()
    .setValue('IND AS 115 - REVENUE FROM CONTRACTS WITH CUSTOMERS')
    .setFontSize(16)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  sheet.setRowHeight(2, 50);
  
  sheet.getRange('B3:E3').merge()
    .setValue('Audit Workpaper - Period-End Book Closure')
    .setFontSize(11)
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  // Client Information Section
  sheet.getRange('B5').setValue('CLIENT INFORMATION').setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange('B6').setValue('Client Name:');
  sheet.getRange('C6').setBackground('#cfe2f3').setNote('INPUT REQUIRED: Enter client/entity name');
  
  sheet.getRange('B7').setValue('Financial Year:');
  sheet.getRange('C7').setBackground('#cfe2f3').setNote('INPUT REQUIRED: Enter FY (e.g., FY 2024-25)');
  
  sheet.getRange('B8').setValue('Reporting Period:');
  sheet.getRange('C8').setBackground('#cfe2f3').setNote('INPUT REQUIRED: Quarter/Year-end (e.g., Q4 FY25)');
  
  sheet.getRange('B9').setValue('Currency:');
  sheet.getRange('C9').setValue('INR').setBackground('#cfe2f3').setNote('INPUT: Default INR, change if needed');
  
  sheet.getRange('B10').setValue('Preparer:');
  sheet.getRange('C10').setBackground('#cfe2f3');
  
  sheet.getRange('B11').setValue('Reviewer:');
  sheet.getRange('C11').setBackground('#cfe2f3');
  
  sheet.getRange('B12').setValue('Date Prepared:');
  sheet.getRange('C12').setFormula('=TODAY()').setNumberFormat('dd-mmm-yyyy');
  
  // Key Metrics Dashboard
  sheet.getRange('B14').setValue('KEY METRICS & CONTROL TOTALS').setFontWeight('bold').setBackground('#e8f0fe');
  
  const metrics = [
    ['Total Contract Value (Period)', '=SUM(\'Contract Register\'!F:F)', 'INR'],
    ['Revenue Recognized (Period)', '=SUM(\'Revenue Recognition\'!H:H)', 'INR'],
    ['Contract Assets', '=SUM(\'Contract Balances\'!E:E)', 'INR'],
    ['Contract Liabilities', '=SUM(\'Contract Balances\'!F:F)', 'INR'],
    ['Accounts Receivable', '=SUM(\'Contract Balances\'!G:G)', 'INR'],
    ['Number of Active Contracts', '=COUNTA(\'Contract Register\'!B:B)-1', 'Count']
  ];
  
  let row = 15;
  metrics.forEach(([label, formula, unit]) => {
    sheet.getRange(row, 2).setValue(label);
    sheet.getRange(row, 3).setFormula(formula).setNumberFormat(unit === 'INR' ? '#,##0.00' : '0');
    sheet.getRange(row, 4).setValue(unit).setFontSize(9).setFontColor('#666666');
    row++;
  });
  
  // Navigation Section
  sheet.getRange('B23').setValue('QUICK NAVIGATION').setFontWeight('bold').setBackground('#e8f0fe');
  
  const navigation = [
    ['Assumptions & Inputs', 'Assumptions'],
    ['Contract Register', 'Contract Register'],
    ['Revenue Recognition', 'Revenue Recognition'],
    ['Contract Balances', 'Contract Balances'],
    ['Performance Obligations', 'Performance Obligations'],
    ['Variable Consideration', 'Variable Consideration'],
    ['Period-End Adjustments', 'Period-End Adjustments'],
    ['IGAAP Reconciliation', 'IGAAP Reconciliation'],
    ['References', 'References'],
    ['Audit Notes', 'Audit Notes']
  ];
  
  row = 24;
  navigation.forEach(([label, sheetName]) => {
    sheet.getRange(row, 2).setValue('→ ' + label).setFontColor('#1a73e8').setFontWeight('bold');
    // Note: Apps Script doesn't support direct hyperlinks to sheets, but users can click sheet tabs
    row++;
  });
  
  // Instructions
  sheet.getRange('B35').setValue('INSTRUCTIONS').setFontWeight('bold').setBackground('#fff3cd');
  sheet.getRange('B36:E40').merge()
    .setValue(
      '1. Complete all BLUE-HIGHLIGHTED cells in the Assumptions sheet\n' +
      '2. Enter contract details in Contract Register\n' +
      '3. Review automated calculations in Revenue Recognition sheet\n' +
      '4. Verify Period-End Adjustments for journal entries\n' +
      '5. Check Audit Notes for control totals and variances\n\n' +
      'All formulas cascade automatically. Do not overwrite formula cells.'
    )
    .setWrap(true)
    .setVerticalAlignment('top')
    .setBackground('#fff3cd');
  
  sheet.setRowHeight(36, 120);
  
  // Freeze header
  sheet.setFrozenRows(4);
}

/**
 * ============================================================================
 * SHEET 2: ASSUMPTIONS & INPUT VARIABLES
 * ============================================================================
 */
function createAssumptionsSheet(ss) {
  let sheet = ss.getSheetByName('Assumptions');
  if (!sheet) {
    sheet = ss.insertSheet('Assumptions', 1);
  } else {
    sheet.clear();
  }
  
  // Set column widths
  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 300);
  
  // Header
  sheet.getRange('A1:E1').merge()
    .setValue('ASSUMPTIONS & INPUT VARIABLES - IND AS 115')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  sheet.setRowHeight(1, 40);
  
  // Section 1: Reporting Parameters
  sheet.getRange('A3:E3').merge()
    .setValue('REPORTING PARAMETERS')
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  
  const headers = ['#', 'Input Variable', 'Value', 'Unit/Format', 'Notes / Purpose'];
  sheet.getRange('A4:E4').setValues([headers])
    .setFontWeight('bold')
    .setBackground('#e8f0fe')
    .setHorizontalAlignment('center');
  
  const reportingParams = [
    [1, 'Reporting Period Start Date', '', 'dd-mmm-yyyy', 'Start date of the reporting period'],
    [2, 'Reporting Period End Date', '', 'dd-mmm-yyyy', 'End date of the reporting period'],
    [3, 'Prior Period Start Date', '', 'dd-mmm-yyyy', 'For comparative figures'],
    [4, 'Prior Period End Date', '', 'dd-mmm-yyyy', 'For comparative figures'],
    [5, 'Functional Currency', 'INR', 'Currency Code', 'Primary currency for reporting'],
    [6, 'Presentation Currency', 'INR', 'Currency Code', 'Currency for financial statements'],
    [7, 'Rounding (in Currency)', '1', 'Actual/000/00000', 'Rounding level: 1=Actual, 1000=Thousands']
  ];
  
  let row = 5;
  reportingParams.forEach((item) => {
    sheet.getRange(row, 1, 1, 5).setValues([item]);
    sheet.getRange(row, 3).setBackground('#cfe2f3'); // Input cell
    row++;
  });
  
  // Section 2: Revenue Recognition Policies
  row++;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('REVENUE RECOGNITION POLICIES')
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  
  row++;
  sheet.getRange(row, 1, 1, 5).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#e8f0fe')
    .setHorizontalAlignment('center');
  
  row++;
  const policyParams = [
    [8, 'Default Revenue Recognition Method', 'Point in Time', 'Dropdown', 'Select: Point in Time / Over Time'],
    [9, 'Over Time Method (if applicable)', 'Input Method', 'Dropdown', 'Output/Input/Time-based'],
    [10, 'Significant Financing Component Threshold', '12', 'Months', 'Period threshold for financing component'],
    [11, 'Discount Rate for Financing Component', '10%', 'Percentage', 'Rate for PV calculations'],
    [12, 'Variable Consideration Constraint', 'Most Likely', 'Dropdown', 'Method: Most Likely / Expected Value'],
    [13, 'Contract Modification Default Treatment', 'Separate Contract', 'Dropdown', 'Separate/Cumulative/Termination'],
    [14, 'Advance Received Treatment', 'Contract Liability', 'Text', 'Default classification']
  ];
  
  policyParams.forEach((item) => {
    sheet.getRange(row, 1, 1, 5).setValues([item]);
    sheet.getRange(row, 3).setBackground('#cfe2f3');
    row++;
  });
  
  // Section 3: Materiality Thresholds
  row++;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('MATERIALITY THRESHOLDS & CONTROL PARAMETERS')
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  
  row++;
  sheet.getRange(row, 1, 1, 5).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#e8f0fe')
    .setHorizontalAlignment('center');
  
  row++;
  const materialityParams = [
    [15, 'Performance Materiality (Amount)', '100000', 'INR', 'Threshold for detailed testing'],
    [16, 'Performance Materiality (%)', '5%', 'Percentage', 'As % of total revenue'],
    [17, 'Clearly Trivial Threshold', '10000', 'INR', 'Below this = clearly trivial'],
    [18, 'Contract Aggregation Threshold', '50000', 'INR', 'Min value for separate tracking'],
    [19, 'Variable Consideration Cap (%)', '50%', 'Percentage', 'Max % of transaction price'],
    [20, 'Unbilled Revenue Review Flag (Days)', '90', 'Days', 'Flag if unbilled > X days']
  ];
  
  materialityParams.forEach((item) => {
    sheet.getRange(row, 1, 1, 5).setValues([item]);
    sheet.getRange(row, 3).setBackground('#cfe2f3');
    row++;
  });
  
  // Section 4: Tax & Regulatory
  row++;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('TAX & REGULATORY PARAMETERS')
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  
  row++;
  sheet.getRange(row, 1, 1, 5).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#e8f0fe')
    .setHorizontalAlignment('center');
  
  row++;
  const taxParams = [
    [21, 'GST Rate (Standard)', '18%', 'Percentage', 'Standard GST rate applicable'],
    [22, 'GST Treatment in Revenue', 'Exclude', 'Dropdown', 'Include/Exclude from revenue'],
    [23, 'TDS Rate (if applicable)', '10%', 'Percentage', 'TDS deducted at source'],
    [24, 'Tax Invoice Basis', 'Accrual', 'Dropdown', 'Accrual/Cash - affects timing']
  ];
  
  taxParams.forEach((item) => {
    sheet.getRange(row, 1, 1, 5).setValues([item]);
    sheet.getRange(row, 3).setBackground('#cfe2f3');
    row++;
  });
  
  // Add data validation for dropdown fields
  addDataValidations(sheet);
  
  // Instructions
  row += 2;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('INSTRUCTIONS: Complete all blue-highlighted cells. These values drive calculations throughout the workbook.')
    .setBackground('#fff3cd')
    .setFontWeight('bold')
    .setWrap(true);
  
  // Freeze
  sheet.setFrozenRows(4);
}

/**
 * Add data validations for dropdown fields in Assumptions
 */
function addDataValidations(sheet) {
  const validations = [
    { range: 'C13', values: ['Point in Time', 'Over Time'] },
    { range: 'C14', values: ['Input Method', 'Output Method', 'Time-based'] },
    { range: 'C17', values: ['Most Likely', 'Expected Value'] },
    { range: 'C18', values: ['Separate Contract', 'Cumulative Catch-up', 'Termination'] },
    { range: 'C27', values: ['Include', 'Exclude'] },
    { range: 'C29', values: ['Accrual', 'Cash'] }
  ];
  
  validations.forEach(v => {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(v.values, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(v.range).setDataValidation(rule);
  });
}

/**
 * ============================================================================
 * SHEET 3: CONTRACT REGISTER
 * ============================================================================
 */
function createContractRegisterSheet(ss) {
  let sheet = ss.getSheetByName('Contract Register');
  if (!sheet) {
    sheet = ss.insertSheet('Contract Register', 2);
  } else {
    sheet.clear();
  }
  
  // Set column widths
  sheet.setColumnWidths(1, 15, 100);
  sheet.setColumnWidth(3, 200); // Customer Name
  sheet.setColumnWidth(5, 150); // Description
  
  // Header
  sheet.getRange('A1:O1').merge()
    .setValue('CONTRACT REGISTER - IND AS 115')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  sheet.setRowHeight(1, 40);
  
  // Instructions
  sheet.getRange('A2:O2').merge()
    .setValue('Enter all contracts for the reporting period. Each contract should have unique ID. Blue cells are inputs.')
    .setBackground('#fff3cd')
    .setWrap(true)
    .setFontStyle('italic');
  
  // Column Headers
  const headers = [
    'Sr No',
    'Contract ID',
    'Customer Name',
    'Contract Date',
    'Contract Description',
    'Total Contract Value (excl GST)',
    'GST Amount',
    'Contract Value (incl GST)',
    'Contract Start Date',
    'Contract End Date',
    'Contract Duration (Months)',
    'Revenue Recognition Pattern',
    '# of Performance Obligations',
    'Contract Status',
    'Notes'
  ];
  
  sheet.getRange('A3:O3').setValues([headers])
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setWrap(true);
  
  sheet.setRowHeight(3, 50);
  
  // Sample Row with Formulas
  sheet.getRange('A4').setValue(1);
  sheet.getRange('B4').setValue('CNT-001').setBackground('#cfe2f3');
  sheet.getRange('C4').setValue('Sample Customer Ltd').setBackground('#cfe2f3');
  sheet.getRange('D4').setValue('01-Apr-2024').setNumberFormat('dd-mmm-yyyy').setBackground('#cfe2f3');
  sheet.getRange('E4').setValue('Software Development Project').setBackground('#cfe2f3');
  sheet.getRange('F4').setValue(1000000).setNumberFormat('#,##0.00').setBackground('#cfe2f3');
  sheet.getRange('G4').setFormula('=F4*Assumptions!$C$26').setNumberFormat('#,##0.00')
    .setNote('Auto-calculated: Contract Value × GST Rate from Assumptions');
  sheet.getRange('H4').setFormula('=F4+G4').setNumberFormat('#,##0.00')
    .setNote('Auto-calculated: Total Contract Value');
  sheet.getRange('I4').setValue('01-Apr-2024').setNumberFormat('dd-mmm-yyyy').setBackground('#cfe2f3');
  sheet.getRange('J4').setValue('31-Mar-2025').setNumberFormat('dd-mmm-yyyy').setBackground('#cfe2f3');
  sheet.getRange('K4').setFormula('=DATEDIF(I4,J4,"M")').setNumberFormat('0')
    .setNote('Auto-calculated: Contract duration in months');
  sheet.getRange('L4').setValue('Over Time').setBackground('#cfe2f3');
  sheet.getRange('M4').setValue(3).setNumberFormat('0').setBackground('#cfe2f3');
  sheet.getRange('N4').setValue('Active').setBackground('#cfe2f3');
  sheet.getRange('O4').setValue('').setBackground('#cfe2f3');
  
  // Add data validation for Revenue Pattern
  const patternRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Point in Time', 'Over Time', 'Mixed'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('L4:L100').setDataValidation(patternRule);
  
  // Add data validation for Contract Status
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Completed', 'Terminated', 'On Hold'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('N4:N100').setDataValidation(statusRule);
  
  // Format additional rows
  for (let i = 5; i <= 50; i++) {
    sheet.getRange(i, 1).setFormula(`=IF(B${i}="","",ROW()-3)`);
    sheet.getRange(i, 2, 1, 1).setBackground('#cfe2f3'); // Contract ID
    sheet.getRange(i, 3, 1, 1).setBackground('#cfe2f3'); // Customer
    sheet.getRange(i, 4, 1, 1).setBackground('#cfe2f3').setNumberFormat('dd-mmm-yyyy'); // Contract Date
    sheet.getRange(i, 5, 1, 1).setBackground('#cfe2f3'); // Description
    sheet.getRange(i, 6, 1, 1).setBackground('#cfe2f3').setNumberFormat('#,##0.00'); // Contract Value
    sheet.getRange(i, 7, 1, 1).setFormula(`=IF(B${i}="","",F${i}*Assumptions!$C$26)`).setNumberFormat('#,##0.00'); // GST
    sheet.getRange(i, 8, 1, 1).setFormula(`=IF(B${i}="","",F${i}+G${i})`).setNumberFormat('#,##0.00'); // Total
    sheet.getRange(i, 9, 1, 1).setBackground('#cfe2f3').setNumberFormat('dd-mmm-yyyy'); // Start Date
    sheet.getRange(i, 10, 1, 1).setBackground('#cfe2f3').setNumberFormat('dd-mmm-yyyy'); // End Date
    sheet.getRange(i, 11, 1, 1).setFormula(`=IF(AND(I${i}<>"",J${i}<>""),DATEDIF(I${i},J${i},"M"),"")`).setNumberFormat('0'); // Duration
    sheet.getRange(i, 12, 1, 1).setBackground('#cfe2f3'); // Pattern
    sheet.getRange(i, 13, 1, 1).setBackground('#cfe2f3').setNumberFormat('0'); // # PO
    sheet.getRange(i, 14, 1, 1).setBackground('#cfe2f3'); // Status
    sheet.getRange(i, 15, 1, 1).setBackground('#cfe2f3'); // Notes
  }
  
  // Summary Section
  const summaryRow = 52;
  sheet.getRange(summaryRow, 1, 1, 5).merge()
    .setValue('SUMMARY TOTALS')
    .setFontWeight('bold')
    .setBackground('#e8f0fe')
    .setHorizontalAlignment('center');
  
  sheet.getRange(summaryRow + 1, 1).setValue('Total Contracts:');
  sheet.getRange(summaryRow + 1, 2).setFormula('=COUNTA(B4:B51)').setNumberFormat('0').setFontWeight('bold');
  
  sheet.getRange(summaryRow + 2, 1).setValue('Total Contract Value:');
  sheet.getRange(summaryRow + 2, 2).setFormula('=SUM(F4:F51)').setNumberFormat('#,##0.00').setFontWeight('bold');
  
  sheet.getRange(summaryRow + 3, 1).setValue('Active Contracts:');
  sheet.getRange(summaryRow + 3, 2).setFormula('=COUNTIF(N4:N51,"Active")').setNumberFormat('0').setFontWeight('bold');
  
  // Freeze
  sheet.setFrozenRows(3);
}

/**
 * ============================================================================
 * SHEET 4: REVENUE RECOGNITION SCHEDULE
 * ============================================================================
 */
function createRevenueRecognitionSheet(ss) {
  let sheet = ss.getSheetByName('Revenue Recognition');
  if (!sheet) {
    sheet = ss.insertSheet('Revenue Recognition', 3);
  } else {
    sheet.clear();
  }
  
  // Set column widths
  sheet.setColumnWidths(1, 18, 110);
  sheet.setColumnWidth(2, 150); // Contract ID
  sheet.setColumnWidth(9, 150); // Calculation basis
  
  // Header
  sheet.getRange('A1:R1').merge()
    .setValue('REVENUE RECOGNITION SCHEDULE - IND AS 115 (5-STEP MODEL)')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  sheet.setRowHeight(1, 40);
  
  // Instructions
  sheet.getRange('A2:R2').merge()
    .setValue('This schedule applies Ind AS 115 5-step model. Formulas link to Contract Register. Review revenue recognized vs billed.')
    .setBackground('#fff3cd')
    .setWrap(true)
    .setFontStyle('italic');
  
  // Column Headers with 5-Step Model Structure
  const headers = [
    'Sr',
    'Contract ID',
    'Customer',
    'STEP 1: Contract Identified?',
    'STEP 2: Performance Obligations',
    'STEP 3: Transaction Price (excl GST)',
    'STEP 4: Allocated Price',
    'STEP 5: Revenue Recognized (Current Period)',
    'Calculation Basis',
    'Progress %',
    'Opening Unearned Revenue',
    'Revenue Recognized YTD',
    'Closing Unearned Revenue',
    'Billed to Date',
    'Contract Asset (Unbilled)',
    'Contract Liability (Advance)',
    'Accounts Receivable',
    'Variance/Notes'
  ];
  
  sheet.getRange('A3:R3').setValues([headers])
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setWrap(true);
  
  sheet.setRowHeight(3, 60);
  
  // Sample Data Row with Formulas
  const startRow = 4;
  
  // Sr No
  sheet.getRange('A4').setValue(1);
  
  // Link to Contract Register
  sheet.getRange('B4').setFormula('=IF(\'Contract Register\'!B4="","",\'Contract Register\'!B4)')
    .setNote('Linked from Contract Register');
  
  sheet.getRange('C4').setFormula('=IF(B4="","",VLOOKUP(B4,\'Contract Register\'!B:C,2,FALSE))')
    .setNote('Auto-lookup customer name');
  
  // Step 1: Contract Identified
  sheet.getRange('D4').setValue('Yes').setBackground('#cfe2f3')
    .setNote('INPUT: Is there an enforceable contract with commercial substance?');
  
  // Step 2: Performance Obligations
  sheet.getRange('E4').setFormula('=IF(B4="","",VLOOKUP(B4,\'Contract Register\'!B:M,12,FALSE))')
    .setNumberFormat('0')
    .setNote('Number of distinct performance obligations');
  
  // Step 3: Transaction Price
  sheet.getRange('F4').setFormula('=IF(B4="","",VLOOKUP(B4,\'Contract Register\'!B:F,5,FALSE))')
    .setNumberFormat('#,##0.00')
    .setNote('Total transaction price excluding GST');
  
  // Step 4: Allocated Price (for single PO = same as transaction price)
  sheet.getRange('G4').setFormula('=IF(B4="","",F4)')
    .setNumberFormat('#,##0.00')
    .setNote('Allocated to this performance obligation');
  
  // Step 5: Revenue Recognized Current Period - INPUT
  sheet.getRange('H4').setValue(250000).setNumberFormat('#,##0.00').setBackground('#cfe2f3')
    .setNote('INPUT REQUIRED: Revenue recognized in current period based on satisfaction of PO');
  
  // Calculation Basis
  sheet.getRange('I4').setValue('Input Method - Costs Incurred').setBackground('#cfe2f3')
    .setNote('INPUT: Method used for revenue recognition');
  
  // Progress %
  sheet.getRange('J4').setFormula('=IF(G4=0,"",H4/G4)')
    .setNumberFormat('0.00%')
    .setNote('Auto: % of performance obligation satisfied');
  
  // Opening Unearned Revenue
  sheet.getRange('K4').setValue(0).setNumberFormat('#,##0.00').setBackground('#cfe2f3')
    .setNote('INPUT: Opening balance of unearned/deferred revenue');
  
  // Revenue Recognized YTD
  sheet.getRange('L4').setFormula('=IF(B4="","",K4+H4)')
    .setNumberFormat('#,##0.00')
    .setNote('Auto: Cumulative revenue recognized');
  
  // Closing Unearned Revenue
  sheet.getRange('M4').setFormula('=IF(B4="","",MAX(0,G4-L4))')
    .setNumberFormat('#,##0.00')
    .setNote('Auto: Remaining unearned revenue (contract liability)');
  
  // Billed to Date
  sheet.getRange('N4').setValue(300000).setNumberFormat('#,##0.00').setBackground('#cfe2f3')
    .setNote('INPUT: Amount billed/invoiced to customer to date');
  
  // Contract Asset (Unbilled Revenue)
  sheet.getRange('O4').setFormula('=IF(B4="","",MAX(0,L4-N4))')
    .setNumberFormat('#,##0.00')
    .setNote('Auto: Revenue recognized but not yet billed');
  
  // Contract Liability (Advance Received)
  sheet.getRange('P4').setFormula('=IF(B4="","",MAX(0,N4-L4))')
    .setNumberFormat('#,##0.00')
    .setNote('Auto: Amount billed but revenue not yet recognized');
  
  // Accounts Receivable
  sheet.getRange('Q4').setFormula('=IF(B4="","",N4-P4)')
    .setNumberFormat('#,##0.00')
    .setNote('Auto: Billed amount less advances (represents A/R)');
  
  // Variance/Notes
  sheet.getRange('R4').setBackground('#cfe2f3')
    .setNote('Document any significant variances or adjustments');
  
  // Apply to additional rows (50 contracts)
  for (let i = 5; i <= 50; i++) {
    sheet.getRange(i, 1).setFormula(`=IF(B${i}="","",ROW()-3)`);
    sheet.getRange(i, 2).setFormula(`=IF('Contract Register'!B${i}="",'Contract Register'!B${i})`);
    sheet.getRange(i, 3).setFormula(`=IF(B${i}="","",IFERROR(VLOOKUP(B${i},'Contract Register'!B:C,2,FALSE),""))`);
    sheet.getRange(i, 4).setBackground('#cfe2f3');
    sheet.getRange(i, 5).setFormula(`=IF(B${i}="","",IFERROR(VLOOKUP(B${i},'Contract Register'!B:M,12,FALSE),""))`).setNumberFormat('0');
    sheet.getRange(i, 6).setFormula(`=IF(B${i}="","",IFERROR(VLOOKUP(B${i},'Contract Register'!B:F,5,FALSE),""))`).setNumberFormat('#,##0.00');
    sheet.getRange(i, 7).setFormula(`=IF(B${i}="","",F${i})`).setNumberFormat('#,##0.00');
    sheet.getRange(i, 8).setBackground('#cfe2f3').setNumberFormat('#,##0.00');
    sheet.getRange(i, 9).setBackground('#cfe2f3');
    sheet.getRange(i, 10).setFormula(`=IF(AND(B${i}<>"",G${i}<>0),H${i}/G${i},"")`).setNumberFormat('0.00%');
    sheet.getRange(i, 11).setBackground('#cfe2f3').setNumberFormat('#,##0.00');
    sheet.getRange(i, 12).setFormula(`=IF(B${i}="","",K${i}+H${i})`).setNumberFormat('#,##0.00');
    sheet.getRange(i, 13).setFormula(`=IF(B${i}="","",MAX(0,G${i}-L${i}))`).setNumberFormat('#,##0.00');
    sheet.getRange(i, 14).setBackground('#cfe2f3').setNumberFormat('#,##0.00');
    sheet.getRange(i, 15).setFormula(`=IF(B${i}="","",MAX(0,L${i}-N${i}))`).setNumberFormat('#,##0.00');
    sheet.getRange(i, 16).setFormula(`=IF(B${i}="","",MAX(0,N${i}-L${i}))`).setNumberFormat('#,##0.00');
    sheet.getRange(i, 17).setFormula(`=IF(B${i}="","",N${i}-P${i})`).setNumberFormat('#,##0.00');
    sheet.getRange(i, 18).setBackground('#cfe2f3');
  }
  
  // Summary Totals
  const summaryRow = 52;
  sheet.getRange(summaryRow, 1, 1, 3).merge()
    .setValue('CONTROL TOTALS')
    .setFontWeight('bold')
    .setBackground('#e8f0fe')
    .setHorizontalAlignment('center');
  
  const summaryItems = [
    ['Total Transaction Price:', '=SUM(F4:F51)', 'F'],
    ['Total Revenue Recognized (Period):', '=SUM(H4:H51)', 'H'],
    ['Total Revenue Recognized (YTD):', '=SUM(L4:L51)', 'L'],
    ['Total Unearned Revenue:', '=SUM(M4:M51)', 'M'],
    ['Total Contract Assets:', '=SUM(O4:O51)', 'O'],
    ['Total Contract Liabilities:', '=SUM(P4:P51)', 'P'],
    ['Total A/R:', '=SUM(Q4:Q51)', 'Q']
  ];
  
  let sumRow = summaryRow + 1;
  summaryItems.forEach(item => {
    sheet.getRange(sumRow, 1, 1, 2).merge().setValue(item[0]);
    sheet.getRange(sumRow, 3).setFormula(item[1]).setNumberFormat('#,##0.00').setFontWeight('bold');
    sumRow++;
  });
  
  // Freeze
  sheet.setFrozenRows(3);
}

/**
 * ============================================================================
 * SHEET 5: CONTRACT BALANCES RECONCILIATION
 * ============================================================================
 */
function createContractBalancesSheet(ss) {
  let sheet = ss.getSheetByName('Contract Balances');
  if (!sheet) {
    sheet = ss.insertSheet('Contract Balances', 4);
  } else {
    sheet.clear();
  }
  
  // Header
  sheet.getRange('A1:K1').merge()
    .setValue('CONTRACT BALANCES RECONCILIATION - IND AS 115')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  sheet.setRowHeight(1, 40);
  
  // Instructions
  sheet.getRange('A2:K2').merge()
    .setValue('Reconciliation of Contract Assets, Contract Liabilities, and Accounts Receivable per Ind AS 115 para 116-118')
    .setBackground('#fff3cd')
    .setWrap(true);
  
  // Column Headers
  const headers = [
    'Contract ID',
    'Customer',
    'Opening Balance\nContract Asset',
    'Revenue Recognized\n(Period)',
    'Billed\n(Period)',
    'Closing Balance\nContract Asset',
    'Closing Balance\nContract Liability',
    'Accounts\nReceivable',
    'Collections\n(Period)',
    'Closing A/R\nBalance',
    'Notes'
  ];
  
  sheet.getRange('A3:K3').setValues([headers])
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setWrap(true);
  
  sheet.setRowHeight(3, 60);
  
  // Set column widths
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidths(3, 8, 120);
  sheet.setColumnWidth(11, 200);
  
  // Link to Revenue Recognition sheet
  for (let i = 4; i <= 50; i++) {
    // Contract ID
    sheet.getRange(i, 1).setFormula(`=IF('Revenue Recognition'!B${i}="",'Revenue Recognition'!B${i})`);
    
    // Customer
    sheet.getRange(i, 2).setFormula(`=IF(A${i}="","",IFERROR(VLOOKUP(A${i},'Contract Register'!B:C,2,FALSE),""))`);
    
    // Opening Contract Asset - INPUT
    sheet.getRange(i, 3).setBackground('#cfe2f3').setNumberFormat('#,##0.00')
      .setNote('INPUT: Opening balance of contract asset');
    
    // Revenue Recognized (link)
    sheet.getRange(i, 4).setFormula(`=IF(A${i}="","",IFERROR(INDEX('Revenue Recognition'!H:H,${i}),""))`).setNumberFormat('#,##0.00');
    
    // Billed (link)
    sheet.getRange(i, 5).setFormula(`=IF(A${i}="","",IFERROR(INDEX('Revenue Recognition'!N:N,${i}),""))`).setNumberFormat('#,##0.00');
    
    // Closing Contract Asset
    sheet.getRange(i, 6).setFormula(`=IF(A${i}="","",MAX(0,C${i}+D${i}-E${i}))`).setNumberFormat('#,##0.00')
      .setNote('Opening + Revenue - Billed');
    
    // Contract Liability (from Revenue Recognition)
    sheet.getRange(i, 7).setFormula(`=IF(A${i}="","",IFERROR(INDEX('Revenue Recognition'!P:P,${i}),""))`).setNumberFormat('#,##0.00');
    
    // Accounts Receivable (Billed - advances)
    sheet.getRange(i, 8).setFormula(`=IF(A${i}="","",E${i}-G${i})`).setNumberFormat('#,##0.00');
    
    // Collections - INPUT
    sheet.getRange(i, 9).setBackground('#cfe2f3').setNumberFormat('#,##0.00')
      .setNote('INPUT: Cash collected from customer in period');
    
    // Closing A/R
    sheet.getRange(i, 10).setFormula(`=IF(A${i}="","",H${i}-I${i})`).setNumberFormat('#,##0.00');
    
    // Notes
    sheet.getRange(i, 11).setBackground('#cfe2f3');
  }
  
  // Summary Section
  const summaryRow = 52;
  sheet.getRange(summaryRow, 1, 1, 2).merge()
    .setValue('SUMMARY TOTALS')
    .setFontWeight('bold')
    .setBackground('#e8f0fe')
    .setHorizontalAlignment('center');
  
  const summaryLabels = [
    'Total Opening Contract Assets:',
    'Total Revenue Recognized:',
    'Total Billed:',
    'Total Closing Contract Assets:',
    'Total Closing Contract Liabilities:',
    'Total Closing Accounts Receivable:'
  ];
  
  const summaryCols = ['C', 'D', 'E', 'F', 'G', 'J'];
  
  let sumRow = summaryRow + 1;
  summaryLabels.forEach((label, idx) => {
    sheet.getRange(sumRow, 1, 1, 2).merge().setValue(label);
    sheet.getRange(sumRow, 3).setFormula(`=SUM(${summaryCols[idx]}4:${summaryCols[idx]}51)`)
      .setNumberFormat('#,##0.00')
      .setFontWeight('bold');
    sumRow++;
  });
  
  // Control Check
  sumRow++;
  sheet.getRange(sumRow, 1, 1, 2).merge()
    .setValue('CONTROL CHECK: Asset + Liability Balance')
    .setFontWeight('bold')
    .setBackground('#fff3cd');
  sheet.getRange(sumRow, 3).setFormula(`=${summaryRow+4}+${summaryRow+5}`)
    .setNumberFormat('#,##0.00')
    .setFontWeight('bold');
  
  sheet.getRange(sumRow+1, 1, 1, 2).merge()
    .setValue('Should Equal: Transaction Price Less Revenue')
    .setFontStyle('italic');
  sheet.getRange(sumRow+1, 3).setFormula(`=SUM('Revenue Recognition'!F4:F51)-SUM('Revenue Recognition'!L4:L51)`)
    .setNumberFormat('#,##0.00');
  
  sheet.getRange(sumRow+2, 1, 1, 2).merge()
    .setValue('Variance:')
    .setFontWeight('bold')
    .setFontColor('#cc0000');
  sheet.getRange(sumRow+2, 3).setFormula(`=${sumRow}-${sumRow+1}`)
    .setNumberFormat('#,##0.00')
    .setFontWeight('bold');
  
  // Freeze
  sheet.setFrozenRows(3);
}

/**
 * ============================================================================
 * SHEET 6: PERFORMANCE OBLIGATIONS TRACKER
 * ============================================================================
 */
function createPerformanceObligationsSheet(ss) {
  let sheet = ss.getSheetByName('Performance Obligations');
  if (!sheet) {
    sheet = ss.insertSheet('Performance Obligations', 5);
  } else {
    sheet.clear();
  }
  
  // Header
  sheet.getRange('A1:L1').merge()
    .setValue('PERFORMANCE OBLIGATIONS TRACKING - IND AS 115')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  sheet.setRowHeight(1, 40);
  
  // Instructions
  sheet.getRange('A2:L2').merge()
    .setValue('Track each distinct performance obligation separately. For contracts with multiple POs, allocate transaction price based on standalone selling prices.')
    .setBackground('#fff3cd')
    .setWrap(true);
  
  // Column Headers
  const headers = [
    'Contract ID',
    'PO ID',
    'Performance Obligation Description',
    'Standalone Selling Price',
    'Allocated Transaction Price',
    'Allocation %',
    'Recognition Pattern',
    'Satisfaction Method',
    '% Complete',
    'Revenue Recognized to Date',
    'Remaining Performance Obligation',
    'Expected Recognition Date'
  ];
  
  sheet.getRange('A3:L3').setValues([headers])
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setWrap(true);
  
  sheet.setRowHeight(3, 50);
  
  // Set column widths
  sheet.setColumnWidths(1, 12, 130);
  sheet.setColumnWidth(3, 250);
  
  // Sample rows with formulas
  const sampleData = [
    ['CNT-001', 'PO-001-1', 'Software License', 300000, '', '', 'Point in Time', 'Transfer of control', 100, '', '', '31-Mar-2025'],
    ['CNT-001', 'PO-001-2', 'Implementation Services', 500000, '', '', 'Over Time', 'Input Method', 50, '', '', '31-Dec-2025'],
    ['CNT-001', 'PO-001-3', 'Annual Maintenance', 200000, '', '', 'Over Time', 'Time Elapsed', 25, '', '', '31-Mar-2026']
  ];
  
  let row = 4;
  sampleData.forEach(data => {
    // Contract ID
    sheet.getRange(row, 1).setValue(data[0]).setBackground('#cfe2f3');
    
    // PO ID
    sheet.getRange(row, 2).setValue(data[1]).setBackground('#cfe2f3');
    
    // Description
    sheet.getRange(row, 3).setValue(data[2]).setBackground('#cfe2f3');
    
    // Standalone Selling Price
    sheet.getRange(row, 4).setValue(data[3]).setNumberFormat('#,##0.00').setBackground('#cfe2f3');
    
    // Allocated Transaction Price (formula to allocate based on SSP)
    sheet.getRange(row, 5).setFormula(
      `=IF(A${row}="","",D${row}/(SUMIF($A$4:$A$100,A${row},$D$4:$D$100))*IFERROR(VLOOKUP(A${row},'Contract Register'!B:F,5,FALSE),0))`
    ).setNumberFormat('#,##0.00')
      .setNote('Auto: Allocated based on relative standalone selling price');
    
    // Allocation %
    sheet.getRange(row, 6).setFormula(
      `=IF(E${row}="","",E${row}/IFERROR(VLOOKUP(A${row},'Contract Register'!B:F,5,FALSE),1))`
    ).setNumberFormat('0.00%');
    
    // Recognition Pattern
    sheet.getRange(row, 7).setValue(data[6]).setBackground('#cfe2f3');
    
    // Satisfaction Method
    sheet.getRange(row, 8).setValue(data[7]).setBackground('#cfe2f3');
    
    // % Complete
    sheet.getRange(row, 9).setValue(data[8]).setNumberFormat('0').setBackground('#cfe2f3')
      .setNote('INPUT: Percentage of PO satisfied (0-100)');
    
    // Revenue Recognized to Date
    sheet.getRange(row, 10).setFormula(`=IF(E${row}="","",E${row}*I${row}/100)`)
      .setNumberFormat('#,##0.00')
      .setNote('Auto: Allocated price × % complete');
    
    // Remaining Performance Obligation
    sheet.getRange(row, 11).setFormula(`=IF(E${row}="","",E${row}-J${row})`)
      .setNumberFormat('#,##0.00')
      .setNote('Auto: Remaining unrecognized revenue');
    
    // Expected Recognition Date
    sheet.getRange(row, 12).setValue(data[11]).setNumberFormat('dd-mmm-yyyy').setBackground('#cfe2f3');
    
    row++;
  });
  
  // Format additional blank rows
  for (let i = row; i <= 100; i++) {
    sheet.getRange(i, 1).setBackground('#cfe2f3');
    sheet.getRange(i, 2).setBackground('#cfe2f3');
    sheet.getRange(i, 3).setBackground('#cfe2f3');
    sheet.getRange(i, 4).setBackground('#cfe2f3').setNumberFormat('#,##0.00');
    sheet.getRange(i, 5).setFormula(
      `=IF(A${i}="","",D${i}/(SUMIF($A$4:$A$100,A${i},$D$4:$D$100))*IFERROR(VLOOKUP(A${i},'Contract Register'!B:F,5,FALSE),0))`
    ).setNumberFormat('#,##0.00');
    sheet.getRange(i, 6).setFormula(
      `=IF(E${i}="","",E${i}/IFERROR(VLOOKUP(A${i},'Contract Register'!B:F,5,FALSE),1))`
    ).setNumberFormat('0.00%');
    sheet.getRange(i, 7).setBackground('#cfe2f3');
    sheet.getRange(i, 8).setBackground('#cfe2f3');
    sheet.getRange(i, 9).setBackground('#cfe2f3').setNumberFormat('0');
    sheet.getRange(i, 10).setFormula(`=IF(E${i}="","",E${i}*I${i}/100)`).setNumberFormat('#,##0.00');
    sheet.getRange(i, 11).setFormula(`=IF(E${i}="","",E${i}-J${i})`).setNumberFormat('#,##0.00');
    sheet.getRange(i, 12).setBackground('#cfe2f3').setNumberFormat('dd-mmm-yyyy');
  }
  
  // Add data validation for Recognition Pattern
  const patternRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Point in Time', 'Over Time'], true)
    .build();
  sheet.getRange('G4:G100').setDataValidation(patternRule);
  
  // Summary
  const summaryRow = 102;
  sheet.getRange(summaryRow, 1, 1, 3).merge()
    .setValue('SUMMARY - REMAINING PERFORMANCE OBLIGATIONS')
    .setFontWeight('bold')
    .setBackground('#e8f0fe');
  
  sheet.getRange(summaryRow+1, 1, 1, 2).merge().setValue('Total Allocated Transaction Price:');
  sheet.getRange(summaryRow+1, 3).setFormula('=SUM(E4:E100)').setNumberFormat('#,##0.00').setFontWeight('bold');
  
  sheet.getRange(summaryRow+2, 1, 1, 2).merge().setValue('Total Revenue Recognized:');
  sheet.getRange(summaryRow+2, 3).setFormula('=SUM(J4:J100)').setNumberFormat('#,##0.00').setFontWeight('bold');
  
  sheet.getRange(summaryRow+3, 1, 1, 2).merge().setValue('Total Remaining PO (Disclosure):');
  sheet.getRange(summaryRow+3, 3).setFormula('=SUM(K4:K100)').setNumberFormat('#,##0.00').setFontWeight('bold')
    .setBackground('#fff3cd')
    .setNote('Required disclosure per Ind AS 115.120');
  
  // Freeze
  sheet.setFrozenRows(3);
}

/**
 * ============================================================================
 * SHEET 7: VARIABLE CONSIDERATION
 * ============================================================================
 */
function createVariableConsiderationSheet(ss) {
  let sheet = ss.getSheetByName('Variable Consideration');
  if (!sheet) {
    sheet = ss.insertSheet('Variable Consideration', 6);
  } else {
    sheet.clear();
  }
  
  // Header
  sheet.getRange('A1:L1').merge()
    .setValue('VARIABLE CONSIDERATION ESTIMATION - IND AS 115')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  sheet.setRowHeight(1, 40);
  
  // Instructions
  sheet.getRange('A2:L2').merge()
    .setValue('Track variable consideration (discounts, penalties, bonuses, incentives). Apply constraint per Ind AS 115.56-58.')
    .setBackground('#fff3cd')
    .setWrap(true);
  
  // Headers
  const headers = [
    'Contract ID',
    'Type of Variable Consideration',
    'Base Transaction Price',
    'Variable Amount (Expected)',
    'Estimation Method',
    'Probability of Occurrence',
    'Constrained Amount',
    'Final Transaction Price',
    'Reason for Constraint',
    'Reassessment Date',
    'Actual Outcome',
    'True-up Adjustment Required'
  ];
  
  sheet.getRange('A3:L3').setValues([headers])
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setWrap(true);
  
  sheet.setRowHeight(3, 50);
  
  // Set column widths
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidths(3, 10, 130);
  
  // Sample Data
  row = 4;
  
  // Contract ID - link
  sheet.getRange('A4').setValue('CNT-001').setBackground('#cfe2f3');
  
  // Type
  sheet.getRange('B4').setValue('Performance Bonus').setBackground('#cfe2f3')
    .setNote('Type: Discount/Refund/Penalty/Bonus/Incentive/Rebate');
  
  // Base Price (link to contract register)
  sheet.getRange('C4').setFormula('=IFERROR(VLOOKUP(A4,\'Contract Register\'!B:F,5,FALSE),"")')
    .setNumberFormat('#,##0.00');
  
  // Variable Amount
  sheet.getRange('D4').setValue(50000).setNumberFormat('#,##0.00').setBackground('#cfe2f3')
    .setNote('INPUT: Expected value of variable consideration');
  
  // Estimation Method
  sheet.getRange('E4').setValue('Most Likely').setBackground('#cfe2f3')
    .setNote('Method: Expected Value / Most Likely Amount');
  
  // Probability
  sheet.getRange('F4').setValue(0.70).setNumberFormat('0%').setBackground('#cfe2f3')
    .setNote('INPUT: Probability of achieving variable consideration');
  
  // Constrained Amount (apply constraint if highly uncertain)
  sheet.getRange('G4').setFormula('=IF(F4>=0.75,D4,0)')
    .setNumberFormat('#,##0.00')
    .setNote('Auto: Include only if probability >= 75% (constraint threshold)');
  
  // Final Transaction Price
  sheet.getRange('H4').setFormula('=C4+G4')
    .setNumberFormat('#,##0.00')
    .setNote('Auto: Base price + constrained variable consideration');
  
  // Reason for Constraint
  sheet.getRange('I4').setFormula('=IF(G4<D4,"Highly uncertain - constraint applied","No constraint")')
    .setBackground('#ffe599');
  
  // Reassessment Date
  sheet.getRange('J4').setValue('31-Mar-2025').setNumberFormat('dd-mmm-yyyy').setBackground('#cfe2f3');
  
  // Actual Outcome
  sheet.getRange('K4').setBackground('#cfe2f3').setNumberFormat('#,##0.00')
    .setNote('INPUT: Actual variable consideration received/settled');
  
  // True-up Required
  sheet.getRange('L4').setFormula('=IF(K4="","",K4-G4)')
    .setNumberFormat('#,##0.00')
    .setNote('Auto: Difference between actual and estimated');
  
  // Format additional rows
  for (let i = 5; i <= 50; i++) {
    sheet.getRange(i, 1).setBackground('#cfe2f3');
    sheet.getRange(i, 2).setBackground('#cfe2f3');
    sheet.getRange(i, 3).setFormula(`=IFERROR(VLOOKUP(A${i},'Contract Register'!B:F,5,FALSE),"")`).setNumberFormat('#,##0.00');
    sheet.getRange(i, 4).setBackground('#cfe2f3').setNumberFormat('#,##0.00');
    sheet.getRange(i, 5).setBackground('#cfe2f3');
    sheet.getRange(i, 6).setBackground('#cfe2f3').setNumberFormat('0%');
    sheet.getRange(i, 7).setFormula(`=IF(A${i}="","",IF(F${i}>=0.75,D${i},0))`).setNumberFormat('#,##0.00');
    sheet.getRange(i, 8).setFormula(`=IF(A${i}="","",C${i}+G${i})`).setNumberFormat('#,##0.00');
    sheet.getRange(i, 9).setFormula(`=IF(A${i}="","",IF(G${i}<D${i},"Highly uncertain - constraint applied","No constraint"))`);
    sheet.getRange(i, 10).setBackground('#cfe2f3').setNumberFormat('dd-mmm-yyyy');
    sheet.getRange(i, 11).setBackground('#cfe2f3').setNumberFormat('#,##0.00');
    sheet.getRange(i, 12).setFormula(`=IF(K${i}="","",K${i}-G${i})`).setNumberFormat('#,##0.00');
  }
  
  // Add data validation
  const methodRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Expected Value', 'Most Likely Amount'], true)
    .build();
  sheet.getRange('E4:E50').setDataValidation(methodRule);
  
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Discount', 'Refund', 'Penalty', 'Bonus', 'Incentive', 'Rebate', 'Credit Note'], true)
    .build();
  sheet.getRange('B4:B50').setDataValidation(typeRule);
  
  // Summary
  const summaryRow = 52;
  sheet.getRange(summaryRow, 1, 1, 2).merge()
    .setValue('SUMMARY')
    .setFontWeight('bold')
    .setBackground('#e8f0fe');
  
  sheet.getRange(summaryRow+1, 1).setValue('Total Variable Consideration (Expected):');
  sheet.getRange(summaryRow+1, 2).setFormula('=SUM(D4:D51)').setNumberFormat('#,##0.00').setFontWeight('bold');
  
  sheet.getRange(summaryRow+2, 1).setValue('Total Variable Consideration (Constrained):');
  sheet.getRange(summaryRow+2, 2).setFormula('=SUM(G4:G51)').setNumberFormat('#,##0.00').setFontWeight('bold');
  
  sheet.getRange(summaryRow+3, 1).setValue('Total True-up Adjustments Required:');
  sheet.getRange(summaryRow+3, 2).setFormula('=SUM(L4:L51)').setNumberFormat('#,##0.00').setFontWeight('bold')
    .setBackground('#fff3cd');
  
  // Freeze
  sheet.setFrozenRows(3);
}

/**
 * ============================================================================
 * SHEET 8: PERIOD-END ADJUSTMENTS & JOURNAL ENTRIES
 * ============================================================================
 */
function createPeriodEndAdjustmentsSheet(ss) {
  let sheet = ss.getSheetByName('Period-End Adjustments');
  if (!sheet) {
    sheet = ss.insertSheet('Period-End Adjustments', 7);
  } else {
    sheet.clear();
  }
  
  // Header
  sheet.getRange('A1:J1').merge()
    .setValue('PERIOD-END ADJUSTMENTS & JOURNAL ENTRIES - IND AS 115')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  sheet.setRowHeight(1, 40);
  
  // Instructions
  sheet.getRange('A2:J2').merge()
    .setValue('Book closure entries for revenue recognition, contract assets/liabilities, and related adjustments. Review before posting to GL.')
    .setBackground('#fff3cd')
    .setWrap(true);
  
  // Column Headers
  const headers = [
    'JE#',
    'Date',
    'Description',
    'Contract ID',
    'Account Head',
    'Account Code',
    'Debit Amount',
    'Credit Amount',
    'Supporting Schedule',
    'Preparer Notes'
  ];
  
  sheet.getRange('A3:J3').setValues([headers])
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setWrap(true);
  
  sheet.setRowHeight(3, 40);
  
  // Set column widths
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 250);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidths(7, 2, 120);
  sheet.setColumnWidth(9, 150);
  sheet.setColumnWidth(10, 200);
  
  // Sample Journal Entries
  let row = 4;
  
  // JE 1: Revenue Recognition
  const je1Data = [
    ['JE-001', '31-Mar-2025', 'To recognize revenue for services rendered as per Ind AS 115', 'CNT-001', 'Contract Asset / Unbilled Revenue', '1234', 250000, '', 'Revenue Recognition!H4', 'Based on % completion: 25%'],
    ['', '', '', '', 'Revenue from Operations', '4001', '', 250000, '', '']
  ];
  
  je1Data.forEach(data => {
    sheet.getRange(row, 1, 1, 10).setValues([data]);
    if (data[0]) {
      sheet.getRange(row, 1).setBackground('#cfe2f3');
      sheet.getRange(row, 2).setBackground('#cfe2f3').setNumberFormat('dd-mmm-yyyy');
      sheet.getRange(row, 3).setBackground('#cfe2f3');
      sheet.getRange(row, 4).setBackground('#cfe2f3');
    }
    sheet.getRange(row, 5).setBackground('#cfe2f3');
    sheet.getRange(row, 6).setBackground('#cfe2f3');
    sheet.getRange(row, 7).setBackground('#cfe2f3').setNumberFormat('#,##0.00');
    sheet.getRange(row, 8).setBackground('#cfe2f3').setNumberFormat('#,##0.00');
    sheet.getRange(row, 9).setBackground('#cfe2f3');
    sheet.getRange(row, 10).setBackground('#cfe2f3');
    row++;
  });
  
  // Blank line
  row++;
  
  // JE 2: Invoice/Billing Entry
  const je2Data = [
    ['JE-002', '31-Mar-2025', 'To record invoice raised to customer', 'CNT-001', 'Accounts Receivable', '1101', 300000, '', 'Contract Balances!E4', 'Invoice# INV-2025-001'],
    ['', '', '', '', 'GST Output (18%)', '2401', '', 48000, '', ''],
    ['', '', '', '', 'Contract Asset / Unbilled Revenue', '1234', '', 250000, '', 'Transfer from unbilled to billed'],
    ['', '', '', '', 'Contract Liability (Advance)', '2301', '', 2000, '', 'Excess billing over revenue']
  ];
  
  je2Data.forEach(data => {
    sheet.getRange(row, 1, 1, 10).setValues([data]);
    if (data[0]) {
      sheet.getRange(row, 1).setBackground('#cfe2f3');
      sheet.getRange(row, 2).setBackground('#cfe2f3').setNumberFormat('dd-mmm-yyyy');
      sheet.getRange(row, 3).setBackground('#cfe2f3');
      sheet.getRange(row, 4).setBackground('#cfe2f3');
    }
    sheet.getRange(row, 5).setBackground('#cfe2f3');
    sheet.getRange(row, 6).setBackground('#cfe2f3');
    sheet.getRange(row, 7).setBackground('#cfe2f3').setNumberFormat('#,##0.00');
    sheet.getRange(row, 8).setBackground('#cfe2f3').setNumberFormat('#,##0.00');
    sheet.getRange(row, 9).setBackground('#cfe2f3');
    sheet.getRange(row, 10).setBackground('#cfe2f3');
    row++;
  });
  
  row++;
  
  // JE 3: Deferred Revenue Recognition
  const je3Data = [
    ['JE-003', '31-Mar-2025', 'To defer unearned revenue to contract liability', 'CNT-002', 'Revenue from Operations', '4001', 100000, '', 'Revenue Recognition!M5', 'Revenue not yet earned'],
    ['', '', '', '', 'Contract Liability (Deferred Revenue)', '2301', '', 100000, '', 'To be recognized in future period']
  ];
  
  je3Data.forEach(data => {
    sheet.getRange(row, 1, 1, 10).setValues([data]);
    if (data[0]) {
      sheet.getRange(row, 1).setBackground('#cfe2f3');
      sheet.getRange(row, 2).setBackground('#cfe2f3').setNumberFormat('dd-mmm-yyyy');
      sheet.getRange(row, 3).setBackground('#cfe2f3');
      sheet.getRange(row, 4).setBackground('#cfe2f3');
    }
    sheet.getRange(row, 5).setBackground('#cfe2f3');
    sheet.getRange(row, 6).setBackground('#cfe2f3');
    sheet.getRange(row, 7).setBackground('#cfe2f3').setNumberFormat('#,##0.00');
    sheet.getRange(row, 8).setBackground('#cfe2f3').setNumberFormat('#,##0.00');
    sheet.getRange(row, 9).setBackground('#cfe2f3');
    sheet.getRange(row, 10).setBackground('#cfe2f3');
    row++;
  });
  
  row++;
  
  // JE 4: Variable Consideration True-up
  const je4Data = [
    ['JE-004', '31-Mar-2025', 'True-up adjustment for variable consideration', 'CNT-001', 'Contract Asset', '1234', 15000, '', 'Variable Consideration!L4', 'Bonus received higher than estimated'],
    ['', '', '', '', 'Revenue from Operations', '4001', '', 15000, '', 'Revenue adjustment']
  ];
  
  je4Data.forEach(data => {
    sheet.getRange(row, 1, 1, 10).setValues([data]);
    if (data[0]) {
      sheet.getRange(row, 1).setBackground('#cfe2f3');
      sheet.getRange(row, 2).setBackground('#cfe2f3').setNumberFormat('dd-mmm-yyyy');
      sheet.getRange(row, 3).setBackground('#cfe2f3');
      sheet.getRange(row, 4).setBackground('#cfe2f3');
    }
    sheet.getRange(row, 5).setBackground('#cfe2f3');
    sheet.getRange(row, 6).setBackground('#cfe2f3');
    sheet.getRange(row, 7).setBackground('#cfe2f3').setNumberFormat('#,##0.00');
    sheet.getRange(row, 8).setBackground('#cfe2f3').setNumberFormat('#,##0.00');
    sheet.getRange(row, 9).setBackground('#cfe2f3');
    sheet.getRange(row, 10).setBackground('#cfe2f3');
    row++;
  });
  
  // Format blank rows for additional entries
  for (let i = row + 2; i <= row + 30; i++) {
    sheet.getRange(i, 1).setBackground('#cfe2f3');
    sheet.getRange(i, 2).setBackground('#cfe2f3').setNumberFormat('dd-mmm-yyyy');
    sheet.getRange(i, 3).setBackground('#cfe2f3');
    sheet.getRange(i, 4).setBackground('#cfe2f3');
    sheet.getRange(i, 5).setBackground('#cfe2f3');
    sheet.getRange(i, 6).setBackground('#cfe2f3');
    sheet.getRange(i, 7).setBackground('#cfe2f3').setNumberFormat('#,##0.00');
    sheet.getRange(i, 8).setBackground('#cfe2f3').setNumberFormat('#,##0.00');
    sheet.getRange(i, 9).setBackground('#cfe2f3');
    sheet.getRange(i, 10).setBackground('#cfe2f3');
  }
  
  // Control Totals Section
  const summaryRow = row + 32;
  sheet.getRange(summaryRow, 1, 1, 3).merge()
    .setValue('CONTROL TOTALS - JOURNAL ENTRIES')
    .setFontWeight('bold')
    .setBackground('#e8f0fe')
    .setHorizontalAlignment('center');
  
  sheet.getRange(summaryRow + 1, 1, 1, 2).merge().setValue('Total Debits:');
  sheet.getRange(summaryRow + 1, 3).setFormula(`=SUM(G4:G${summaryRow-1})`)
    .setNumberFormat('#,##0.00')
    .setFontWeight('bold');
  
  sheet.getRange(summaryRow + 2, 1, 1, 2).merge().setValue('Total Credits:');
  sheet.getRange(summaryRow + 2, 3).setFormula(`=SUM(H4:H${summaryRow-1})`)
    .setNumberFormat('#,##0.00')
    .setFontWeight('bold');
  
  sheet.getRange(summaryRow + 3, 1, 1, 2).merge().setValue('Out of Balance:');
  sheet.getRange(summaryRow + 3, 3).setFormula(`=${summaryRow+1}-${summaryRow+2}`)
    .setNumberFormat('#,##0.00')
    .setFontWeight('bold')
    .setBackground('#f4cccc')
    .setNote('Should be ZERO. Non-zero indicates unbalanced entries.');
  
  // Conditional formatting for out of balance
  const balanceRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberNotEqualTo(0)
    .setBackground('#cc0000')
    .setFontColor('#ffffff')
    .setRanges([sheet.getRange(summaryRow + 3, 3)])
    .build();
  
  const rules = sheet.getConditionalFormatRules();
  rules.push(balanceRule);
  sheet.setConditionalFormatRules(rules);
  
  // Freeze
  sheet.setFrozenRows(3);
}

/**
 * ============================================================================
 * SHEET 9: IGAAP vs IND AS RECONCILIATION
 * ============================================================================
 */
function createIGAAPReconciliationSheet(ss) {
  let sheet = ss.getSheetByName('IGAAP Reconciliation');
  if (!sheet) {
    sheet = ss.insertSheet('IGAAP Reconciliation', 8);
  } else {
    sheet.clear();
  }
  
  // Header
  sheet.getRange('A1:H1').merge()
    .setValue('IGAAP vs IND AS 115 RECONCILIATION')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  sheet.setRowHeight(1, 40);
  
  // Instructions
  sheet.getRange('A2:H2').merge()
    .setValue('Reconciliation of revenue under Old IGAAP (Ind AS 18/11) vs New Ind AS 115. Document all transitional adjustments.')
    .setBackground('#fff3cd')
    .setWrap(true);
  
  // Column Headers
  const headers = [
    'Item / Contract ID',
    'IGAAP Revenue\n(Ind AS 18/11)',
    'Ind AS 115 Adjustment',
    'Ind AS 115 Revenue',
    'Reason for Adjustment',
    'Ind AS 115 Reference',
    'Impact on P&L',
    'Impact on Balance Sheet'
  ];
  
  sheet.getRange('A3:H3').setValues([headers])
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setWrap(true);
  
  sheet.setRowHeight(3, 50);
  
  // Set column widths
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidths(2, 4, 130);
  sheet.setColumnWidth(5, 250);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidths(7, 2, 150);
  
  // Sample Reconciliation Items
  const reconData = [
    [
      'CNT-001 - Software Project',
      1000000,
      -250000,
      750000,
      'Under IGAAP: Full revenue recognized on billing. Under Ind AS 115: Revenue recognized based on % completion.',
      'Para 35 - Over time recognition',
      -250000,
      'Deferred Revenue +250000'
    ],
    [
      'CNT-002 - License + Services',
      800000,
      50000,
      850000,
      'Under IGAAP: Bundled as single unit. Under Ind AS 115: Separated into distinct performance obligations with different timing.',
      'Para 27-30 - Distinct POs',
      50000,
      'Contract Asset +50000'
    ],
    [
      'CNT-003 - Variable Consideration',
      500000,
      -100000,
      400000,
      'Under IGAAP: Full variable consideration included. Under Ind AS 115: Constraint applied due to uncertainty.',
      'Para 56-58 - Constraint',
      -100000,
      'Contract Liability +100000'
    ],
    [
      'CNT-004 - Significant Financing',
      1200000,
      -80000,
      1120000,
      'Under IGAAP: No adjustment for financing component. Under Ind AS 115: PV adjustment for extended credit period (18 months).',
      'Para 60-65 - Financing',
      -80000,
      'Interest Expense +80000'
    ]
  ];
  
  let row = 4;
  reconData.forEach(data => {
    sheet.getRange(row, 1).setValue(data[0]).setBackground('#cfe2f3');
    sheet.getRange(row, 2).setValue(data[1]).setNumberFormat('#,##0.00').setBackground('#cfe2f3')
      .setNote('INPUT: Revenue as per old IGAAP accounting policy');
    sheet.getRange(row, 3).setValue(data[2]).setNumberFormat('#,##0.00').setBackground('#cfe2f3')
      .setNote('INPUT: Adjustment amount (positive or negative)');
    sheet.getRange(row, 4).setFormula(`=B${row}+C${row}`).setNumberFormat('#,##0.00')
      .setNote('Auto: Revenue as per Ind AS 115');
    sheet.getRange(row, 5).setValue(data[4]).setBackground('#cfe2f3').setWrap(true);
    sheet.getRange(row, 6).setValue(data[5]).setBackground('#cfe2f3');
    sheet.getRange(row, 7).setFormula(`=C${row}`).setNumberFormat('#,##0.00')
      .setNote('Impact on Profit & Loss');
    sheet.getRange(row, 8).setValue(data[7]).setBackground('#cfe2f3');
    
    sheet.setRowHeight(row, 50);
    row++;
  });
  
  // Format additional blank rows
  for (let i = row; i <= row + 20; i++) {
    sheet.getRange(i, 1).setBackground('#cfe2f3');
    sheet.getRange(i, 2).setBackground('#cfe2f3').setNumberFormat('#,##0.00');
    sheet.getRange(i, 3).setBackground('#cfe2f3').setNumberFormat('#,##0.00');
    sheet.getRange(i, 4).setFormula(`=IF(B${i}="","",B${i}+C${i})`).setNumberFormat('#,##0.00');
    sheet.getRange(i, 5).setBackground('#cfe2f3');
    sheet.getRange(i, 6).setBackground('#cfe2f3');
    sheet.getRange(i, 7).setFormula(`=IF(C${i}="","",C${i})`).setNumberFormat('#,##0.00');
    sheet.getRange(i, 8).setBackground('#cfe2f3');
  }
  
  // Summary Section
  const summaryRow = row + 22;
  sheet.getRange(summaryRow, 1, 1, 2).merge()
    .setValue('RECONCILIATION SUMMARY')
    .setFontWeight('bold')
    .setBackground('#e8f0fe')
    .setHorizontalAlignment('center');
  
  sheet.getRange(summaryRow + 1, 1).setValue('Total Revenue - IGAAP Basis:');
  sheet.getRange(summaryRow + 1, 2).setFormula(`=SUM(B4:B${summaryRow-1})`)
    .setNumberFormat('#,##0.00')
    .setFontWeight('bold');
  
  sheet.getRange(summaryRow + 2, 1).setValue('Total Ind AS 115 Adjustments:');
  sheet.getRange(summaryRow + 2, 2).setFormula(`=SUM(C4:C${summaryRow-1})`)
    .setNumberFormat('#,##0.00')
    .setFontWeight('bold')
    .setBackground('#fff3cd');
  
  sheet.getRange(summaryRow + 3, 1).setValue('Total Revenue - Ind AS 115 Basis:');
  sheet.getRange(summaryRow + 3, 2).setFormula(`=SUM(D4:D${summaryRow-1})`)
    .setNumberFormat('#,##0.00')
    .setFontWeight('bold');
  
  sheet.getRange(summaryRow + 5, 1).setValue('Net Impact on P&L:');
  sheet.getRange(summaryRow + 5, 2).setFormula(`=SUM(G4:G${summaryRow-1})`)
    .setNumberFormat('#,##0.00')
    .setFontWeight('bold')
    .setBackground('#f4cccc')
    .setNote('Total impact on profit/loss from Ind AS 115 adoption');
  
  // Freeze
  sheet.setFrozenRows(3);
}

/**
 * ============================================================================
 * SHEET 10: REFERENCES
 * ============================================================================
 */
function createReferencesSheet(ss) {
  let sheet = ss.getSheetByName('References');
  if (!sheet) {
    sheet = ss.insertSheet('References', 9);
  } else {
    sheet.clear();
  }
  
  // Header
  sheet.getRange('A1:D1').merge()
    .setValue('IND AS 115 - REFERENCES & CITATIONS')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  sheet.setRowHeight(1, 40);
  
  // Set column widths
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 400);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 300);
  
  // Section 1: Key Paragraphs
  sheet.getRange('A3:D3').merge()
    .setValue('KEY IND AS 115 PARAGRAPHS')
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  
  const keyParas = [
    ['Paragraph', 'Description', 'Page', 'Notes'],
    ['Para 9-21', 'Step 1: Identify the Contract', '', 'Contract must be enforceable, approved, rights identified, payment terms clear, commercial substance'],
    ['Para 22-30', 'Step 2: Identify Performance Obligations', '', 'Distinct goods/services or series of distinct goods/services'],
    ['Para 31-35', 'Step 3: Determine Transaction Price', '', 'Variable consideration, significant financing, non-cash consideration, consideration payable'],
    ['Para 36-45', 'Step 4: Allocate Transaction Price', '', 'Based on standalone selling prices; use adjusted market assessment/cost plus/residual approach'],
    ['Para 31-45', 'Step 5: Recognize Revenue', '', 'When/as entity satisfies performance obligation by transferring control'],
    ['Para 56-58', 'Variable Consideration Constraint', '', 'Include only to extent highly probable that reversal will not occur'],
    ['Para 60-65', 'Significant Financing Component', '', 'Adjust transaction price for time value of money if significant'],
    ['Para 106-109', 'Contract Assets & Liabilities', '', 'Asset: Right to consideration conditional on something other than time. Liability: Obligation to transfer goods/services'],
    ['Para 110-129', 'Disclosure Requirements', '', 'Disaggregation, performance obligations, transaction price, contract balances']
  ];
  
  sheet.getRange(4, 1, keyParas.length, 4).setValues(keyParas);
  sheet.getRange('A4:D4').setFontWeight('bold').setBackground('#e8f0fe');
  
  // Section 2: Common Industry Applications
  let row = 4 + keyParas.length + 2;
  sheet.getRange(row, 1, 1, 4).merge()
    .setValue('COMMON INDUSTRY APPLICATIONS')
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  
  row++;
  const industries = [
    ['Industry', 'Typical Revenue Pattern', '', 'Key Considerations'],
    ['Software/SaaS', 'License: Point in Time; Services: Over Time', '', 'Separate POs, Subscription renewals, Implementation vs license'],
    ['Construction', 'Over Time (Input/Output Method)', '', 'Cost-to-cost, milestones, variations, retention money'],
    ['Consulting', 'Over Time (Time-based)', '', 'Time & materials, fixed fee, retainers'],
    ['Manufacturing', 'Point in Time (Delivery)', '', 'Right of return, warranties, trade discounts'],
    ['Real Estate', 'Over Time / Point in Time', '', 'Significant financing, payment plans, customization'],
    ['Telecom', 'Over Time (Subscription)', '', 'Device + service bundles, loyalty points, roaming'],
    ['Healthcare', 'Mixed (Service + Goods)', '', 'Insurance claims, co-pays, capitation arrangements']
  ];
  
  sheet.getRange(row, 1, industries.length, 4).setValues(industries);
  sheet.getRange(row, 1, 1, 4).setFontWeight('bold').setBackground('#e8f0fe');
  
  // Section 3: Key Terms & Definitions
  row += industries.length + 2;
  sheet.getRange(row, 1, 1, 4).merge()
    .setValue('KEY TERMS & DEFINITIONS')
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  
  row++;
  const terms = [
    ['Term', 'Definition', '', ''],
    ['Performance Obligation', 'Promise to transfer distinct good/service or series of distinct goods/services', '', ''],
    ['Transaction Price', 'Amount of consideration entity expects to be entitled to in exchange for transferring promised goods/services', '', ''],
    ['Contract Asset', 'Entity\'s right to consideration in exchange for goods/services that the entity has transferred to customer (conditional on something other than passage of time)', '', ''],
    ['Contract Liability', 'Entity\'s obligation to transfer goods/services to customer for which entity has received consideration', '', ''],
    ['Standalone Selling Price', 'Price at which entity would sell promised good/service separately to customer', '', ''],
    ['Control', 'Ability to direct use of and obtain substantially all remaining benefits from asset', '', ''],
    ['Input Method', 'Recognizes revenue on basis of entity\'s efforts toward satisfying performance obligation (e.g., costs incurred, labor hours)', '', ''],
    ['Output Method', 'Recognizes revenue on basis of direct measurements of value transferred to customer (e.g., units produced, milestones)', '', '']
  ];
  
  sheet.getRange(row, 1, terms.length, 4).setValues(terms);
  sheet.getRange(row, 1, 1, 4).setFontWeight('bold').setBackground('#e8f0fe');
  
  // Wrap text in definition column
  sheet.getRange(row, 2, terms.length, 1).setWrap(true);
  
  // Section 4: Audit Considerations
  row += terms.length + 2;
  sheet.getRange(row, 1, 1, 4).merge()
    .setValue('AUDIT CONSIDERATIONS & ASSERTIONS')
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  
  row++;
  const assertions = [
    ['Assertion', 'Audit Procedure', '', 'Evidence'],
    ['Occurrence/Existence', 'Verify contracts are valid and revenue transactions occurred', '', 'Signed contracts, customer confirmations, proof of delivery'],
    ['Completeness', 'Ensure all revenue is captured', '', 'Reconcile billing to revenue, review unbilled revenue'],
    ['Accuracy/Valuation', 'Verify revenue measured correctly per Ind AS 115', '', 'Recalculate transaction price, test allocations, review constraints'],
    ['Cut-off', 'Verify revenue recorded in correct period', '', 'Test transactions around period-end, review % completion'],
    ['Classification', 'Verify proper classification of contract balances', '', 'Review contract asset vs A/R, deferred revenue classification'],
    ['Presentation', 'Verify appropriate disclosure per Ind AS 115.110-129', '', 'Review financial statement disclosures, disaggregation'],
    ['Rights & Obligations', 'Verify entity has rights to consideration', '', 'Review contract terms, payment terms, cancellation clauses']
  ];
  
  sheet.getRange(row, 1, assertions.length, 4).setValues(assertions);
  sheet.getRange(row, 1, 1, 4).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange(row+1, 2, assertions.length-1, 1).setWrap(true);
  
  // Freeze
  sheet.setFrozenRows(3);
}

/**
 * ============================================================================
 * SHEET 11: AUDIT NOTES & CONTROL TOTALS
 * ============================================================================
 */
function createAuditNotesSheet(ss) {
  let sheet = ss.getSheetByName('Audit Notes');
  if (!sheet) {
    sheet = ss.insertSheet('Audit Notes', 10);
  } else {
    sheet.clear();
  }
  
  // Header
  sheet.getRange('A1:E1').merge()
    .setValue('AUDIT NOTES & CONTROL TOTALS - IND AS 115')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  sheet.setRowHeight(1, 40);
  
  // Set column widths
  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 300);
  
  // Section 1: Control Totals
  sheet.getRange('A3:E3').merge()
    .setValue('CONTROL TOTALS & RECONCILIATION')
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  
  const headers = ['#', 'Control Point', 'Calculated Total', 'GL Balance', 'Variance'];
  sheet.getRange('A4:E4').setValues([headers])
    .setFontWeight('bold')
    .setBackground('#e8f0fe');
  
  const controls = [
    [1, 'Total Contract Value (Current Period)', '=SUM(\'Contract Register\'!F:F)', '', ''],
    [2, 'Total Revenue Recognized (Period)', '=SUM(\'Revenue Recognition\'!H:H)', '', ''],
    [3, 'Total Revenue Recognized (YTD)', '=SUM(\'Revenue Recognition\'!L:L)', '', ''],
    [4, 'Total Contract Assets (Unbilled Revenue)', '=SUM(\'Contract Balances\'!F:F)', '', ''],
    [5, 'Total Contract Liabilities (Deferred Revenue)', '=SUM(\'Contract Balances\'!G:G)', '', ''],
    [6, 'Total Accounts Receivable', '=SUM(\'Contract Balances\'!J:J)', '', ''],
    [7, 'Total Remaining Performance Obligations', '=SUM(\'Performance Obligations\'!K:K)', '', ''],
    [8, 'Journal Entry Totals - Debits', '=SUM(\'Period-End Adjustments\'!G:G)', '', ''],
    [9, 'Journal Entry Totals - Credits', '=SUM(\'Period-End Adjustments\'!H:H)', '', ''],
    [10, 'JE Balance Check (should be zero)', '=ABS(C12-C13)', '', '']
  ];
  
  let row = 5;
  controls.forEach(item => {
    sheet.getRange(row, 1).setValue(item[0]);
    sheet.getRange(row, 2).setValue(item[1]);
    sheet.getRange(row, 3).setFormula(item[2]).setNumberFormat('#,##0.00').setBackground('#d9ead3');
    sheet.getRange(row, 4).setBackground('#cfe2f3').setNumberFormat('#,##0.00')
      .setNote('INPUT: Enter actual GL balance for comparison');
    sheet.getRange(row, 5).setFormula(`=IF(D${row}="","",C${row}-D${row})`).setNumberFormat('#,##0.00');
    row++;
  });
  
  // Conditional formatting for variances
  const varianceRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberNotEqualTo(0)
    .setBackground('#f4cccc')
    .setRanges([sheet.getRange('E5:E14')])
    .build();
  
  let rules = sheet.getConditionalFormatRules();
  rules.push(varianceRule);
  sheet.setConditionalFormatRules(rules);
  
  // Section 2: Key Review Points
  row += 2;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('KEY REVIEW POINTS & AUDIT PROCEDURES')
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  
  row++;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('Reviewer: Document completion of key audit procedures below')
    .setBackground('#fff3cd')
    .setFontStyle('italic');
  
  row++;
  const reviewHeaders = ['#', 'Review Point', 'Status', 'Reviewer Initials', 'Comments'];
  sheet.getRange(row, 1, 1, 5).setValues([reviewHeaders])
    .setFontWeight('bold')
    .setBackground('#e8f0fe');
  
  row++;
  const reviewPoints = [
    [1, 'Verified all contracts in register are valid and enforceable', '', '', ''],
    [2, 'Confirmed performance obligations are distinct per Ind AS 115.27', '', '', ''],
    [3, 'Reviewed transaction price allocation methodology', '', '', ''],
    [4, 'Tested revenue recognition calculations for accuracy', '', '', ''],
    [5, 'Verified contract asset/liability classifications', '', '', ''],
    [6, 'Reviewed variable consideration constraints', '', '', ''],
    [7, 'Tested cut-off procedures around period-end', '', '', ''],
    [8, 'Verified journal entries are complete and accurate', '', '', ''],
    [9, 'Reconciled workpaper totals to general ledger', '', '', ''],
    [10, 'Reviewed disclosure requirements per Ind AS 115.110-129', '', '', ''],
    [11, 'Assessed IGAAP to Ind AS transition adjustments', '', '', ''],
    [12, 'Obtained management representations', '', '', '']
  ];
  
  reviewPoints.forEach(item => {
    sheet.getRange(row, 1).setValue(item[0]);
    sheet.getRange(row, 2).setValue(item[1]).setWrap(true);
    sheet.getRange(row, 3).setBackground('#cfe2f3')
      .setNote('Enter: Complete / In Progress / Not Started / N/A');
    sheet.getRange(row, 4).setBackground('#cfe2f3');
    sheet.getRange(row, 5).setBackground('#cfe2f3');
    row++;
  });
  
  // Add data validation for Status column
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Complete', 'In Progress', 'Not Started', 'N/A'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(row - reviewPoints.length, 3, reviewPoints.length, 1).setDataValidation(statusRule);
  
  // Section 3: Significant Observations
  row += 2;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('SIGNIFICANT OBSERVATIONS & EXCEPTIONS')
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  
  row++;
  const obsHeaders = ['Date', 'Observation', 'Impact', 'Resolution', 'Status'];
  sheet.getRange(row, 1, 1, 5).setValues([obsHeaders])
    .setFontWeight('bold')
    .setBackground('#e8f0fe');
  
  row++;
  // Format 10 blank rows for observations
  for (let i = 0; i < 10; i++) {
    sheet.getRange(row + i, 1).setBackground('#cfe2f3').setNumberFormat('dd-mmm-yyyy');
    sheet.getRange(row + i, 2).setBackground('#cfe2f3').setWrap(true);
    sheet.getRange(row + i, 3).setBackground('#cfe2f3').setWrap(true);
    sheet.getRange(row + i, 4).setBackground('#cfe2f3').setWrap(true);
    sheet.getRange(row + i, 5).setBackground('#cfe2f3');
  }
  
  // Add status validation
  const obsStatusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Open', 'Closed', 'Pending', 'Noted'], true)
    .build();
  sheet.getRange(row, 5, 10, 1).setDataValidation(obsStatusRule);
  
  // Section 4: Sign-off
  row += 12;
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue('WORKPAPER SIGN-OFF')
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  
  row++;
  const signoffs = [
    ['Prepared by:', '', 'Date:', ''],
    ['Reviewed by:', '', 'Date:', ''],
    ['Approved by:', '', 'Date:', '']
  ];
  
  signoffs.forEach(item => {
    sheet.getRange(row, 1).setValue(item[0]).setFontWeight('bold');
    sheet.getRange(row, 2).setBackground('#cfe2f3');
    sheet.getRange(row, 3).setValue(item[2]).setFontWeight('bold');
    sheet.getRange(row, 4).setBackground('#cfe2f3').setNumberFormat('dd-mmm-yyyy');
    row++;
  });
  
  // Freeze
  sheet.setFrozenRows(4);
}

/**
 * ============================================================================
 * SETUP NAMED RANGES
 * ============================================================================
 */
function setupNamedRanges(ss) {
  try {
    // Key input ranges for easy reference in formulas
    ss.setNamedRange('ReportingPeriodStart', ss.getRange('Assumptions!C5'));
    ss.setNamedRange('ReportingPeriodEnd', ss.getRange('Assumptions!C6'));
    ss.setNamedRange('FunctionalCurrency', ss.getRange('Assumptions!C9'));
    ss.setNamedRange('GSTRate', ss.getRange('Assumptions!C26'));
    ss.setNamedRange('MaterialityThreshold', ss.getRange('Assumptions!C20'));
    
    SpreadsheetApp.getActiveSpreadsheet().toast('Named ranges created successfully', 'Setup', 2);
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Warning: Some named ranges could not be created', 'Setup', 3);
  }
}

/**
 * ============================================================================
 * FINAL FORMATTING & PROTECTION
 * ============================================================================
 */
function finalFormatting(ss) {
  // Set theme colors and final touches
  const allSheets = ss.getSheets();
  
  allSheets.forEach(sheet => {
    // Set grid lines
    sheet.setHiddenGridlines(false);
    
    // Protect formulas (optional - uncomment if needed)
    // const protection = sheet.protect().setDescription('Protect formula cells');
    // protection.setWarningOnly(true);
  });
  
  SpreadsheetApp.getActiveSpreadsheet().toast('Formatting complete!', 'Success', 2);
}

/**
 * ============================================================================
 * ON OPEN TRIGGER - Create menu automatically
 * ============================================================================
 */
// onOpen() is handled by common/utilities.gs

/**
 * ============================================================================
 * END OF SCRIPT
 * ============================================================================
 * 
 * USAGE NOTES FOR AUDITORS:
 * 
 * 1. INPUT CELLS: All blue-highlighted cells require user input
 * 2. FORMULA CELLS: All other cells with values are formula-driven - do not overwrite
 * 3. NAVIGATION: Use the custom menu "Ind AS 115 Navigator" for quick access
 * 4. MAINTENANCE: To add more contracts, simply fill rows in Contract Register
 * 5. CONTROL TOTALS: Always verify control totals in Audit Notes sheet
 * 6. SIGN-OFF: Complete review checklist in Audit Notes before finalizing
 * 
 * COMPLIANCE CHECKLIST:
 * ☑ Ind AS 115 5-step model implemented
 * ☑ Contract assets/liabilities properly classified
 * ☑ Variable consideration constraint applied
 * ☑ Performance obligations tracked separately
 * ☑ Period-end adjustment journal entries prepared
 * ☑ IGAAP vs Ind AS reconciliation documented
 * ☑ Audit trail and control totals maintained
 * ☑ Professional formatting and self-explanatory layout
 * 
 * For questions or enhancements, document in Audit Notes sheet.
 * ============================================================================
 */