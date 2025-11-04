/**
 * @name indas109
 * @version 1.1.0
 * @built 2025-11-04T10:11:10.767Z
 * @description Standalone script. Do not edit directly - edit source files in src/ folder.
 * 
 * This file is auto-generated from:
 * - Common utilities (src/common/*.gs)
 * - Workbook-specific code (src/workbooks/indas109.gs)
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
 * ═══════════════════════════════════════════════════════════════════════════
 * IGAAP-IND AS 109 AUDIT BUILDER
 * Financial Instruments - Period Book Closure Workings
 * ═══════════════════════════════════════════════════════════════════════════
 * 
 * PURPOSE: Automate creation of comprehensive Ind AS 109 compliant audit
 *          workings for period-end book closure entries covering:
 *          - Classification & Measurement
 *          - Fair Value Adjustments (FVTPL/FVOCI)
 *          - Expected Credit Loss (ECL) Impairment
 *          - Amortized Cost calculations
 *          - Journal Entries & Reconciliation
 * 
 * COMPLIANCE: Ind AS 109 - Financial Instruments
 * AUTHOR: IGAAP-Ind AS Audit Builder
 * VERSION: 1.0
 * 
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ═══════════════════════════════════════════════════════════════════════════
// MAIN EXECUTION FUNCTION
// ═══════════════════════════════════════════════════════════════════════════

function createIndAS109Workbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Set workbook type for menu detection
  setWorkbookType('INDAS109');
  
  // Clear existing sheets (optional - comment out if you want to keep existing)
  clearExistingSheets(ss);
  
  // Create all sheets in sequence
  Logger.log('Creating Ind AS 109 Working Papers...');
  
  createCoverSheet(ss);
  createInputVariablesSheet(ss);
  createInstrumentsRegisterSheet(ss);
  createClassificationMatrixSheet(ss);
  createFairValueWorkingsSheet(ss);
  createECLImpairmentSheet(ss);
  createAmortizationScheduleSheet(ss);
  createPeriodEndEntriesSheet(ss);
  createReconciliationSheet(ss);
  createReferencesSheet(ss);
  createAuditNotesSheet(ss);
  
  // Set up named ranges
  setupNamedRanges(ss);

  // Apply final professional formatting
  finalizeWorkingPapers(ss);

  // Activate Cover sheet
  ss.setActiveSheet(ss.getSheetByName('Cover'));

  SpreadsheetApp.getUi().alert(
    'Ind AS 109 Working Papers Created Successfully!',
    '✓ All sheets created with formulas\n' +
    '✓ Input cells marked in light blue\n' +
    '✓ Navigation buttons added\n' +
    '✓ Audit trail embedded\n\n' +
    'Start by filling the "Input_Variables" sheet.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  Logger.log('Ind AS 109 Working Papers creation completed.');
}

// ═══════════════════════════════════════════════════════════════════════════
// WORKBOOK-SPECIFIC CONFIGURATION
// ═══════════════════════════════════════════════════════════════════════════

// Column mappings for Ind AS 109 workbook
const COLS = {
  INSTRUMENTS_REGISTER: {
    ID: 1,
    NAME: 2,
    TYPE: 3,
    COUNTERPARTY: 4,
    ISSUE_DATE: 5,
    MATURITY_DATE: 6,
    FACE_VALUE: 7,
    COUPON_RATE: 8,
    EIR: 9,
    OPENING_BALANCE: 10,
    CURRENCY: 11,
    SECURITY_TYPE: 12,
    CREDIT_RATING: 13,
    DPD: 14,
    SPPI_TEST: 15,
    BUSINESS_MODEL: 16,
    DESIGNATED_FVTPL: 17,
    FVOCI_EQUITY: 18,
    COUPON_FREQ: 19,
    SIMPLIFIED_ECL: 20
  }
};

// Key row numbers
const ROWS = {
  INPUT_VARS: {
    REPORTING_DATE: 4,
    PREV_REPORTING_DATE: 5,
    RISK_FREE_RATE: 6,
    DAYS_IN_YEAR: 7,
    DAYS_IN_PERIOD: 8
  }
};

// ═══════════════════════════════════════════════════════════════════════════
// 1. COVER SHEET
// ═══════════════════════════════════════════════════════════════════════════

function createCoverSheet(ss) {
  const sheet = ss.insertSheet('Cover', 0);
  sheet.setTabColor('#1a237e');
  
  // Set column widths
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidths(2, 4, 180);
  
  // Title
  formatHeader(sheet, 2, 1, 5, 'IND AS 109 - FINANCIAL INSTRUMENTS', '#0d47a1');
  formatHeader(sheet, 3, 1, 5, 'Period Book Closure Working Papers', '#1565c0');
  
  // Company details section
  sheet.getRange('A5').setValue('Company Name:').setFontWeight('bold');
  formatInputCell(sheet.getRange('B5:D5').merge());
  
  sheet.getRange('A6').setValue('Financial Year:').setFontWeight('bold');
  formatInputCell(sheet.getRange('B6:D6').merge());
  
  sheet.getRange('A7').setValue('Period End Date:').setFontWeight('bold');
  formatInputCell(sheet.getRange('B7:D7').merge());
  formatDate(sheet.getRange('B7'));
  
  sheet.getRange('A8').setValue('Reporting Currency:').setFontWeight('bold');
  formatInputCell(sheet.getRange('B8:D8').merge().setValue('INR'));
  
  // Summary metrics section
  formatHeader(sheet, 10, 1, 5, 'KEY METRICS SUMMARY', '#283593');
  
  const metrics = [
    ['Metric', 'Opening Balance', 'Movements', 'Closing Balance', 'Reference'],
    ['Financial Assets at Amortized Cost', '=Reconciliation!B8', '=Reconciliation!C8', '=Reconciliation!D8', '=Reconciliation!E8'],
    ['Financial Assets at FVTPL', '=Reconciliation!B9', '=Reconciliation!C9', '=Reconciliation!D9', '=Reconciliation!E9'],
    ['Financial Assets at FVOCI', '=Reconciliation!B10', '=Reconciliation!C10', '=Reconciliation!D10', '=Reconciliation!E10'],
    ['Total Financial Assets', '=SUM(B12:B14)', '=SUM(C12:C14)', '=SUM(D12:D14)', ''],
    ['', '', '', '', ''],
    ['ECL Provision - Stage 1', '=Reconciliation!B17', '=Reconciliation!C17', '=Reconciliation!D17', '=Reconciliation!E17'],
    ['ECL Provision - Stage 2', '=Reconciliation!B18', '=Reconciliation!C18', '=Reconciliation!D18', '=Reconciliation!E18'],
    ['ECL Provision - Stage 3', '=Reconciliation!B19', '=Reconciliation!C19', '=Reconciliation!D19', '=Reconciliation!E19'],
    ['Total ECL Provision', '=SUM(B17:B19)', '=SUM(C17:C19)', '=SUM(D17:D19)', ''],
    ['', '', '', '', ''],
    ['Net Financial Assets', '=B15-B20', '=C15-C20', '=D15-D20', '']
  ];
  
  formatSubHeader(sheet, 11, 1, metrics[0], '#3949ab');
  
  let row = 12;
  metrics.slice(1).forEach(rowData => {
    rowData.forEach((value, col) => {
      sheet.getRange(row, col + 1).setValue(value);
    });
    if (rowData[0].includes('Total') || rowData[0].includes('Net')) {
      sheet.getRange(row, 1, 1, 5).setFontWeight('bold').setBackground('#e8eaf6');
    }
    row++;
  });
  
  formatCurrency(sheet.getRange('B12:D22'));
  
  // Navigation section
  formatHeader(sheet, 24, 1, 5, 'NAVIGATION & WORKFLOW', '#283593');
  
  const navigation = [
    ['Sheet Name', 'Purpose', 'Action'],
    ['Input_Variables', 'Master assumptions & parameters', 'START HERE'],
    ['Instruments_Register', 'List of all financial instruments', 'Review & Update'],
    ['Classification_Matrix', 'Ind AS 109 classification logic', 'Auto-populated'],
    ['Fair_Value_Workings', 'FVTPL & FVOCI adjustments', 'Review'],
    ['ECL_Impairment', 'Expected Credit Loss calculations', 'Review'],
    ['Amortization_Schedule', 'EIR-based amortization', 'Review'],
    ['Period_End_Entries', 'Journal entries for book closure', 'EXTRACT FOR POSTING'],
    ['Reconciliation', 'Opening to closing balances', 'Verify'],
    ['Audit_Notes', 'Control totals & assertions', 'Review']
  ];
  
  formatSubHeader(sheet, 25, 1, navigation[0], '#3949ab');
  
  row = 26;
  navigation.slice(1).forEach(rowData => {
    rowData.forEach((value, col) => {
      const cell = sheet.getRange(row, col + 1);
      cell.setValue(value);
      if (col === 2 && (value.includes('START') || value.includes('EXTRACT'))) {
        cell.setBackground('#fff3e0').setFontWeight('bold').setFontColor('#e65100');
      }
    });
    row++;
  });
  
  // Compliance note
  sheet.getRange('A36').setValue('COMPLIANCE NOTE:').setFontWeight('bold').setFontColor('#b71c1c');
  sheet.getRange('A37:E39').merge()
       .setValue('This workbook implements Ind AS 109 requirements for classification, measurement, and impairment of financial instruments. ' +
                 'All calculations follow SPPI (Solely Payments of Principal and Interest) and business model tests. ' +
                 'ECL impairment uses the three-stage approach as mandated by the standard.')
       .setWrap(true)
       .setVerticalAlignment('top')
       .setBorder(true, true, true, true, false, false, '#d32f2f', SpreadsheetApp.BorderStyle.SOLID);
  
  // Freeze header
  sheet.setFrozenRows(3);
  
  Logger.log('Cover sheet created.');
}

// ═══════════════════════════════════════════════════════════════════════════
// 2. INPUT VARIABLES SHEET
// ═══════════════════════════════════════════════════════════════════════════

function createInputVariablesSheet(ss) {
  const sheet = ss.insertSheet('Input_Variables');
  sheet.setTabColor('#4caf50');
  
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidths(2, 3, 150);
  sheet.setColumnWidth(5, 300);
  
  formatHeader(sheet, 1, 1, 5, 'INPUT VARIABLES & ASSUMPTIONS', '#2e7d32');
  
  sheet.getRange('A2:E2').setValues([['Variable Name', 'Value', 'Unit/Type', 'Valid Range', 'Notes']]);
  formatSubHeader(sheet, 2, 1, ['Variable Name', 'Value', 'Unit/Type', 'Valid Range', 'Notes'], '#388e3c');
  
  const inputs = [
    // Section: General Parameters
    ['GENERAL PARAMETERS', '', '', '', ''],
    ['Reporting Date', '', 'Date', '', 'Period end date for book closure'],
    ['Previous Reporting Date', '', 'Date', '', 'Previous period end date'],
    ['Risk-Free Rate', 0.0675, 'Percentage', '0% - 20%', 'G-Sec 10Y rate for discounting'],
    ['Days in Year', 365, 'Days', '365 or 360', 'Day count convention'],
    ['Days in Current Period', 365, 'Days', '30-365', 'CORRECTED: Actual days in reporting period (30=monthly, 90=quarterly, 365=annual)'],
    ['', '', '', '', ''],
    
    // Section: ECL Parameters
    ['ECL IMPAIRMENT PARAMETERS', '', '', '', ''],
    ['PD - Stage 1 (Performing)', 0.005, 'Percentage', '0% - 5%', 'Probability of Default - 12 month'],
    ['PD - Stage 2 (Underperforming)', 0.15, 'Percentage', '5% - 30%', 'Probability of Default - Lifetime'],
    ['PD - Stage 3 (NPA)', 0.85, 'Percentage', '50% - 100%', 'Probability of Default - Lifetime'],
    ['LGD - Secured', 0.25, 'Percentage', '10% - 40%', 'Loss Given Default - Secured assets'],
    ['LGD - Unsecured', 0.65, 'Percentage', '40% - 90%', 'Loss Given Default - Unsecured'],
    ['DPD Threshold Stage 2', 30, 'Days', '30 - 90', 'Days Past Due for Stage 2 transfer'],
    ['DPD Threshold Stage 3', 90, 'Days', '90 - 180', 'Days Past Due for Stage 3 (NPA)'],
    ['', '', '', '', ''],
    
    // Section: Fair Value Parameters
    ['FAIR VALUE PARAMETERS', '', '', '', ''],
    ['Equity Risk Premium', 0.08, 'Percentage', '5% - 12%', 'Market risk premium for equity valuation'],
    ['Credit Spread - AAA', 0.0025, 'Percentage', '0% - 1%', 'Credit spread for AAA rated instruments'],
    ['Credit Spread - AA', 0.005, 'Percentage', '0% - 2%', 'Credit spread for AA rated instruments'],
    ['Credit Spread - A', 0.01, 'Percentage', '0% - 3%', 'Credit spread for A rated instruments'],
    ['Credit Spread - BBB', 0.025, 'Percentage', '1% - 5%', 'Credit spread for BBB rated instruments'],
    ['FX Rate USD/INR', 83.25, 'Rate', '70 - 100', 'Foreign exchange rate if applicable'],
    ['', '', '', '', ''],
    
    // Section: Materiality Thresholds
    ['MATERIALITY & THRESHOLDS', '', '', '', ''],
    ['Materiality Threshold', 50000, 'Currency', '', 'Minimum amount for separate disclosure'],
    ['Rounding Factor', 1, 'Currency', '1, 1000, 100000', '1=Exact, 1000=Thousands, 100000=Lakhs']
  ];
  
  let row = 3;
  inputs.forEach(input => {
    if (input[0] === '' || input[0].includes('PARAMETERS') || input[0].includes('THRESHOLDS')) {
      // Section header
      if (input[0] !== '') {
        sheet.getRange(row, 1, 1, 5).merge()
             .setValue(input[0])
             .setBackground('#66bb6a')
             .setFontColor('#ffffff')
             .setFontWeight('bold')
             .setHorizontalAlignment('left');
      }
    } else {
      // Data row
      sheet.getRange(row, 1).setValue(input[0]);
      
      const valueCell = sheet.getRange(row, 2);
      valueCell.setValue(input[1]);
      formatInputCell(valueCell, '#e1f5e1');
      
      sheet.getRange(row, 3).setValue(input[2]);
      sheet.getRange(row, 4).setValue(input[3]).setFontStyle('italic').setFontColor('#666666');
      sheet.getRange(row, 5).setValue(input[4]).setWrap(true);
      
      // Format based on type
      if (input[2] === 'Date') {
        formatDate(valueCell);
      } else if (input[2] === 'Percentage') {
        formatPercentage(valueCell);
      } else if (input[2] === 'Currency') {
        formatCurrency(valueCell);
      }
    }
    row++;
  });
  
  // Add data validation for dates
  sheet.getRange('B4').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText('Enter the period end date')
    .build());
  
  sheet.getRange('B5').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText('Enter the previous period end date')
    .build());
  
  // Instructions
  sheet.getRange('A' + (row + 2)).setValue('INSTRUCTIONS:').setFontWeight('bold').setFontColor('#2e7d32');
  sheet.getRange('A' + (row + 3) + ':E' + (row + 5)).merge()
       .setValue('1. Fill all light green cells with appropriate values\n' +
                 '2. Ensure dates are in DD-MMM-YYYY format\n' +
                 '3. Percentages should be entered as decimals (e.g., 5% = 0.05)\n' +
                 '4. All other sheets will auto-calculate based on these inputs\n' +
                 '5. Review "Valid Range" column to ensure inputs are reasonable')
       .setWrap(true)
       .setVerticalAlignment('top');
  
  sheet.setFrozenRows(2);
  
  Logger.log('Input Variables sheet created.');
}

// ═══════════════════════════════════════════════════════════════════════════
// 3. INSTRUMENTS REGISTER SHEET
// ═══════════════════════════════════════════════════════════════════════════

function createInstrumentsRegisterSheet(ss) {
  const sheet = ss.insertSheet('Instruments_Register');
  sheet.setTabColor('#ff9800');
  
  // Set column widths
  sheet.setColumnWidth(1, 80);  // ID
  sheet.setColumnWidth(2, 200); // Instrument Name
  sheet.setColumnWidth(3, 120); // Type
  sheet.setColumnWidth(4, 120); // Counterparty
  sheet.setColumnWidth(5, 100); // Issue Date
  sheet.setColumnWidth(6, 100); // Maturity Date
  sheet.setColumnWidths(7, 3, 120); // Face Value, Coupon Rate, EIR
  sheet.setColumnWidth(10, 120); // Current Balance
  sheet.setColumnWidth(11, 100); // Currency
  sheet.setColumnWidth(12, 100); // Security
  sheet.setColumnWidth(13, 100); // Credit Rating
  sheet.setColumnWidth(14, 100); // DPD
  sheet.setColumnWidth(15, 150); // SPPI Test
  sheet.setColumnWidth(16, 150); // Business Model
  sheet.setColumnWidth(17, 130); // Designated at FVTPL
  sheet.setColumnWidth(18, 130); // FVOCI Equity Election
  sheet.setColumnWidth(19, 120); // Coupon Frequency
  sheet.setColumnWidth(20, 120); // Simplified ECL

  formatHeader(sheet, 1, 1, 20, 'FINANCIAL INSTRUMENTS REGISTER', '#e65100');

  const headers = [
    'ID',
    'Instrument Name',
    'Type',
    'Counterparty',
    'Issue Date',
    'Maturity Date',
    'Face Value (₹)',
    'Coupon Rate',
    'EIR',
    'Opening Balance (₹)',
    'Currency',
    'Security Type',
    'Credit Rating',
    'DPD (Days)',
    'SPPI Test Result',
    'Business Model',
    'Designated at FVTPL',
    'FVOCI Equity Election',
    'Coupon Frequency',
    'Simplified ECL'
  ];
  
  formatSubHeader(sheet, 2, 1, headers, '#f57c00');
  
  // Sample data with formulas
  const sampleData = [
    ['FI001', 'Term Loan - ABC Ltd', 'Loan', 'ABC Limited', '=DATE(2023,4,1)', '=DATE(2028,3,31)', 10000000, 0.09, 0.095, 9500000, 'INR', 'Secured', 'AA', 0, 'Pass', 'Hold to Collect', 'No', 'No', 'Annual', 'No'],
    ['FI002', 'Corporate Bond - XYZ Corp', 'Bond', 'XYZ Corporation', '=DATE(2022,1,15)', '=DATE(2027,1,14)', 5000000, 0.085, 0.088, 4950000, 'INR', 'Unsecured', 'A', 0, 'Pass', 'Hold to Collect', 'No', 'No', 'Semi-Annual', 'No'],
    ['FI003', 'Equity - TechCo', 'Equity', 'TechCo Ltd', '=DATE(2021,6,1)', '', 2000000, 0, 0, 2500000, 'INR', 'Equity', 'Not Rated', 0, 'Fail', 'Other (Trading)', 'No', 'No', 'Not Applicable', 'N/A'],
    ['FI004', 'Trade Receivable - Client A', 'Receivable', 'Client A', '=DATE(2024,10,1)', '=DATE(2025,1,31)', 750000, 0, 0.12, 750000, 'INR', 'Unsecured', 'BBB', 15, 'Pass', 'Hold to Collect', 'No', 'No', 'Not Applicable', 'Yes'],
    ['FI005', 'Govt Security - 10Y', 'G-Sec', 'Government of India', '=DATE(2020,7,1)', '=DATE(2030,6,30)', 3000000, 0.0675, 0.068, 3050000, 'INR', 'Sovereign', 'AAA', 0, 'Pass', 'Hold to Collect & Sell', 'No', 'No', 'Semi-Annual', 'No'],
    ['FI006', 'Mutual Fund Units', 'Mutual Fund', 'HDFC Balanced Fund', '=DATE(2023,3,1)', '', 1000000, 0, 0, 1150000, 'INR', 'Units', 'Not Rated', 0, 'Fail', 'Other (Trading)', 'No', 'No', 'Not Applicable', 'N/A'],
    ['FI007', 'Loan - Stressed Account', 'Loan', 'DEF Enterprises', '=DATE(2021,9,1)', '=DATE(2026,8,31)', 2000000, 0.11, 0.115, 1800000, 'INR', 'Secured', 'B', 120, 'Pass', 'Hold to Collect', 'No', 'No', 'Quarterly', 'No']
  ];
  
  let row = 3;
  sampleData.forEach(data => {
    data.forEach((value, col) => {
      const cell = sheet.getRange(row, col + 1);
      if (typeof value === 'string' && value.startsWith('=')) {
        cell.setFormula(value);
      } else {
        cell.setValue(value);
      }
    });
    
    // Format input cells (most cells are inputs)
    formatInputCell(sheet.getRange(row, 2, 1, 19), '#fff3e0');

    row++;
  });

  // Format columns
  formatDate(sheet.getRange('E3:F' + (row - 1)));
  formatCurrency(sheet.getRange('G3:G' + (row - 1)));
  formatPercentage(sheet.getRange('H3:I' + (row - 1)));
  formatCurrency(sheet.getRange('J3:J' + (row - 1)));
  
  // Add data validation for dropdowns
  const typeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Loan', 'Bond', 'Debenture', 'Equity', 'Mutual Fund', 'G-Sec', 'T-Bill', 'Receivable', 'Derivative', 'Other'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange('C3:C250').setDataValidation(typeValidation);
  
  const securityValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Secured', 'Unsecured', 'Equity', 'Sovereign', 'Units', 'Not Applicable'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange('L3:L250').setDataValidation(securityValidation);
  
  const ratingValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['AAA', 'AA+', 'AA', 'AA-', 'A+', 'A', 'A-', 'BBB+', 'BBB', 'BBB-', 'BB', 'B', 'C', 'D', 'Not Rated'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange('M3:M250').setDataValidation(ratingValidation);
  
  const sppiValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pass', 'Fail', 'Not Applicable'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange('O3:O250').setDataValidation(sppiValidation);
  
  const businessModelValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Hold to Collect', 'Hold to Collect & Sell', 'Other (Trading)'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange('P3:P250').setDataValidation(businessModelValidation);

  // Add "Designated at FVTPL" validation (Q column)
  const fvtplDesignationValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange('Q3:Q250').setDataValidation(fvtplDesignationValidation);

  // Add "FVOCI Equity Election" validation (R column)
  const fvociEquityValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange('R3:R250').setDataValidation(fvociEquityValidation);

  // Add "Coupon Frequency" validation (S column)
  const couponFrequencyValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Annual', 'Semi-Annual', 'Quarterly', 'Monthly', 'Not Applicable'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange('S3:S250').setDataValidation(couponFrequencyValidation);

  // Add "Simplified ECL" validation (T column)
  const simplifiedECLValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No', 'N/A'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange('T3:T250').setDataValidation(simplifiedECLValidation);

  // Add notes
  sheet.getRange('A' + (row + 2)).setValue('INSTRUCTIONS:').setFontWeight('bold').setFontColor('#e65100');
  sheet.getRange('A' + (row + 3) + ':T' + (row + 6)).merge()
       .setValue('• Add all financial instruments held as of reporting date\n' +
                 '• SPPI Test: "Pass" if cash flows are solely payments of principal and interest\n' +
                 '• Business Model: Select "Hold to Collect", "Hold to Collect & Sell", or "Other (Trading)"\n' +
                 '• Designated at FVTPL: Use "Yes" only when irrevocably designating to eliminate accounting mismatch (fair value option)\n' +
                 '• FVOCI Equity Election: "Yes" for equity investments irrevocably designated at FVOCI (no P&L recycling on disposal, per Ind AS 109.5.7.5)\n' +
                 '• Coupon Frequency: Select frequency for interest/coupon payments (affects interim period cash flow calculations)\n' +
                 '• Simplified ECL: "Yes" for trade receivables without significant financing component (lifetime ECL from day 1, per Ind AS 109.5.5.15)\n' +
                 '• DPD (Days Past Due): Enter 0 if current, otherwise number of days overdue\n' +
                 '• All fields in orange background are required for accurate classification\n' +
                 '• EIR should include all transaction costs and fees for amortized cost instruments')
       .setWrap(true)
       .setVerticalAlignment('top')
       .setBackground('#fff3e0')
       .setBorder(true, true, true, true, false, false);

  sheet.setFrozenRows(2);
  // Note: setFrozenColumns removed to avoid conflict with merged header cells

  Logger.log('Instruments Register sheet created.');
}

// ═══════════════════════════════════════════════════════════════════════════
// 4. CLASSIFICATION MATRIX SHEET
// ═══════════════════════════════════════════════════════════════════════════

function createClassificationMatrixSheet(ss) {
  const sheet = ss.insertSheet('Classification_Matrix');
  sheet.setTabColor('#9c27b0');
  
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 130);
  sheet.setColumnWidth(7, 150);
  sheet.setColumnWidth(8, 180);
  sheet.setColumnWidth(9, 250);

  formatHeader(sheet, 1, 1, 9, 'IND AS 109 CLASSIFICATION MATRIX', '#6a1b9a');

  const headers = [
    'ID',
    'Instrument Name',
    'SPPI Test',
    'Business Model',
    'Designated at FVTPL',
    'FVOCI Equity Election',
    'Classification',
    'Measurement',
    'Ind AS 109 Reference'
  ];
  
  formatSubHeader(sheet, 2, 1, headers, '#7b1fa2');
  
  // Classification logic formulas
  // OPTIMIZED: Batch operations to reduce API calls from ~900 to ~2
  const numRows = 100;
  const startRow = 3;
  const numCols = 9;

  // Create 2D array for batch operations
  const formulaArray = [];

  // Template formulas for row 3
  const formulaTemplate = [
    '=IF(ISBLANK(Instruments_Register!A3),"",Instruments_Register!A3)',
    '=IF(ISBLANK(Instruments_Register!A3),"",Instruments_Register!B3)',
    '=IF(ISBLANK(Instruments_Register!A3),"",Instruments_Register!O3)',
    '=IF(ISBLANK(Instruments_Register!A3),"",Instruments_Register!P3)',
    '=IF(ISBLANK(Instruments_Register!A3),"",Instruments_Register!Q3)',
    '=IF(ISBLANK(Instruments_Register!A3),"",Instruments_Register!R3)',
    '=IF(ISBLANK(Instruments_Register!A3),"",IF(E3="Yes","FVTPL",IF(F3="Yes","FVOCI",IF(C3="Fail","FVTPL",IF(AND(C3="Pass",D3="Hold to Collect"),"Amortized Cost",IF(AND(C3="Pass",D3="Hold to Collect & Sell"),"FVOCI","FVTPL"))))))',
    '=IF(ISBLANK(Instruments_Register!A3),"",IF(G3="Amortized Cost","Amortized Cost using EIR",IF(G3="FVOCI",IF(F3="Yes","Fair Value - OCI (Equity - no recycling)","Fair Value - OCI (Debt - recycling)"),IF(G3="FVTPL","Fair Value through P&L","Review Required"))))',
    '=IF(ISBLANK(Instruments_Register!A3),"",IF(G3="Amortized Cost","Para 4.1.2 - AC if SPPI passed & HTC",IF(G3="FVOCI",IF(F3="Yes","Para 5.7.5 - FVOCI equity (irrevocable election)","Para 4.1.2A - FVOCI debt (SPPI passed & HTC&S)"),IF(G3="FVTPL","Para 4.1.4 - Default FVTPL or fair value option",""))))'
  ];

  // Populate array in memory (fast)
  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    const rowFormulas = formulaTemplate.map(formula => formula.replace(/3/g, row.toString()));
    formulaArray.push(rowFormulas);
  }

  // Write all formulas in batch (fast - only 1 API call!)
  sheet.getRange(startRow, 1, numRows, numCols).setFormulas(formulaArray);

  // Conditional formatting for classification (reduced range for performance)
  const classificationRange = sheet.getRange('G3:G250');

  // Amortized Cost - Green
  const acRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Amortized Cost')
    .setBackground('#c8e6c9')
    .setFontColor('#2e7d32')
    .setRanges([classificationRange])
    .build();

  // FVOCI - Blue
  const fvociRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('FVOCI')
    .setBackground('#bbdefb')
    .setFontColor('#1565c0')
    .setRanges([classificationRange])
    .build();

  // FVTPL - Orange
  const fvtplRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('FVTPL')
    .setBackground('#ffe0b2')
    .setFontColor('#e65100')
    .setRanges([classificationRange])
    .build();
  
  const rules = sheet.getConditionalFormatRules();
  rules.push(acRule, fvociRule, fvtplRule);
  sheet.setConditionalFormatRules(rules);
  
  // Summary section
  const summaryRow = 13;
  formatHeader(sheet, summaryRow, 1, 9, 'CLASSIFICATION SUMMARY', '#7b1fa2');

  sheet.getRange(summaryRow + 1, 1, 1, 2).merge().setValue('Classification Category').setFontWeight('bold');
  sheet.getRange(summaryRow + 1, 3).setValue('Count').setFontWeight('bold');
  sheet.getRange(summaryRow + 1, 4).setValue('Total Balance (₹)').setFontWeight('bold');
  sheet.getRange(summaryRow + 1, 5).setValue('% of Total').setFontWeight('bold');

  const summaryData = [
    ['Amortized Cost', '=COUNTIF(G3:G250,"Amortized Cost")', '=SUMIF(G3:G250,"Amortized Cost",Instruments_Register!J3:J250)', '=D15/D18'],
    ['FVOCI', '=COUNTIF(G3:G250,"FVOCI")', '=SUMIF(G3:G250,"FVOCI",Instruments_Register!J3:J250)', '=D16/D18'],
    ['FVTPL', '=COUNTIF(G3:G250,"FVTPL")', '=SUMIF(G3:G250,"FVTPL",Instruments_Register!J3:J250)', '=D17/D18'],
    ['Total', '=SUM(C15:C17)', '=SUM(D15:D17)', '1']
  ];
  
  let row = summaryRow + 2;
  summaryData.forEach(data => {
    sheet.getRange(row, 1, 1, 2).merge().setValue(data[0]);
    sheet.getRange(row, 3).setFormula(data[1]);
    sheet.getRange(row, 4).setFormula(data[2]);
    sheet.getRange(row, 5).setFormula(data[3]);
    
    if (data[0] === 'Total') {
      sheet.getRange(row, 1, 1, 5).setFontWeight('bold').setBackground('#e1bee7');
    }
    row++;
  });
  
  formatCurrency(sheet.getRange('D15:D18'));
  formatPercentage(sheet.getRange('E15:E18'));
  
  // Add explanation
  sheet.getRange('A' + (row + 2)).setValue('CLASSIFICATION LOGIC:').setFontWeight('bold').setFontColor('#6a1b9a');
  sheet.getRange('A' + (row + 3) + ':I' + (row + 7)).merge()
       .setValue('Ind AS 109 Classification Decision Tree (Corrected per Ind AS 109):\n\n' +
                 '1. Designated at FVTPL = "Yes" → FVTPL (fair value option to eliminate accounting mismatch)\n' +
                 '2. FVOCI Equity Election = "Yes" → FVOCI (equity - no recycling on disposal per Para 5.7.5)\n' +
                 '3. SPPI Test Fails → FVTPL (mandatory)\n' +
                 '4. SPPI Pass + Business Model "Hold to Collect" → Amortized Cost\n' +
                 '5. SPPI Pass + Business Model "Hold to Collect & Sell" → FVOCI (debt - recycling)\n' +
                 '6. Default → FVTPL\n\n' +
                 'SPPI = Solely Payments of Principal and Interest\n' +
                 'Note: FVTPL is not a business model; it is a measurement category. Business models are: Hold to Collect, Hold to Collect & Sell, or Other (Trading).')
       .setWrap(true)
       .setVerticalAlignment('top')
       .setBackground('#f3e5f5')
       .setBorder(true, true, true, true, false, false);
  
  sheet.setFrozenRows(2);
  
  Logger.log('Classification Matrix sheet created.');
}

// ═══════════════════════════════════════════════════════════════════════════
// 5. FAIR VALUE WORKINGS SHEET
// ═══════════════════════════════════════════════════════════════════════════

function createFairValueWorkingsSheet(ss) {
  const sheet = ss.insertSheet('Fair_Value_Workings');
  sheet.setTabColor('#00bcd4');
  
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 130);
  sheet.setColumnWidth(7, 130);
  sheet.setColumnWidth(8, 150);
  sheet.setColumnWidth(9, 150);
  
  formatHeader(sheet, 1, 1, 9, 'FAIR VALUE ADJUSTMENTS - FVTPL & FVOCI', '#0097a7');
  
  const headers = [
    'ID',
    'Instrument Name',
    'Classification',
    'Opening Balance (₹)',
    'Fair Value - Period End (₹)',
    'Fair Value Gain/(Loss) (₹)',
    'Impact on P&L (₹)',
    'Impact on OCI (₹)',
    'Level in Fair Value Hierarchy'
  ];
  
  formatSubHeader(sheet, 2, 1, headers, '#00acc1');
  
  // Formulas for fair value instruments only (extended to 100 rows)
  // OPTIMIZED: Batch operations to reduce API calls from ~900 to ~5
  const numRows = 100;
  const startRow = 3;
  const numCols = 9;

  // Create 2D arrays for batch operations
  const formulaArray = [];
  const backgroundArray = Array(numRows).fill(null).map(() => Array(numCols).fill(null));

  // Populate arrays in memory (fast)
  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    const formulas = [
      `=IF(ISBLANK(Instruments_Register!A${row}),"",Instruments_Register!A${row})`,
      `=IF(ISBLANK(Instruments_Register!A${row}),"",Instruments_Register!B${row})`,
      `=IF(ISBLANK(Instruments_Register!A${row}),"",Classification_Matrix!G${row})`,
      `=IF(ISBLANK(Instruments_Register!A${row}),"",Instruments_Register!J${row})`,
      // Fair Value calculation
      // CORRECTED: Removed RANDBETWEEN placeholder. Fair value should be based on:
      // - Market prices for listed securities, OR
      // - Valuation models (DCF, comparable multiples, etc.), OR
      // - Manual override input by user
      // Default: Opening balance (to be updated by user with actual fair value)
      `=IF(ISBLANK(Instruments_Register!A${row}),"",IF(OR(C${row}="FVTPL",C${row}="FVOCI"),D${row},0))`,
      `=IF(ISBLANK(Instruments_Register!A${row}),"",IF(OR(C${row}="FVTPL",C${row}="FVOCI"),E${row}-D${row},0))`,
      `=IF(ISBLANK(Instruments_Register!A${row}),"",IF(C${row}="FVTPL",F${row},0))`,
      `=IF(ISBLANK(Instruments_Register!A${row}),"",IF(C${row}="FVOCI",F${row},0))`,
      `=IF(ISBLANK(Instruments_Register!A${row}),"",IF(OR(C${row}="FVTPL",C${row}="FVOCI"),"Level 2 - Observable Inputs","-"))`
    ];

    formulaArray.push(formulas);

    // Mark Fair Value (Column E, index 4) as input cell for manual override
    backgroundArray[i][4] = '#e0f2f1';
  }

  // Write all formulas and backgrounds in batch (fast - only 2 API calls!)
  const dataRange = sheet.getRange(startRow, 1, numRows, numCols);
  dataRange.setFormulas(formulaArray);
  dataRange.setBackgrounds(backgroundArray);

  // Batch format currency ranges (reduced from :1000 to :250 for performance)
  formatCurrency(sheet.getRange('D3:H250'));
  
  // Conditional formatting for gains/losses (reduced range for performance)
  const gainLossRange = sheet.getRange('F3:H250');

  const gainRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#c8e6c9')
    .setFontColor('#2e7d32')
    .setRanges([gainLossRange])
    .build();

  const lossRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground('#ffcdd2')
    .setFontColor('#c62828')
    .setRanges([gainLossRange])
    .build();

  // Conditional formatting for stale fair values (Period End = Opening and Opening > 0)
  // This highlights cells that need user attention for fair value updates
  const staleFVRange = sheet.getRange('E3:E250');
  const staleFVRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(E3=D3,D3>0,NOT(ISBLANK(D3)))')
    .setBackground('#fff9c4')
    .setFontColor('#f57f17')
    .setRanges([staleFVRange])
    .build();

  const rules = sheet.getConditionalFormatRules();
  rules.push(gainRule, lossRule, staleFVRule);
  sheet.setConditionalFormatRules(rules);
  
  // Summary section
  const summaryRow = 13;
  formatHeader(sheet, summaryRow, 1, 9, 'FAIR VALUE MOVEMENTS SUMMARY', '#00acc1');
  
  sheet.getRange(summaryRow + 1, 1, 1, 3).merge().setValue('Category').setFontWeight('bold');
  sheet.getRange(summaryRow + 1, 4).setValue('Total Gain/(Loss)').setFontWeight('bold');
  sheet.getRange(summaryRow + 1, 5).setValue('To P&L').setFontWeight('bold');
  sheet.getRange(summaryRow + 1, 6).setValue('To OCI').setFontWeight('bold');
  
  const summaryData = [
    ['FVTPL Instruments', '=SUMIF(C3:C250,"FVTPL",F3:F250)', '=SUMIF(C3:C250,"FVTPL",G3:G250)', '0'],
    ['FVOCI Instruments', '=SUMIF(C3:C250,"FVOCI",F3:F250)', '0', '=SUMIF(C3:C250,"FVOCI",H3:H250)'],
    ['Total Fair Value Movement', '=SUM(D15:D16)', '=SUM(E15:E16)', '=SUM(F15:F16)']
  ];
  
  let row = summaryRow + 2;
  summaryData.forEach(data => {
    sheet.getRange(row, 1, 1, 3).merge().setValue(data[0]);
    sheet.getRange(row, 4).setFormula(data[1]);
    sheet.getRange(row, 5).setFormula(data[2]);
    sheet.getRange(row, 6).setFormula(data[3]);
    
    if (data[0].includes('Total')) {
      sheet.getRange(row, 1, 1, 6).setFontWeight('bold').setBackground('#b2ebf2');
    }
    row++;
  });
  
  formatCurrency(sheet.getRange('D15:F17'));
  
  // Input section for manual fair values
  formatHeader(sheet, row + 2, 1, 9, 'MANUAL FAIR VALUE INPUTS (Override auto-calculation if needed)', '#00acc1');
  
  sheet.getRange(row + 3, 1).setValue('ID');
  sheet.getRange(row + 3, 2).setValue('Instrument Name');
  sheet.getRange(row + 3, 3).setValue('Market Quote / Valuation');
  sheet.getRange(row + 3, 4).setValue('Valuation Date');
  sheet.getRange(row + 3, 5).setValue('Source / Basis');
  
  formatSubHeader(sheet, row + 3, 1, ['ID', 'Instrument Name', 'Market Quote / Valuation', 'Valuation Date', 'Source / Basis'], '#00acc1');
  
  // Add a few input rows
  for (let i = 0; i < 5; i++) {
    formatInputCell(sheet.getRange(row + 4 + i, 3), '#e0f7fa');
    formatInputCell(sheet.getRange(row + 4 + i, 4), '#e0f7fa');
    formatInputCell(sheet.getRange(row + 4 + i, 5), '#e0f7fa');
    formatDate(sheet.getRange(row + 4 + i, 4));
  }
  
  // Add notes
  sheet.getRange('A' + (row + 10)).setValue('VALUATION NOTES:').setFontWeight('bold').setFontColor('#0097a7');
  sheet.getRange('A' + (row + 11) + ':I' + (row + 14)).merge()
       .setValue('Fair Value Hierarchy (Ind AS 113):\n' +
                 '• Level 1: Quoted prices in active markets for identical assets\n' +
                 '• Level 2: Observable inputs other than Level 1 prices (e.g., interest rates, yield curves)\n' +
                 '• Level 3: Unobservable inputs (requires significant management judgment)\n\n' +
                 'FVTPL: All fair value changes recognized in P&L immediately\n' +
                 'FVOCI: Fair value changes in OCI; reclassified to P&L on derecognition (debt) or never reclassified (equity)')
       .setWrap(true)
       .setVerticalAlignment('top')
       .setBackground('#e0f7fa')
       .setBorder(true, true, true, true, false, false);
  
  sheet.setFrozenRows(2);
  
  Logger.log('Fair Value Workings sheet created.');
}

// ═══════════════════════════════════════════════════════════════════════════
// 6. ECL IMPAIRMENT SHEET
// ═══════════════════════════════════════════════════════════════════════════

function createECLImpairmentSheet(ss) {
  const sheet = ss.insertSheet('ECL_Impairment');
  sheet.setTabColor('#f44336');
  
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidths(6, 4, 120);
  sheet.setColumnWidth(10, 130);
  sheet.setColumnWidth(11, 130);
  sheet.setColumnWidth(12, 130);
  
  formatHeader(sheet, 1, 1, 12, 'EXPECTED CREDIT LOSS (ECL) IMPAIRMENT - THREE STAGE MODEL', '#c62828');
  
  const headers = [
    'ID',
    'Instrument Name',
    'DPD',
    'Gross Carrying Amount (₹)',
    'ECL Stage',
    'PD (%)',
    'LGD (%)',
    'EAD (₹)',
    'ECL Amount (₹)',
    'Opening Provision (₹)',
    'Movement (₹)',
    'Closing Provision (₹)'
  ];
  
  formatSubHeader(sheet, 2, 1, headers, '#d32f2f');
  
  // ECL calculation formulas (extended to 100 rows)
  // OPTIMIZED: Batch operations to reduce API calls from ~1300 to ~5
  const numRows = 100;
  const startRow = 3;
  const numCols = 12;

  // Create 2D arrays for batch operations
  const formulaArray = [];
  const backgroundArray = Array(numRows).fill(null).map(() => Array(numCols).fill(null));

  // Populate arrays in memory (fast)
  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    const formulas = [
      `=IF(ISBLANK(Instruments_Register!A${row}),"",Instruments_Register!A${row})`,
      `=IF(ISBLANK(Instruments_Register!A${row}),"",Instruments_Register!B${row})`,
      `=IF(ISBLANK(Instruments_Register!A${row}),"",Instruments_Register!N${row})`,
      `=IF(ISBLANK(Instruments_Register!A${row}),"",Instruments_Register!J${row})`,
      // Stage determination based on DPD, classification, and simplified ECL approach
      // CORRECTED: Implements simplified approach per Ind AS 109.5.5.15 for trade receivables
      // Simplified approach = Lifetime ECL from day 1 (no Stage 1/12-month bucket)
      // B$16 = DPD Threshold Stage 3 (90 days), B$15 = DPD Threshold Stage 2 (30 days)
      `=IF(ISBLANK(Instruments_Register!A${row}),"",IF(Classification_Matrix!F${row}="Amortized Cost",IF(OR(Instruments_Register!S${row}="Yes",Instruments_Register!C${row}="Receivable"),IF(C${row}>=Input_Variables!$B$16,"Stage 3","Simplified (Lifetime)"),IF(C${row}>=Input_Variables!$B$16,"Stage 3",IF(C${row}>=Input_Variables!$B$15,"Stage 2","Stage 1"))),"N/A"))`,
      // PD based on stage
      // Simplified (Lifetime) uses lifetime PD (same as Stage 2)
      `=IF(ISBLANK(Instruments_Register!A${row}),"",IF(E${row}="Stage 1",Input_Variables!$B$10,IF(OR(E${row}="Stage 2",E${row}="Simplified (Lifetime)"),Input_Variables!$B$11,IF(E${row}="Stage 3",Input_Variables!$B$12,0))))`,
      // LGD based on security type
      `=IF(ISBLANK(Instruments_Register!A${row}),"",IF(Instruments_Register!L${row}="Secured",Input_Variables!$B$13,Input_Variables!$B$14))`,
      // EAD (Exposure at Default) - typically equal to carrying amount
      `=IF(ISBLANK(Instruments_Register!A${row}),"",D${row})`,
      // ECL = EAD × PD × LGD discounted using effective interest rate
      `=IF(ISBLANK(Instruments_Register!A${row}),"",IF(E${row}<>"N/A",LET(discountRate,IF(Instruments_Register!I${row}="",RiskFreeRate,Instruments_Register!I${row}),remainingYears,IF(E${row}="Stage 1",1,IF(Instruments_Register!F${row}="",1,MAX(0,(Instruments_Register!F${row}-ReportingDate)/Input_Variables!$B$7))),discountFactor,IF(discountRate=0,1,(1+discountRate)^(-remainingYears)),H${row}*F${row}*G${row}*discountFactor),0))`,
      // Opening provision - INPUT CELL (should be closing provision from prior period)
      // CORRECTED: Changed from arbitrary 50% to input cell for proper period-to-period continuity
      // Users must enter opening ECL provision from prior period's closing balance
      `=IF(ISBLANK(Instruments_Register!A${row}),"",0)`,
      // Movement
      `=IF(ISBLANK(Instruments_Register!A${row}),"",I${row}-J${row})`,
      // Closing provision
      `=IF(ISBLANK(Instruments_Register!A${row}),"",I${row})`
    ];

    formulaArray.push(formulas);

    // Mark Opening Provision (Column J, index 9) as input cell
    backgroundArray[i][9] = '#ffebee';
  }

  // Write all formulas and backgrounds in batch (fast - only 2 API calls!)
  const dataRange = sheet.getRange(startRow, 1, numRows, numCols);
  dataRange.setFormulas(formulaArray);
  dataRange.setBackgrounds(backgroundArray);

  // Batch format entire column ranges (reduced from :1000 to :250 for performance)
  formatCurrency(sheet.getRange('D3:D250'));
  formatPercentage(sheet.getRange('F3:G250'));
  formatCurrency(sheet.getRange('H3:L250'));

  // Conditional formatting for stages (reduced range for performance)
  const stageRange = sheet.getRange('E3:E250');
  
  const stage1Rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Stage 1')
    .setBackground('#c8e6c9')
    .setFontColor('#2e7d32')
    .setRanges([stageRange])
    .build();
  
  const stage2Rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Stage 2')
    .setBackground('#fff9c4')
    .setFontColor('#f57f17')
    .setRanges([stageRange])
    .build();
  
  const stage3Rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Stage 3')
    .setBackground('#ffcdd2')
    .setFontColor('#c62828')
    .setRanges([stageRange])
    .build();

  const simplifiedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Simplified (Lifetime)')
    .setBackground('#e1bee7')
    .setFontColor('#6a1b9a')
    .setRanges([stageRange])
    .build();

  const rules = sheet.getConditionalFormatRules();
  rules.push(stage1Rule, stage2Rule, stage3Rule, simplifiedRule);
  sheet.setConditionalFormatRules(rules);
  
  // Summary section
  const summaryRow = 13;
  formatHeader(sheet, summaryRow, 1, 12, 'ECL PROVISION SUMMARY BY STAGE', '#d32f2f');
  
  sheet.getRange(summaryRow + 1, 1, 1, 2).merge().setValue('ECL Stage').setFontWeight('bold');
  sheet.getRange(summaryRow + 1, 3).setValue('Count').setFontWeight('bold');
  sheet.getRange(summaryRow + 1, 4).setValue('Gross Carrying Amount').setFontWeight('bold');
  sheet.getRange(summaryRow + 1, 5).setValue('ECL Rate').setFontWeight('bold');
  sheet.getRange(summaryRow + 1, 6).setValue('Opening Provision').setFontWeight('bold');
  sheet.getRange(summaryRow + 1, 7).setValue('Movement').setFontWeight('bold');
  sheet.getRange(summaryRow + 1, 8).setValue('Closing Provision').setFontWeight('bold');
  sheet.getRange(summaryRow + 1, 9).setValue('Coverage Ratio').setFontWeight('bold');
  
  const summaryData = [
    ['Stage 1 - Performing', '=COUNTIF(E3:E250,"Stage 1")', '=SUMIF(E3:E250,"Stage 1",D3:D250)', '=IF(C15>0,G15/C15,0)', '=SUMIF(E3:E250,"Stage 1",J3:J250)', '=SUMIF(E3:E250,"Stage 1",K3:K250)', '=SUMIF(E3:E250,"Stage 1",L3:L250)', '=IF(C15>0,G15/C15,0)'],
    ['Stage 2 - Underperforming', '=COUNTIF(E3:E250,"Stage 2")', '=SUMIF(E3:E250,"Stage 2",D3:D250)', '=IF(C16>0,G16/C16,0)', '=SUMIF(E3:E250,"Stage 2",J3:J250)', '=SUMIF(E3:E250,"Stage 2",K3:K250)', '=SUMIF(E3:E250,"Stage 2",L3:L250)', '=IF(C16>0,G16/C16,0)'],
    ['Stage 3 - Credit Impaired', '=COUNTIF(E3:E250,"Stage 3")', '=SUMIF(E3:E250,"Stage 3",D3:D250)', '=IF(C17>0,G17/C17,0)', '=SUMIF(E3:E250,"Stage 3",J3:J250)', '=SUMIF(E3:E250,"Stage 3",K3:K250)', '=SUMIF(E3:E250,"Stage 3",L3:L250)', '=IF(C17>0,G17/C17,0)'],
    ['Simplified (Lifetime ECL)', '=COUNTIF(E3:E250,"Simplified (Lifetime)")', '=SUMIF(E3:E250,"Simplified (Lifetime)",D3:D250)', '=IF(C18>0,G18/C18,0)', '=SUMIF(E3:E250,"Simplified (Lifetime)",J3:J250)', '=SUMIF(E3:E250,"Simplified (Lifetime)",K3:K250)', '=SUMIF(E3:E250,"Simplified (Lifetime)",L3:L250)', '=IF(C18>0,G18/C18,0)'],
    ['Total', '=SUM(B15:B18)', '=SUM(C15:C18)', '=IF(C19>0,G19/C19,0)', '=SUM(E15:E18)', '=SUM(F15:F18)', '=SUM(G15:G18)', '=IF(C19>0,G19/C19,0)']
  ];
  
  let row = summaryRow + 2;
  summaryData.forEach(data => {
    sheet.getRange(row, 1, 1, 2).merge().setValue(data[0]);
    for (let i = 1; i < data.length; i++) {
      sheet.getRange(row, i + 2).setFormula(data[i]);
    }
    
    if (data[0] === 'Total') {
      sheet.getRange(row, 1, 1, 9).setFontWeight('bold').setBackground('#ffcdd2');
    }
    row++;
  });
  
  formatCurrency(sheet.getRange('C15:G18'));
  formatPercentage(sheet.getRange('D15:D18'));
  formatPercentage(sheet.getRange('H15:H18'));
  
  // Add ECL methodology notes
  sheet.getRange('A' + (row + 2)).setValue('ECL CALCULATION METHODOLOGY:').setFontWeight('bold').setFontColor('#c62828');
  sheet.getRange('A' + (row + 3) + ':L' + (row + 9)).merge()
       .setValue('Expected Credit Loss = EAD × PD × LGD\n\n' +
                 'Where:\n' +
                 '• EAD (Exposure at Default) = Gross carrying amount of financial asset\n' +
                 '• PD (Probability of Default) = Likelihood of default occurring (12-month for Stage 1, lifetime for Stage 2 & 3)\n' +
                 '• LGD (Loss Given Default) = % of EAD that will be lost if default occurs\n' +
                 '• Discounting: Expected losses are discounted using the instrument\'s EIR to present value (Ind AS 109.B5.5.29)\n\n' +
                 'Stage Transfer Criteria:\n' +
                 '• Stage 1 → Stage 2: Significant increase in credit risk (typically DPD > 30 days)\n' +
                 '• Stage 2 → Stage 3: Objective evidence of impairment (typically DPD > 90 days or NPA classification)\n' +
                 '• Backward movement possible if credit risk decreases\n\n' +
                 'Ind AS 109 requires forward-looking information to be incorporated into ECL estimates.')
       .setWrap(true)
       .setVerticalAlignment('top')
       .setBackground('#ffebee')
       .setBorder(true, true, true, true, false, false);
  
  sheet.setFrozenRows(2);
  
  Logger.log('ECL Impairment sheet created.');
}

// ═══════════════════════════════════════════════════════════════════════════
// 7. AMORTIZATION SCHEDULE SHEET
// ═══════════════════════════════════════════════════════════════════════════

function createAmortizationScheduleSheet(ss) {
  const sheet = ss.insertSheet('Amortization_Schedule');
  sheet.setTabColor('#4caf50');
  
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidths(6, 5, 130);
  sheet.setColumnWidth(11, 130);
  
  formatHeader(sheet, 1, 1, 11, 'AMORTIZED COST - EFFECTIVE INTEREST RATE METHOD', '#388e3c');
  
  const headers = [
    'ID',
    'Instrument Name',
    'Opening Amortized Cost (₹)',
    'EIR (%)',
    'Days',
    'Interest Income (₹)',
    'Cash Received (₹)',
    'Amortization (₹)',
    'Impairment Charge (₹)',
    'Other Adjustments (₹)',
    'Closing Amortized Cost (₹)'
  ];
  
  formatSubHeader(sheet, 2, 1, headers, '#4caf50');
  
  // Amortization calculations for Amortized Cost instruments only (extended to 100 rows)
  // OPTIMIZED: Batch operations to reduce API calls from ~1100 to ~5
  const numRows = 100;
  const startRow = 3;
  const numCols = 11;

  // Create 2D arrays for batch operations
  const formulaArray = [];
  const backgroundArray = Array(numRows).fill(null).map(() => Array(numCols).fill(null));

  // Populate arrays in memory (fast)
  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    const formulas = [
      `=IF(ISBLANK(Instruments_Register!A${row}),"",Instruments_Register!A${row})`,
      `=IF(ISBLANK(Instruments_Register!A${row}),"",Instruments_Register!B${row})`,
      `=IF(ISBLANK(Instruments_Register!A${row}),"",Instruments_Register!J${row})`,
      `=IF(ISBLANK(Instruments_Register!A${row}),"",Instruments_Register!I${row})`,
      // Days in period - CORRECTED: Use actual period days (B$8) not full year (B$7)
      // This allows for interim reporting (monthly, quarterly, etc.)
      `=IF(ISBLANK(Instruments_Register!A${row}),"",Input_Variables!$B$8)`,
      // Interest income = Opening Balance × EIR × (Days / Days in Year)
      // CORRECTED: For Stage 3 (credit-impaired), calculate interest on net carrying amount (gross - ECL provision)
      // Per Ind AS 109.5.4.1, interest revenue for credit-impaired assets = net carrying amount × EIR
      `=IF(ISBLANK(Instruments_Register!A${row}),"",IF(Classification_Matrix!F${row}="Amortized Cost",IF(ECL_Impairment!E${row}="Stage 3",(C${row}-ECL_Impairment!J${row})*D${row}*(E${row}/Input_Variables!$B$7),C${row}*D${row}*(E${row}/Input_Variables!$B$7)),0))`,
      // Cash received (coupon payment) - CORRECTED to account for coupon frequency and period
      // Formula: Face Value × Coupon Rate × Frequency Factor × (Days in Period / Days in Year)
      // Frequency Factor: Annual=1, Semi-Annual=0.5, Quarterly=0.25, Monthly=1/12
      `=IF(ISBLANK(Instruments_Register!A${row}),"",IF(Classification_Matrix!F${row}="Amortized Cost",Instruments_Register!G${row}*Instruments_Register!H${row}*IF(Instruments_Register!R${row}="Semi-Annual",0.5,IF(Instruments_Register!R${row}="Quarterly",0.25,IF(Instruments_Register!R${row}="Monthly",1/12,IF(Instruments_Register!R${row}="Annual",1,0))))*(E${row}/Input_Variables!$B$7),0))`,
      // Amortization = Interest Income - Cash Received
      `=IF(ISBLANK(Instruments_Register!A${row}),"",F${row}-G${row})`,
      // Impairment charge from ECL sheet
      `=IF(ISBLANK(Instruments_Register!A${row}),"",IF(Classification_Matrix!F${row}="Amortized Cost",ECL_Impairment!K${row},0))`,
      // Other adjustments (input cell for manual entries)
      `=IF(ISBLANK(Instruments_Register!A${row}),"",0)`,
      // Closing = Opening + Interest - Cash - Impairment + Adjustments
      `=IF(ISBLANK(Instruments_Register!A${row}),"",IF(Classification_Matrix!F${row}="Amortized Cost",C${row}+F${row}-G${row}-I${row}+J${row},0))`
    ];

    formulaArray.push(formulas);

    // Mark "Other Adjustments" (Column J, index 9) as input cell
    backgroundArray[i][9] = '#e8f5e9';
  }

  // Write all formulas and backgrounds in batch (fast - only 2 API calls!)
  const dataRange = sheet.getRange(startRow, 1, numRows, numCols);
  dataRange.setFormulas(formulaArray);
  dataRange.setBackgrounds(backgroundArray);

  // Batch format currency and percentage ranges (reduced from :1000 to :250 for performance)
  formatCurrency(sheet.getRange('C3:C250'));
  formatPercentage(sheet.getRange('D3:D250'));
  formatCurrency(sheet.getRange('F3:K250'));
  
  // Summary section
  const summaryRow = 13;
  formatHeader(sheet, summaryRow, 1, 11, 'AMORTIZED COST SUMMARY', '#4caf50');
  
  sheet.getRange(summaryRow + 1, 1, 1, 2).merge().setValue('Description').setFontWeight('bold');
  sheet.getRange(summaryRow + 1, 3).setValue('Amount (₹)').setFontWeight('bold');
  
  const summaryData = [
    ['Total Opening Amortized Cost', '=SUMIF(Classification_Matrix!F3:F250,"Amortized Cost",C3:C250)'],
    ['Add: Interest Income (EIR basis)', '=SUM(F3:F250)'],
    ['Less: Cash Receipts', '=SUM(G3:G250)'],
    ['Less: Impairment Charge', '=SUM(I3:I250)'],
    ['Add/(Less): Other Adjustments', '=SUM(J3:J250)'],
    ['Total Closing Amortized Cost', '=SUMIF(Classification_Matrix!F3:F250,"Amortized Cost",K3:K250)'],
    ['', ''],
    ['Verification (should be zero)', '=C15+C16-C17-C18+C19-C20']
  ];
  
  let row = summaryRow + 2;
  summaryData.forEach(data => {
    sheet.getRange(row, 1, 1, 2).merge().setValue(data[0]);
    if (data.length > 1) {
      sheet.getRange(row, 3).setFormula(data[1]);
    }
    
    if (data[0].includes('Total') || data[0].includes('Verification')) {
      sheet.getRange(row, 1, 1, 3).setFontWeight('bold').setBackground('#c8e6c9');
    }
    row++;
  });
  
  formatCurrency(sheet.getRange('C15:C22'));
  
  // Conditional formatting for verification
  const verifyRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberNotBetween(-100, 100)
    .setBackground('#ffcdd2')
    .setFontColor('#c62828')
    .setRanges([sheet.getRange('C22')])
    .build();
  
  const rules = sheet.getConditionalFormatRules();
  rules.push(verifyRule);
  sheet.setConditionalFormatRules(rules);
  
  // Add methodology notes
  sheet.getRange('A' + (row + 2)).setValue('EFFECTIVE INTEREST RATE (EIR) METHOD:').setFontWeight('bold').setFontColor('#388e3c');
  sheet.getRange('A' + (row + 3) + ':K' + (row + 7)).merge()
       .setValue('The effective interest rate (EIR) is the rate that exactly discounts estimated future cash flows through the expected life ' +
                 'of the financial instrument to the gross carrying amount.\n\n' +
                 'Interest Income Recognition:\n' +
                 '• Interest income = Opening amortized cost × EIR × (Time proportion)\n' +
                 '• For credit-impaired assets (Stage 3), interest income is calculated on net carrying amount (gross - ECL provision)\n\n' +
                 'Amortization:\n' +
                 '• Premium amortization = Interest income < Cash received → reduces carrying amount\n' +
                 '• Discount accretion = Interest income > Cash received → increases carrying amount\n\n' +
                 'Per Ind AS 109, EIR includes all fees, transaction costs, premiums, and discounts that are integral to the instrument.')
       .setWrap(true)
       .setVerticalAlignment('top')
       .setBackground('#e8f5e9')
       .setBorder(true, true, true, true, false, false);
  
  sheet.setFrozenRows(2);
  
  Logger.log('Amortization Schedule sheet created.');
}

// ═══════════════════════════════════════════════════════════════════════════
// 8. PERIOD END ENTRIES SHEET
// ═══════════════════════════════════════════════════════════════════════════

function createPeriodEndEntriesSheet(ss) {
  const sheet = ss.insertSheet('Period_End_Entries');
  sheet.setTabColor('#ff5722');
  
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 250);
  
  formatHeader(sheet, 1, 1, 6, 'PERIOD-END JOURNAL ENTRIES FOR BOOK CLOSURE', '#d84315');
  
  sheet.getRange('A2').setValue('Extract these entries for posting in books of accounts').setFontStyle('italic').setFontColor('#bf360c');
  
  // Entry 1: Fair Value Adjustments - FVTPL
  formatHeader(sheet, 4, 1, 6, 'ENTRY 1: FAIR VALUE ADJUSTMENTS - FVTPL', '#f4511e');
  
  const entry1Headers = ['Entry No.', 'Account Head', 'Narration', 'Debit (₹)', 'Credit (₹)', 'Ind AS Reference'];
  formatSubHeader(sheet, 5, 1, entry1Headers, '#ff5722');
  
  sheet.getRange('A6').setValue('JE001');
  sheet.getRange('B6').setValue('Financial Assets - FVTPL');
  sheet.getRange('C6').setValue('Fair value adjustment for period');
  sheet.getRange('D6').setFormula('=IF(Fair_Value_Workings!E15>0,Fair_Value_Workings!E15,0)');
  sheet.getRange('E6').setFormula('=IF(Fair_Value_Workings!E15<0,ABS(Fair_Value_Workings!E15),0)');
  sheet.getRange('F6').setValue('Ind AS 109.5.7.1 - FVTPL measurement');
  
  sheet.getRange('A7').setValue('JE001');
  sheet.getRange('B7').setValue('   Gain/(Loss) on Fair Valuation - P&L');
  sheet.getRange('C7').setValue('Fair value adjustment for period');
  sheet.getRange('D7').setFormula('=IF(Fair_Value_Workings!E15<0,ABS(Fair_Value_Workings!E15),0)');
  sheet.getRange('E7').setFormula('=IF(Fair_Value_Workings!E15>0,Fair_Value_Workings!E15,0)');
  sheet.getRange('F7').setValue('Ind AS 109.5.7.1');
  
  // Entry 2: Fair Value Adjustments - FVOCI
  formatHeader(sheet, 9, 1, 6, 'ENTRY 2: FAIR VALUE ADJUSTMENTS - FVOCI', '#f4511e');
  formatSubHeader(sheet, 10, 1, entry1Headers, '#ff5722');
  
  sheet.getRange('A11').setValue('JE002');
  sheet.getRange('B11').setValue('Financial Assets - FVOCI');
  sheet.getRange('C11').setValue('Fair value adjustment through OCI');
  sheet.getRange('D11').setFormula('=IF(Fair_Value_Workings!F16>0,Fair_Value_Workings!F16,0)');
  sheet.getRange('E11').setFormula('=IF(Fair_Value_Workings!F16<0,ABS(Fair_Value_Workings!F16),0)');
  sheet.getRange('F11').setValue('Ind AS 109.5.7.5 - FVOCI measurement');
  
  sheet.getRange('A12').setValue('JE002');
  sheet.getRange('B12').setValue('   OCI - Fair Value Reserve');
  sheet.getRange('C12').setValue('Fair value adjustment through OCI');
  sheet.getRange('D12').setFormula('=IF(Fair_Value_Workings!F16<0,ABS(Fair_Value_Workings!F16),0)');
  sheet.getRange('E12').setFormula('=IF(Fair_Value_Workings!F16>0,Fair_Value_Workings!F16,0)');
  sheet.getRange('F12').setValue('Ind AS 109.5.7.5');
  
  // Entry 3: Interest Income - EIR method
  formatHeader(sheet, 14, 1, 6, 'ENTRY 3: INTEREST INCOME - EFFECTIVE INTEREST RATE METHOD', '#f4511e');
  formatSubHeader(sheet, 15, 1, entry1Headers, '#ff5722');
  
  sheet.getRange('A16').setValue('JE003');
  sheet.getRange('B16').setValue('Financial Assets - Amortized Cost');
  sheet.getRange('C16').setValue('Interest income on EIR basis');
  sheet.getRange('D16').setFormula('=Amortization_Schedule!C16');
  sheet.getRange('E16').setValue('');
  sheet.getRange('F16').setValue('Ind AS 109.5.4.1 - EIR method');
  
  sheet.getRange('A17').setValue('JE003');
  sheet.getRange('B17').setValue('   Interest Income - Financial Assets');
  sheet.getRange('C17').setValue('Interest income on EIR basis');
  sheet.getRange('D17').setValue('');
  sheet.getRange('E17').setFormula('=Amortization_Schedule!C16');
  sheet.getRange('F17').setValue('Ind AS 109.5.4.1');
  
  // Entry 4: ECL Provision
  formatHeader(sheet, 19, 1, 6, 'ENTRY 4: EXPECTED CREDIT LOSS PROVISION', '#f4511e');
  formatSubHeader(sheet, 20, 1, entry1Headers, '#ff5722');
  
  sheet.getRange('A21').setValue('JE004');
  sheet.getRange('B21').setValue('Impairment Loss on Financial Assets - P&L');
  sheet.getRange('C21').setValue('ECL provision for the period');
  sheet.getRange('D21').setFormula('=IF(ECL_Impairment!F18>0,ECL_Impairment!F18,0)');
  sheet.getRange('E21').setFormula('=IF(ECL_Impairment!F18<0,ABS(ECL_Impairment!F18),0)');
  sheet.getRange('F21').setValue('Ind AS 109.5.5 - Impairment');
  
  sheet.getRange('A22').setValue('JE004');
  sheet.getRange('B22').setValue('   ECL Provision - Financial Assets');
  sheet.getRange('C22').setValue('ECL provision for the period');
  sheet.getRange('D22').setFormula('=IF(ECL_Impairment!F18<0,ABS(ECL_Impairment!F18),0)');
  sheet.getRange('E22').setFormula('=IF(ECL_Impairment!F18>0,ECL_Impairment!F18,0)');
  sheet.getRange('F22').setValue('Ind AS 109.5.5');
  
  // Entry 5: Premium/Discount Amortization
  formatHeader(sheet, 24, 1, 6, 'ENTRY 5: PREMIUM/DISCOUNT AMORTIZATION', '#f4511e');
  formatSubHeader(sheet, 25, 1, entry1Headers, '#ff5722');
  
  sheet.getRange('A26').setValue('JE005');
  sheet.getRange('B26').setValue('Interest Income - Financial Assets');
  sheet.getRange('C26').setValue('Net premium amortization adjustment');
  sheet.getRange('D26').setFormula('=IF(SUM(Amortization_Schedule!H3:H250)<0,ABS(SUM(Amortization_Schedule!H3:H250)),0)');
  sheet.getRange('E26').setFormula('=IF(SUM(Amortization_Schedule!H3:H250)>0,SUM(Amortization_Schedule!H3:H250),0)');
  sheet.getRange('F26').setValue('Ind AS 109.5.4.1 - EIR adjustment');
  
  sheet.getRange('A27').setValue('JE005');
  sheet.getRange('B27').setValue('   Financial Assets - Amortized Cost');
  sheet.getRange('C27').setValue('Net premium amortization adjustment');
  sheet.getRange('D27').setFormula('=IF(SUM(Amortization_Schedule!H3:H250)>0,SUM(Amortization_Schedule!H3:H250),0)');
  sheet.getRange('E27').setFormula('=IF(SUM(Amortization_Schedule!H3:H250)<0,ABS(SUM(Amortization_Schedule!H3:H250)),0)');
  sheet.getRange('F27').setValue('Ind AS 109.5.4.1');
  
  // Summary of entries
  formatHeader(sheet, 30, 1, 6, 'SUMMARY OF PERIOD-END ENTRIES', '#f4511e');
  
  sheet.getRange('A31:B31').merge().setValue('Entry').setFontWeight('bold');
  sheet.getRange('C31').setValue('Description').setFontWeight('bold');
  sheet.getRange('D31').setValue('Total Debit').setFontWeight('bold');
  sheet.getRange('E31').setValue('Total Credit').setFontWeight('bold');
  sheet.getRange('F31').setValue('P&L Impact').setFontWeight('bold');
  
  formatSubHeader(sheet, 31, 1, ['Entry', '', 'Description', 'Total Debit', 'Total Credit', 'P&L Impact'], '#ff5722');
  
  const summaryData = [
    ['JE001', 'FVTPL Fair Value', '=D6+D7', '=E6+E7', '=Fair_Value_Workings!E15'],
    ['JE002', 'FVOCI Fair Value', '=D11+D12', '=E11+E12', '0'],
    ['JE003', 'Interest Income (EIR)', '=D16', '=E17', '=Amortization_Schedule!C16'],
    ['JE004', 'ECL Provision', '=D21+D22', '=E21+E22', '=-ECL_Impairment!F18'],
    ['JE005', 'Premium/Discount Amortization', '=D26+D27', '=E26+E27', '=-SUM(Amortization_Schedule!H3:H250)'],
    ['TOTAL', '', '=SUM(D32:D36)', '=SUM(E32:E36)', '=SUM(F32:F36)']
  ];
  
  let row = 32;
  summaryData.forEach(data => {
    sheet.getRange(row, 1).setValue(data[0]);
    sheet.getRange(row, 2).setValue('');
    sheet.getRange(row, 3).setValue(data[1]);
    sheet.getRange(row, 4).setFormula(data[2]);
    sheet.getRange(row, 5).setFormula(data[3]);
    if (data.length > 4) {
      sheet.getRange(row, 6).setFormula(data[4]);
    }
    
    if (data[0] === 'TOTAL') {
      sheet.getRange(row, 1, 1, 6).setFontWeight('bold').setBackground('#ffccbc');
    }
    row++;
  });
  
  formatCurrency(sheet.getRange('D6:E250'));
  formatCurrency(sheet.getRange('D32:F37'));
  
  // Balancing check
  sheet.getRange('A' + (row + 1)).setValue('BALANCING CHECK:').setFontWeight('bold').setFontColor('#bf360c');
  sheet.getRange('A' + (row + 2)).setValue('Difference (should be zero):');
  sheet.getRange('B' + (row + 2)).setFormula('=D37-E37').setFontWeight('bold');
  formatCurrency(sheet.getRange('B' + (row + 2)));
  
  const balanceRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberNotBetween(-10, 10)
    .setBackground('#ffcdd2')
    .setFontColor('#c62828')
    .setRanges([sheet.getRange('B' + (row + 2))])
    .build();
  
  const rules = sheet.getConditionalFormatRules();
  rules.push(balanceRule);
  sheet.setConditionalFormatRules(rules);
  
  // Instructions
  sheet.getRange('A' + (row + 4)).setValue('POSTING INSTRUCTIONS:').setFontWeight('bold').setFontColor('#bf360c');
  sheet.getRange('A' + (row + 5) + ':F' + (row + 9)).merge()
       .setValue('1. Review each journal entry for accuracy and completeness\n' +
                 '2. Ensure debit and credit totals match (see balancing check above)\n' +
                 '3. Post entries in the general ledger with appropriate narration\n' +
                 '4. Attach this working paper as supporting documentation\n' +
                 '5. Update trial balance and prepare financial statements\n' +
                 '6. Ensure proper disclosure in notes to accounts per Ind AS 107 (Financial Instruments: Disclosures)')
       .setWrap(true)
       .setVerticalAlignment('top')
       .setBackground('#ffebee')
       .setBorder(true, true, true, true, false, false);
  
  sheet.setFrozenRows(2);
  
  Logger.log('Period End Entries sheet created.');
}

// ═══════════════════════════════════════════════════════════════════════════
// 9. RECONCILIATION SHEET
// ═══════════════════════════════════════════════════════════════════════════

function createReconciliationSheet(ss) {
  const sheet = ss.insertSheet('Reconciliation');
  sheet.setTabColor('#795548');
  
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidths(2, 4, 150);
  sheet.setColumnWidth(6, 250);
  
  formatHeader(sheet, 1, 1, 6, 'OPENING TO CLOSING BALANCE RECONCILIATION', '#5d4037');
  
  const headers = ['Particulars', 'Opening Balance (₹)', 'Additions/(Movements) (₹)', 'Closing Balance (₹)', 'Reference', 'Notes'];
  formatSubHeader(sheet, 2, 1, headers, '#6d4c41');
  
  // Financial Assets section
  formatHeader(sheet, 4, 1, 6, 'FINANCIAL ASSETS BY CLASSIFICATION', '#6d4c41');
  
  const assetsData = [
    ['Description', 'Opening', 'Movement', 'Closing', 'Sheet', 'Notes'],
    ['', '', '', '', '', ''],
    ['Financial Assets at Amortized Cost', '=Classification_Matrix!D15', '=Amortization_Schedule!C16-Amortization_Schedule!C17-Amortization_Schedule!C18', '=B8+C8', 'Amortization_Schedule', 'Interest income less cash and impairment'],
    ['Financial Assets at FVTPL', '=Classification_Matrix!D16', '=Fair_Value_Workings!F15', '=B9+C9', 'Fair_Value_Workings', 'Fair value movements to P&L'],
    ['Financial Assets at FVOCI', '=Classification_Matrix!D17', '=Fair_Value_Workings!F16', '=B10+C10', 'Fair_Value_Workings', 'Fair value movements to OCI'],
    ['Sub-total: Gross Financial Assets', '=SUM(B8:B10)', '=SUM(C8:C10)', '=SUM(D8:D10)', '', ''],
    ['', '', '', '', '', ''],
    
    // Provisions section
    ['PROVISIONS & IMPAIRMENT', '', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['ECL Provision - Stage 1', '=ECL_Impairment!E15', '=ECL_Impairment!F15', '=ECL_Impairment!G15', 'ECL_Impairment', '12-month ECL'],
    ['ECL Provision - Stage 2', '=ECL_Impairment!E16', '=ECL_Impairment!F16', '=ECL_Impairment!G16', 'ECL_Impairment', 'Lifetime ECL - performing'],
    ['ECL Provision - Stage 3', '=ECL_Impairment!E17', '=ECL_Impairment!F17', '=ECL_Impairment!G17', 'ECL_Impairment', 'Lifetime ECL - impaired'],
    ['Sub-total: Total ECL Provision', '=SUM(B17:B19)', '=SUM(C17:C19)', '=SUM(D17:D19)', '', ''],
    ['', '', '', '', '', ''],
    
    // Net position
    ['NET FINANCIAL ASSETS', '', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['Net Financial Assets (Gross - ECL)', '=B11-B20', '=C11-C20', '=D11-D20', '', 'Carrying amount in balance sheet'],
    ['', '', '', '', '', ''],
    
    // Equity section
    ['EQUITY COMPONENTS', '', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['OCI - Fair Value Reserve (FVOCI)', '', '=Fair_Value_Workings!H16', '=C28', '', 'Cumulative fair value changes'],
    ['Retained Earnings Impact', '', '=Cover!C15+Cover!C20', '=C29', '', 'Net P&L impact of all entries']
  ];
  
  let row = 5;
  assetsData.forEach(data => {
    data.forEach((value, col) => {
      const cell = sheet.getRange(row, col + 1);
      if (typeof value === 'string' && value.startsWith('=')) {
        cell.setFormula(value);
      } else {
        cell.setValue(value);
      }
      
      // Bold formatting for sub-totals and headers
      if (data[0].includes('Sub-total') || data[0].includes('NET FINANCIAL') || data[0].includes('PROVISIONS') || data[0].includes('EQUITY')) {
        cell.setFontWeight('bold');
        if (data[0].includes('Sub-total') || data[0].includes('NET FINANCIAL')) {
          sheet.getRange(row, 1, 1, 6).setBackground('#d7ccc8');
        }
      }
    });
    row++;
  });
  
  formatCurrency(sheet.getRange('B8:D30'));
  
  // Verification section
  formatHeader(sheet, row + 2, 1, 6, 'RECONCILIATION VERIFICATION', '#6d4c41');
  
  sheet.getRange(row + 3, 1).setValue('Control Total: Opening + Movement - Closing').setFontWeight('bold');
  sheet.getRange(row + 3, 2).setFormula('=(B11-B20)+(C11-C20)-(D11-D20)');
  sheet.getRange(row + 3, 3).setValue('(Should be zero)').setFontStyle('italic');
  
  formatCurrency(sheet.getRange(row + 3, 2));
  
  const verifyRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberNotBetween(-100, 100)
    .setBackground('#ffcdd2')
    .setFontColor('#c62828')
    .setRanges([sheet.getRange(row + 3, 2)])
    .build();
  
  const rules = sheet.getConditionalFormatRules();
  rules.push(verifyRule);
  sheet.setConditionalFormatRules(rules);
  
  // P&L impact summary
  formatHeader(sheet, row + 5, 1, 6, 'PROFIT & LOSS IMPACT SUMMARY', '#6d4c41');
  
  const plData = [
    ['Revenue/Income Items', '', '', '', '', ''],
    ['Interest Income (EIR method)', '=Amortization_Schedule!C16', 'Revenue from financial assets'],
    ['Fair Value Gain - FVTPL', '=IF(Fair_Value_Workings!E15>0,Fair_Value_Workings!E15,0)', 'Unrealized gains on FVTPL instruments'],
    ['', '', ''],
    ['Expense Items', '', ''],
    ['Fair Value Loss - FVTPL', '=IF(Fair_Value_Workings!E15<0,ABS(Fair_Value_Workings!E15),0)', 'Unrealized losses on FVTPL instruments'],
    ['Impairment Loss - ECL', '=ECL_Impairment!F18', 'Expected credit loss provision'],
    ['', '', ''],
    ['Net Impact on P&L', '=B' + (row + 7) + '+B' + (row + 8) + '-B' + (row + 11) + '-B' + (row + 12), 'To be reflected in P&L for the period']
  ];
  
  let plRow = row + 6;
  plData.forEach(data => {
    if (data[0] !== '') {
      sheet.getRange(plRow, 1).setValue(data[0]);
      if (data.length > 1 && data[1].startsWith('=')) {
        sheet.getRange(plRow, 2).setFormula(data[1]);
      }
      if (data.length > 2) {
        sheet.getRange(plRow, 3).setValue(data[2]);
      }
      
      if (data[0].includes('Items') || data[0].includes('Net Impact')) {
        sheet.getRange(plRow, 1).setFontWeight('bold');
      }
      if (data[0].includes('Net Impact')) {
        sheet.getRange(plRow, 1, 1, 3).setBackground('#d7ccc8').setFontWeight('bold');
      }
    }
    plRow++;
  });
  
  formatCurrency(sheet.getRange('B' + (row + 7) + ':B' + plRow));
  
  // Add reconciliation notes
  sheet.getRange('A' + (plRow + 2)).setValue('RECONCILIATION NOTES:').setFontWeight('bold').setFontColor('#5d4037');
  sheet.getRange('A' + (plRow + 3) + ':F' + (plRow + 6)).merge()
       .setValue('This reconciliation provides a complete trail from opening to closing balances for all financial instruments.\n\n' +
                 '• All movements are traced to specific journal entries in the Period_End_Entries sheet\n' +
                 '• Control totals ensure mathematical accuracy of the working papers\n' +
                 '• P&L impact summary shows the total effect on profit or loss for the period\n' +
                 '• OCI movements affect equity but not profit or loss (until reclassification/disposal)')
       .setWrap(true)
       .setVerticalAlignment('top')
       .setBackground('#efebe9')
       .setBorder(true, true, true, true, false, false);
  
  sheet.setFrozenRows(2);
  
  Logger.log('Reconciliation sheet created.');
}

// ═══════════════════════════════════════════════════════════════════════════
// 10. REFERENCES SHEET
// ═══════════════════════════════════════════════════════════════════════════

function createReferencesSheet(ss) {
  const sheet = ss.insertSheet('References');
  sheet.setTabColor('#607d8b');
  
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 600);
  
  formatHeader(sheet, 1, 1, 2, 'IND AS 109 - KEY PROVISIONS & REFERENCES', '#455a64');
  
  const references = [
    ['SECTION', 'PROVISION/REQUIREMENT'],
    ['', ''],
    ['Classification & Measurement', ''],
    ['Para 4.1.1', 'An entity shall classify financial assets as subsequently measured at amortized cost, fair value through other comprehensive income (FVOCI), or fair value through profit or loss (FVTPL).'],
    ['Para 4.1.2', 'A financial asset shall be measured at amortized cost if both: (a) the asset is held within a business model whose objective is to hold assets to collect contractual cash flows, and (b) the contractual terms give rise on specified dates to cash flows that are solely payments of principal and interest (SPPI).'],
    ['Para 4.1.2A', 'A financial asset shall be measured at FVOCI if both: (a) the asset is held within a business model whose objective is achieved by both collecting contractual cash flows and selling financial assets, and (b) the contractual terms meet the SPPI test.'],
    ['Para 4.1.4', 'Financial assets that do not meet the criteria for amortized cost or FVOCI shall be measured at FVTPL.'],
    ['', ''],
    
    ['Business Model Assessment', ''],
    ['Para 4.1.2B', 'The business model assessment is based on reasonably expected scenarios without taking "worst case" or "stress case" scenarios into account.'],
    ['Para B4.1.2A', 'Factors to consider include: how performance is evaluated, how managers are compensated, the risks that affect performance, and the frequency and volume of sales.'],
    ['', ''],
    
    ['SPPI Test', ''],
    ['Para 4.1.3', 'Contractual cash flows that are SPPI are consistent with a basic lending arrangement. Interest includes only consideration for time value of money, credit risk, other basic lending risks, and profit margin.'],
    ['Para B4.1.9A', 'Contractual terms that change the timing or amount of cash flows in a manner inconsistent with SPPI fail the test (e.g., leverage features, caps/floors on interest).'],
    ['', ''],
    
    ['Effective Interest Rate', ''],
    ['Para 5.4.1', 'Interest revenue shall be calculated using the effective interest rate (EIR) method. EIR exactly discounts estimated future cash flows through the expected life to the gross carrying amount.'],
    ['Para B5.4.1', 'EIR calculation includes all fees and points paid or received between parties, transaction costs, and all premiums or discounts that are integral to EIR.'],
    ['', ''],
    
    ['Impairment - ECL Model', ''],
    ['Para 5.5.1', 'An entity shall recognize a loss allowance for expected credit losses (ECL) on financial assets measured at amortized cost and debt instruments at FVOCI.'],
    ['Para 5.5.3', 'General approach requires 12-month ECL for Stage 1, lifetime ECL for Stage 2 (significant increase in credit risk), and lifetime ECL for Stage 3 (credit impaired).'],
    ['Para 5.5.9', 'Significant increase in credit risk is assessed by comparing the risk of default at reporting date with the risk at initial recognition.'],
    ['Para 5.5.11', 'An entity may assume credit risk has increased significantly when contractual payments are more than 30 days past due (rebuttable presumption).'],
    ['Para B5.5.37', 'ECL is the probability-weighted estimate of credit losses. It shall reflect: (a) an unbiased and probability-weighted amount, (b) time value of money, and (c) reasonable and supportable information.'],
    ['', ''],
    
    ['Fair Value Measurement', ''],
    ['Para 5.1.1', 'On initial recognition, fair value is normally the transaction price. Subsequently, fair value is measured in accordance with Ind AS 113.'],
    ['Para 5.7.5', 'Gains and losses on debt instruments at FVOCI shall be recognized in OCI, except for interest revenue, ECL, and foreign exchange gains/losses which are recognized in P&L.'],
    ['Para 5.7.1', 'A gain or loss on a financial asset measured at FVTPL shall be recognized in profit or loss.'],
    ['', ''],
    
    ['Hedge Accounting', ''],
    ['Para 6.4.1', 'A hedging relationship qualifies for hedge accounting only if: (a) it consists of eligible items, (b) there is formal designation and documentation, and (c) the hedge meets effectiveness requirements.'],
    ['Para 6.4.1(c)', 'Ind AS 109 effectiveness requirements (principles-based): (i) economic relationship between hedged item and hedging instrument, (ii) credit risk does not dominate value changes, and (iii) hedge ratio consistent with risk management objective. No fixed 80-125% rule.'],
    ['', ''],
    
    ['Derecognition', ''],
    ['Para 3.2.3', 'An entity shall derecognize a financial asset when: (a) the contractual rights to cash flows expire, or (b) it transfers the asset and substantially all risks and rewards of ownership.'],
    ['Para 3.3.1', 'An entity shall derecognize a financial liability when it is extinguished (i.e., when the obligation is discharged, cancelled, or expires).'],
    ['', ''],
    
    ['Disclosure Requirements', ''],
    ['Ind AS 107', 'Financial Instruments: Disclosures requires extensive qualitative and quantitative disclosures including: significance of financial instruments, nature and extent of risks, credit risk concentrations, ECL methodologies, fair value hierarchy, and hedge accounting.']
  ];
  
  let row = 2;
  references.forEach(ref => {
    if (ref[0] === '' && ref[1] === '') {
      // Empty row for spacing
      row++;
    } else if (ref[1] === '') {
      // Section header
      formatHeader(sheet, row, 1, 2, ref[0], '#546e7a');
      row++;
    } else if (ref[0] === 'SECTION') {
      // Column headers
      formatSubHeader(sheet, row, 1, ref, '#607d8b');
      row++;
    } else {
      // Content row
      sheet.getRange(row, 1).setValue(ref[0]).setFontWeight('bold').setVerticalAlignment('top');
      sheet.getRange(row, 2).setValue(ref[1]).setWrap(true).setVerticalAlignment('top');
      sheet.setRowHeight(row, 60);
      row++;
    }
  });
  
  // Add disclaimer
  formatHeader(sheet, row + 2, 1, 2, 'IMPORTANT NOTES', '#546e7a');
  
  sheet.getRange('A' + (row + 3) + ':B' + (row + 6)).merge()
       .setValue('1. This summary is for quick reference only. Please refer to the complete text of Ind AS 109 for detailed requirements.\n\n' +
                 '2. The standard is subject to amendments and interpretations issued by ICAI from time to time.\n\n' +
                 '3. Professional judgment is required in applying these principles to specific facts and circumstances.\n\n' +
                 '4. Consult with qualified auditors and financial advisors for complex transactions.')
       .setWrap(true)
       .setVerticalAlignment('top')
       .setBackground('#eceff1')
       .setFontStyle('italic')
       .setBorder(true, true, true, true, false, false);
  
  sheet.setFrozenRows(2);
  
  Logger.log('References sheet created.');
}

// ═══════════════════════════════════════════════════════════════════════════
// 11. AUDIT NOTES SHEET
// ═══════════════════════════════════════════════════════════════════════════

function createAuditNotesSheet(ss) {
  const sheet = ss.insertSheet('Audit_Notes');
  sheet.setTabColor('#3f51b5');
  
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidths(2, 5, 150);
  
  formatHeader(sheet, 1, 1, 6, 'AUDIT NOTES & CONTROL CHECKS', '#303f9f');
  
  // Control Totals Section
  formatHeader(sheet, 3, 1, 6, 'CONTROL TOTALS & MATHEMATICAL ACCURACY', '#3949ab');
  
  const controlHeaders = ['Control Check', 'Calculated Value', 'Expected Value', 'Variance', 'Status', 'Action Required'];
  formatSubHeader(sheet, 4, 1, controlHeaders, '#3f51b5');
  
  const controlChecks = [
    ['1. Total Journal Entries Balance', '=Period_End_Entries!D37-Period_End_Entries!E37', '0', '=B5-C5', '=IF(ABS(D5)<100,"✓ Pass","✗ FAIL")', 'If fail, review journal entries'],
    ['2. Reconciliation: Opening + Movement = Closing', '=(Reconciliation!B11-Reconciliation!B20)+(Reconciliation!C11-Reconciliation!C20)-(Reconciliation!D11-Reconciliation!D20)', '0', '=B6-C6', '=IF(ABS(D6)<100,"✓ Pass","✗ FAIL")', 'If fail, check reconciliation formulas'],
    ['3. Amortized Cost Verification', '=Amortization_Schedule!C15+Amortization_Schedule!C16-Amortization_Schedule!C17-Amortization_Schedule!C18+Amortization_Schedule!C19-Amortization_Schedule!C20', '0', '=B7-C7', '=IF(ABS(D7)<100,"✓ Pass","✗ FAIL")', 'If fail, review amortization calculations'],
    ['4. Classification Count Reconciliation', '=Classification_Matrix!C18', '=COUNTA(Instruments_Register!A3:A250)-COUNTBLANK(Instruments_Register!A3:A250)', '=B8-C8', '=IF(D8=0,"✓ Pass","✗ FAIL")', 'All instruments must be classified'],
    ['5. ECL Coverage Check - Stage 3', '=ECL_Impairment!H17', '', '', '=IF(B9>0.5,"✓ Adequate","⚠ Review")', 'Stage 3 coverage should exceed 50%']
  ];
  
  let row = 5;
  controlChecks.forEach(check => {
    check.forEach((value, col) => {
      const cell = sheet.getRange(row, col + 1);
      if (typeof value === 'string' && value.startsWith('=')) {
        cell.setFormula(value);
      } else {
        cell.setValue(value);
      }
    });
    row++;
  });
  
  formatCurrency(sheet.getRange('B5:D9'));
  
  // Conditional formatting for status
  const statusRange = sheet.getRange('E5:E9');
  const passRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('Pass')
    .setBackground('#c8e6c9')
    .setFontColor('#2e7d32')
    .setRanges([statusRange])
    .build();
  
  const failRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('FAIL')
    .setBackground('#ffcdd2')
    .setFontColor('#c62828')
    .setRanges([statusRange])
    .build();
  
  const reviewRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('Review')
    .setBackground('#fff9c4')
    .setFontColor('#f57f17')
    .setRanges([statusRange])
    .build();
  
  let rules = sheet.getConditionalFormatRules();
  rules.push(passRule, failRule, reviewRule);
  sheet.setConditionalFormatRules(rules);
  
  // Audit Assertions Section
  formatHeader(sheet, row + 2, 1, 6, 'AUDIT ASSERTIONS - IND AS 109 COMPLIANCE', '#3949ab');
  
  const assertionHeaders = ['Assertion', 'Requirement', 'Evidence/Check', 'Status', 'Comments'];
  formatSubHeader(sheet, row + 3, 1, assertionHeaders, '#3f51b5');
  
  const assertions = [
    ['Completeness', 'All financial instruments are recorded', 'Count of instruments in register vs. ledger', 'Pending', 'Verify with GL'],
    ['Accuracy', 'Amounts are mathematically correct', 'All control totals pass', '=IF(COUNTIF(E5:E9,"✗ FAIL")=0,"Verified","Issues Found")', 'Auto-verified'],
    ['Classification', 'SPPI test & business model correctly applied', 'Review Classification_Matrix sheet', 'Pending', 'Review with management'],
    ['Valuation', 'Fair values per Ind AS 113 hierarchy', 'Fair value sources documented', 'Pending', 'Obtain external valuations'],
    ['Impairment', 'ECL provisions per 3-stage model', 'ECL workings complete with PD/LGD/EAD', 'Pending', 'Review staging logic'],
    ['Measurement', 'EIR method correctly applied', 'Amortization schedule reviewed', 'Pending', 'Test EIR calculations'],
    ['Presentation', 'Balance sheet presentation per Ind AS 32', 'Assets net of ECL provision', 'Pending', 'Review FS presentation'],
    ['Disclosure', 'Ind AS 107 disclosures complete', 'Disclosure checklist prepared', 'Pending', 'Prepare disclosure notes']
  ];
  
  row = row + 4;
  assertions.forEach(assertion => {
    assertion.forEach((value, col) => {
      const cell = sheet.getRange(row, col + 1);
      if (typeof value === 'string' && value.startsWith('=')) {
        cell.setFormula(value);
      } else {
        cell.setValue(value);
      }
    });
    row++;
  });
  
  // Risk Areas Section
  formatHeader(sheet, row + 2, 1, 6, 'KEY RISK AREAS & AUDIT FOCUS', '#3949ab');
  
  const riskHeaders = ['Risk Area', 'Description', 'Mitigating Control', 'Audit Procedure', 'Priority'];
  formatSubHeader(sheet, row + 3, 1, riskHeaders, '#3f51b5');
  
  const risks = [
    ['Classification Errors', 'Incorrect SPPI or business model assessment', 'Review by finance manager and documentation', 'Test sample of instruments for classification', 'High'],
    ['Fair Value Measurement', 'Unreliable or outdated valuations', 'Use of independent valuers', 'Verify valuation sources and recalculate', 'High'],
    ['ECL Estimation', 'Inadequate or excessive provisioning', 'Credit committee review of staging', 'Test PD/LGD assumptions and staging logic', 'High'],
    ['Stage Migration', 'Failure to identify credit deterioration', 'Monthly DPD monitoring', 'Review stage transfer analysis', 'Medium'],
    ['EIR Calculation', 'Omission of fees/costs in EIR', 'Checklist of items to include in EIR', 'Recalculate EIR for material instruments', 'Medium'],
    ['Hedge Accounting', 'Ineffective hedges not identified', 'Quarterly effectiveness testing', 'Review hedge effectiveness calculations', 'Medium'],
    ['Documentation', 'Insufficient support for judgments', 'Policy to document all key decisions', 'Review policy docs and memos', 'Low']
  ];
  
  row = row + 4;
  risks.forEach(risk => {
    risk.forEach((value, col) => {
      sheet.getRange(row, col + 1).setValue(value);
    });
    
    // Color code by priority
    const priority = risk[4];
    const priorityCell = sheet.getRange(row, 5);
    if (priority === 'High') {
      priorityCell.setBackground('#ffcdd2').setFontColor('#c62828');
    } else if (priority === 'Medium') {
      priorityCell.setBackground('#fff9c4').setFontColor('#f57f17');
    } else {
      priorityCell.setBackground('#c8e6c9').setFontColor('#2e7d32');
    }
    
    row++;
  });
  
  // Materiality Assessment
  formatHeader(sheet, row + 2, 1, 6, 'MATERIALITY ASSESSMENT', '#3949ab');
  
  sheet.getRange(row + 3, 1).setValue('Overall Materiality:').setFontWeight('bold');
  formatInputCell(sheet.getRange(row + 3, 2), '#e3f2fd');
  sheet.getRange(row + 3, 3).setValue('(Typically 1-5% of appropriate benchmark)');
  formatCurrency(sheet.getRange(row + 3, 2));
  
  sheet.getRange(row + 4, 1).setValue('Performance Materiality:').setFontWeight('bold');
  sheet.getRange(row + 4, 2).setFormula('=B' + (row + 3) + '*0.75');
  sheet.getRange(row + 4, 3).setValue('(Typically 75% of overall materiality)');
  formatCurrency(sheet.getRange(row + 4, 2));
  
  sheet.getRange(row + 5, 1).setValue('Clearly Trivial Threshold:').setFontWeight('bold');
  sheet.getRange(row + 5, 2).setFormula('=B' + (row + 3) + '*0.05');
  sheet.getRange(row + 5, 3).setValue('(Typically 5% of overall materiality)');
  formatCurrency(sheet.getRange(row + 5, 2));
  
  // Audit Conclusion
  formatHeader(sheet, row + 7, 1, 6, 'AUDIT CONCLUSION & SIGN-OFF', '#3949ab');
  
  sheet.getRange(row + 8, 1).setValue('Prepared by:').setFontWeight('bold');
  formatInputCell(sheet.getRange(row + 8, 2, 1, 2).merge());
  
  sheet.getRange(row + 9, 1).setValue('Date:').setFontWeight('bold');
  formatInputCell(sheet.getRange(row + 9, 2, 1, 2).merge());
  formatDate(sheet.getRange(row + 9, 2));
  
  sheet.getRange(row + 10, 1).setValue('Reviewed by:').setFontWeight('bold');
  formatInputCell(sheet.getRange(row + 10, 2, 1, 2).merge());
  
  sheet.getRange(row + 11, 1).setValue('Date:').setFontWeight('bold');
  formatInputCell(sheet.getRange(row + 11, 2, 1, 2).merge());
  formatDate(sheet.getRange(row + 11, 2));
  
  sheet.getRange(row + 13, 1).setValue('Conclusion:').setFontWeight('bold');
  formatInputCell(sheet.getRange(row + 13, 2, 4, 5).merge());
  
  // Final notes
  sheet.getRange('A' + (row + 18)).setValue('AUDIT TRAIL NOTE:').setFontWeight('bold').setFontColor('#303f9f');
  sheet.getRange('A' + (row + 19) + ':F' + (row + 22)).merge()
       .setValue('This working paper provides a complete audit trail for all Ind AS 109 period-end adjustments:\n\n' +
                 '• All formulas are transparent and traceable\n' +
                 '• Control totals ensure mathematical accuracy\n' +
                 '• All judgments and estimates are documented\n' +
                 '• Evidence supporting key assumptions is referenced\n' +
                 '• This workbook should be retained as part of the audit file')
       .setWrap(true)
       .setVerticalAlignment('top')
       .setBackground('#e8eaf6')
       .setBorder(true, true, true, true, false, false);
  
  sheet.setFrozenRows(2);
  
  Logger.log('Audit Notes sheet created.');
}

// ═══════════════════════════════════════════════════════════════════════════
// NAMED RANGES SETUP
// ═══════════════════════════════════════════════════════════════════════════

function setupNamedRanges(ss) {
  try {
    // Input Variables
    ss.setNamedRange('ReportingDate', ss.getRange('Input_Variables!B4'));
    ss.setNamedRange('RiskFreeRate', ss.getRange('Input_Variables!B6'));
    ss.setNamedRange('PD_Stage1', ss.getRange('Input_Variables!B10'));
    ss.setNamedRange('PD_Stage2', ss.getRange('Input_Variables!B11'));
    ss.setNamedRange('PD_Stage3', ss.getRange('Input_Variables!B12'));
    
    // Instruments
    ss.setNamedRange('InstrumentsList', ss.getRange('Instruments_Register!A3:P250'));
    
    Logger.log('Named ranges created successfully.');
  } catch (error) {
    Logger.log('Error creating named ranges: ' + error);
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// FINALIZATION FUNCTION
// ═══════════════════════════════════════════════════════════════════════════

function finalizeWorkingPapers(ss) {
  /**
   * Apply professional formatting to all sheets per 109 guide standards
   * - Hide gridlines for clean, sleek professional appearance
   * - Set consistent fonts and sizing
   * - Freeze header rows
   */
  ss.getSheets().forEach(sheet => {
    // Hide gridlines for professional clean appearance
    sheet.setHiddenGridlines(true);

    // Set professional Arial font throughout
    sheet.getDataRange().setFontFamily("Arial").setFontSize(10);

    // Freeze top rows for better navigation
    if (sheet.getMaxRows() > 3) {
      sheet.setFrozenRows(3);
    }
  });

  Logger.log('Working papers finalized - gridlines hidden for professional appearance');
}

// ═══════════════════════════════════════════════════════════════════════════
// MENU FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════

// onOpen() is handled by common/utilities.gs - auto-detects workbook type

function refreshFormulas() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Recalculating all formulas...', 'Refresh', 3);
  SpreadsheetApp.flush();
  SpreadsheetApp.getActiveSpreadsheet().toast('All formulas refreshed!', 'Complete', 2);
}

function exportJournalEntries() {
  SpreadsheetApp.getUi().alert(
    'Export Journal Entries',
    'Navigate to the "Period_End_Entries" sheet and copy the journal entries.\n\n' +
    'You can then paste them into your accounting system or export as CSV.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function showHelp() {
  const helpText = `
IND AS 109 AUDIT BUILDER - USER GUIDE

GETTING STARTED:
1. Fill the "Input_Variables" sheet with your parameters
2. Complete the "Instruments_Register" with all financial instruments
3. Review auto-populated sheets for accuracy
4. Extract journal entries from "Period_End_Entries" sheet

KEY FEATURES:
• Automatic classification per Ind AS 109 logic
• Three-stage ECL impairment calculations
• Fair value adjustments for FVTPL and FVOCI
• EIR-based amortization schedules
• Complete audit trail with control checks

COMPLIANCE:
All calculations follow Ind AS 109 requirements including:
• Classification (SPPI test + Business model)
• Measurement (Amortized cost, FVTPL, FVOCI)
• Impairment (Expected Credit Loss - 3 stages)
• Derecognition criteria

SUPPORT:
Review the "References" sheet for Ind AS 109 provisions.
Check "Audit_Notes" for control totals and assertions.

For complex instruments, consult with qualified auditors.
  `;
  
  SpreadsheetApp.getUi().alert('Help & Documentation', helpText, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ═══════════════════════════════════════════════════════════════════════════
// END OF SCRIPT
// ═══════════════════════════════════════════════════════════════════════════