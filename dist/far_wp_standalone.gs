/**
 * @name far_wp
 * @version 1.1.0
 * @built 2025-11-04T10:11:10.758Z
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
 * 7. Click "Audit Tools" > "Create Fixed Assets Workbook"
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

/**
 * Main function to setup the entire Fixed Assets audit workpaper
 */
function createFixedAssetsWorkbook() {
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

  // Clear existing sheets using standardized utility function
  clearExistingSheets(ss);

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
  const sheet = getOrCreateSheet(ss, "FA-Index", 0, '#1f4e78');
  
  // Set column widths
  setColumnWidths(sheet, [50, 350, 150, 250]);
  
  // Main header
  formatHeader(sheet, 1, 1, 4, "FIXED ASSETS AUDIT WORKPAPER", COLORS.HEADER_BG);
  
  // Client information section
  const clientInputs = [
    {label: "Client Name:", type: "text"},
    {label: "Engagement:", type: "text"},
    {label: "Date:", type: "date"}
  ];

  // Add period end, preparer, and reviewer fields with proper formatting
  clientInputs.forEach((input, index) => {
    const row = 3 + index;
    const labelRange = sheet.getRange(row, 1);
    const valueRange = sheet.getRange(row, 2);
    const label2Range = sheet.getRange(row, 3);
    const value2Range = sheet.getRange(row, 4);

    // Format input cells
    safeRangeFormat(labelRange, {fontWeight: "bold", background: COLORS.INPUT_BG});
    safeRangeFormat(valueRange, {background: "#ffffff"});
    safeRangeFormat(label2Range, {fontWeight: "bold", background: COLORS.INPUT_BG});
    safeRangeFormat(value2Range, {background: "#ffffff"});

    labelRange.setValue(input.label);
    label2Range.setValue(index === 0 ? "Period End:" : (index === 1 ? "Prepared By:" : "Reviewed By:"));
  });
  
  // Table of Contents Header
  createSectionHeader(sheet, 7, "TABLE OF CONTENTS", 1, 4);
  
  // Index data
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
  
  createDataTable(sheet, 8, 1, ["Ref", "Workpaper Description", "Preparer", "Reviewer"], indexData, {
    borders: true,
    headerBg: COLORS.SUBHEADER_BG
  });
  
  // Format reference column using safe formatting
  const refColRange = sheet.getRange(9, 1, indexData.length, 1);
  safeRangeFormat(refColRange, {background: COLORS.CALC_BG, fontWeight: "bold"});
  
  // Freeze header rows
  freezeHeaders(sheet, 8);
}

/**
 * Creates the Summary & Conclusion sheet (FA-1)
 */
function createSummarySheet(ss) {
  const sheet = getOrCreateSheet(ss, "FA-1 Summary", null, '#4472c4');
  
  // Set column widths
  setColumnWidths(sheet, [100, 300, 150, 150, 150]);
  
  // Header
  createWorkpaperHeader(sheet, "FA-1", "FIXED ASSETS - SUMMARY & CONCLUSION");
  
  // Summary of Balances
  createSectionHeader(sheet, 5, "SUMMARY OF FIXED ASSETS BALANCES", 1, 5);
  
  const summaryData = [
    ["Gross Fixed Assets", safeFormula("'FA-2 Roll Forward'!E7", "0"), safeFormula("'FA-2 Roll Forward'!F7", "0"), safeFormula("'FA-2 Roll Forward'!G7", "0"), safeFormula("'FA-2 Roll Forward'!H7", "0")],
    ["Accumulated Depreciation", safeFormula("'FA-3 Depreciation'!E15", "0"), safeFormula("'FA-3 Depreciation'!F15", "0"), safeFormula("'FA-3 Depreciation'!G15", "0"), safeFormula("'FA-3 Depreciation'!H15", "0")],
    ["Net Fixed Assets", "=B7-B8", "=C7-C8", "=D7-D8", "=E7-E8"]
  ];
  
  createDataTable(sheet, 6, 1, ["Description", "Beginning Balance", "Additions", "Disposals", "Ending Balance"], summaryData, {borders: true});
  
  // Format numbers and totals row
  formatCurrency(sheet.getRange("B7:E9"));
  const totalsRange = sheet.getRange("B9:E9");
  safeRangeFormat(totalsRange, {fontWeight: "bold", background: COLORS.TOTAL_BG});
  
  // Audit Procedures Summary
  createSectionHeader(sheet, 12, "AUDIT PROCEDURES PERFORMED", 1, 5);
  
  const procedures = [
    ["1", "Obtained and reviewed fixed asset roll forward", "FA-2", ""],
    ["2", "Tested additions for proper authorization and capitalization", "FA-4", ""],
    ["3", "Tested disposals and verified removal from records", "FA-5", ""],
    ["4", "Performed physical verification of selected assets", "FA-6", ""],
    ["5", "Tested completeness of fixed asset recording", "FA-7", ""],
    ["6", "Recalculated depreciation expense", "FA-3", ""],
    ["7", "Reviewed financial statement presentation and disclosure", "FA-8", ""]
  ];
  
  createDataTable(sheet, 13, 1, ["#", "Procedure", "Ref", "Conclusion"], procedures, {borders: true});
  
  // Conclusion section
  const conclusionRow = 13 + procedures.length + 2;
  createSectionHeader(sheet, conclusionRow, "AUDIT CONCLUSION", 1, 5);
  
  sheet.getRange(conclusionRow + 1, 1, 1, 5).merge()
    .setValue("Based on the audit procedures performed, we conclude that:")
    .setWrap(true);
  
  sheet.getRange(conclusionRow + 2, 1, 4, 5).merge()
    .setValue("[Enter conclusion here - e.g., 'Fixed assets are fairly stated in all material respects...']")
    .setWrap(true)
    .setVerticalAlignment("top")
    .setBackground(COLORS.INPUT_BG);
  sheet.setRowHeights(conclusionRow + 2, 4, 25);
  
  // Sign-off
  createSignOffSection(sheet, conclusionRow + 7, 1);
}

/**
 * Creates the Roll Forward sheet (FA-2)
 */
function createRollForwardSheet(ss) {
  const sheet = getOrCreateSheet(ss, "FA-2 Roll Forward", null, '#5b9bd5');
  
  // Set column widths
  setColumnWidths(sheet, [50, 200, 100, 120, 120, 120, 120, 120]);
  
  createWorkpaperHeader(sheet, "FA-2", "FIXED ASSETS ROLL FORWARD");
  
  // Asset categories with formulas
  const categories = [
    ["", "Land", "N/A", "Land - not depreciated"],
    ["", "Buildings", "39 years", "Office buildings and improvements"],
    ["", "Machinery & Equipment", "5-10 years", "Manufacturing equipment"],
    ["", "Furniture & Fixtures", "7 years", "Office furniture and fixtures"],
    ["", "Vehicles", "5 years", "Company vehicles"],
    ["", "Computer Equipment", "3-5 years", "Computers, servers, IT equipment"],
    ["", "Leasehold Improvements", "Lease term", "Improvements to leased property"]
  ];
  
  createDataTable(sheet, 5, 1, 
    ["Ref", "Asset Category", "Useful Life", "Description", "Beginning Balance", "Additions", "Disposals", "Ending Balance"],
    categories, 
    {borders: true, headerHeight: 40}
  );
  
  // Add ending balance formulas
  for (let i = 0; i < categories.length; i++) {
    const row = 6 + i;
    sheet.getRange(row, 8).setFormula(safeFormula(`E${row}+F${row}-G${row}`, "0"));
  }
  
  // Total row
  const totalRow = 6 + categories.length;
  createTotalsSection(sheet, totalRow, 2, [
    {label: "TOTAL GROSS FIXED ASSETS", formula: safeSumFormula(`E6:E${totalRow-1}`), format: 'currency'},
  ], '');
  
  sheet.getRange(totalRow, 6).setFormula(safeSumFormula(`F6:F${totalRow-1}`))
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
  const sheet = getOrCreateSheet(ss, "FA-3 Depreciation", null, '#70ad47');
  setColumnWidths(sheet, Array(8).fill(120));
  
  createWorkpaperHeader(sheet, "FA-3", "ACCUMULATED DEPRECIATION & DEPRECIATION EXPENSE");
  
  const categories = [
    ["Land", "N/A"],
    ["Buildings", "Straight-Line"],
    ["Machinery & Equipment", "Straight-Line"],
    ["Furniture & Fixtures", "Straight-Line"],
    ["Vehicles", "Straight-Line"],
    ["Computer Equipment", "Straight-Line"],
    ["Leasehold Improvements", "Straight-Line"]
  ];
  
  createDataTable(sheet, 5, 1,
    ["Asset Category", "Method", "Beginning Balance", "Current Year Expense", "Disposals", "Ending Balance", "Recalc", "Variance"],
    categories, {borders: true, headerHeight: 40}
  );
  
  // Add formulas
  for (let i = 0; i < categories.length; i++) {
    const row = 6 + i;
    sheet.getRange(row, 6).setFormula(safeFormula(`C${row}+D${row}-E${row}`, "0"));
    sheet.getRange(row, 8).setFormula(safeFormula(`F${row}-G${row}`, "0"));
  }
  
  // Total row
  const totalRow = 6 + categories.length;
  const totals = [
    {label: "TOTAL ACCUMULATED DEPRECIATION", formula: safeSumFormula(`C6:C${totalRow-1}`), format: 'currency'}
  ];
  
  sheet.getRange(totalRow, 1).setValue("TOTAL ACCUMULATED DEPRECIATION").setFontWeight("bold");
  for (let col = 3; col <= 8; col++) {
    sheet.getRange(totalRow, col).setFormula(safeSumFormula(`${String.fromCharCode(64+col)}6:${String.fromCharCode(64+col)}${totalRow-1}`))
      .setFontWeight("bold");
  }
  sheet.getRange(totalRow, 1, 1, 8).setBackground(COLORS.TOTAL_BG);
  
  // Depreciation calculation section
  const calcRow = totalRow + 2;
  createSectionHeader(sheet, calcRow, "DEPRECIATION EXPENSE RECALCULATION", 1, 8);
  
  createDataTable(sheet, calcRow + 1, 1,
    ["Asset Category", "Gross Assets", "Useful Life", "Method", "Calculated Expense", "Per Client", "Variance", "Notes"],
    [], {borders: true}
  );
  
  formatCurrency(sheet.getRange(6, 3, categories.length + 1, 6));
  
  createSignOffSection(sheet, calcRow + 12, 1);
  freezeHeaders(sheet, 5);
}

/**
 * Creates the Additions Testing sheet (FA-4)
 */
function createAdditionsTestingSheet(ss) {
  const sheet = getOrCreateSheet(ss, "FA-4 Additions", null, '#ffc000');
  setColumnWidths(sheet, Array(10).fill(110));
  
  createWorkpaperHeader(sheet, "FA-4", "ADDITIONS TESTING");
  
  // Testing objective
  createInstructionsSection(sheet, 5, 1, 10, "OBJECTIVE", 
    "Test additions to verify proper authorization, occurrence, and capitalization");
  
  // Sample selection
  createSectionHeader(sheet, 7, "SAMPLE SELECTION", 1, 3);
  
  const sampleInputs = [
    {label: "Total Additions:", value: safeFormula("'FA-2 Roll Forward'!F13", "0"), type: 'currency'},
    {label: "Sample Size:", value: 25, type: 'number'},
    {label: "Sample Coverage:", value: safeFormula("SUM(G13:G37)/B8", "0"), type: 'percentage'}
  ];
  
  createInputSection(sheet, 8, 1, 2, sampleInputs);
  
  // Testing table
  const sampleCount = 25;
  createDataTable(sheet, 12, 1, 
    ["Date", "Description", "Category", "Vendor", "Invoice #", "Amount", "Authorization", "Capitalization", "Classification", "Conclusion"],
    [], 
    {borders: true, headerHeight: 40}
  );
  
  // Apply validations
  applyMultipleValidations(sheet, [
    {range: `G13:I${12+sampleCount}`, type: 'CHECK_MARKS'},
    {range: `J13:J${12+sampleCount}`, type: 'PASS_FAIL_NOTE'}
  ]);
  
  // Format amount column
  formatCurrency(sheet.getRange(13, 6, sampleCount, 1));
  
  // Total row
  const totalRow = 13 + sampleCount;
  createTotalsSection(sheet, totalRow, 1, [
    {label: "TOTAL TESTED", formula: safeSumFormula(`F13:F${totalRow-1}`), format: 'currency'}
  ], '');
  
  // Exceptions section
  createInstructionsSection(sheet, totalRow + 2, 1, 10, "EXCEPTIONS & NOTES",
    "[Document any exceptions, unusual items, or additional notes]");
  
  createSignOffSection(sheet, totalRow + 6, 1);
  freezeHeaders(sheet, 12);
}

/**
 * Creates the Disposals Testing sheet (FA-5)
 */
function createDisposalsTestingSheet(ss) {
  const sheet = getOrCreateSheet(ss, "FA-5 Disposals", null, '#f4b084');
  setColumnWidths(sheet, Array(10).fill(110));
  
  createWorkpaperHeader(sheet, "FA-5", "DISPOSALS & RETIREMENTS TESTING");
  
  // Testing objective
  createInstructionsSection(sheet, 5, 1, 10, "OBJECTIVE",
    "Test disposals to verify proper authorization, removal from records, and gain/loss calculation");
  
  // Sample selection
  createSectionHeader(sheet, 7, "SAMPLE SELECTION", 1, 3);
  
  const sampleInputs = [
    {label: "Total Disposals:", value: safeFormula("'FA-2 Roll Forward'!G13", "0"), type: 'currency'},
    {label: "Sample Size:", value: 15, type: 'number'}
  ];
  
  createInputSection(sheet, 8, 1, 2, sampleInputs);
  
  // Testing table
  const sampleCount = 15;
  createDataTable(sheet, 12, 1,
    ["Date", "Asset Description", "Category", "Original Cost", "Accum. Depr.", "Net Book Value", "Proceeds", "Gain/(Loss)", "Authorization", "Conclusion"],
    [],
    {borders: true, headerHeight: 40}
  );
  
  // Add formulas for calculated columns
  for (let i = 0; i < sampleCount; i++) {
    const row = 13 + i;
    sheet.getRange(row, 6).setFormula(safeFormula(`D${row}-E${row}`, "0"));
    sheet.getRange(row, 8).setFormula(safeFormula(`G${row}-F${row}`, "0"));
  }
  
  // Apply validations using common helpers
  applyMultipleValidations(sheet, [
    {range: `I13:I${12+sampleCount}`, type: 'CHECK_MARKS'},
    {range: `J13:J${12+sampleCount}`, type: 'PASS_FAIL_NOTE'}
  ]);

  // Format numbers as currency
  formatCurrency(sheet.getRange(13, 4, sampleCount, 5));
  
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
  const sheet = getOrCreateSheet(ss, "FA-6 Existence", null, '#9dc3e6');
  setColumnWidths(sheet, Array(9).fill(120));
  
  createWorkpaperHeader(sheet, "FA-6", "PHYSICAL EXISTENCE VERIFICATION");
  
  createInstructionsSection(sheet, 5, 1, 9, "OBJECTIVE",
    "Verify physical existence of selected fixed assets");
  
  createSectionHeader(sheet, 7, "SAMPLE SELECTION", 1, 3);
  
  const sampleInputs = [
    {label: "Total Fixed Assets:", value: safeFormula("'FA-2 Roll Forward'!H13", "0"), type: 'currency'},
    {label: "Items Selected:", value: 30, type: 'number'}
  ];
  
  createInputSection(sheet, 8, 1, 2, sampleInputs);
  
  const sampleCount = 30;
  createDataTable(sheet, 12, 1,
    ["Asset ID", "Description", "Category", "Location", "Book Value", "Observed?", "Condition", "Tag #", "Notes"],
    [], {borders: true, headerHeight: 40}
  );
  
  applyMultipleValidations(sheet, [
    {range: `F13:F${12+sampleCount}`, type: 'LOCATION_STATUS'},
    {range: `G13:G${12+sampleCount}`, type: 'CONDITION_PHYSICAL'}
  ]);
  
  formatCurrency(sheet.getRange(13, 5, sampleCount, 1));
  
  const summaryRow = 13 + sampleCount + 2;
  createSectionHeader(sheet, summaryRow, "VERIFICATION SUMMARY", 1, 9);
  
  const summaryInputs = [
    {label: "Assets Physically Verified:", value: safeFormula(`COUNTIF(F13:F${13+sampleCount-1},"✓ Yes")`, "0")},
    {label: "Assets Not Located:", value: safeFormula(`COUNTIF(F13:F${13+sampleCount-1},"Unable to locate")`, "0")},
    {label: "Verification Rate:", value: safeFormula(`B${summaryRow+1}/${sampleCount}`, "0"), type: 'percentage'}
  ];
  
  createInputSection(sheet, summaryRow + 1, 1, 2, summaryInputs);
  
  createSignOffSection(sheet, summaryRow + 6, 1);
  freezeHeaders(sheet, 12);
}

/**
 * Creates the Completeness Testing sheet (FA-7)
 */
function createCompletenessTestingSheet(ss) {
  const sheet = getOrCreateSheet(ss, "FA-7 Completeness", null, '#a9d08e');
  setColumnWidths(sheet, Array(8).fill(130));
  
  createWorkpaperHeader(sheet, "FA-7", "COMPLETENESS TESTING");
  
  createInstructionsSection(sheet, 5, 1, 8, "OBJECTIVE",
    "Test that all qualifying expenditures have been properly capitalized");
  
  // Procedure 1
  createSectionHeader(sheet, 7, "PROCEDURE 1: REVIEW REPAIR & MAINTENANCE EXPENSES", 1, 8);
  
  const repairRows = 15;
  createDataTable(sheet, 8, 1,
    ["Date", "Vendor", "Description", "Amount", "Nature", "Capitalize?", "Adjustment", "Notes"],
    [], {borders: true}
  );
  
  applyMultipleValidations(sheet, [
    {range: `E9:E${8+repairRows}`, type: 'REPAIR_TYPE'},
    {range: `F9:F${8+repairRows}`, type: 'YES_NO'}
  ]);
  
  formatCurrency(sheet.getRange(9, 4, repairRows, 1));
  formatCurrency(sheet.getRange(9, 7, repairRows, 1));
  
  // Procedure 2
  const cipRow = 9 + repairRows + 2;
  createSectionHeader(sheet, cipRow, "PROCEDURE 2: CONSTRUCTION IN PROGRESS REVIEW", 1, 8);
  
  const cipRows = 10;
  createDataTable(sheet, cipRow + 1, 1,
    ["Project", "Start Date", "Status", "Costs to Date", "Ready for Use?", "Transfer to FA?", "Notes", ""],
    [], {borders: true}
  );
  
  applyMultipleValidations(sheet, [
    {range: `C${cipRow+2}:C${cipRow+1+cipRows}`, type: 'custom', values: ['In Progress', 'Complete', 'On Hold']},
    {range: `E${cipRow+2}:E${cipRow+1+cipRows}`, type: 'custom', values: ['Yes', 'No', 'Partial']},
    {range: `F${cipRow+2}:F${cipRow+1+cipRows}`, type: 'YES_NO_NA'}
  ]);
  
  formatCurrency(sheet.getRange(cipRow + 2, 4, cipRows, 1));
  
  createSignOffSection(sheet, cipRow + cipRows + 4, 1);
  freezeHeaders(sheet, 8);
}

/**
 * Creates the Disclosure sheet (FA-8)
 */
function createDisclosureSheet(ss) {
  const sheet = getOrCreateSheet(ss, "FA-8 Disclosure", null, '#c5e0b4');
  setColumnWidths(sheet, Array(5).fill(200));
  
  createWorkpaperHeader(sheet, "FA-8", "PRESENTATION & DISCLOSURE CHECKLIST");
  
  createSectionHeader(sheet, 5, "DISCLOSURE REQUIREMENTS CHECKLIST", 1, 5);
  
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
  
  createDataTable(sheet, 6, 1, ["Requirement", "Yes/No/N/A", "Reference", "Notes", ""], disclosures, {borders: true});
  
  applyValidationList(sheet.getRange(`B7:B${6+disclosures.length}`), 'YES_NO_NA');
  
  // Financial Statement Presentation
  const fsRow = 7 + disclosures.length + 2;
  createSectionHeader(sheet, fsRow, "FINANCIAL STATEMENT PRESENTATION", 1, 5);
  
  const fsData = [
    ["Fixed Assets (Gross)", safeFormula("'FA-2 Roll Forward'!H13", "0"), "=", ""],
    ["Less: Accumulated Depreciation", safeFormula("'FA-3 Depreciation'!F15", "0"), "=", ""],
    ["Fixed Assets (Net)", `=B${fsRow+1}-B${fsRow+2}`, "=", ""],
    ["", "", "", ""],
    ["Depreciation Expense (P&L)", safeFormula("'FA-3 Depreciation'!D15", "0"), "=", ""]
  ];
  
  createDataTable(sheet, fsRow + 1, 1, ["Description", "Amount", "Tie", "Notes"], fsData, {borders: true});
  
  formatCurrency(sheet.getRange(fsRow + 1, 2, fsData.length, 1));
  formatCurrency(sheet.getRange(fsRow + 1, 4, fsData.length, 1));
  
  createSignOffSection(sheet, fsRow + 8, 1);
  freezeHeaders(sheet, 6);
}

/**
 * Creates the Conclusion sheet (FA-9)
 */
function createConclusionSheet(ss) {
  const sheet = getOrCreateSheet(ss, "FA-9 Conclusion", null, '#70ad47');
  setColumnWidths(sheet, Array(6).fill(150));
  
  createWorkpaperHeader(sheet, "FA-9", "AUDIT CONCLUSION & SIGN-OFF");
  
  createSectionHeader(sheet, 5, "SUMMARY OF AUDIT PROCEDURES", 1, 6);
  
  const procedures = [
    ["Obtained and agreed fixed asset roll forward to general ledger", "FA-2", "✓"],
    ["Tested additions for proper authorization and capitalization", "FA-4", ""],
    ["Tested disposals and verified gain/loss calculations", "FA-5", ""],
    ["Performed physical verification of selected assets", "FA-6", ""],
    ["Tested completeness of fixed asset recording", "FA-7", ""],
    ["Recalculated depreciation expense", "FA-3", ""],
    ["Reviewed financial statement presentation and disclosure", "FA-8", ""]
  ];
  
  createDataTable(sheet, 6, 1, ["Procedure", "Reference", "Complete"], procedures, {borders: true});
  
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
 * Uses safe formatting helpers from common/errorHandling.gs
 */
function createWorkpaperHeader(sheet, reference, title) {
  // Title row with safe formatting
  const titleRange = sheet.getRange("A1:H1");
  titleRange.merge();
  safeRangeFormat(titleRange, {
    fontSize: FONT_SIZES.title,
    fontWeight: "bold",
    background: COLORS.header,
    fontColor: "#ffffff"
  });
  // Set alignment after safe format
  titleRange.setHorizontalAlignment("center");
  titleRange.setVerticalAlignment("middle");
  titleRange.setValue(title);
  sheet.setRowHeight(1, 35);

  // Reference and metadata with safe formatting
  const refLabelRange = sheet.getRange("A2");
  const refValueRange = sheet.getRange("B2");
  const prepByLabelRange = sheet.getRange("D2");
  const prepByValueRange = sheet.getRange("E2");
  const dateLabelRange = sheet.getRange("F2");
  const dateValueRange = sheet.getRange("G2");

  refLabelRange.setValue("Reference:");
  safeRangeFormat(refLabelRange, {fontWeight: "bold"});

  refValueRange.setValue(reference);
  safeRangeFormat(refValueRange, {fontWeight: "bold", background: COLORS.referenceCell});

  prepByLabelRange.setValue("Prepared By:");
  safeRangeFormat(prepByLabelRange, {fontWeight: "bold"});
  safeRangeFormat(prepByValueRange, {background: "#ffffff"});

  dateLabelRange.setValue("Date:");
  safeRangeFormat(dateLabelRange, {fontWeight: "bold"});
  safeRangeFormat(dateValueRange, {background: "#ffffff"});

  // Row 3: Client and Reviewed By
  const clientLabelRange = sheet.getRange("A3");
  const clientValueRange = sheet.getRange("B3:C3");
  const revByLabelRange = sheet.getRange("D3");
  const revByValueRange = sheet.getRange("E3");
  const dateLabelRange2 = sheet.getRange("F3");
  const dateValueRange2 = sheet.getRange("G3");

  clientLabelRange.setValue("Client:");
  safeRangeFormat(clientLabelRange, {fontWeight: "bold"});

  clientValueRange.merge();
  safeRangeFormat(clientValueRange, {background: "#ffffff"});

  revByLabelRange.setValue("Reviewed By:");
  safeRangeFormat(revByLabelRange, {fontWeight: "bold"});
  safeRangeFormat(revByValueRange, {background: "#ffffff"});

  dateLabelRange2.setValue("Date:");
  safeRangeFormat(dateLabelRange2, {fontWeight: "bold"});
  safeRangeFormat(dateValueRange2, {background: "#ffffff"});
}

/**
 * Helper function to create sign-off section
 * Uses safe formatting helpers from common/errorHandling.gs
 */
function createSignOffSection(sheet, startRow) {
  // Header row
  const headerRange = sheet.getRange(startRow, 1, 1, 8);
  headerRange.merge();
  safeRangeFormat(headerRange, {
    fontWeight: "bold",
    background: COLORS.sectionHeader,
    fontColor: "#ffffff"
  });
  headerRange.setHorizontalAlignment("center");
  headerRange.setValue("PREPARER & REVIEWER SIGN-OFF");

  // Prepared By row
  startRow++;
  const prepByLabelRange = sheet.getRange(startRow, 1);
  const prepByValueRange = sheet.getRange(startRow, 2);
  const dateLabel1Range = sheet.getRange(startRow, 3);
  const dateValue1Range = sheet.getRange(startRow, 4);

  prepByLabelRange.setValue("Prepared By:");
  safeRangeFormat(prepByLabelRange, {fontWeight: "bold"});
  safeRangeFormat(prepByValueRange, {background: COLORS.preparer});

  dateLabel1Range.setValue("Date:");
  safeRangeFormat(dateLabel1Range, {fontWeight: "bold"});
  safeRangeFormat(dateValue1Range, {background: COLORS.preparer});

  // Reviewed By row
  startRow++;
  const revByLabelRange = sheet.getRange(startRow, 1);
  const revByValueRange = sheet.getRange(startRow, 2);
  const dateLabel2Range = sheet.getRange(startRow, 3);
  const dateValue2Range = sheet.getRange(startRow, 4);

  revByLabelRange.setValue("Reviewed By:");
  safeRangeFormat(revByLabelRange, {fontWeight: "bold"});
  safeRangeFormat(revByValueRange, {background: COLORS.reviewer});

  dateLabel2Range.setValue("Date:");
  safeRangeFormat(dateLabel2Range, {fontWeight: "bold"});
  safeRangeFormat(dateValue2Range, {background: COLORS.reviewer});
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