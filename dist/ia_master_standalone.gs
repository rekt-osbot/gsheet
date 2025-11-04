/**
 * @name ia_master
 * @version 1.1.0
 * @built 2025-11-04T10:11:10.760Z
 * @description Standalone script. Do not edit directly - edit source files in src/ folder.
 * 
 * This file is auto-generated from:
 * - Common utilities (src/common/*.gs)
 * - Workbook-specific code (src/workbooks/ia_master.gs)
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
 * Internal Audit Master Program - FY2025-26
 * Central coordination workbook for entire IA programme
 */

function createIAMasterWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setWorkbookType('IA_MASTER');
  clearExistingSheets(ss);
  
  createCoverSheet_IAMaster(ss);
  createTeamAllocationSheet_IA(ss);
  createWorkpaperIndexSheet_IA(ss);
  createProgressDashboardSheet_IA(ss);
  createFindingsTrackerSheet_IA(ss);
  
  const tempSheet = ss.getSheetByName('_temp_sheet_');
  if (tempSheet) ss.deleteSheet(tempSheet);
  ss.getSheetByName('Cover').activate();
  
  SpreadsheetApp.getUi().alert('IA Master Program Created!', 'Ready to track all audit phases.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function createCoverSheet_IAMaster(ss) {
  const sheet = getOrCreateSheet(ss, 'Cover', 0, '#1a237e');
  createStandardHeader(sheet, 'INTERNAL AUDIT PROGRAMME FY2025-26', 'SNVA TravelTech Pvt. Ltd.', 1, 4);
  
  const inputs = [
    {label: 'Audit Period:', value: 'April 2025 - March 2026'},
    {label: 'Team:', value: '3 Members + IA Manager'},
    {label: 'Total Tests:', value: '163 (H1: 69, Q3: 33, Q4: 61)'},
    {label: 'Last Updated:', value: '=TODAY()', type: 'date'}
  ];
  createInputSection(sheet, 4, 1, 2, inputs);
  
  createSectionHeader(sheet, 9, 'AUDIT PHASES', 1, 4);
  const phaseData = [
    ['H1 Review', 'Oct-Dec 2025', 'Revenue, P2P, Treasury, Tax, Systems', '69'],
    ['Q3 Review', 'Jan-Feb 2026', 'Payroll & HR', '33'],
    ['Q4 Review', 'Apr-May 2026', 'Fixed Assets, Close, IFC', '61']
  ];
  createDataTable(sheet, 10, 1, ['Phase', 'Timeline', 'Focus', 'Tests'], phaseData);
  
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 300);
  sheet.setColumnWidth(4, 80);
}

function createTeamAllocationSheet_IA(ss) {
  const sheet = getOrCreateSheet(ss, 'Team Allocation', 1, '#283593');
  createStandardHeader(sheet, 'TEAM ALLOCATION MATRIX', '', 1, 5);
  
  let row = 3;
  createSectionHeader(sheet, row++, 'H1 REVIEW', 1, 5);
  const h1 = [
    ['Revenue/OTC', 'PRIMARY', '', '', 'H1-REV'],
    ['P2P', '', 'PRIMARY', '', 'H1-P2P'],
    ['Taxation', '', 'PRIMARY', '', 'H1-TAX'],
    ['Treasury', 'PRIMARY', '', '', 'H1-TRY'],
    ['Systems', '', '', 'PRIMARY', 'H1-SYS']
  ];
  createDataTable(sheet, row, 1, ['Area', 'TM1', 'TM2', 'TM3', 'Prefix'], h1);
  row += h1.length + 2;
  
  createSectionHeader(sheet, row++, 'Q3 REVIEW', 1, 5);
  const q3 = [['Payroll/HR', '', '', 'PRIMARY', 'Q3-PAY']];
  createDataTable(sheet, row, 1, ['Area', 'TM1', 'TM2', 'TM3', 'Prefix'], q3);
  row += q3.length + 2;
  
  createSectionHeader(sheet, row++, 'Q4 REVIEW', 1, 5);
  const q4 = [
    ['Fixed Assets', '', 'PRIMARY', '', 'Q4-FAR'],
    ['Close/MIS', 'PRIMARY', '', '', 'Q4-FSC'],
    ['IFC', '', '', 'PRIMARY', 'Q4-IFC']
  ];
  createDataTable(sheet, row, 1, ['Area', 'TM1', 'TM2', 'TM3', 'Prefix'], q4);
  
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);
}

function createWorkpaperIndexSheet_IA(ss) {
  const sheet = getOrCreateSheet(ss, 'Workpaper Index', 2, '#3949ab');
  createStandardHeader(sheet, 'WORKPAPER INDEX', '', 1, 7);
  
  const headers = ['WP Index', 'Phase', 'Area', 'Description', 'By', 'Status', 'Date'];
  createDataTable(sheet, 3, 1, headers, []);
  
  applyValidationList(sheet.getRange('F4:F200'), 'STATUS');
  applyValidationList(sheet.getRange('E4:E200'), 'TEAM_MEMBER');
  
  freezeHeaders(sheet, 3);
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 300);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 100);
}

function createProgressDashboardSheet_IA(ss) {
  const sheet = getOrCreateSheet(ss, 'Progress Dashboard', 3, '#5e35b1');
  createStandardHeader(sheet, 'PROGRESS DASHBOARD', '', 1, 5);
  
  let row = 3;
  createSectionHeader(sheet, row++, 'OVERALL STATUS', 1, 3);
  const overall = [
    ['H1 Review', 'Not Started', '0%'],
    ['Q3 Review', 'Not Started', '0%'],
    ['Q4 Review', 'Not Started', '0%']
  ];
  createDataTable(sheet, row, 1, ['Phase', 'Status', '%'], overall);
  row += overall.length + 2;
  
  createSectionHeader(sheet, row++, 'H1 PROGRESS', 1, 5);
  const h1 = [
    ['Revenue/OTC', '13', '0', '13'],
    ['P2P', '14', '0', '14'],
    ['Taxation', '12', '0', '12'],
    ['Treasury', '8', '0', '8'],
    ['Systems', '12', '0', '12']
  ];
  createDataTable(sheet, row, 1, ['Area', 'Total', 'Done', 'Pending'], h1);
  
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);
}

function createFindingsTrackerSheet_IA(ss) {
  const sheet = getOrCreateSheet(ss, 'Findings Tracker', 4, '#6a1b9a');
  createStandardHeader(sheet, 'FINDINGS TRACKER', '', 1, 9);
  
  const headers = ['ID', 'Phase', 'Area', 'Finding', 'Severity', 'Recommendation', 'Response', 'Owner', 'Status'];
  createDataTable(sheet, 3, 1, headers, []);
  
  applyValidationList(sheet.getRange('E4:E200'), 'SEVERITY');
  applyValidationList(sheet.getRange('I4:I200'), 'FINDING_STATUS');
  
  addConditionalFormatting(sheet, 'E4:E200', 'Critical', '#ea4335', '#ffffff');
  addConditionalFormatting(sheet, 'E4:E200', 'High', '#f4cccc', '#000000');
  
  freezeHeaders(sheet, 3);
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 250);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 200);
  sheet.setColumnWidth(7, 200);
  sheet.setColumnWidth(8, 120);
  sheet.setColumnWidth(9, 100);
}
