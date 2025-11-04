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
