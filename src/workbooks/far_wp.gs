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