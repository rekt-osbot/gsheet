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
// UTILITY FUNCTIONS
// ============================================================================

function clearExistingSheets(ss) {
  const sheets = ss.getSheets();
  sheets.forEach(sheet => {
    if (sheet.getName() !== "Sheet1") {
      ss.deleteSheet(sheet);
    }
  });
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
    ["Deferred Tax Assets (DTA)", "=Movement_Analysis!F30", "=Movement_Analysis!C30", "=Movement_Analysis!I30", "=IF(Movement_Analysis!C30<>0,Movement_Analysis!I30/Movement_Analysis!C30,\"-\")"],
    ["Deferred Tax Liabilities (DTL)", "=Movement_Analysis!F31", "=Movement_Analysis!C31", "=Movement_Analysis!I31", "=IF(Movement_Analysis!C31<>0,Movement_Analysis!I31/Movement_Analysis!C31,\"-\")"],
    ["Net DTA/(DTL)", "=Movement_Analysis!F32", "=Movement_Analysis!C32", "=Movement_Analysis!I32", "=IF(Movement_Analysis!C32<>0,Movement_Analysis!I32/Movement_Analysis!C32,\"-\")"],
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
    ["Temp_Differences", "H7:H50", "Additions", "New temporary differences in current year", "Number", "50,000"],
    ["Temp_Differences", "I7:I50", "Reversals", "Reversals of prior period differences", "Number", "30,000"],
    ["Temp_Differences", "J7:J50", "Rate Change Impact", "Impact of tax rate changes", "Number", "0"],
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
  setColumnWidths(sheet, [40, 250, 150, 120, 120, 120, 80, 120, 120, 120, 150]);
  
  // Header
  sheet.getRange("A1:K1").merge()
    .setValue("TEMPORARY DIFFERENCES INPUT SCHEDULE")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT);
  
  sheet.getRange("A2:K2").merge()
    .setValue("Enter all temporary differences between book and tax basis of assets and liabilities")
    .setFontSize(10)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.SUBHEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT);
  
  // Instructions
  sheet.getRange("A4:K4").merge()
    .setValue("INSTRUCTIONS: Enter temporary differences line by line. System will auto-calculate DTA/DTL. Yellow cells are inputs.")
    .setFontSize(10)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.INFO_BG);
  
  // Column headers
  const headers = [
    ["Sr.", "Temporary Difference Item", "Category", "Tax Base (A)", "Book Base (B)", "Temp Diff (C=B-A)", "Nature", "Additions (CY)", "Reversals (CY)", "Rate Change Impact", "Remarks"]
  ];
  sheet.getRange(6, 1, 1, 11).setValues(headers);
  sheet.getRange("A6:K6").setBackground(COLORS.HEADER_BG)
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
      75000,
      25000,
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
      0,
      500000,
      0,
      "DTA subject to probability assessment"
    ]
  ];
  
  sheet.getRange(7, 1, sampleData.length, 11).setValues(sampleData);
  
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
  sheet.getRange("H7:J50").setBackground(COLORS.INPUT_BG);
  sheet.getRange("K7:K50").setBackground("#ffffff");
  
  // Number formatting
  sheet.getRange("D7:F50").setNumberFormat("#,##0");
  sheet.getRange("H7:J50").setNumberFormat("#,##0");
  
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
  
  // Borders
  sheet.getRange("A6:K52").setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  
  // Add cell notes for guidance
  sheet.getRange("D7").setNote("Tax Base:\nThe amount attributed to this item for tax purposes.\n\nFor assets: Tax WDV or NIL if fully written off\nFor liabilities: Tax-deductible amount");
  
  sheet.getRange("E7").setNote("Book Base:\nCarrying amount as per books/financial statements.\n\nFor assets: NBV as per books\nFor liabilities: Provision amount in books");
  
  sheet.getRange("F7").setNote("Temporary Difference:\nDifference between book base and tax base.\n\nPositive = Taxable temporary difference (creates DTL)\nNegative = Deductible temporary difference (creates DTA)");
  
  sheet.getRange("H7").setNote("Additions:\nNew temporary differences arising in the current year that will reverse in future periods.");
  
  sheet.getRange("I7").setNote("Reversals:\nReversal of temporary differences from prior years that reversed in the current year.");
  
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
    sheet.getRange(row, 10).setFormula(`=Temp_Differences!K${srcRow}`);
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
  
  // DTA Section
  sheet.getRange("A7:I7").merge()
    .setValue("DEFERRED TAX ASSETS")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(COLORS.SECTION_BG);
  
  const dtaItems = [
    [1, "Provision for Doubtful Debts", "=Assumptions!B28*0.2", "=SUMIF(Temp_Differences!G7:G50,\"DTA\",Temp_Differences!H7:H50)*Assumptions!B14*0.2", "=SUMIF(Temp_Differences!G7:G50,\"DTA\",Temp_Differences!I7:I50)*Assumptions!B14*0.2", "=C8+D8-E8", "=F8-C8", 0, "=I8"],
    [2, "Employee Benefits Provisions", "=Assumptions!B28*0.3", "=SUMIF(Temp_Differences!G7:G50,\"DTA\",Temp_Differences!H7:H50)*Assumptions!B14*0.3", "=SUMIF(Temp_Differences!G7:G50,\"DTA\",Temp_Differences!I7:I50)*Assumptions!B14*0.3", "=C9+D9-E9", "=F9-C9", 0, "=I9"],
    [3, "Disallowances u/s 43B", "=Assumptions!B28*0.25", "=SUMIF(Temp_Differences!G7:G50,\"DTA\",Temp_Differences!H7:H50)*Assumptions!B14*0.25", "=SUMIF(Temp_Differences!G7:G50,\"DTA\",Temp_Differences!I7:I50)*Assumptions!B14*0.25", "=C10+D10-E10", "=F10-C10", 0, "=I10"],
    [4, "Carry Forward Losses (Subject to Recognition)", "=Assumptions!B28*0.25", "=IF(Assumptions!B34=\"Yes\",SUMIF(Temp_Differences!C7:C50,\"Tax Losses\",Temp_Differences!H7:H50)*Assumptions!B14,0)", "=SUMIF(Temp_Differences!C7:C50,\"Tax Losses\",Temp_Differences!I7:I50)*Assumptions!B14", "=C11+D11-E11", "=F11-C11", 0, "=I11"]
  ];
  
  sheet.getRange(8, 1, dtaItems.length, 9).setValues(dtaItems);
  
  // DTA Subtotal
  sheet.getRange("A12:B12").merge()
    .setValue("Subtotal - DTA")
    .setFontWeight("bold")
    .setBackground(COLORS.CALC_BG);
  
  for (let col = 3; col <= 9; col++) {
    const colLetter = String.fromCharCode(64 + col);
    sheet.getRange(12, col).setFormula(`=SUM(${colLetter}8:${colLetter}11)`)
      .setFontWeight("bold")
      .setBackground(COLORS.CALC_BG);
  }
  
  // DTL Section
  sheet.getRange("A14:I14").merge()
    .setValue("DEFERRED TAX LIABILITIES")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(COLORS.SECTION_BG);
  
  const dtlItems = [
    [5, "Depreciation Differences", "=Assumptions!B29*0.8", "=SUMIF(Temp_Differences!G7:G50,\"DTL\",Temp_Differences!H7:H50)*Assumptions!B14*0.8", "=SUMIF(Temp_Differences!G7:G50,\"DTL\",Temp_Differences!I7:I50)*Assumptions!B14*0.8", "=C15+D15-E15", "=F15-C15", 0, "=I15"],
    [6, "Other Timing Differences", "=Assumptions!B29*0.2", "=SUMIF(Temp_Differences!G7:G50,\"DTL\",Temp_Differences!H7:H50)*Assumptions!B14*0.2", "=SUMIF(Temp_Differences!G7:G50,\"DTL\",Temp_Differences!I7:I50)*Assumptions!B14*0.2", "=C16+D16-E16", "=F16-C16", 0, "=I16"]
  ];
  
  sheet.getRange(15, 1, dtlItems.length, 9).setValues(dtlItems);
  
  // DTL Subtotal
  sheet.getRange("A17:B17").merge()
    .setValue("Subtotal - DTL")
    .setFontWeight("bold")
    .setBackground(COLORS.CALC_BG);
  
  for (let col = 3; col <= 9; col++) {
    const colLetter = String.fromCharCode(64 + col);
    sheet.getRange(17, col).setFormula(`=SUM(${colLetter}15:${colLetter}16)`)
      .setFontWeight("bold")
      .setBackground(COLORS.CALC_BG);
  }
  
  // Net Position
  sheet.getRange("A19:I19").merge()
    .setValue("NET POSITION")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(COLORS.SECTION_BG);
  
  const netRows = [
    ["", "Net DTA/(DTL) - Before Netting", "=C12-C17", "=D12-D17", "=E12-E17", "=F12-F17", "=G12-G17", "=H12-H17", "=I12-I17"]
  ];
  
  sheet.getRange(20, 1, netRows.length, 9).setValues(netRows);
  sheet.getRange("A20:B20").merge().setFontWeight("bold");
  
  // Netting section (if applicable)
  sheet.getRange("A22:I22").merge()
    .setValue("NETTING OF DTA AND DTL (If Applicable per Framework)")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(COLORS.SECTION_BG);
  
  sheet.getRange("A23:B23").merge()
    .setValue("Apply Netting?")
    .setFontWeight("bold");
  sheet.getRange("C23").setFormula("=Assumptions!B36")
    .setFontWeight("bold")
    .setFontColor("#d32f2f");
  
  const nettingRows = [
    ["", "DTA (Gross)", "=IF(C23=\"Yes\",C12,C12)", "", "", "=IF(C23=\"Yes\",F12,F12)", "", "", ""],
    ["", "DTL (Gross)", "=IF(C23=\"Yes\",C17,C17)", "", "", "=IF(C23=\"Yes\",F17,F17)", "", "", ""],
    ["", "Amount Netted", "=IF(C23=\"Yes\",MIN(C24,C25),0)", "", "", "=IF(C23=\"Yes\",MIN(F24,F25),0)", "", "", ""],
    ["", "", "", "", "", "", "", "", ""]
  ];
  
  sheet.getRange(24, 1, nettingRows.length, 9).setValues(nettingRows);
  sheet.getRange("A24:B27").mergeAcross().setFontWeight("bold");
  
  // Final Presentation
  sheet.getRange("A28:I28").merge()
    .setValue("FINAL PRESENTATION (Balance Sheet)")
    .setFontWeight("bold")
    .setFontSize(12)
    .setHorizontalAlignment("center")
    .setBackground(COLORS.GRAND_TOTAL_BG);
  
  const finalRows = [
    ["", "Deferred Tax Assets", "=C12", "", "", "=F12", "", "", "=F30-C30"],
    ["", "Deferred Tax Liabilities", "=C17", "", "", "=F17", "", "", "=F31-C31"],
    ["", "Net DTA/(DTL)", "=C30-C31", "", "", "=F30-F31", "=G30-G31", "=H30-H31", "=F32-C32"]
  ];
  
  sheet.getRange(30, 1, finalRows.length, 9).setValues(finalRows);
  sheet.getRange("A30:B32").merge().setFontWeight("bold").setBackground(COLORS.GRAND_TOTAL_BG);
  sheet.getRange("C30:I32").setFontWeight("bold").setBackground(COLORS.GRAND_TOTAL_BG);
  
  // Journal Entry section
  sheet.getRange("A35:I35").merge()
    .setValue("PERIOD CLOSURE JOURNAL ENTRY")
    .setFontSize(14)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT);
  
  const jeHeaders = [["Particulars", "Debit (₹)", "Credit (₹)"]];
  sheet.getRange(37, 2, 1, 3).setValues(jeHeaders);
  sheet.getRange("B37:D37").setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  
  const jeEntries = [
    ["Deferred Tax Asset A/c                              Dr.", "=IF(I30>0,I30,\"\")", ""],
    ["     To Deferred Tax Expense A/c", "", "=IF(I30>0,I30,\"\")"],
    ["(Being DTA created/increased during the year)", "", ""],
    ["", "", ""],
    ["Deferred Tax Expense A/c                           Dr.", "=IF(I31>0,I31,\"\")", ""],
    ["     To Deferred Tax Liability A/c", "", "=IF(I31>0,I31,\"\")"],
    ["(Being DTL created/increased during the year)", "", ""],
    ["", "", ""],
    ["TOTAL", "=SUM(C38,C42)", "=SUM(D39,D43)"]
  ];
  
  sheet.getRange(38, 2, jeEntries.length, 3).setValues(jeEntries);
  sheet.getRange("B46:D46").setBackground(COLORS.TOTAL_BG).setFontWeight("bold");
  
  // Number formatting
  sheet.getRange("C8:I32").setNumberFormat("#,##0");
  sheet.getRange("C38:D46").setNumberFormat("#,##0");
  
  // Borders
  sheet.getRange("A5:I32").setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("B37:D46").setBorder(true, true, true, true, true, true, COLORS.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  
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
    ["Deferred Tax Expense/(Income)", "=Movement_Analysis!G32", "Net change in deferred tax"],
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
    ["Deferred Tax Assets (Net)", "=Movement_Analysis!F30", "=Movement_Analysis!C30", "Note X - Deferred Taxation"],
    ["", "", "", ""],
    ["NON-CURRENT LIABILITIES", "", "", ""],
    ["Deferred Tax Liabilities (Net)", "=Movement_Analysis!F31", "=Movement_Analysis!C31", "Note X - Deferred Taxation"],
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
    ["Per Movement Analysis Schedule", "=Movement_Analysis!F30", "=Movement_Analysis!F31", "=Movement_Analysis!F32"],
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
    ["DTA (Gross before netting)", "=Movement_Analysis!C24", ""],
    ["DTL (Gross before netting)", "=Movement_Analysis!C25", ""],
    ["Amount Netted", "=Movement_Analysis!C26", ""],
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
    ["Net Deferred Tax Asset/(Liability)", "=Movement_Analysis!F32", "=Movement_Analysis!C32"]
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
    ["DTA per Movement Schedule", "=Movement_Analysis!F30", "=BS_Reconciliation!C7", "=IF(C7=D7,\"✓ OK\",\"⚠ MISMATCH\")", "Must match BS presentation"],
    ["DTL per Movement Schedule", "=Movement_Analysis!F31", "=BS_Reconciliation!C10", "=IF(C8=D8,\"✓ OK\",\"⚠ MISMATCH\")", "Must match BS presentation"],
    ["Net DTA/(DTL)", "=Movement_Analysis!F32", "=BS_Reconciliation!C12", "=IF(C9=D9,\"✓ OK\",\"⚠ MISMATCH\")", "Must match BS presentation"],
    ["", "", "", "", ""],
    ["Deferred Tax Expense per Movement", "=Movement_Analysis!G32", "=PL_Reconciliation!C7", "=IF(C11=D11,\"✓ OK\",\"⚠ MISMATCH\")", "Must match P&L reconciliation"],
    ["", "", "", "", ""],
    ["Total Tax Expense Check", "=PL_Reconciliation!C9", "=PL_Reconciliation!C6+PL_Reconciliation!C7", "=IF(C13=D13,\"✓ OK\",\"⚠ MISMATCH\")", "Current + Deferred"],
    ["", "", "", "", ""],
    ["Opening + Movement = Closing?", "=Movement_Analysis!F32", "=Movement_Analysis!C32+Movement_Analysis!I32", "=IF(C15=D15,\"✓ OK\",\"⚠ MISMATCH\")", "Movement reconciliation check"]
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
    ss.setNamedRange("Closing_DTA", ss.getSheetByName("Movement_Analysis").getRange("F30"));
    ss.setNamedRange("Closing_DTL", ss.getSheetByName("Movement_Analysis").getRange("F31"));
    ss.setNamedRange("Net_DTA_DTL", ss.getSheetByName("Movement_Analysis").getRange("F32"));
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
    sheet.hideGridlines(true);

    // Auto-resize rows for better visibility
    sheet.autoResizeRows(1, sheet.getMaxRows());
  });

  Logger.log("Final formatting applied - gridlines hidden for professional appearance");
}

// ============================================================================
// MENU FUNCTIONS
// ============================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔧 DT Audit Builder')
    .addItem('🆕 Create New Workbook', 'createDeferredTaxWorkbook')
    .addSeparator()
    .addItem('📊 Refresh All Formulas', 'refreshAllFormulas')
    .addItem('🔄 Recalculate Control Totals', 'recalculateControls')
    .addSeparator()
    .addItem('📝 Add New Temp Difference Row', 'addTempDiffRow')
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