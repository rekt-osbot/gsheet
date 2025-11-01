/**
 * ═══════════════════════════════════════════════════════════════════════════
 * IGAAP-IND AS 116 LEASE ACCOUNTING WORKINGS BUILDER
 * ═══════════════════════════════════════════════════════════════════════════
 * 
 * Purpose: Automated creation of Ind AS 116 compliant lease accounting workings
 *          with period book closure journal entries
 * 
 * Created: November 2025
 * Standard: Ind AS 116 - Leases (effective from 1 April 2019)
 * 
 * Key Features:
 * - Right-of-Use (ROU) Asset capitalization and depreciation
 * - Lease Liability measurement using effective interest method
 * - Automatic journal entry generation for period closure
 * - Full audit trail with control totals
 * - IGAAP vs Ind AS 116 reconciliation
 * 
 * Usage: Run 'createIndAS116Workbook()' from Script Editor
 * ═══════════════════════════════════════════════════════════════════════════
 */

function createIndAS116Workbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Clear existing sheets except default
  clearExistingSheets(ss);
  
  // Create all sheets in order
  createCoverSheet(ss);
  createAssumptionsSheet(ss);
  createLeaseRegisterSheet(ss);
  createROUAssetScheduleSheet(ss);
  createLeaseLiabilityScheduleSheet(ss);
  createPaymentScheduleSheet(ss);
  createPeriodMovementsSheet(ss);
  createJournalEntriesSheet(ss);
  createReconciliationSheet(ss);
  createReferencesSheet(ss);
  createAuditTrailSheet(ss);
  
  // Set up named ranges for key inputs
  setupNamedRanges(ss);
  
  // Final formatting and protection
  finalizeWorkbook(ss);
  
  SpreadsheetApp.getUi().alert(
    '✓ Ind AS 116 Workbook Created Successfully!\n\n' +
    'Please navigate to the Cover sheet and start with the Assumptions sheet.\n' +
    'Input cells are highlighted in light blue.\n\n' +
    'Refer to Audit Trail sheet for control checks.'
  );
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 * UTILITY FUNCTIONS
 * ═══════════════════════════════════════════════════════════════════════════
 */

function clearExistingSheets(ss) {
  const sheets = ss.getSheets();
  sheets.forEach(sheet => {
    if (sheet.getName() !== 'Sheet1') {
      ss.deleteSheet(sheet);
    }
  });
}

function setupNamedRanges(ss) {
  try {
    // Key named ranges for formulas
    const assumptions = ss.getSheetByName('Assumptions');
    
    ss.setNamedRange('ReportingDate', assumptions.getRange('C4'));
    ss.setNamedRange('CompanyName', assumptions.getRange('C3'));
    ss.setNamedRange('BaseCurrency', assumptions.getRange('C5'));
    ss.setNamedRange('DefaultIBR', assumptions.getRange('C8'));
    
  } catch (e) {
    Logger.log('Named ranges setup: ' + e);
  }
}

function finalizeWorkbook(ss) {
  // Set Cover sheet as active
  ss.getSheetByName('Cover').activate();
  
  // Freeze header rows on all sheets
  ss.getSheets().forEach(sheet => {
    if (sheet.getMaxRows() > 3) {
      sheet.setFrozenRows(3);
    }
  });
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 * SHEET 1: COVER / DASHBOARD
 * ═══════════════════════════════════════════════════════════════════════════
 */

function createCoverSheet(ss) {
  let sheet = ss.getSheetByName('Cover');
  if (!sheet) {
    sheet = ss.insertSheet('Cover', 0);
  }
  
  sheet.clear();
  
  // Header Section
  sheet.getRange('A1:K1').merge()
    .setValue('IND AS 116 - LEASE ACCOUNTING WORKINGS')
    .setBackground('#1a472a')
    .setFontColor('#ffffff')
    .setFontSize(18)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  sheet.setRowHeight(1, 50);
  
  sheet.getRange('A2:K2').merge()
    .setValue('Right-of-Use Assets & Lease Liabilities | Period Book Closure Entries')
    .setBackground('#2e7d54')
    .setFontColor('#ffffff')
    .setFontSize(11)
    .setHorizontalAlignment('center');
  
  // Company Info Section
  sheet.getRange('A4').setValue('Company:')
    .setFontWeight('bold');
  sheet.getRange('B4').setFormula('=Assumptions!C3')
    .setFontWeight('bold')
    .setFontSize(12);
  
  sheet.getRange('A5').setValue('Reporting Period:')
    .setFontWeight('bold');
  sheet.getRange('B5').setFormula('=TEXT(Assumptions!C4,"DD-MMM-YYYY")')
    .setFontWeight('bold');
  
  sheet.getRange('A6').setValue('Currency:')
    .setFontWeight('bold');
  sheet.getRange('B6').setFormula('=Assumptions!C5')
    .setFontWeight('bold');
  
  // Key Metrics Dashboard
  sheet.getRange('A8:K8').merge()
    .setValue('KEY METRICS SUMMARY')
    .setBackground('#e8f0fe')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  const metricsHeaders = [
    ['Metric', 'Opening Balance', 'Additions', 'Depreciation/Interest', 'Payments', 'Closing Balance', 'Variance %']
  ];
  
  sheet.getRange('A9:G9').setValues(metricsHeaders)
    .setBackground('#d9d9d9')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // ROU Assets Row
  sheet.getRange('A10').setValue('Right-of-Use Assets');
  sheet.getRange('B10').setFormula('=IFERROR(\'ROU Asset Schedule\'!C6,0)')
    .setNumberFormat('#,##0');
  sheet.getRange('C10').setFormula('=IFERROR(SUM(\'ROU Asset Schedule\'!E7:E50),0)')
    .setNumberFormat('#,##0');
  sheet.getRange('D10').setFormula('=IFERROR(SUM(\'ROU Asset Schedule\'!H7:H50),0)')
    .setNumberFormat('#,##0');
  sheet.getRange('E10').setValue('-')
    .setHorizontalAlignment('center');
  sheet.getRange('F10').setFormula('=B10+C10-D10')
    .setNumberFormat('#,##0')
    .setFontWeight('bold');
  sheet.getRange('G10').setFormula('=IFERROR((F10-B10)/B10,0)')
    .setNumberFormat('0.0%');
  
  // Lease Liabilities Row
  sheet.getRange('A11').setValue('Lease Liabilities');
  sheet.getRange('B11').setFormula('=IFERROR(\'Lease Liability Schedule\'!C6,0)')
    .setNumberFormat('#,##0');
  sheet.getRange('C11').setFormula('=IFERROR(SUM(\'Lease Liability Schedule\'!E7:E50),0)')
    .setNumberFormat('#,##0');
  sheet.getRange('D11').setFormula('=IFERROR(SUM(\'Lease Liability Schedule\'!G7:G50),0)')
    .setNumberFormat('#,##0');
  sheet.getRange('E11').setFormula('=IFERROR(SUM(\'Lease Liability Schedule\'!F7:F50),0)')
    .setNumberFormat('#,##0');
  sheet.getRange('F11').setFormula('=B11+C11+D11-E11')
    .setNumberFormat('#,##0')
    .setFontWeight('bold');
  sheet.getRange('G11').setFormula('=IFERROR((F11-B11)/B11,0)')
    .setNumberFormat('0.0%');
  
  // Net Impact Row
  sheet.getRange('A12').setValue('Net Impact on Equity')
    .setFontWeight('bold');
  sheet.getRange('F12').setFormula('=F10-F11')
    .setNumberFormat('#,##0')
    .setFontWeight('bold')
    .setBackground('#fff2cc');
  
  // Navigation Section
  sheet.getRange('A14:K14').merge()
    .setValue('NAVIGATION & QUICK LINKS')
    .setBackground('#e8f0fe')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  const navButtons = [
    ['Sheet Name', 'Description', 'Go To'],
    ['Assumptions', 'Input key parameters (IBR, lease terms, dates)', '→'],
    ['Lease Register', 'Master list of all leases', '→'],
    ['ROU Asset Schedule', 'Capitalization & depreciation tracker', '→'],
    ['Lease Liability Schedule', 'Liability amortization with interest', '→'],
    ['Payment Schedule', 'Track actual payments vs schedule', '→'],
    ['Period Movements', 'Monthly/quarterly activity summary', '→'],
    ['Journal Entries', 'Auto-generated period closure entries', '→'],
    ['Reconciliation', 'Balance sheet tie-outs & control totals', '→'],
    ['Audit Trail', 'Control checks & audit assertions', '→']
  ];
  
  sheet.getRange(15, 1, navButtons.length, navButtons[0].length).setValues(navButtons);
  sheet.getRange('A15:C15').setBackground('#d9d9d9').setFontWeight('bold');
  
  // Format navigation buttons
  for (let i = 16; i <= 24; i++) {
    sheet.getRange('C' + i)
      .setBackground('#4285f4')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }
  
  // Instructions
  sheet.getRange('A26:K26').merge()
    .setValue('INSTRUCTIONS')
    .setBackground('#fce5cd')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  const instructions = [
    ['1. START with the Assumptions sheet - fill in all light blue cells (company info, IBR rates, period dates)'],
    ['2. POPULATE Lease Register with details of all leases (lease terms, payment schedules, commencement dates)'],
    ['3. REVIEW ROU Asset and Lease Liability schedules - formulas calculate automatically'],
    ['4. VERIFY Payment Schedule against actual bank statements'],
    ['5. CHECK Period Movements for month-wise activity'],
    ['6. EXTRACT Journal Entries sheet for booking in ERP/accounting system'],
    ['7. RECONCILE using Reconciliation sheet - ensure all control totals match'],
    ['8. AUDIT using Audit Trail sheet - review all assertions and control checks']
  ];
  
  sheet.getRange(27, 1, instructions.length, 1).setValues(instructions);
  sheet.getRange('A27:A34').setWrap(true).setVerticalAlignment('top');
  
  // Control Totals Section
  sheet.getRange('A36:K36').merge()
    .setValue('CONTROL TOTALS & CHECKS')
    .setBackground('#e8f0fe')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.getRange('A37').setValue('ROU Asset = Lease Liability (Initial):');
  sheet.getRange('B37').setFormula('=IF(\'Audit Trail\'!D10="PASS","✓ PASS","✗ FAIL")')
    .setFontWeight('bold');
  
  sheet.getRange('A38').setValue('Depreciation + Interest = Rent (IGAAP):');
  sheet.getRange('B38').setFormula('=IF(\'Audit Trail\'!D11="PASS","✓ PASS","✗ FAIL")')
    .setFontWeight('bold');
  
  sheet.getRange('A39').setValue('Closing Balance Reconciliation:');
  sheet.getRange('B39').setFormula('=IF(\'Audit Trail\'!D12="PASS","✓ PASS","✗ FAIL")')
    .setFontWeight('bold');
  
  // Conditional formatting for control checks
  const passRange = sheet.getRange('B37:B39');
  const passRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('PASS')
    .setBackground('#d9ead3')
    .setFontColor('#38761d')
    .setRanges([passRange])
    .build();
  
  const failRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('FAIL')
    .setBackground('#f4cccc')
    .setFontColor('#cc0000')
    .setRanges([passRange])
    .build();
  
  const rules = sheet.getConditionalFormatRules();
  rules.push(passRule, failRule);
  sheet.setConditionalFormatRules(rules);
  
  // Column widths
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 100);
  
  // Borders
  sheet.getRange('A9:G12').setBorder(true, true, true, true, true, true);
  sheet.getRange('A15:C24').setBorder(true, true, true, true, true, true);
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 * SHEET 2: ASSUMPTIONS
 * ═══════════════════════════════════════════════════════════════════════════
 */

function createAssumptionsSheet(ss) {
  let sheet = ss.getSheetByName('Assumptions');
  if (!sheet) {
    sheet = ss.insertSheet('Assumptions', 1);
  }
  
  sheet.clear();
  
  // Header
  sheet.getRange('A1:H1').merge()
    .setValue('ASSUMPTIONS & INPUT PARAMETERS')
    .setBackground('#1a472a')
    .setFontColor('#ffffff')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.setRowHeight(1, 40);
  
  sheet.getRange('A2:H2').merge()
    .setValue('All input cells are highlighted in light blue - please fill these carefully')
    .setBackground('#cfe2f3')
    .setFontStyle('italic')
    .setHorizontalAlignment('center');
  
  // Company Information Section
  sheet.getRange('A3:H3').merge()
    .setValue('COMPANY INFORMATION')
    .setBackground('#d9ead3')
    .setFontWeight('bold');
  
  sheet.getRange('A4').setValue('Company Name:').setFontWeight('bold');
  sheet.getRange('C4').setValue('[Enter Company Name]')
    .setBackground('#cfe2f3')
    .setNote('INPUT REQUIRED: Full legal name of the entity');
  
  sheet.getRange('A5').setValue('Reporting Date:').setFontWeight('bold');
  sheet.getRange('C5').setValue(new Date())
    .setNumberFormat('dd-mmm-yyyy')
    .setBackground('#cfe2f3')
    .setNote('INPUT REQUIRED: Period end date for these workings');
  
  sheet.getRange('A6').setValue('Base Currency:').setFontWeight('bold');
  sheet.getRange('C6').setValue('INR')
    .setBackground('#cfe2f3')
    .setNote('INPUT REQUIRED: Functional currency (INR, USD, etc.)');
  
  sheet.getRange('A7').setValue('Currency Unit:').setFontWeight('bold');
  sheet.getRange('C7').setValue('Actuals')
    .setBackground('#cfe2f3')
    .setNote('INPUT REQUIRED: Actuals, Thousands, Lakhs, Crores, Millions');
  
  // Discount Rate Section
  sheet.getRange('A9:H9').merge()
    .setValue('DISCOUNT RATES (Incremental Borrowing Rate - IBR)')
    .setBackground('#d9ead3')
    .setFontWeight('bold');
  
  sheet.getRange('A10:H10').setValues([[
    'Lease Category', 'IBR % p.a.', 'Basis of Rate', 'Effective Date', 'Source/Documentation', '', '', ''
  ]]).setBackground('#d9d9d9').setFontWeight('bold').setHorizontalAlignment('center');
  
  const ibrData = [
    ['Property/Real Estate', 0.095, 'Bank borrowing rate + spread', '01-Apr-2024', 'Avg secured loan rate from bank'],
    ['Vehicles', 0.085, 'Auto loan rate', '01-Apr-2024', 'Bank auto loan schedule'],
    ['Equipment/Machinery', 0.090, 'Term loan rate', '01-Apr-2024', 'Corporate borrowing rate'],
    ['IT Equipment', 0.088, 'Short-term facility rate', '01-Apr-2024', 'Working capital facility']
  ];
  
  sheet.getRange(11, 1, ibrData.length, 5).setValues(ibrData);
  sheet.getRange('B11:B14').setBackground('#cfe2f3')
    .setNumberFormat('0.00%')
    .setNote('INPUT REQUIRED: Enter IBR as decimal (9.5% = 0.095)');
  
  sheet.getRange('C11:E14').setBackground('#cfe2f3');
  
  // Default IBR
  sheet.getRange('A16').setValue('Default IBR (if specific rate not available):')
    .setFontWeight('bold');
  sheet.getRange('C16').setValue(0.09)
    .setNumberFormat('0.00%')
    .setBackground('#cfe2f3')
    .setNote('INPUT REQUIRED: Default discount rate for all leases');
  
  // Period Information
  sheet.getRange('A18:H18').merge()
    .setValue('PERIOD INFORMATION')
    .setBackground('#d9ead3')
    .setFontWeight('bold');
  
  sheet.getRange('A19').setValue('Transition Date (Ind AS 116 first time adoption):')
    .setFontWeight('bold');
  sheet.getRange('C19').setValue('01-Apr-2019')
    .setBackground('#cfe2f3')
    .setNote('Date of first-time adoption of Ind AS 116');
  
  sheet.getRange('A20').setValue('Current Period Start Date:')
    .setFontWeight('bold');
  sheet.getRange('C20').setValue('01-Apr-2024')
    .setBackground('#cfe2f3')
    .setNote('INPUT REQUIRED: Start of current accounting period');
  
  sheet.getRange('A21').setValue('Current Period End Date:')
    .setFontWeight('bold');
  sheet.getRange('C21').setValue('31-Mar-2025')
    .setBackground('#cfe2f3')
    .setNote('INPUT REQUIRED: End of current accounting period');
  
  sheet.getRange('A22').setValue('Closure Frequency:')
    .setFontWeight('bold');
  sheet.getRange('C22').setValue('Monthly')
    .setBackground('#cfe2f3')
    .setNote('INPUT REQUIRED: Monthly, Quarterly, Half-Yearly, Annual');
  
  // Depreciation Policy
  sheet.getRange('A24:H24').merge()
    .setValue('DEPRECIATION POLICY FOR ROU ASSETS')
    .setBackground('#d9ead3')
    .setFontWeight('bold');
  
  sheet.getRange('A25:H25').setValues([[
    'Asset Category', 'Depreciation Method', 'Depreciation Period', 'Notes', '', '', '', ''
  ]]).setBackground('#d9d9d9').setFontWeight('bold');
  
  const depreciationData = [
    ['Property', 'Straight Line', 'Shorter of lease term or useful life', 'Over lease term as ownership not transferred'],
    ['Vehicles', 'Straight Line', 'Shorter of lease term or useful life', 'Typically over lease term (3-5 years)'],
    ['Equipment', 'Straight Line', 'Shorter of lease term or useful life', 'Over lease term'],
  ];
  
  sheet.getRange(26, 1, depreciationData.length, 4).setValues(depreciationData);
  sheet.getRange('C26:C28').setBackground('#cfe2f3');
  
  // Materiality Thresholds
  sheet.getRange('A30:H30').merge()
    .setValue('MATERIALITY & RECOGNITION EXEMPTIONS (Ind AS 116)')
    .setBackground('#d9ead3')
    .setFontWeight('bold');
  
  sheet.getRange('A31').setValue('Low Value Asset Threshold:')
    .setFontWeight('bold');
  sheet.getRange('C31').setValue(400000)
    .setNumberFormat('#,##0')
    .setBackground('#cfe2f3')
    .setNote('INPUT: Leases below this value may be expensed (typically USD 5,000 = ₹4 lakh)');
  
  sheet.getRange('A32').setValue('Short-term Lease Threshold (months):')
    .setFontWeight('bold');
  sheet.getRange('C32').setValue(12)
    .setBackground('#cfe2f3')
    .setNote('INPUT: Leases 12 months or less may be expensed');
  
  sheet.getRange('A33').setValue('Apply Low Value Exemption?')
    .setFontWeight('bold');
  sheet.getRange('C33').setValue('Yes')
    .setBackground('#cfe2f3')
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes', 'No'])
      .build())
    .setNote('Choose Yes or No');
  
  sheet.getRange('A34').setValue('Apply Short-term Exemption?')
    .setFontWeight('bold');
  sheet.getRange('C34').setValue('Yes')
    .setBackground('#cfe2f3')
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes', 'No'])
      .build());
  
  // Chart of Accounts Mapping
  sheet.getRange('A36:H36').merge()
    .setValue('CHART OF ACCOUNTS MAPPING')
    .setBackground('#d9ead3')
    .setFontWeight('bold');
  
  sheet.getRange('A37:D37').setValues([[
    'Description', 'Account Code', 'Account Name', 'Dr/Cr'
  ]]).setBackground('#d9d9d9').setFontWeight('bold');
  
  const coaData = [
    ['ROU Asset', '1510', 'Right-of-Use Assets', 'Dr'],
    ['Accumulated Depreciation - ROU', '1519', 'Accumulated Depreciation - ROU Assets', 'Cr'],
    ['Lease Liability - Current', '2210', 'Current Lease Liabilities', 'Cr'],
    ['Lease Liability - Non-Current', '2510', 'Non-Current Lease Liabilities', 'Cr'],
    ['Depreciation Expense', '6210', 'Depreciation - ROU Assets', 'Dr'],
    ['Interest Expense', '7110', 'Interest on Lease Liabilities', 'Dr'],
    ['Rent Expense (IGAAP)', '6310', 'Rent Expense', 'Dr']
  ];
  
  sheet.getRange(38, 1, coaData.length, 4).setValues(coaData);
  sheet.getRange('B38:C44').setBackground('#cfe2f3')
    .setNote('INPUT: Update account codes per your chart of accounts');
  
  // Column widths
  sheet.setColumnWidth(1, 300);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 200);
  
  // Borders
  sheet.getRange('A10:E14').setBorder(true, true, true, true, true, true);
  sheet.getRange('A25:D28').setBorder(true, true, true, true, true, true);
  sheet.getRange('A37:D44').setBorder(true, true, true, true, true, true);
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 * SHEET 3: LEASE REGISTER
 * ═══════════════════════════════════════════════════════════════════════════
 */

function createLeaseRegisterSheet(ss) {
  let sheet = ss.getSheetByName('Lease Register');
  if (!sheet) {
    sheet = ss.insertSheet('Lease Register', 2);
  }
  
  sheet.clear();
  
  // Header
  sheet.getRange('A1:P1').merge()
    .setValue('LEASE REGISTER - MASTER LIST')
    .setBackground('#1a472a')
    .setFontColor('#ffffff')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.setRowHeight(1, 40);
  
  sheet.getRange('A2:P2').merge()
    .setValue('Complete details of all leases - Input cells highlighted in light blue')
    .setBackground('#cfe2f3')
    .setFontStyle('italic')
    .setHorizontalAlignment('center');
  
  // Column Headers
  const headers = [
    'Lease ID',
    'Lease Description',
    'Lessor Name',
    'Asset Category',
    'Commencement Date',
    'End Date',
    'Lease Term (Months)',
    'Monthly Payment',
    'Payment Frequency',
    'Total Lease Payments',
    'IBR %',
    'PV of Lease Payments',
    'Initial Direct Costs',
    'ROU Asset Value',
    'Exemption Applied',
    'Notes'
  ];
  
  sheet.getRange('A3:P3').setValues([headers])
    .setBackground('#d9d9d9')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setWrap(true);
  
  sheet.setRowHeight(3, 50);
  
  // Sample Data (to be replaced by user)
  const sampleData = [
    [
      'L001',
      'Corporate Office - 5th Floor, Commercial Complex',
      'ABC Properties Ltd',
      'Property',
      '01-Apr-2023',
      '31-Mar-2028',
      60,
      150000,
      'Monthly',
      '=H4*G4',
      '=VLOOKUP(D4,Assumptions!$A$11:$B$14,2,FALSE)',
      '=PV(K4/12,G4,-H4,0,1)',
      50000,
      '=L4+M4',
      'No',
      'Includes maintenance charges'
    ],
    [
      'L002',
      'Warehouse Facility - Industrial Area',
      'XYZ Logistics Pvt Ltd',
      'Property',
      '01-Jul-2023',
      '30-Jun-2026',
      36,
      80000,
      'Monthly',
      '=H5*G5',
      '=VLOOKUP(D5,Assumptions!$A$11:$B$14,2,FALSE)',
      '=PV(K5/12,G5,-H5,0,1)',
      0,
      '=L5+M5',
      'No',
      'Excludes utilities'
    ],
    [
      'L003',
      'Company Vehicle - Toyota Fortuner',
      'LeasePlan India',
      'Vehicles',
      '15-Jan-2024',
      '14-Jan-2027',
      36,
      35000,
      'Monthly',
      '=H6*G6',
      '=VLOOKUP(D6,Assumptions!$A$11:$B$14,2,FALSE)',
      '=PV(K6/12,G6,-H6,0,1)',
      0,
      '=L6+M6',
      'No',
      'Including insurance'
    ],
    [
      'L004',
      'IT Server Equipment',
      'Dell Financial Services',
      'Equipment',
      '01-Apr-2024',
      '31-Mar-2027',
      36,
      25000,
      'Monthly',
      '=H7*G7',
      '=VLOOKUP(D7,Assumptions!$A$11:$B$14,2,FALSE)',
      '=PV(K7/12,G7,-H7,0,1)',
      15000,
      '=L7+M7',
      'No',
      'Production servers'
    ],
    [
      'L005',
      'Office Printer - High Volume',
      'Xerox India',
      'IT Equipment',
      '01-Oct-2024',
      '30-Sep-2025',
      12,
      8000,
      'Monthly',
      '=H8*G8',
      '=VLOOKUP(D8,Assumptions!$A$11:$B$14,2,FALSE)',
      '=PV(K8/12,G8,-H8,0,1)',
      0,
      '=L8+M8',
      'Short-term',
      'Below 12 months - may be expensed'
    ]
  ];
  
  sheet.getRange(4, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  
  // Format input columns (light blue)
  const inputColumns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'M', 'O', 'P'];
  inputColumns.forEach(col => {
    sheet.getRange(col + '4:' + col + '20').setBackground('#cfe2f3');
  });
  
  // Format calculated columns (light yellow)
  sheet.getRange('J4:L20').setBackground('#fff2cc');
  sheet.getRange('N4:N20').setBackground('#fff2cc');
  
  // Number formats
  sheet.getRange('E4:F20').setNumberFormat('dd-mmm-yyyy');
  sheet.getRange('G4:G20').setNumberFormat('0');
  sheet.getRange('H4:H20').setNumberFormat('#,##0');
  sheet.getRange('J4:J20').setNumberFormat('#,##0');
  sheet.getRange('K4:K20').setNumberFormat('0.00%');
  sheet.getRange('L4:L20').setNumberFormat('#,##0');
  sheet.getRange('M4:M20').setNumberFormat('#,##0');
  sheet.getRange('N4:N20').setNumberFormat('#,##0');
  
  // Data validation for dropdowns
  const categoryValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Property', 'Vehicles', 'Equipment', 'IT Equipment', 'Other'])
    .build();
  sheet.getRange('D4:D20').setDataValidation(categoryValidation);
  
  const frequencyValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Monthly', 'Quarterly', 'Half-Yearly', 'Annual'])
    .build();
  sheet.getRange('I4:I20').setDataValidation(frequencyValidation);
  
  const exemptionValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['No', 'Low Value', 'Short-term', 'Both'])
    .build();
  sheet.getRange('O4:O20').setDataValidation(exemptionValidation);
  
  // Summary Row
  sheet.getRange('A21').setValue('TOTAL:').setFontWeight('bold');
  sheet.getRange('J21').setFormula('=SUM(J4:J20)')
    .setNumberFormat('#,##0')
    .setFontWeight('bold')
    .setBackground('#d9ead3');
  sheet.getRange('L21').setFormula('=SUM(L4:L20)')
    .setNumberFormat('#,##0')
    .setFontWeight('bold')
    .setBackground('#d9ead3');
  sheet.getRange('N21').setFormula('=SUM(N4:N20)')
    .setNumberFormat('#,##0')
    .setFontWeight('bold')
    .setBackground('#d9ead3');
  
  // Notes section
  sheet.getRange('A23').setValue('NOTES:').setFontWeight('bold');
  sheet.getRange('A24:P24').merge()
    .setValue('• Lease ID: Unique identifier for each lease\n' +
              '• Payment Frequency: Adjust formulas if not monthly\n' +
              '• IBR %: Automatically pulled from Assumptions sheet based on Asset Category\n' +
              '• PV Calculation: Uses Excel PV function with periodic discounting\n' +
              '• Exemption: Mark if low-value or short-term exemption applied per Ind AS 116')
    .setWrap(true)
    .setVerticalAlignment('top')
    .setBackground('#fef7e0');
  
  sheet.setRowHeight(24, 100);
  
  // Column widths
  sheet.setColumnWidth(1, 80);   // Lease ID
  sheet.setColumnWidth(2, 250);  // Description
  sheet.setColumnWidth(3, 180);  // Lessor
  sheet.setColumnWidth(4, 120);  // Category
  sheet.setColumnWidth(5, 120);  // Dates
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 100);  // Term
  sheet.setColumnWidth(8, 120);  // Payment
  sheet.setColumnWidth(9, 120);  // Frequency
  sheet.setColumnWidth(10, 120); // Total Payments
  sheet.setColumnWidth(11, 80);  // IBR
  sheet.setColumnWidth(12, 130); // PV
  sheet.setColumnWidth(13, 120); // Initial Costs
  sheet.setColumnWidth(14, 120); // ROU Value
  sheet.setColumnWidth(15, 120); // Exemption
  sheet.setColumnWidth(16, 200); // Notes
  
  // Borders
  sheet.getRange('A3:P21').setBorder(true, true, true, true, true, true);
  
  // Cell notes
  sheet.getRange('J4').setNote('Formula: Monthly Payment × Lease Term (Months)');
  sheet.getRange('K4').setNote('Formula: VLOOKUP from Assumptions sheet based on Asset Category');
  sheet.getRange('L4').setNote('Formula: PV(IBR/12, Term, -Payment, 0, 1) - Present Value of all lease payments');
  sheet.getRange('N4').setNote('Formula: PV of Lease Payments + Initial Direct Costs = Opening ROU Asset Value');
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 * SHEET 4: ROU ASSET SCHEDULE
 * ═══════════════════════════════════════════════════════════════════════════
 */

function createROUAssetScheduleSheet(ss) {
  let sheet = ss.getSheetByName('ROU Asset Schedule');
  if (!sheet) {
    sheet = ss.insertSheet('ROU Asset Schedule', 3);
  }
  
  sheet.clear();
  
  // Header
  sheet.getRange('A1:L1').merge()
    .setValue('RIGHT-OF-USE (ROU) ASSET SCHEDULE')
    .setBackground('#1a472a')
    .setFontColor('#ffffff')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.setRowHeight(1, 40);
  
  sheet.getRange('A2:L2').merge()
    .setValue('Capitalization, Depreciation & Carrying Value (Ind AS 116 Para 23-25)')
    .setBackground('#2e7d54')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  // Summary Section
  sheet.getRange('A3').setValue('Reporting Date:').setFontWeight('bold');
  sheet.getRange('C3').setFormula('=TEXT(Assumptions!C5,"DD-MMM-YYYY")').setFontWeight('bold');
  
  sheet.getRange('A4').setValue('Currency:').setFontWeight('bold');
  sheet.getRange('C4').setFormula('=Assumptions!C6').setFontWeight('bold');
  
  sheet.getRange('A5:L5').merge()
    .setValue('SUMMARY')
    .setBackground('#e8f0fe')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.getRange('A6').setValue('Opening Balance:');
  sheet.getRange('C6').setFormula('=SUM(D7:D50)')
    .setNumberFormat('#,##0')
    .setFontWeight('bold');
  
  sheet.getRange('E6').setValue('Additions:');
  sheet.getRange('F6').setFormula('=SUM(E7:E50)')
    .setNumberFormat('#,##0')
    .setFontWeight('bold');
  
  sheet.getRange('H6').setValue('Depreciation:');
  sheet.getRange('I6').setFormula('=SUM(H7:H50)')
    .setNumberFormat('#,##0')
    .setFontWeight('bold');
  
  sheet.getRange('K6').setValue('Closing Balance:');
  sheet.getRange('L6').setFormula('=C6+F6-I6')
    .setNumberFormat('#,##0')
    .setFontWeight('bold')
    .setBackground('#d9ead3');
  
  // Column Headers
  const headers = [
    'Lease ID',
    'Description',
    'Category',
    'Opening Balance\n(Cost)',
    'Additions\n(Current Period)',
    'Accumulated Dep.\n(Opening)',
    'Net Opening\nBalance',
    'Depreciation\n(Current Period)',
    'Accumulated Dep.\n(Closing)',
    'Net Closing\nBalance',
    'Remaining Life\n(Months)',
    'Annual Dep. Rate'
  ];
  
  sheet.getRange('A7:L7').setValues([headers])
    .setBackground('#d9d9d9')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setWrap(true);
  
  sheet.setRowHeight(7, 60);
  
  // Data rows - Link to Lease Register
  for (let row = 8; row <= 12; row++) {
    const dataRow = row - 4; // Maps to Lease Register row
    
    sheet.getRange('A' + row).setFormula(`=IF('Lease Register'!A${dataRow}<>"", 'Lease Register'!A${dataRow}, "")`);
    sheet.getRange('B' + row).setFormula(`=IF('Lease Register'!A${dataRow}<>"", 'Lease Register'!B${dataRow}, "")`);
    sheet.getRange('C' + row).setFormula(`=IF('Lease Register'!A${dataRow}<>"", 'Lease Register'!D${dataRow}, "")`);
    
    // Opening Balance (Cost) - ROU Asset value from Lease Register
    sheet.getRange('D' + row).setFormula(
      `=IF('Lease Register'!A${dataRow}<>"", IF('Lease Register'!O${dataRow}="No", 'Lease Register'!N${dataRow}, 0), "")`
    );
    
    // Additions (if commencement date is in current period)
    sheet.getRange('E' + row).setFormula(
      `=IF('Lease Register'!A${dataRow}<>"", ` +
      `IF(AND('Lease Register'!E${dataRow}>=Assumptions!$C$20, 'Lease Register'!E${dataRow}<=Assumptions!$C$21), D${row}, 0), "")`
    );

    // Accumulated Depreciation (Opening) - INPUT CELL
    // CORRECTED: Changed from hardcoded 0 to input cell
    // Users must enter opening accumulated depreciation from prior period's closing balance
    // For first-time adoption or new leases, this would be 0; for subsequent periods, copy prior closing
    const openingAccumDepCell = sheet.getRange('F' + row);
    openingAccumDepCell.setValue(0);
    openingAccumDepCell.setBackground('#cfe2f3')
      .setNote('INPUT REQUIRED: Enter opening accumulated depreciation from prior period closing balance. Enter 0 for new leases.');

    // Net Opening Balance
    sheet.getRange('G' + row).setFormula(`=IF(A${row}<>"", D${row}-F${row}, "")`);
    
    // Depreciation (Current Period) - Straight line over lease term
    sheet.getRange('H' + row).setFormula(
      `=IF('Lease Register'!A${dataRow}<>"", ` +
      `IF('Lease Register'!O${dataRow}="No", ` +
      `D${row}/'Lease Register'!G${dataRow}*` +
      `DATEDIF(MAX('Lease Register'!E${dataRow},Assumptions!$C$20), ` +
      `MIN('Lease Register'!F${dataRow},Assumptions!$C$21), "M"), 0), "")`
    );
    
    // Accumulated Depreciation (Closing)
    sheet.getRange('I' + row).setFormula(`=IF(A${row}<>"", F${row}+H${row}, "")`);
    
    // Net Closing Balance
    sheet.getRange('J' + row).setFormula(`=IF(A${row}<>"", D${row}-I${row}, "")`);
    
    // Remaining Life
    sheet.getRange('K' + row).setFormula(
      `=IF('Lease Register'!A${dataRow}<>"", ` +
      `DATEDIF(Assumptions!$C$21, 'Lease Register'!F${dataRow}, "M"), "")`
    );
    
    // Annual Depreciation Rate
    sheet.getRange('L' + row).setFormula(
      `=IF('Lease Register'!A${dataRow}<>"", IF(D${row}>0, (12/('Lease Register'!G${dataRow})), 0), "")`
    );
  }
  
  // Total row
  sheet.getRange('A13').setValue('TOTAL:').setFontWeight('bold');
  sheet.getRange('D13').setFormula('=SUM(D8:D12)')
    .setNumberFormat('#,##0')
    .setFontWeight('bold')
    .setBackground('#d9ead3');
  sheet.getRange('E13').setFormula('=SUM(E8:E12)')
    .setNumberFormat('#,##0')
    .setFontWeight('bold')
    .setBackground('#d9ead3');
  sheet.getRange('F13').setFormula('=SUM(F8:F12)')
    .setNumberFormat('#,##0')
    .setFontWeight('bold')
    .setBackground('#d9ead3');
  sheet.getRange('G13').setFormula('=SUM(G8:G12)')
    .setNumberFormat('#,##0')
    .setFontWeight('bold')
    .setBackground('#d9ead3');
  sheet.getRange('H13').setFormula('=SUM(H8:H12)')
    .setNumberFormat('#,##0')
    .setFontWeight('bold')
    .setBackground('#d9ead3');
  sheet.getRange('I13').setFormula('=SUM(I8:I12)')
    .setNumberFormat('#,##0')
    .setFontWeight('bold')
    .setBackground('#d9ead3');
  sheet.getRange('J13').setFormula('=SUM(J8:J12)')
    .setNumberFormat('#,##0')
    .setFontWeight('bold')
    .setBackground('#d9ead3');
  
  // Format numbers
  sheet.getRange('D8:L12').setNumberFormat('#,##0');
  sheet.getRange('L8:L12').setNumberFormat('0.00%');
  
  // Notes Section
  sheet.getRange('A15').setValue('NOTES & CALCULATIONS:').setFontWeight('bold');
  
  const notes = [
    ['• ROU Asset Initial Recognition:', 'Cost = PV of Lease Payments + Initial Direct Costs (Ind AS 116 Para 24)'],
    ['• Depreciation Method:', 'Straight-line over shorter of lease term or useful life (Ind AS 116 Para 32)'],
    ['• Depreciation Formula:', '(ROU Asset Cost / Lease Term in Months) × Months in Period'],
    ['• Period Calculation:', 'Only months falling within current accounting period are considered'],
    ['• Exemptions:', 'Short-term and low-value leases are excluded from ROU asset recognition'],
    ['• Journal Entry:', 'Dr. Depreciation Expense, Cr. Accumulated Depreciation - ROU Asset']
  ];
  
  sheet.getRange(16, 1, notes.length, 2).setValues(notes);
  sheet.getRange('A16:B21').setBackground('#fef7e0')
    .setWrap(true)
    .setVerticalAlignment('top');
  
  // Column widths
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 130);
  sheet.setColumnWidth(7, 130);
  sheet.setColumnWidth(8, 130);
  sheet.setColumnWidth(9, 130);
  sheet.setColumnWidth(10, 130);
  sheet.setColumnWidth(11, 120);
  sheet.setColumnWidth(12, 120);
  
  // Borders
  sheet.getRange('A7:L13').setBorder(true, true, true, true, true, true);
  
  // Conditional formatting - highlight any negative balances
  const negativeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground('#f4cccc')
    .setRanges([sheet.getRange('G8:J12')])
    .build();
  
  const rules = sheet.getConditionalFormatRules();
  rules.push(negativeRule);
  sheet.setConditionalFormatRules(rules);
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 * SHEET 5: LEASE LIABILITY SCHEDULE
 * ═══════════════════════════════════════════════════════════════════════════
 */

function createLeaseLiabilityScheduleSheet(ss) {
  let sheet = ss.getSheetByName('Lease Liability Schedule');
  if (!sheet) {
    sheet = ss.insertSheet('Lease Liability Schedule', 4);
  }
  
  sheet.clear();
  
  // Header
  sheet.getRange('A1:K1').merge()
    .setValue('LEASE LIABILITY SCHEDULE')
    .setBackground('#1a472a')
    .setFontColor('#ffffff')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.setRowHeight(1, 40);
  
  sheet.getRange('A2:K2').merge()
    .setValue('Amortization using Effective Interest Method (Ind AS 116 Para 36)')
    .setBackground('#2e7d54')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  // Summary Section
  sheet.getRange('A3').setValue('Reporting Date:').setFontWeight('bold');
  sheet.getRange('C3').setFormula('=TEXT(Assumptions!C5,"DD-MMM-YYYY")').setFontWeight('bold');
  
  sheet.getRange('A4').setValue('Currency:').setFontWeight('bold');
  sheet.getRange('C4').setFormula('=Assumptions!C6').setFontWeight('bold');
  
  sheet.getRange('A5:K5').merge()
    .setValue('SUMMARY')
    .setBackground('#e8f0fe')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.getRange('A6').setValue('Opening Liability:');
  sheet.getRange('C6').setFormula('=SUM(D7:D50)')
    .setNumberFormat('#,##0')
    .setFontWeight('bold');
  
  sheet.getRange('E6').setValue('Additions:');
  sheet.getRange('F6').setFormula('=SUM(E7:E50)')
    .setNumberFormat('#,##0')
    .setFontWeight('bold');
  
  sheet.getRange('G6').setValue('Interest:');
  sheet.getRange('H6').setFormula('=SUM(G7:G50)')
    .setNumberFormat('#,##0')
    .setFontWeight('bold');
  
  sheet.getRange('I6').setValue('Payments:');
  sheet.getRange('J6').setFormula('=SUM(F7:F50)')
    .setNumberFormat('#,##0')
    .setFontWeight('bold');
  
  sheet.getRange('K6').setValue('Closing Liability:');
  sheet.getRange('K6').offset(0, 1).setFormula('=C6+F6+H6-J6')
    .setNumberFormat('#,##0')
    .setFontWeight('bold')
    .setBackground('#d9ead3');
  
  // Column Headers
  const headers = [
    'Lease ID',
    'Description',
    'Category',
    'Opening\nLiability',
    'Additions\n(New Leases)',
    'Payments\n(Current Period)',
    'Interest Expense\n(Current Period)',
    'Closing\nLiability',
    'Current\nPortion',
    'Non-Current\nPortion',
    'Effective\nIBR %'
  ];
  
  sheet.getRange('A7:K7').setValues([headers])
    .setBackground('#d9d9d9')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setWrap(true);
  
  sheet.setRowHeight(7, 60);
  
  // Data rows
  for (let row = 8; row <= 12; row++) {
    const dataRow = row - 4;
    
    sheet.getRange('A' + row).setFormula(`=IF('Lease Register'!A${dataRow}<>"", 'Lease Register'!A${dataRow}, "")`);
    sheet.getRange('B' + row).setFormula(`=IF('Lease Register'!A${dataRow}<>"", 'Lease Register'!B${dataRow}, "")`);
    sheet.getRange('C' + row).setFormula(`=IF('Lease Register'!A${dataRow}<>"", 'Lease Register'!D${dataRow}, "")`);
    
    // Opening Liability - same as PV from Lease Register
    sheet.getRange('D' + row).setFormula(
      `=IF('Lease Register'!A${dataRow}<>"", IF('Lease Register'!O${dataRow}="No", 'Lease Register'!L${dataRow}, 0), "")`
    );
    
    // Additions (if new lease in current period)
    sheet.getRange('E' + row).setFormula(
      `=IF('Lease Register'!A${dataRow}<>"", ` +
      `IF(AND('Lease Register'!E${dataRow}>=Assumptions!$C$20, 'Lease Register'!E${dataRow}<=Assumptions!$C$21), D${row}, 0), "")`
    );
    
    // Payments (months in current period × monthly payment)
    sheet.getRange('F' + row).setFormula(
      `=IF('Lease Register'!A${dataRow}<>"", ` +
      `IF('Lease Register'!O${dataRow}="No", ` +
      `'Lease Register'!H${dataRow}*` +
      `DATEDIF(MAX('Lease Register'!E${dataRow},Assumptions!$C$20), ` +
      `MIN('Lease Register'!F${dataRow},Assumptions!$C$21), "M"), 0), "")`
    );
    
    // Interest Expense - IMPROVED but still simplified
    // CORRECTED: Using average balance method for better accuracy
    // Proper effective interest method would require monthly breakdown
    // Formula: ((Opening + Closing pre-interest) / 2) × IBR × (Months / 12)
    // Closing pre-interest = Opening + Additions - Payments
    // This is more accurate than using only opening balance
    sheet.getRange('G' + row).setFormula(
      `=IF('Lease Register'!A${dataRow}<>"", ` +
      `IF('Lease Register'!O${dataRow}="No", ` +
      `((D${row}+E${row})+(D${row}+E${row}-F${row}))/2*'Lease Register'!K${dataRow}*` +
      `DATEDIF(MAX('Lease Register'!E${dataRow},Assumptions!$C$20), ` +
      `MIN('Lease Register'!F${dataRow},Assumptions!$C$21), "M")/12, 0), "")`
    );
    
    // Closing Liability
    sheet.getRange('H' + row).setFormula(`=IF(A${row}<>"", D${row}+E${row}+G${row}-F${row}, "")`);
    
    // Current Portion - IMPROVED calculation
    // CORRECTED: More accurate approximation than MIN(Closing, 12 × Monthly Payment)
    // Current portion = Principal repayment in next 12 months
    // Approximation: Next 12 months payments - Avg interest on reducing balance
    // Formula: MIN(Closing, 12×Payment - H×IBR×0.5)
    // The 0.5 factor approximates average balance over 12 months
    sheet.getRange('I' + row).setFormula(
      `=IF('Lease Register'!A${dataRow}<>"", ` +
      `MIN(H${row}, MAX(0, 'Lease Register'!H${dataRow}*12 - H${row}*'Lease Register'!K${dataRow}*0.5)), "")`
    );
    
    // Non-Current Portion
    sheet.getRange('J' + row).setFormula(`=IF(A${row}<>"", H${row}-I${row}, "")`);
    
    // Effective IBR
    sheet.getRange('K' + row).setFormula(`=IF('Lease Register'!A${dataRow}<>"", 'Lease Register'!K${dataRow}, "")`);
  }
  
  // Total row
  sheet.getRange('A13').setValue('TOTAL:').setFontWeight('bold');
  const totalCols = ['D', 'E', 'F', 'G', 'H', 'I', 'J'];
  totalCols.forEach(col => {
    sheet.getRange(col + '13').setFormula(`=SUM(${col}8:${col}12)`)
      .setNumberFormat('#,##0')
      .setFontWeight('bold')
      .setBackground('#d9ead3');
  });
  
  // Format numbers
  sheet.getRange('D8:J12').setNumberFormat('#,##0');
  sheet.getRange('K8:K12').setNumberFormat('0.00%');
  
  // Notes Section
  sheet.getRange('A15').setValue('NOTES & CALCULATIONS:').setFontWeight('bold');
  
  const notes = [
    ['• Initial Measurement:', 'Lease Liability = PV of future lease payments using IBR (Ind AS 116 Para 26)'],
    ['• Subsequent Measurement:', 'Liability increases by interest, decreases by payments (Para 36)'],
    ['• Interest Calculation (IMPROVED):', 'Interest = Average Balance × IBR × (Months/12). Uses average of opening and closing liability for more accurate effective interest method approximation.'],
    ['• Effective Interest Method:', 'True EIM requires monthly calculations. This workbook uses average balance method as practical approximation.'],
    ['• Current vs Non-Current (IMPROVED):', 'Current portion = Principal repayment in next 12 months. Calculated as (Next 12 months payments - Interest on avg balance). More accurate than simple MIN(Closing, 12×Payment) proxy.'],
    ['• Journal Entries:', 'Dr. Lease Liability / Cr. Cash (payment)'],
    ['', 'Dr. Interest Expense / Cr. Lease Liability (interest accretion)']
  ];
  
  sheet.getRange(16, 1, notes.length, 2).setValues(notes);
  sheet.getRange('A16:B22').setBackground('#fef7e0')
    .setWrap(true)
    .setVerticalAlignment('top');
  
  // Column widths
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 130);
  sheet.setColumnWidth(7, 130);
  sheet.setColumnWidth(8, 130);
  sheet.setColumnWidth(9, 130);
  sheet.setColumnWidth(10, 130);
  sheet.setColumnWidth(11, 100);
  
  // Borders
  sheet.getRange('A7:K13').setBorder(true, true, true, true, true, true);
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 * SHEET 6-11: Additional supporting schedules
 * ═══════════════════════════════════════════════════════════════════════════
 */

function createPaymentScheduleSheet(ss) {
  let sheet = ss.getSheetByName('Payment Schedule');
  if (!sheet) {
    sheet = ss.insertSheet('Payment Schedule', 5);
  }
  
  sheet.clear();
  
  sheet.getRange('A1:H1').merge()
    .setValue('LEASE PAYMENT SCHEDULE')
    .setBackground('#1a472a')
    .setFontColor('#ffffff')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.setRowHeight(1, 40);
  
  sheet.getRange('A2:H2').merge()
    .setValue('Track actual payments vs scheduled payments')
    .setBackground('#cfe2f3')
    .setHorizontalAlignment('center');
  
  const headers = ['Lease ID', 'Month', 'Scheduled Payment', 'Actual Payment', 'Variance', 'Payment Date', 'Payment Reference', 'Notes'];
  sheet.getRange('A3:H3').setValues([headers])
    .setBackground('#d9d9d9')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Sample structure
  sheet.getRange('A4').setValue('L001');
  sheet.getRange('B4').setValue('Apr-2024');
  sheet.getRange('C4').setValue(150000).setNumberFormat('#,##0');
  sheet.getRange('D4').setValue(150000).setNumberFormat('#,##0').setBackground('#cfe2f3');
  sheet.getRange('E4').setFormula('=D4-C4').setNumberFormat('#,##0');
  sheet.getRange('F4').setBackground('#cfe2f3');
  sheet.getRange('G4').setBackground('#cfe2f3');
  sheet.getRange('H4').setBackground('#cfe2f3');
  
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 150);
  sheet.setColumnWidth(8, 200);
  
  sheet.getRange('A3:H20').setBorder(true, true, true, true, true, true);
}

function createPeriodMovementsSheet(ss) {
  let sheet = ss.getSheetByName('Period Movements');
  if (!sheet) {
    sheet = ss.insertSheet('Period Movements', 6);
  }
  
  sheet.clear();
  
  sheet.getRange('A1:F1').merge()
    .setValue('PERIOD MOVEMENTS SUMMARY')
    .setBackground('#1a472a')
    .setFontColor('#ffffff')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  const headers = ['Period', 'ROU Asset Additions', 'Depreciation', 'Lease Liability Additions', 'Interest Expense', 'Payments Made'];
  sheet.getRange('A3:F3').setValues([headers])
    .setBackground('#d9d9d9')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.getRange('A4').setValue('FY 2024-25');
  sheet.getRange('B4').setFormula('=SUM(\'ROU Asset Schedule\'!E8:E12)').setNumberFormat('#,##0');
  sheet.getRange('C4').setFormula('=SUM(\'ROU Asset Schedule\'!H8:H12)').setNumberFormat('#,##0');
  sheet.getRange('D4').setFormula('=SUM(\'Lease Liability Schedule\'!E8:E12)').setNumberFormat('#,##0');
  sheet.getRange('E4').setFormula('=SUM(\'Lease Liability Schedule\'!G8:G12)').setNumberFormat('#,##0');
  sheet.getRange('F4').setFormula('=SUM(\'Lease Liability Schedule\'!F8:F12)').setNumberFormat('#,##0');
  
  sheet.setColumnWidth(1, 120);
  for (let i = 2; i <= 6; i++) {
    sheet.setColumnWidth(i, 150);
  }
  
  sheet.getRange('A3:F10').setBorder(true, true, true, true, true, true);
}

function createJournalEntriesSheet(ss) {
  let sheet = ss.getSheetByName('Journal Entries');
  if (!sheet) {
    sheet = ss.insertSheet('Journal Entries', 7);
  }
  
  sheet.clear();
  
  // Header
  sheet.getRange('A1:H1').merge()
    .setValue('PERIOD BOOK CLOSURE JOURNAL ENTRIES')
    .setBackground('#1a472a')
    .setFontColor('#ffffff')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.setRowHeight(1, 40);
  
  sheet.getRange('A2:H2').merge()
    .setValue('Auto-generated entries for period closure - Copy these to your ERP system')
    .setBackground('#fff2cc')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Entry 1: ROU Asset Recognition
  sheet.getRange('A4:H4').merge()
    .setValue('JOURNAL ENTRY 1: Recognition of ROU Assets (New Leases)')
    .setBackground('#d9ead3')
    .setFontWeight('bold');
  
  const je1Headers = ['Date', 'Account Code', 'Account Name', 'Description', 'Debit', 'Credit', 'Lease ID', 'Notes'];
  sheet.getRange('A5:H5').setValues([je1Headers])
    .setBackground('#d9d9d9')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.getRange('A6').setFormula('=Assumptions!C21').setNumberFormat('dd-mmm-yyyy');
  sheet.getRange('B6').setFormula('=Assumptions!B38');
  sheet.getRange('C6').setFormula('=Assumptions!C38');
  sheet.getRange('D6').setValue('Recognition of Right-of-Use Assets');
  sheet.getRange('E6').setFormula('=SUM(\'ROU Asset Schedule\'!E8:E12)').setNumberFormat('#,##0');
  sheet.getRange('F6').setValue('').setNumberFormat('#,##0');
  sheet.getRange('G6').setValue('Various');
  sheet.getRange('H6').setValue('Ind AS 116 Para 23-24');
  
  sheet.getRange('A7').setFormula('=Assumptions!C21').setNumberFormat('dd-mmm-yyyy');
  sheet.getRange('B7').setFormula('=Assumptions!B41');
  sheet.getRange('C7').setFormula('=Assumptions!C41');
  sheet.getRange('D7').setValue('Recognition of Lease Liability');
  sheet.getRange('E7').setValue('').setNumberFormat('#,##0');
  sheet.getRange('F7').setFormula('=SUM(\'Lease Liability Schedule\'!E8:E12)').setNumberFormat('#,##0');
  sheet.getRange('G7').setValue('Various');
  sheet.getRange('H7').setValue('Ind AS 116 Para 26');
  
  // Entry 2: Depreciation
  sheet.getRange('A9:H9').merge()
    .setValue('JOURNAL ENTRY 2: Depreciation of ROU Assets')
    .setBackground('#d9ead3')
    .setFontWeight('bold');
  
  sheet.getRange('A10:H10').setValues([je1Headers])
    .setBackground('#d9d9d9')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.getRange('A11').setFormula('=Assumptions!C21').setNumberFormat('dd-mmm-yyyy');
  sheet.getRange('B11').setFormula('=Assumptions!B42');
  sheet.getRange('C11').setFormula('=Assumptions!C42');
  sheet.getRange('D11').setValue('Depreciation on ROU Assets - Current Period');
  sheet.getRange('E11').setFormula('=SUM(\'ROU Asset Schedule\'!H8:H12)').setNumberFormat('#,##0');
  sheet.getRange('F11').setValue('').setNumberFormat('#,##0');
  sheet.getRange('G11').setValue('Various');
  sheet.getRange('H11').setValue('Ind AS 116 Para 32');
  
  sheet.getRange('A12').setFormula('=Assumptions!C21').setNumberFormat('dd-mmm-yyyy');
  sheet.getRange('B12').setFormula('=Assumptions!B39');
  sheet.getRange('C12').setFormula('=Assumptions!C39');
  sheet.getRange('D12').setValue('Accumulated Depreciation - ROU Assets');
  sheet.getRange('E12').setValue('').setNumberFormat('#,##0');
  sheet.getRange('F12').setFormula('=SUM(\'ROU Asset Schedule\'!H8:H12)').setNumberFormat('#,##0');
  sheet.getRange('G12').setValue('Various');
  sheet.getRange('H12').setValue('Ind AS 116 Para 32');
  
  // Entry 3: Interest
  sheet.getRange('A14:H14').merge()
    .setValue('JOURNAL ENTRY 3: Interest Expense on Lease Liabilities')
    .setBackground('#d9ead3')
    .setFontWeight('bold');
  
  sheet.getRange('A15:H15').setValues([je1Headers])
    .setBackground('#d9d9d9')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.getRange('A16').setFormula('=Assumptions!C21').setNumberFormat('dd-mmm-yyyy');
  sheet.getRange('B16').setFormula('=Assumptions!B43');
  sheet.getRange('C16').setFormula('=Assumptions!C43');
  sheet.getRange('D16').setValue('Interest on Lease Liabilities - Current Period');
  sheet.getRange('E16').setFormula('=SUM(\'Lease Liability Schedule\'!G8:G12)').setNumberFormat('#,##0');
  sheet.getRange('F16').setValue('').setNumberFormat('#,##0');
  sheet.getRange('G16').setValue('Various');
  sheet.getRange('H16').setValue('Ind AS 116 Para 36');
  
  sheet.getRange('A17').setFormula('=Assumptions!C21').setNumberFormat('dd-mmm-yyyy');
  sheet.getRange('B17').setFormula('=Assumptions!B41');
  sheet.getRange('C17').setFormula('=Assumptions!C41');
  sheet.getRange('D17').setValue('Lease Liability (Interest Accretion)');
  sheet.getRange('E17').setValue('').setNumberFormat('#,##0');
  sheet.getRange('F17').setFormula('=SUM(\'Lease Liability Schedule\'!G8:G12)').setNumberFormat('#,##0');
  sheet.getRange('G17').setValue('Various');
  sheet.getRange('H17').setValue('Ind AS 116 Para 36');
  
  // Entry 4: Payments
  sheet.getRange('A19:H19').merge()
    .setValue('JOURNAL ENTRY 4: Lease Payments Made')
    .setBackground('#d9ead3')
    .setFontWeight('bold');
  
  sheet.getRange('A20:H20').setValues([je1Headers])
    .setBackground('#d9d9d9')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.getRange('A21').setFormula('=Assumptions!C21').setNumberFormat('dd-mmm-yyyy');
  sheet.getRange('B21').setFormula('=Assumptions!B41');
  sheet.getRange('C21').setFormula('=Assumptions!C41');
  sheet.getRange('D21').setValue('Reduction of Lease Liability - Payments Made');
  sheet.getRange('E21').setFormula('=SUM(\'Lease Liability Schedule\'!F8:F12)').setNumberFormat('#,##0');
  sheet.getRange('F21').setValue('').setNumberFormat('#,##0');
  sheet.getRange('G21').setValue('Various');
  sheet.getRange('H21').setValue('Ind AS 116 Para 36');
  
  sheet.getRange('A22').setValue('Various');
  sheet.getRange('B22').setValue('1010');
  sheet.getRange('C22').setValue('Cash/Bank');
  sheet.getRange('D22').setValue('Lease Payments Made During Period');
  sheet.getRange('E22').setValue('').setNumberFormat('#,##0');
  sheet.getRange('F22').setFormula('=SUM(\'Lease Liability Schedule\'!F8:F12)').setNumberFormat('#,##0');
  sheet.getRange('G22').setValue('Various');
  sheet.getRange('H22').setValue('Actual payment dates');
  
  // Summary Section
  sheet.getRange('A24:H24').merge()
    .setValue('ENTRY VALIDATION & CONTROL TOTALS')
    .setBackground('#e8f0fe')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.getRange('A25').setValue('Total Debits:').setFontWeight('bold');
  sheet.getRange('B25').setFormula('=SUM(E6:E22)').setNumberFormat('#,##0').setFontWeight('bold');
  
  sheet.getRange('A26').setValue('Total Credits:').setFontWeight('bold');
  sheet.getRange('B26').setFormula('=SUM(F6:F22)').setNumberFormat('#,##0').setFontWeight('bold');
  
  sheet.getRange('A27').setValue('Difference (Should be 0):').setFontWeight('bold');
  sheet.getRange('B27').setFormula('=B25-B26').setNumberFormat('#,##0').setFontWeight('bold');
  
  // Conditional formatting
  const balanceRange = sheet.getRange('B27');
  const zeroRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(0)
    .setBackground('#d9ead3')
    .setFontColor('#38761d')
    .setRanges([balanceRange])
    .build();
  
  const nonZeroRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberNotEqualTo(0)
    .setBackground('#f4cccc')
    .setFontColor('#cc0000')
    .setRanges([balanceRange])
    .build();
  
  let rules = sheet.getConditionalFormatRules();
  rules.push(zeroRule, nonZeroRule);
  sheet.setConditionalFormatRules(rules);
  
  // Column widths
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 280);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 180);
  
  // Borders
  sheet.getRange('A5:H7').setBorder(true, true, true, true, true, true);
  sheet.getRange('A10:H12').setBorder(true, true, true, true, true, true);
  sheet.getRange('A15:H17').setBorder(true, true, true, true, true, true);
  sheet.getRange('A20:H22').setBorder(true, true, true, true, true, true);
}

function createReconciliationSheet(ss) {
  let sheet = ss.getSheetByName('Reconciliation');
  if (!sheet) {
    sheet = ss.insertSheet('Reconciliation', 8);
  }
  
  sheet.clear();
  
  sheet.getRange('A1:F1').merge()
    .setValue('RECONCILIATION & CONTROL TOTALS')
    .setBackground('#1a472a')
    .setFontColor('#ffffff')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.setRowHeight(1, 40);
  
  // ROU Asset Reconciliation
  sheet.getRange('A3:F3').merge()
    .setValue('ROU ASSET RECONCILIATION')
    .setBackground('#d9ead3')
    .setFontWeight('bold');
  
  const rouHeaders = ['Description', 'Amount (₹)', '', '', '', ''];
  sheet.getRange('A4:F4').setValues([rouHeaders])
    .setBackground('#d9d9d9')
    .setFontWeight('bold');
  
  sheet.getRange('A5').setValue('Opening Balance (Cost)');
  sheet.getRange('B5').setFormula('=\'ROU Asset Schedule\'!D13').setNumberFormat('#,##0');
  
  sheet.getRange('A6').setValue('Add: Additions during period');
  sheet.getRange('B6').setFormula('=\'ROU Asset Schedule\'!E13').setNumberFormat('#,##0');
  
  sheet.getRange('A7').setValue('Less: Depreciation during period');
  sheet.getRange('B7').setFormula('=-\'ROU Asset Schedule\'!H13').setNumberFormat('#,##0');
  
  sheet.getRange('A8').setValue('Closing Balance (Net Book Value)').setFontWeight('bold');
  sheet.getRange('B8').setFormula('=B5+B6+B7').setNumberFormat('#,##0').setFontWeight('bold').setBackground('#d9ead3');
  
  // Lease Liability Reconciliation
  sheet.getRange('A10:F10').merge()
    .setValue('LEASE LIABILITY RECONCILIATION')
    .setBackground('#d9ead3')
    .setFontWeight('bold');
  
  sheet.getRange('A11:F11').setValues([rouHeaders])
    .setBackground('#d9d9d9')
    .setFontWeight('bold');
  
  sheet.getRange('A12').setValue('Opening Balance');
  sheet.getRange('B12').setFormula('=\'Lease Liability Schedule\'!D13').setNumberFormat('#,##0');
  
  sheet.getRange('A13').setValue('Add: New lease liabilities');
  sheet.getRange('B13').setFormula('=\'Lease Liability Schedule\'!E13').setNumberFormat('#,##0');
  
  sheet.getRange('A14').setValue('Add: Interest expense');
  sheet.getRange('B14').setFormula('=\'Lease Liability Schedule\'!G13').setNumberFormat('#,##0');
  
  sheet.getRange('A15').setValue('Less: Payments made');
  sheet.getRange('B15').setFormula('=-\'Lease Liability Schedule\'!F13').setNumberFormat('#,##0');
  
  sheet.getRange('A16').setValue('Closing Balance').setFontWeight('bold');
  sheet.getRange('B16').setFormula('=B12+B13+B14+B15').setNumberFormat('#,##0').setFontWeight('bold').setBackground('#d9ead3');
  
  sheet.getRange('A17').setValue('  - Current Portion');
  sheet.getRange('B17').setFormula('=\'Lease Liability Schedule\'!I13').setNumberFormat('#,##0');
  
  sheet.getRange('A18').setValue('  - Non-Current Portion');
  sheet.getRange('B18').setFormula('=\'Lease Liability Schedule\'!J13').setNumberFormat('#,##0');
  
  // IGAAP vs Ind AS 116 Impact
  sheet.getRange('A20:F20').merge()
    .setValue('IGAAP VS IND AS 116 IMPACT ANALYSIS')
    .setBackground('#d9ead3')
    .setFontWeight('bold');
  
  const impactHeaders = ['Item', 'IGAAP Treatment', 'Ind AS 116 Treatment', 'Impact (₹)', '', ''];
  sheet.getRange('A21:F21').setValues([impactHeaders])
    .setBackground('#d9d9d9')
    .setFontWeight('bold');
  
  sheet.getRange('A22').setValue('Rent Expense');
  sheet.getRange('B22').setFormula('=\'Lease Liability Schedule\'!F13').setNumberFormat('#,##0');
  sheet.getRange('C22').setValue('Eliminated');
  sheet.getRange('D22').setFormula('=-B22').setNumberFormat('#,##0');
  
  sheet.getRange('A23').setValue('Depreciation Expense');
  sheet.getRange('B23').setValue('Nil');
  sheet.getRange('C23').setFormula('=\'ROU Asset Schedule\'!H13').setNumberFormat('#,##0');
  sheet.getRange('D23').setFormula('=C23').setNumberFormat('#,##0');
  
  sheet.getRange('A24').setValue('Interest Expense');
  sheet.getRange('B24').setValue('Nil');
  sheet.getRange('C24').setFormula('=\'Lease Liability Schedule\'!G13').setNumberFormat('#,##0');
  sheet.getRange('D24').setFormula('=C24').setNumberFormat('#,##0');
  
  sheet.getRange('A25').setValue('Net P&L Impact').setFontWeight('bold');
  sheet.getRange('D25').setFormula('=SUM(D22:D24)').setNumberFormat('#,##0').setFontWeight('bold').setBackground('#fff2cc');
  
  sheet.getRange('A27').setValue('Balance Sheet Impact:').setFontWeight('bold');
  sheet.getRange('A28').setValue('ROU Assets recognized');
  sheet.getRange('D28').setFormula('=\'ROU Asset Schedule\'!J13').setNumberFormat('#,##0');
  
  sheet.getRange('A29').setValue('Lease Liabilities recognized');
  sheet.getRange('D29').setFormula('=\'Lease Liability Schedule\'!H13').setNumberFormat('#,##0');
  
  sheet.getRange('A30').setValue('Net Impact on Equity').setFontWeight('bold');
  sheet.getRange('D30').setFormula('=D28-D29').setNumberFormat('#,##0').setFontWeight('bold').setBackground('#fff2cc');
  
  // Column widths
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);
  
  // Borders
  sheet.getRange('A4:B8').setBorder(true, true, true, true, true, true);
  sheet.getRange('A11:B18').setBorder(true, true, true, true, true, true);
  sheet.getRange('A21:D30').setBorder(true, true, true, true, true, true);
}

function createReferencesSheet(ss) {
  let sheet = ss.getSheetByName('Ind AS 116 References');
  if (!sheet) {
    sheet = ss.insertSheet('Ind AS 116 References', 9);
  }
  
  sheet.clear();
  
  sheet.getRange('A1:D1').merge()
    .setValue('IND AS 116 - LEASES: KEY REFERENCES')
    .setBackground('#1a472a')
    .setFontColor('#ffffff')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  const headers = ['Para #', 'Topic', 'Key Requirement', 'Application in this Workbook'];
  sheet.getRange('A3:D3').setValues([headers])
    .setBackground('#d9d9d9')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  const references = [
    ['Para 9', 'Lease Identification', 'Contract conveys right to control use of identified asset', 'All leases in Lease Register meet this criteria'],
    ['Para 22', 'Initial Recognition', 'Recognize ROU asset and lease liability at commencement', 'See ROU Asset & Lease Liability schedules'],
    ['Para 24', 'ROU Asset Measurement', 'Initial = Lease liability + initial direct costs + prepayments - incentives', 'Column N in Lease Register'],
    ['Para 26', 'Lease Liability Measurement', 'Initial = PV of lease payments not paid, discounted using IBR', 'Column L in Lease Register (PV formula)'],
    ['Para 32', 'ROU Asset Depreciation', 'Straight-line over shorter of lease term or useful life', 'ROU Asset Schedule - Column H'],
    ['Para 36', 'Lease Liability Remeasurement', 'Increase by interest, decrease by payments', 'Lease Liability Schedule'],
    ['Para 47', 'Presentation - Balance Sheet', 'Separately or with disclosure of line items', 'See Reconciliation sheet'],
    ['Para 49', 'Presentation - P&L', 'Depreciation separate from interest', 'Journal Entries (JE 2 and JE 3)'],
    ['Para 53', 'Disclosures', 'Maturity analysis of lease liabilities required', 'Current vs Non-Current split provided'],
    ['Para 5', 'Short-term Lease Exemption', 'Leases ≤ 12 months may be expensed', 'Applied per Assumptions sheet'],
    ['Para 6', 'Low-value Asset Exemption', 'Low-value leases may be expensed', 'Threshold in Assumptions sheet'],
    ['Para C8', 'Transition - Modified Retrospective', 'Recognize cumulative effect in retained earnings', 'Not shown - for first-time adoption only']
  ];
  
  sheet.getRange(4, 1, references.length, 4).setValues(references);
  sheet.getRange('A4:D15').setWrap(true).setVerticalAlignment('top');
  
  sheet.getRange('A17:D17').merge()
    .setValue('KEY DEFINITIONS')
    .setBackground('#d9ead3')
    .setFontWeight('bold');
  
  const definitions = [
    ['Term', 'Definition', '', ''],
    ['Lease', 'Contract that conveys right to control use of identified asset for a period in exchange for consideration', '', ''],
    ['ROU Asset', 'Lessee\'s right to use an underlying asset over the lease term', '', ''],
    ['Lease Liability', 'Lessee\'s obligation to make lease payments, measured at PV', '', ''],
    ['IBR', 'Incremental Borrowing Rate - rate lessee would pay to borrow funds to obtain similar asset', '', ''],
    ['Lease Term', 'Non-cancellable period + periods covered by extension/termination options reasonably certain to exercise', '', ''],
    ['Commencement Date', 'Date lessor makes underlying asset available for use by lessee', '', '']
  ];
  
  sheet.getRange(18, 1, definitions.length, 4).setValues(definitions);
  sheet.getRange('A18:D18').setBackground('#d9d9d9').setFontWeight('bold');
  sheet.getRange('A19:D25').setWrap(true).setVerticalAlignment('top');
  
  // Column widths
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 300);
  sheet.setColumnWidth(4, 250);
  
  // Borders
  sheet.getRange('A3:D15').setBorder(true, true, true, true, true, true);
  sheet.getRange('A18:D25').setBorder(true, true, true, true, true, true);
}

function createAuditTrailSheet(ss) {
  let sheet = ss.getSheetByName('Audit Trail');
  if (!sheet) {
    sheet = ss.insertSheet('Audit Trail', 10);
  }
  
  sheet.clear();
  
  sheet.getRange('A1:E1').merge()
    .setValue('AUDIT TRAIL & CONTROL CHECKS')
    .setBackground('#1a472a')
    .setFontColor('#ffffff')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.setRowHeight(1, 40);
  
  sheet.getRange('A2:E2').merge()
    .setValue('Automated control totals and audit assertions')
    .setBackground('#cfe2f3')
    .setHorizontalAlignment('center');
  
  // Control Checks Section
  sheet.getRange('A4:E4').merge()
    .setValue('CONTROL TOTALS')
    .setBackground('#d9ead3')
    .setFontWeight('bold');
  
  const controlHeaders = ['Control Check', 'Expected', 'Actual', 'Status', 'Notes'];
  sheet.getRange('A5:E5').setValues([controlHeaders])
    .setBackground('#d9d9d9')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Control 1: Initial recognition
  sheet.getRange('A6').setValue('Initial: ROU Asset = Lease Liability');
  sheet.getRange('B6').setFormula('=\'Lease Register\'!N21').setNumberFormat('#,##0');
  sheet.getRange('C6').setFormula('=\'Lease Register\'!L21').setNumberFormat('#,##0');
  sheet.getRange('D6').setFormula('=IF(ABS(B6-C6)<100,"PASS","FAIL")').setFontWeight('bold');
  sheet.getRange('E6').setValue('At commencement, ROU asset should equal PV of lease payments + initial costs');
  
  // Control 2: P&L impact
  sheet.getRange('A7').setValue('P&L: Depreciation + Interest = Approx. Rent');
  sheet.getRange('B7').setFormula('=\'Lease Liability Schedule\'!F13').setNumberFormat('#,##0');
  sheet.getRange('C7').setFormula('=\'ROU Asset Schedule\'!H13+\'Lease Liability Schedule\'!G13').setNumberFormat('#,##0');
  sheet.getRange('D7').setFormula('=IF(ABS((C7-B7)/B7)<0.15,"PASS","FAIL")').setFontWeight('bold');
  sheet.getRange('E7').setValue('Total expense under Ind AS 116 should approximate straight-line rent over lease term');
  
  // Control 3: Balance sheet reconciliation
  sheet.getRange('A8').setValue('Balance Sheet: Assets reconcile to schedules');
  sheet.getRange('B8').setFormula('=\'ROU Asset Schedule\'!J13').setNumberFormat('#,##0');
  sheet.getRange('C8').setFormula('=Reconciliation!B8').setNumberFormat('#,##0');
  sheet.getRange('D8').setFormula('=IF(B8=C8,"PASS","FAIL")').setFontWeight('bold');
  sheet.getRange('E8').setValue('ROU Asset closing balance must tie to reconciliation');
  
  // Control 4: Journal entries balanced
  sheet.getRange('A9').setValue('Journal Entries: Debits = Credits');
  sheet.getRange('B9').setFormula('=\'Journal Entries\'!B25').setNumberFormat('#,##0');
  sheet.getRange('C9').setFormula('=\'Journal Entries\'!B26').setNumberFormat('#,##0');
  sheet.getRange('D9').setFormula('=IF(B9=C9,"PASS","FAIL")').setFontWeight('bold');
  sheet.getRange('E9').setValue('All journal entries must balance');
  
  // Control 5: Current vs non-current split
  sheet.getRange('A10').setValue('Liability Split: Current + Non-Current = Total');
  sheet.getRange('B10').setFormula('=\'Lease Liability Schedule\'!H13').setNumberFormat('#,##0');
  sheet.getRange('C10').setFormula('=\'Lease Liability Schedule\'!I13+\'Lease Liability Schedule\'!J13').setNumberFormat('#,##0');
  sheet.getRange('D10').setFormula('=IF(B10=C10,"PASS","FAIL")').setFontWeight('bold');
  sheet.getRange('E10').setValue('Current and non-current portions must sum to total liability');
  
  // Audit Assertions Section
  sheet.getRange('A12:E12').merge()
    .setValue('AUDIT ASSERTIONS CHECKLIST')
    .setBackground('#d9ead3')
    .setFontWeight('bold');
  
  const assertionHeaders = ['Assertion', 'Account', 'Check Performed', 'Result', 'Supporting Evidence'];
  sheet.getRange('A13:E13').setValues([assertionHeaders])
    .setBackground('#d9d9d9')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  const assertions = [
    ['Existence', 'ROU Assets', 'All leases in register have valid contracts', 'VERIFY', 'Lease agreements on file'],
    ['Completeness', 'Lease Liabilities', 'All leases captured in register', 'VERIFY', 'Cross-check with rent expense ledger'],
    ['Valuation', 'ROU Assets', 'PV calculated using appropriate IBR', 'AUTO-CHECK', 'IBR sourced from treasury/bank'],
    ['Valuation', 'Lease Liabilities', 'Future payments discounted correctly', 'AUTO-CHECK', 'PV formula in Lease Register'],
    ['Accuracy', 'Depreciation', 'Calculated per straight-line method', 'AUTO-CHECK', 'Formula in ROU Asset Schedule'],
    ['Accuracy', 'Interest Expense', 'Calculated using effective interest method', 'AUTO-CHECK', 'Formula in Lease Liability Schedule'],
    ['Classification', 'Current vs Non-Current', 'Split based on maturity < 12 months', 'AUTO-CHECK', 'Formula in Lease Liability Schedule'],
    ['Presentation', 'Journal Entries', 'Entries formatted per chart of accounts', 'VERIFY', 'Match to COA in Assumptions'],
    ['Disclosure', 'Notes to Accounts', 'Lease commitments and maturities disclosed', 'VERIFY', 'Prepare disclosure note separately']
  ];
  
  sheet.getRange(14, 1, assertions.length, 5).setValues(assertions);
  
  // Summary Section
  sheet.getRange('A24:E24').merge()
    .setValue('AUDIT TRAIL SUMMARY')
    .setBackground('#e8f0fe')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.getRange('A25').setValue('Total Control Checks:').setFontWeight('bold');
  sheet.getRange('B25').setValue(5);
  
  sheet.getRange('A26').setValue('Passed:').setFontWeight('bold');
  sheet.getRange('B26').setFormula('=COUNTIF(D6:D10,"PASS")').setFontWeight('bold');
  
  sheet.getRange('A27').setValue('Failed:').setFontWeight('bold');
  sheet.getRange('B27').setFormula('=COUNTIF(D6:D10,"FAIL")').setFontWeight('bold');
  
  sheet.getRange('A29').setValue('Overall Status:').setFontWeight('bold');
  sheet.getRange('B29').setFormula('=IF(B27=0,"✓ ALL CHECKS PASSED","✗ REVIEW FAILURES")')
    .setFontWeight('bold')
    .setFontSize(12);
  
  // Conditional formatting
  const passRange = sheet.getRange('D6:D10');
  const passRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('PASS')
    .setBackground('#d9ead3')
    .setFontColor('#38761d')
    .setRanges([passRange])
    .build();
  
  const failRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('FAIL')
    .setBackground('#f4cccc')
    .setFontColor('#cc0000')
    .setRanges([passRange])
    .build();
  
  const statusRange = sheet.getRange('B29');
  const allPassRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('PASSED')
    .setBackground('#d9ead3')
    .setFontColor('#38761d')
    .setRanges([statusRange])
    .build();
  
  const someFailRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('REVIEW')
    .setBackground('#f4cccc')
    .setFontColor('#cc0000')
    .setRanges([statusRange])
    .build();
  
  let rules = sheet.getConditionalFormatRules();
  rules.push(passRule, failRule, allPassRule, someFailRule);
  sheet.setConditionalFormatRules(rules);
  
  // Column widths
  sheet.setColumnWidth(1, 300);
  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 280);
  
  // Borders
  sheet.getRange('A5:E10').setBorder(true, true, true, true, true, true);
  sheet.getRange('A13:E22').setBorder(true, true, true, true, true, true);
  
  // Wrap text
  sheet.getRange('E6:E10').setWrap(true).setVerticalAlignment('top');
  sheet.getRange('C14:E22').setWrap(true).setVerticalAlignment('top');
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 * MAIN MENU INTEGRATION
 * ═══════════════════════════════════════════════════════════════════════════
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Ind AS 116 Builder')
    .addItem('🔨 Create Ind AS 116 Workbook', 'createIndAS116Workbook')
    .addSeparator()
    .addItem('📖 User Guide', 'showUserGuide')
    .addItem('ℹ️ About', 'showAbout')
    .addToUi();
}

function showUserGuide() {
  const html = HtmlService.createHtmlOutput(`
    <h2>Ind AS 116 Lease Workings - User Guide</h2>
    <p><strong>Step 1:</strong> Run "Create Ind AS 116 Workbook" from the menu</p>
    <p><strong>Step 2:</strong> Fill in the Assumptions sheet (light blue cells)</p>
    <p><strong>Step 3:</strong> Complete the Lease Register with all lease details</p>
    <p><strong>Step 4:</strong> Review auto-calculated schedules</p>
    <p><strong>Step 5:</strong> Extract journal entries for your accounting system</p>
    <p><strong>Step 6:</strong> Verify control totals in Audit Trail sheet</p>
    <br>
    <p>All formulas are automated - only input required in light blue cells!</p>
  `).setWidth(500).setHeight(350);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'User Guide');
}

function showAbout() {
  const html = HtmlService.createHtmlOutput(`
    <h2>Ind AS 116 Audit Builder</h2>
    <p><strong>Version:</strong> 1.0</p>
    <p><strong>Standard:</strong> Ind AS 116 - Leases</p>
    <p><strong>Effective:</strong> 1 April 2019</p>
    <br>
    <p>This tool creates production-ready lease accounting workings compliant with Ind AS 116.</p>
    <p>Features:</p>
    <ul>
      <li>Automated ROU asset and lease liability calculations</li>
      <li>Period book closure journal entries</li>
      <li>IGAAP vs Ind AS 116 reconciliation</li>
      <li>Built-in control totals and audit checks</li>
    </ul>
    <br>
    <p><em>For audit and professional use</em></p>
  `).setWidth(500).setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'About Ind AS 116 Builder');
}