/**
 * ICFR WORKPAPER BUILDER - PROCURE-TO-PAY (P2P) PROCESS
 * 
 * This script creates a comprehensive ICFR audit workpaper for the P2P cycle
 * including RCM, Test of Design, Test of Operating Effectiveness, and Dashboard
 * 
 * HOW TO USE:
 * 1. Open a new Google Sheet
 * 2. Go to Extensions > Apps Script
 * 3. Paste this code
 * 4. Run the function: createICFRP2PWorkbook()
 * 5. Authorize the script when prompted
 */

// ============================================================================
// WORKBOOK-SPECIFIC CONFIGURATION
// ============================================================================

// Column mappings for ICFR P2P workbook
const COLS = {
  RCM: {
    CONTROL_ID: 1,
    PROCESS: 2,
    RISK: 3,
    CONTROL_ACTIVITY: 4,
    CONTROL_TYPE: 5,
    FREQUENCY: 6,
    OWNER: 7,
    KEY_CONTROL: 8
  },
  TEST_OF_DESIGN: {
    CONTROL_ID: 1,
    CONTROL_DESC: 2,
    DESIGN_PROCEDURE: 3,
    EVIDENCE: 4,
    CONCLUSION: 5,
    TESTER: 6,
    DATE: 7
  }
};

function createICFRP2PWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Set workbook type for menu detection
  setWorkbookType('ICFR_P2P');
  
  // Clear existing sheets except first one
  const sheets = ss.getSheets();
  for (let i = sheets.length - 1; i > 0; i--) {
    ss.deleteSheet(sheets[i]);
  }
  
  // Create all sheets
  createCoverSheet(ss);
  createReferencesSheet(ss);
  createRCMSummary(ss);
  createTestOfDesign(ss);
  createTestOfEffectiveness(ss);
  createDashboard(ss);
  
  // Set Cover Sheet as active
  ss.setActiveSheet(ss.getSheetByName('Cover Sheet'));
  
  SpreadsheetApp.getUi().alert('✅ P2P ICFR Workpaper Created Successfully!\n\nAll sheets have been generated with proper formatting and linkages.');
}

// ==================== COVER SHEET ====================
function createCoverSheet(ss) {
  let sheet = ss.getSheets()[0];
  sheet.setName('Cover Sheet');
  sheet.clear();
  
  // Set column widths
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 350);
  
  // Header
  sheet.getRange('A1:B1').merge()
    .setValue('INTERNAL CONTROLS OVER FINANCIAL REPORTING')
    .setBackground('#1a237e')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 40);
  
  sheet.getRange('A2:B2').merge()
    .setValue('PROCURE-TO-PAY (P2P) PROCESS')
    .setBackground('#283593')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(2, 30);
  
  // Workpaper Details
  const details = [
    ['', ''],
    ['WORKPAPER DETAILS', ''],
    ['Client Name:', ''],
    ['Audit Period:', ''],
    ['Workpaper Code:', 'WP-P2P-001'],
    ['Version:', '1.0'],
    ['', ''],
    ['PREPARER INFORMATION', ''],
    ['Prepared By:', ''],
    ['Designation:', ''],
    ['Date Prepared:', new Date()],
    ['', ''],
    ['REVIEWER INFORMATION', ''],
    ['Reviewed By:', ''],
    ['Designation:', ''],
    ['Date Reviewed:', ''],
    ['', ''],
    ['AUDIT SCOPE', ''],
    ['Process Scope:', 'Complete Procure-to-Pay Cycle'],
    ['Sub-processes:', 'Vendor Management, Purchase Requisition, Purchase Order, Goods Receipt, Invoice Processing, Payment Processing'],
    ['Applicable Standards:', 'Ind AS, Companies Act 2013, IGAAP'],
    ['Control Framework:', 'COSO 2013']
  ];
  
  sheet.getRange(3, 1, details.length, 2).setValues(details);
  
  // Format section headers
  sheet.getRange('A4').setBackground('#e3f2fd').setFontWeight('bold');
  sheet.getRange('A9').setBackground('#e3f2fd').setFontWeight('bold');
  sheet.getRange('A13').setBackground('#e3f2fd').setFontWeight('bold');
  sheet.getRange('A18').setBackground('#e3f2fd').setFontWeight('bold');
  
  // Format input cells
  const inputCells = ['B5', 'B6', 'B10', 'B11', 'B15', 'B16', 'B17'];
  inputCells.forEach(cell => {
    sheet.getRange(cell).setBackground('#fff9c4');
  });
  
  // Format date cell
  sheet.getRange('B12').setNumberFormat('dd-mmm-yyyy');
  
  // Add borders
  sheet.getRange(3, 1, details.length, 2).setBorder(true, true, true, true, true, true);
  
  // Footer
  sheet.getRange('A26:B26').merge()
    .setValue('This workpaper documents the design and operating effectiveness of P2P controls')
    .setFontSize(9)
    .setFontStyle('italic')
    .setHorizontalAlignment('center');
}

// ==================== REFERENCES & ASSERTIONS SHEET ====================
function createReferencesSheet(ss) {
  const sheet = ss.insertSheet('References & Assertions');
  
  // Set column widths
  sheet.setColumnWidths(1, 4, 200);
  
  // Header
  sheet.getRange('A1').setValue('FINANCIAL STATEMENT ASSERTIONS & IND AS REFERENCES');
  sheet.getRange('A1:D1')
    .merge()
    .setBackground('#1a237e')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(1, 35);
  
  // Assertions Table
  sheet.getRange('A3').setValue('FINANCIAL STATEMENT ASSERTIONS')
    .setBackground('#283593')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  
  const assertions = [
    ['Assertion', 'Definition', 'P2P Example Control', 'Ind AS Reference'],
    ['Existence / Occurrence', 'Assets, liabilities and equity interests exist, and recorded transactions have occurred', 'Three-way match (PO, GRN, Invoice) before payment', 'Ind AS 1, Ind AS 2'],
    ['Completeness', 'All transactions and accounts that should be presented are included', 'All invoices are recorded in the period received; Accrual for GRN without invoice', 'Ind AS 1, Ind AS 37'],
    ['Valuation / Allocation', 'Assets, liabilities, and equity are recorded at appropriate amounts', 'Invoice pricing matches approved PO rates; Foreign exchange revaluation for imports', 'Ind AS 2, Ind AS 21'],
    ['Rights & Obligations', 'Entity holds or controls rights to assets and liabilities are obligations', 'Vendor master maintained with proper due diligence; Contract review before commitment', 'Ind AS 1, Ind AS 37'],
    ['Presentation', 'Components are properly classified, described and disclosed', 'Expense classification per nature/function; Related party disclosure', 'Ind AS 1, Ind AS 24'],
    ['Accuracy', 'Amounts are recorded accurately and at correct values', 'Automated PO/invoice matching; GL coding validation', 'Ind AS 1, Ind AS 8'],
    ['Cutoff', 'Transactions are recorded in correct accounting period', 'Period-end accrual process; GRN-based inventory recognition', 'Ind AS 1, Ind AS 2']
  ];
  
  sheet.getRange(4, 1, assertions.length, 4).setValues(assertions);
  
  // Format assertions table
  sheet.getRange(4, 1, 1, 4)
    .setBackground('#3f51b5')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.getRange(5, 1, assertions.length - 1, 4)
    .setBorder(true, true, true, true, true, true);
  
  // Alternate row colors
  for (let i = 5; i <= 4 + assertions.length - 1; i++) {
    if ((i - 5) % 2 === 0) {
      sheet.getRange(i, 1, 1, 4).setBackground('#f5f5f5');
    }
  }
  
  // Ind AS References
  sheet.getRange('A' + (assertions.length + 6)).setValue('RELEVANT IND AS STANDARDS FOR P2P')
    .setBackground('#283593')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  
  const indAS = [
    ['Ind AS', 'Title', 'P2P Relevance'],
    ['Ind AS 1', 'Presentation of Financial Statements', 'Proper classification of expenses, liabilities, accruals'],
    ['Ind AS 2', 'Inventories', 'Procurement valuation, FIFO/weighted average costing'],
    ['Ind AS 8', 'Accounting Policies, Changes in Estimates and Errors', 'Consistency in expense recognition and valuation'],
    ['Ind AS 10', 'Events After Reporting Period', 'Post-period adjustments for invoices/credits'],
    ['Ind AS 16', 'Property, Plant & Equipment', 'Capitalization of qualifying expenditure'],
    ['Ind AS 21', 'Effects of Changes in Foreign Exchange Rates', 'Import purchases and payables in foreign currency'],
    ['Ind AS 24', 'Related Party Disclosures', 'Transactions with related party vendors'],
    ['Ind AS 37', 'Provisions, Contingent Liabilities', 'Accruals for goods/services received not invoiced'],
    ['Ind AS 38', 'Intangible Assets', 'Procurement of software, licenses, IP'],
    ['Ind AS 115', 'Revenue from Contracts with Customers', 'Purchase returns, discounts, rebates (contra-revenue view)']
  ];
  
  const indASStartRow = assertions.length + 7;
  sheet.getRange(indASStartRow, 1, indAS.length, 3).setValues(indAS);
  
  // Format Ind AS table
  sheet.getRange(indASStartRow, 1, 1, 3)
    .setBackground('#3f51b5')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.getRange(indASStartRow, 1, indAS.length, 3)
    .setBorder(true, true, true, true, true, true);
  
  // Alternate row colors
  for (let i = indASStartRow + 1; i < indASStartRow + indAS.length; i++) {
    if ((i - indASStartRow - 1) % 2 === 0) {
      sheet.getRange(i, 1, 1, 3).setBackground('#f5f5f5');
    }
  }
  
  // Freeze header
  sheet.setFrozenRows(1);
  
  // Named ranges for lookups
  ss.setNamedRange('Assertions', sheet.getRange(5, 1, assertions.length - 1, 1));
  ss.setNamedRange('IndAS_List', sheet.getRange(indASStartRow + 1, 1, indAS.length - 1, 1));
}

// ==================== RCM SUMMARY SHEET ====================
function createRCMSummary(ss) {
  const sheet = ss.insertSheet('RCM Summary');
  
  // Set column widths
  const widths = [40, 120, 150, 150, 100, 250, 120, 100, 120, 180, 150, 150, 150];
  for (let i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }
  
  // Header
  sheet.getRange('A1').setValue('RISK CONTROL MATRIX - PROCURE-TO-PAY PROCESS');
  sheet.getRange('A1:M1')
    .merge()
    .setBackground('#1a237e')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(1, 35);
  
  // Column headers
  const headers = [
    ['#', 'Process', 'Sub-Process', 'Risk Description', 'Risk Category', 'Control Description', 
     'Control Type', 'Frequency', 'Control Owner', 'Evidence Source', 'Financial Assertion', 'Ind AS Reference', 'Link to Narrative']
  ];
  
  sheet.getRange(2, 1, 1, 13).setValues(headers);
  sheet.getRange(2, 1, 1, 13)
    .setBackground('#3f51b5')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(2, 45);
  
  // P2P Controls Data
  const controls = [
    [1, 'P2P', 'Vendor Management', 'Unauthorized or fraudulent vendors added to master', 'Access Control', 
     'Segregation of duties: Vendor creation requires approval from Finance Manager. System enforces maker-checker for new vendors', 
     'Preventive', 'Each Transaction', 'Finance Manager', 'Vendor master change log, approval workflow', 
     'Existence / Occurrence', 'Ind AS 1, Ind AS 24', 'Section 3.1'],
    
    [2, 'P2P', 'Vendor Management', 'Duplicate vendor records leading to payment errors', 'Data Integrity', 
     'System validation checks for duplicate PAN, GSTIN, and bank account numbers before vendor creation', 
     'Preventive', 'Each Transaction', 'IT Systems', 'System validation rules, error logs', 
     'Accuracy', 'Ind AS 1', 'Section 3.1'],
    
    [3, 'P2P', 'Vendor Management', 'Payments to non-compliant or blacklisted vendors', 'Compliance', 
     'Quarterly review of vendor master against blacklist databases and compliance status verification', 
     'Detective', 'Quarterly', 'Procurement Head', 'Vendor compliance review report', 
     'Rights & Obligations', 'Ind AS 24, Ind AS 37', 'Section 3.2'],
    
    [4, 'P2P', 'Purchase Requisition', 'Unauthorized or inappropriate purchases', 'Authorization', 
     'All purchase requisitions require approval based on delegation of authority matrix. Budget checks performed', 
     'Preventive', 'Each Transaction', 'Department Head', 'Approved PR with budget validation', 
     'Existence / Occurrence', 'Ind AS 1, Ind AS 16', 'Section 4.1'],
    
    [5, 'P2P', 'Purchase Order', 'POs created without valid PR or exceed approved amounts', 'Authorization', 
     'System restricts PO creation without approved PR reference. PO amount cannot exceed PR+10% tolerance', 
     'Preventive', 'Each Transaction', 'ERP System', 'PO-PR linkage report, system controls', 
     'Valuation / Allocation', 'Ind AS 2, Ind AS 16', 'Section 4.2'],
    
    [6, 'P2P', 'Purchase Order', 'Incorrect pricing, terms, or vendor details in PO', 'Accuracy', 
     'PO reviewer validates pricing against rate contracts, terms, and vendor details before approval', 
     'Preventive', 'Each Transaction', 'Procurement Team', 'Approved PO with review checklist', 
     'Accuracy', 'Ind AS 2, Ind AS 21', 'Section 4.2'],
    
    [7, 'P2P', 'Goods Receipt', 'Recording of goods not physically received or quality issues', 'Completeness', 
     'Physical verification and quality inspection required before GRN creation. GRN linked to PO in system', 
     'Preventive', 'Each Transaction', 'Warehouse Manager', 'GRN with inspection report, PO reference', 
     'Existence / Occurrence', 'Ind AS 2', 'Section 5.1'],
    
    [8, 'P2P', 'Goods Receipt', 'GRN quantity variance or specification mismatch', 'Accuracy', 
     'System alerts for GRN quantity >5% variance from PO. Specifications verified per PO', 
     'Detective', 'Each Transaction', 'ERP System / QA', 'Variance report, QA inspection log', 
     'Accuracy', 'Ind AS 2', 'Section 5.1'],
    
    [9, 'P2P', 'Invoice Processing', 'Duplicate invoice processing and payment', 'Duplication Control', 
     'System checks invoice number uniqueness per vendor. OCR scans for duplicate invoices', 
     'Preventive', 'Each Transaction', 'AP System', 'System duplication check log', 
     'Completeness', 'Ind AS 1, Ind AS 37', 'Section 6.1'],
    
    [10, 'P2P', 'Invoice Processing', 'Invoices processed without valid PO/GRN', 'Authorization', 
     'Three-way match (PO-GRN-Invoice) mandatory for >₹50,000. Tolerance: ±2% for amount variance', 
     'Preventive', 'Each Transaction', 'AP Team / System', 'Three-way match exception report', 
     'Existence / Occurrence', 'Ind AS 1, Ind AS 2', 'Section 6.1'],
    
    [11, 'P2P', 'Invoice Processing', 'Incorrect GL coding of expenses', 'Classification', 
     'Invoice GL coding auto-populated from PO category. Manual coding requires Finance approval', 
     'Preventive', 'Each Transaction', 'Finance Team', 'GL coding validation report', 
     'Presentation', 'Ind AS 1', 'Section 6.2'],
    
    [12, 'P2P', 'Invoice Processing', 'Incorrect period of expense recognition', 'Cutoff', 
     'Month-end cutoff review: GRNs without invoices accrued; invoices dated post-period reversed', 
     'Detective', 'Monthly', 'Finance Manager', 'Accrual schedule, cutoff review checklist', 
     'Cutoff', 'Ind AS 1, Ind AS 37', 'Section 6.3'],
    
    [13, 'P2P', 'Payment Processing', 'Unauthorized or fraudulent payment releases', 'Authorization', 
     'Payment batch requires dual approval (Finance Manager + CFO for >₹10L). Bank portal has MFA', 
     'Preventive', 'Each Transaction', 'Finance Manager / CFO', 'Payment approval log, bank records', 
     'Existence / Occurrence', 'Ind AS 1, Ind AS 7', 'Section 7.1'],
    
    [14, 'P2P', 'Payment Processing', 'Payment to incorrect vendor or bank account', 'Accuracy', 
     'Vendor bank details verified against master before payment. Change in bank details requires signed authorization', 
     'Preventive', 'Each Transaction', 'AP Team', 'Bank detail verification log', 
     'Accuracy', 'Ind AS 1', 'Section 7.1'],
    
    [15, 'P2P', 'Payment Processing', 'Duplicate payments or overpayments', 'Completeness', 
     'Pre-payment review of open items. System flags invoices already paid or exceeding due amount', 
     'Detective', 'Each Transaction', 'AP System / Team', 'Payment review checklist, system alerts', 
     'Completeness', 'Ind AS 1', 'Section 7.2'],
    
    [16, 'P2P', 'Payment Processing', 'Foreign currency payment valuation errors', 'Valuation', 
     'FX rate auto-fetched from authorized source. Variance >1% requires manual review and approval', 
     'Preventive', 'Each Transaction', 'Treasury / System', 'FX rate application log', 
     'Valuation / Allocation', 'Ind AS 21', 'Section 7.3'],
    
    [17, 'P2P', 'Vendor Reconciliation', 'Unreconciled vendor balances or aging discrepancies', 'Reconciliation', 
     'Monthly vendor statement reconciliation for top 80% vendors by value. Disputes logged and resolved', 
     'Detective', 'Monthly', 'AP Team', 'Vendor reconciliation statements, dispute log', 
     'Completeness', 'Ind AS 1, Ind AS 37', 'Section 8.1'],
    
    [18, 'P2P', 'Related Party Transactions', 'Non-disclosure or improper approval of RPTs', 'Compliance', 
     'Related party vendors flagged in master. RPT transactions require Board/Audit Committee approval per materiality', 
     'Preventive', 'Each Transaction', 'Company Secretary / Legal', 'RPT approval records, board minutes', 
     'Presentation', 'Ind AS 24', 'Section 9.1']
  ];
  
  sheet.getRange(3, 1, controls.length, 13).setValues(controls);
  
  // Add data validation
  const riskCategories = ['Access Control', 'Authorization', 'Accuracy', 'Completeness', 'Duplication Control', 
                         'Data Integrity', 'Compliance', 'Classification', 'Cutoff', 'Reconciliation', 'Valuation'];
  const controlTypes = ['Preventive', 'Detective'];
  const frequencies = ['Each Transaction', 'Daily', 'Weekly', 'Monthly', 'Quarterly', 'Annually'];
  
  const riskCategoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(riskCategories, true)
    .build();
  sheet.getRange(3, 5, controls.length, 1).setDataValidation(riskCategoryRule);
  
  const controlTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(controlTypes, true)
    .build();
  sheet.getRange(3, 7, controls.length, 1).setDataValidation(controlTypeRule);
  
  const frequencyRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(frequencies, true)
    .build();
  sheet.getRange(3, 8, controls.length, 1).setDataValidation(frequencyRule);
  
  // Format data area
  sheet.getRange(3, 1, controls.length, 13).setBorder(true, true, true, true, true, true);
  
  // Alternate row colors
  for (let i = 3; i < 3 + controls.length; i++) {
    if ((i - 3) % 2 === 0) {
      sheet.getRange(i, 1, 1, 13).setBackground('#f5f5f5');
    }
  }
  
  // Freeze headers
  sheet.setFrozenRows(2);

  // Named range for Control IDs
  ss.setNamedRange('RCM_Controls', sheet.getRange(3, 1, controls.length, 13));
  ss.setNamedRange('Control_IDs', sheet.getRange(3, 1, controls.length, 1));
  
  // Add instructions comment
  sheet.getRange('A2').setNote('RCM INSTRUCTIONS:\n\n' +
    '1. Each control must have unique ID\n' +
    '2. Link controls to financial assertions\n' +
    '3. Ensure Ind AS references are accurate\n' +
    '4. Update control owner and evidence source\n' +
    '5. Use dropdowns for standardized entries');
}

// ==================== TEST OF DESIGN SHEET ====================
function createTestOfDesign(ss) {
  const sheet = ss.insertSheet('Test of Design (ToD)');
  
  // Set column widths
  const widths = [60, 250, 250, 200, 200, 200, 150];
  for (let i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }
  
  // Header
  sheet.getRange('A1').setValue('TEST OF DESIGN - P2P CONTROLS');
  sheet.getRange('A1:G1')
    .merge()
    .setBackground('#1a237e')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(1, 35);
  
  // Column headers
  const headers = [
    ['Control ID', 'Control Description', 'Design Evaluation Procedure', 
     'Evidence Reviewed', 'Design Assessment', 'Exceptions / Findings', 'Conclusion']
  ];
  
  sheet.getRange(2, 1, 1, 7).setValues(headers);
  sheet.getRange(2, 1, 1, 7)
    .setBackground('#3f51b5')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(2, 45);
  
  // Pull Control IDs and descriptions from RCM using formulas
  const rcmSheet = ss.getSheetByName('RCM Summary');
  const numControls = 18;
  
  for (let i = 0; i < numControls; i++) {
    const row = i + 3;
    
    // Control ID formula
    sheet.getRange(row, 1).setFormula(`='RCM Summary'!A${row}`);
    
    // Control Description formula
    sheet.getRange(row, 2).setFormula(`='RCM Summary'!F${row}`);
    
    // Pre-fill design evaluation procedure based on control type
    const procedures = [
      'Walkthrough with process owner, review of authorization matrix and system access controls',
      'Review of system configuration for duplicate check logic and validation rules',
      'Review of compliance review SOP, sample vendor compliance reports from last quarter',
      'Walkthrough of PR approval workflow, review delegation of authority matrix and budget integration',
      'Review of system configuration for PR-PO linkage and tolerance settings',
      'Interview with procurement team, review sample POs with rate contract validation evidence',
      'Walkthrough of GRN process, review inspection checklist and sample GRNs',
      'Review of system variance settings and sample variance reports with resolution',
      'Review of system duplicate invoice check configuration and OCR functionality',
      'Walkthrough of three-way match process, review system settings and exception handling procedures',
      'Review of GL coding logic in ERP, sample manual invoice coding approvals',
      'Review of month-end cutoff procedures, sample accrual schedules from prior periods',
      'Review of payment approval matrix, bank portal MFA settings, sample payment approvals',
      'Review of vendor master maintenance SOP, sample bank detail change requests with authorization',
      'Walkthrough of pre-payment review process, review system duplicate payment controls',
      'Review of FX rate source integration, sample FX rate variance reviews',
      'Review of vendor reconciliation SOP, sample reconciliation statements and dispute logs',
      'Review of RPT identification process in vendor master, sample board approvals for material RPTs'
    ];
    
    sheet.getRange(row, 3).setValue(procedures[i]);
  }
  
  // Add data validation for Design Assessment
  const assessmentRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Effective', 'Partially Effective', 'Ineffective', 'Not Tested'], true)
    .build();
  sheet.getRange(3, 5, numControls, 1).setDataValidation(assessmentRule);
  
  // Highlight input cells
  sheet.getRange(3, 3, numControls, 1).setBackground('#fff9c4'); // Procedures (editable)
  sheet.getRange(3, 4, numControls, 3).setBackground('#fff9c4'); // Evidence, Assessment, Exceptions
  
  // Format data area
  sheet.getRange(3, 1, numControls, 7).setBorder(true, true, true, true, true, true);
  
  // Conditional formatting for Design Assessment
  const effectiveRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Effective')
    .setBackground('#c8e6c9')
    .setRanges([sheet.getRange(3, 5, numControls, 1)])
    .build();
  
  const partialRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Partially Effective')
    .setBackground('#fff9c4')
    .setRanges([sheet.getRange(3, 5, numControls, 1)])
    .build();
  
  const ineffectiveRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Ineffective')
    .setBackground('#ffcdd2')
    .setRanges([sheet.getRange(3, 5, numControls, 1)])
    .build();
  
  const rules = sheet.getConditionalFormatRules();
  rules.push(effectiveRule, partialRule, ineffectiveRule);
  sheet.setConditionalFormatRules(rules);
  
  // Alternate row colors
  for (let i = 3; i < 3 + numControls; i++) {
    if ((i - 3) % 2 === 0) {
      sheet.getRange(i, 1, 1, 7).setBackground('#f5f5f5');
    }
  }
  
  // Freeze headers
  sheet.setFrozenRows(2);

  // Named range
  ss.setNamedRange('ToD_Results', sheet.getRange(3, 1, numControls, 7));
  
  // Add instructions
  sheet.getRange('A' + (numControls + 4)).setValue('ToD INSTRUCTIONS:')
    .setFontWeight('bold')
    .setBackground('#e3f2fd');
  sheet.getRange('A' + (numControls + 5)).setValue(
    '1. Review control design through walkthroughs and documentation\n' +
    '2. Document specific evidence reviewed\n' +
    '3. Assess if control is designed effectively to address the risk\n' +
    '4. Note any design deficiencies or gaps\n' +
    '5. Effective design is prerequisite for ToE testing'
  ).setWrap(true);
  sheet.setRowHeight(numControls + 5, 80);
}

// ==================== TEST OF OPERATING EFFECTIVENESS SHEET ====================
function createTestOfEffectiveness(ss) {
  const sheet = ss.insertSheet('Test of Effectiveness (ToE)');
  
  // Set column widths
  const widths = [60, 250, 120, 80, 250, 150, 250, 250, 150];
  for (let i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }
  
  // Header
  sheet.getRange('A1').setValue('TEST OF OPERATING EFFECTIVENESS - P2P CONTROLS');
  sheet.getRange('A1:I1')
    .merge()
    .setBackground('#1a237e')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(1, 35);
  
  // Column headers
  const headers = [
    ['Control ID', 'Control Description', 'Population Period', 'Sample Size', 
     'Testing Steps Performed', 'Sample Reference', 'Test Results', 'Exceptions Found', 'ToE Conclusion']
  ];
  
  sheet.getRange(2, 1, 1, 9).setValues(headers);
  sheet.getRange(2, 1, 1, 9)
    .setBackground('#3f51b5')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(2, 45);
  
  // Pull Control IDs and descriptions from RCM
  const numControls = 18;
  
  for (let i = 0; i < numControls; i++) {
    const row = i + 3;
    
    // Control ID formula
    sheet.getRange(row, 1).setFormula(`='RCM Summary'!A${row}`);
    
    // Control Description formula
    sheet.getRange(row, 2).setFormula(`='RCM Summary'!F${row}`);
    
    // Default population period
    sheet.getRange(row, 3).setValue('Apr-Sep 2024 (6 months)');
    
    // Sample sizes based on control frequency
    const sampleSizes = [25, 25, 4, 25, 25, 25, 25, 25, 25, 25, 25, 6, 25, 25, 25, 25, 6, 4];
    sheet.getRange(row, 4).setValue(sampleSizes[i]);
    
    // Testing steps based on control nature
    const testingSteps = [
      'Selected 25 new vendors added in period. Verified approval from Finance Manager, checked maker-checker workflow completion',
      'Tested 25 vendor creation instances for duplicate check execution. Verified system rejection of duplicates',
      'Reviewed 4 quarterly vendor compliance reports. Verified blacklist screening performed and documented',
      'Selected 25 PRs across amount bands. Verified approval per DOA matrix, budget availability check performed',
      'Selected 25 POs. Verified PR linkage, checked PO amount within PR+10% tolerance, confirmed system enforcement',
      'Selected 25 POs. Reperformed pricing verification against rate contracts, validated terms and vendor details match',
      'Selected 25 GRNs. Verified physical inspection sign-off, QA report attached, PO reference present',
      'Selected 25 GRNs with variances. Verified system alerts generated, variance resolution documented',
      'Tested system on 25 invoices for duplicate check. Attempted to enter duplicate invoice number (rejected)',
      'Selected 25 high-value invoices. Verified three-way match performed, variances within tolerance, approvals obtained',
      'Selected 25 manually coded invoices. Verified Finance approval obtained, GL code matches PO category logic',
      'Reviewed 6 month-end cutoff checklists. Verified GRN-without-invoice accrued, post-period invoices identified',
      'Selected 25 payment batches. Verified dual approval (Finance Manager + CFO for >₹10L), confirmed bank MFA logs',
      'Selected 25 payments. Verified vendor bank details match master, changes supported by signed authorization',
      'Selected 25 payment runs. Verified pre-payment review performed, no duplicate/excess amounts flagged',
      'Selected 25 FX payments. Verified rate source, reperformed rate variance calculation, confirmed approvals for >1% variance',
      'Reviewed 6 monthly vendor reconciliation reports. Verified completion for top 80% vendors, disputes tracked',
      'Reviewed 4 quarters of RPT approvals. Verified Board/AC approval per materiality threshold, minutes documented'
    ];
    
    sheet.getRange(row, 5).setValue(testingSteps[i]);
  }
  
  // Add data validation for ToE Conclusion
  const conclusionRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Operating Effectively', 'Operating Partially', 'Not Operating Effectively', 'Not Tested'], true)
    .build();
  sheet.getRange(3, 9, numControls, 1).setDataValidation(conclusionRule);
  
  // Highlight input cells
  sheet.getRange(3, 3, numControls, 1).setBackground('#fff9c4'); // Population Period
  sheet.getRange(3, 6, numControls, 4).setBackground('#fff9c4'); // Sample Reference, Results, Exceptions, Conclusion
  
  // Format data area
  sheet.getRange(3, 1, numControls, 9).setBorder(true, true, true, true, true, true);
  
  // Conditional formatting for ToE Conclusion
  const effectiveRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Operating Effectively')
    .setBackground('#c8e6c9')
    .setRanges([sheet.getRange(3, 9, numControls, 1)])
    .build();
  
  const partialRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Operating Partially')
    .setBackground('#fff9c4')
    .setRanges([sheet.getRange(3, 9, numControls, 1)])
    .build();
  
  const notEffectiveRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Not Operating Effectively')
    .setBackground('#ffcdd2')
    .setRanges([sheet.getRange(3, 9, numControls, 1)])
    .build();
  
  const rules = sheet.getConditionalFormatRules();
  rules.push(effectiveRule, partialRule, notEffectiveRule);
  sheet.setConditionalFormatRules(rules);
  
  // Alternate row colors
  for (let i = 3; i < 3 + numControls; i++) {
    if ((i - 3) % 2 === 0) {
      sheet.getRange(i, 1, 1, 9).setBackground('#f5f5f5');
    }
  }
  
  // Freeze headers
  sheet.setFrozenRows(2);

  // Named range
  ss.setNamedRange('ToE_Results', sheet.getRange(3, 1, numControls, 9));
  
  // Add instructions
  sheet.getRange('A' + (numControls + 4)).setValue('ToE INSTRUCTIONS:')
    .setFontWeight('bold')
    .setBackground('#e3f2fd');
  sheet.getRange('A' + (numControls + 5)).setValue(
    '1. Define population and testing period\n' +
    '2. Determine sample size based on control frequency and risk\n' +
    '3. Select samples using random/systematic/judgmental sampling\n' +
    '4. Perform testing procedures on each sample item\n' +
    '5. Document all exceptions and rate of deviation\n' +
    '6. Conclude on operating effectiveness'
  ).setWrap(true);
  sheet.setRowHeight(numControls + 5, 100);
  
  // Add exception rate summary
  sheet.getRange('A' + (numControls + 7)).setValue('EXCEPTION RATE SUMMARY')
    .setFontWeight('bold')
    .setBackground('#283593')
    .setFontColor('#ffffff');
  
  sheet.getRange('A' + (numControls + 8)).setValue('Total Controls Tested:');
  sheet.getRange('B' + (numControls + 8)).setFormula(`=COUNTA(I3:I${2+numControls})`);
  
  sheet.getRange('A' + (numControls + 9)).setValue('Controls Operating Effectively:');
  sheet.getRange('B' + (numControls + 9)).setFormula(`=COUNTIF(I3:I${2+numControls},"Operating Effectively")`);
  
  sheet.getRange('A' + (numControls + 10)).setValue('Controls with Exceptions:');
  sheet.getRange('B' + (numControls + 10)).setFormula(`=COUNTIF(I3:I${2+numControls},"Operating Partially")+COUNTIF(I3:I${2+numControls},"Not Operating Effectively")`);
  
  sheet.getRange('A' + (numControls + 11)).setValue('Effectiveness Rate:');
  sheet.getRange('B' + (numControls + 11)).setFormula(`=IF(B${numControls+8}>0,B${numControls+9}/B${numControls+8},0)`);
  sheet.getRange('B' + (numControls + 11)).setNumberFormat('0.0%');
}

// ==================== DASHBOARD SHEET ====================
function createDashboard(ss) {
  const sheet = ss.insertSheet('Dashboard');
  sheet.setTabColor('#4caf50');
  
  // Set column widths
  for (let i = 1; i <= 8; i++) {
    sheet.setColumnWidth(i, 150);
  }
  
  // Title
  sheet.getRange('A1').setValue('P2P ICFR CONTROL EFFECTIVENESS DASHBOARD');
  sheet.getRange('A1:H1')
    .merge()
    .setBackground('#1a237e')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(1, 40);
  
  // Summary Metrics Section
  sheet.getRange('A3').setValue('EXECUTIVE SUMMARY');
  sheet.getRange('A3:H3')
    .merge()
    .setBackground('#283593')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  const summaryLabels = [
    ['Total Controls in RCM:', '', 'Controls Tested (ToD):', '', 'Controls Tested (ToE):', ''],
    ['Design Effective:', '', 'Operating Effectively:', '', 'Overall Effectiveness:', '']
  ];
  
  sheet.getRange('A4:F5').setValues(summaryLabels);
  
  // Formulas for summary metrics
  sheet.getRange('B4').setFormula(`=COUNTA('RCM Summary'!A3:A20)`);
  sheet.getRange('D4').setFormula(`=COUNTIF('Test of Design (ToD)'!E3:E20,"<>Not Tested")`);
  sheet.getRange('F4').setFormula(`=COUNTIF('Test of Effectiveness (ToE)'!I3:I20,"<>Not Tested")`);
  
  sheet.getRange('B5').setFormula(`=COUNTIF('Test of Design (ToD)'!E3:E20,"Effective")`);
  sheet.getRange('D5').setFormula(`=COUNTIF('Test of Effectiveness (ToE)'!I3:I20,"Operating Effectively")`);
  sheet.getRange('F5').setFormula(`=IF(B4>0,D5/B4,0)`);
  sheet.getRange('F5').setNumberFormat('0.0%');
  
  // Format summary section
  sheet.getRange('A4:F5').setBorder(true, true, true, true, true, true);
  sheet.getRange('B4:B5').setBackground('#e3f2fd').setHorizontalAlignment('center').setFontWeight('bold');
  sheet.getRange('D4:D5').setBackground('#e3f2fd').setHorizontalAlignment('center').setFontWeight('bold');
  sheet.getRange('F4:F5').setBackground('#e3f2fd').setHorizontalAlignment('center').setFontWeight('bold');
  
  // ToD Summary by Assessment
  sheet.getRange('A7').setValue('TEST OF DESIGN - SUMMARY');
  sheet.getRange('A7:D7')
    .merge()
    .setBackground('#283593')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  const todHeaders = [['Design Assessment', 'Count', 'Percentage', 'Status']];
  sheet.getRange('A8:D8').setValues(todHeaders);
  sheet.getRange('A8:D8')
    .setBackground('#3f51b5')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  const todRows = [
    ['Effective', '', '', 'PASS'],
    ['Partially Effective', '', '', 'REVIEW'],
    ['Ineffective', '', '', 'FAIL'],
    ['Not Tested', '', '', 'PENDING']
  ];

  sheet.getRange('A9:D12').setValues(todRows);

  // Set formulas for column B (Count)
  sheet.getRange('B9:B12').setFormulas([
    [`=COUNTIF('Test of Design (ToD)'!E3:E20,"Effective")`],
    [`=COUNTIF('Test of Design (ToD)'!E3:E20,"Partially Effective")`],
    [`=COUNTIF('Test of Design (ToD)'!E3:E20,"Ineffective")`],
    [`=COUNTIF('Test of Design (ToD)'!E3:E20,"Not Tested")`]
  ]);

  // Set formulas for column C (Percentage)
  sheet.getRange('C9:C12').setFormulas([
    [`=IF(B4>0,B9/B4,0)`],
    [`=IF(B4>0,B10/B4,0)`],
    [`=IF(B4>0,B11/B4,0)`],
    [`=IF(B4>0,B12/B4,0)`]
  ]);
  sheet.getRange('C9:C12').setNumberFormat('0.0%');
  
  // Conditional formatting for ToD
  const todPass = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('PASS')
    .setBackground('#c8e6c9')
    .setRanges([sheet.getRange('D9')])
    .build();
  
  const todReview = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('REVIEW')
    .setBackground('#fff9c4')
    .setRanges([sheet.getRange('D10')])
    .build();
  
  const todFail = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('FAIL')
    .setBackground('#ffcdd2')
    .setRanges([sheet.getRange('D11')])
    .build();
  
  let rules = [todPass, todReview, todFail];
  
  // ToE Summary by Assessment
  sheet.getRange('F7').setValue('TEST OF EFFECTIVENESS - SUMMARY');
  sheet.getRange('F7:I7')
    .merge()
    .setBackground('#283593')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  const toeHeaders = [['Operating Status', 'Count', 'Percentage', 'Status']];
  sheet.getRange('F8:I8').setValues(toeHeaders);
  sheet.getRange('F8:I8')
    .setBackground('#3f51b5')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  const toeRows = [
    ['Operating Effectively', '', '', 'PASS'],
    ['Operating Partially', '', '', 'REVIEW'],
    ['Not Operating Effectively', '', '', 'FAIL'],
    ['Not Tested', '', '', 'PENDING']
  ];

  sheet.getRange('F9:I12').setValues(toeRows);

  // Set formulas for column G (Count)
  sheet.getRange('G9:G12').setFormulas([
    [`=COUNTIF('Test of Effectiveness (ToE)'!I3:I20,"Operating Effectively")`],
    [`=COUNTIF('Test of Effectiveness (ToE)'!I3:I20,"Operating Partially")`],
    [`=COUNTIF('Test of Effectiveness (ToE)'!I3:I20,"Not Operating Effectively")`],
    [`=COUNTIF('Test of Effectiveness (ToE)'!I3:I20,"Not Tested")`]
  ]);

  // Set formulas for column H (Percentage)
  sheet.getRange('H9:H12').setFormulas([
    [`=IF(B4>0,G9/B4,0)`],
    [`=IF(B4>0,G10/B4,0)`],
    [`=IF(B4>0,G11/B4,0)`],
    [`=IF(B4>0,G12/B4,0)`]
  ]);
  sheet.getRange('H9:H12').setNumberFormat('0.0%');
  
  // Conditional formatting for ToE
  const toePass = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('PASS')
    .setBackground('#c8e6c9')
    .setRanges([sheet.getRange('I9')])
    .build();
  
  const toeReview = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('REVIEW')
    .setBackground('#fff9c4')
    .setRanges([sheet.getRange('I10')])
    .build();
  
  const toeFail = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('FAIL')
    .setBackground('#ffcdd2')
    .setRanges([sheet.getRange('I11')])
    .build();
  
  rules.push(toePass, toeReview, toeFail);
  sheet.setConditionalFormatRules(rules);
  
  // Process-wise Summary
  sheet.getRange('A14').setValue('CONTROL EFFECTIVENESS BY SUB-PROCESS');
  sheet.getRange('A14:H14')
    .merge()
    .setBackground('#283593')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  const processHeaders = [['Sub-Process', 'Total Controls', 'Design Effective', 'Operating Effectively', 'Design %', 'Operating %', 'Overall Status']];
  sheet.getRange('A15:G15').setValues(processHeaders);
  sheet.getRange('A15:G15')
    .setBackground('#3f51b5')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  const subProcesses = [
    'Vendor Management',
    'Purchase Requisition',
    'Purchase Order',
    'Goods Receipt',
    'Invoice Processing',
    'Payment Processing',
    'Vendor Reconciliation',
    'Related Party Transactions'
  ];
  
  for (let i = 0; i < subProcesses.length; i++) {
    const row = 16 + i;
    const process = subProcesses[i];
    
    sheet.getRange(row, 1).setValue(process);
    
    // Total controls formula
    sheet.getRange(row, 2).setFormula(`=COUNTIF('RCM Summary'!C3:C20,"${process}")`);
    
    // Design effective - count from ToD where subprocess matches and assessment is "Effective"
    sheet.getRange(row, 3).setFormula(
      `=SUMPRODUCT(('RCM Summary'!C3:C20="${process}")*('Test of Design (ToD)'!E3:E20="Effective"))`
    );
    
    // Operating effectively
    sheet.getRange(row, 4).setFormula(
      `=SUMPRODUCT(('RCM Summary'!C3:C20="${process}")*('Test of Effectiveness (ToE)'!I3:I20="Operating Effectively"))`
    );
    
    // Design %
    sheet.getRange(row, 5).setFormula(`=IF(B${row}>0,C${row}/B${row},0)`);
    sheet.getRange(row, 5).setNumberFormat('0.0%');
    
    // Operating %
    sheet.getRange(row, 6).setFormula(`=IF(B${row}>0,D${row}/B${row},0)`);
    sheet.getRange(row, 6).setNumberFormat('0.0%');
    
    // Overall Status
    sheet.getRange(row, 7).setFormula(
      `=IF(AND(E${row}>=0.9,F${row}>=0.9),"✓ EFFECTIVE",IF(OR(E${row}<0.7,F${row}<0.7),"✗ NEEDS ATTENTION","⚠ REVIEW REQUIRED"))`
    );
  }
  
  // Format process table
  sheet.getRange('A16:G' + (15 + subProcesses.length)).setBorder(true, true, true, true, true, true);
  
  // Conditional formatting for overall status
  const effectiveStatusRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('EFFECTIVE')
    .setBackground('#c8e6c9')
    .setRanges([sheet.getRange('G16:G' + (15 + subProcesses.length))])
    .build();
  
  const reviewStatusRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('REVIEW')
    .setBackground('#fff9c4')
    .setRanges([sheet.getRange('G16:G' + (15 + subProcesses.length))])
    .build();
  
  const attentionStatusRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('ATTENTION')
    .setBackground('#ffcdd2')
    .setRanges([sheet.getRange('G16:G' + (15 + subProcesses.length))])
    .build();
  
  const statusRules = sheet.getConditionalFormatRules();
  statusRules.push(effectiveStatusRule, reviewStatusRule, attentionStatusRule);
  sheet.setConditionalFormatRules(statusRules);
  
  // Key Findings Section
  const findingsRow = 16 + subProcesses.length + 2;
  sheet.getRange('A' + findingsRow).setValue('KEY FINDINGS & ACTION ITEMS');
  sheet.getRange('A' + findingsRow + ':H' + findingsRow)
    .merge()
    .setBackground('#283593')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  const findingsData = [
    ['Priority', 'Finding Category', 'Description', 'Action Required'],
    ['HIGH', 'Design Ineffective', 'Controls requiring immediate design remediation', '=COUNTIF(\'Test of Design (ToD)\'!E3:E20,"Ineffective") & " controls to redesign"'],
    ['HIGH', 'Not Operating', 'Controls failing operating effectiveness', '=COUNTIF(\'Test of Effectiveness (ToE)\'!I3:I20,"Not Operating Effectively") & " controls to fix"'],
    ['MEDIUM', 'Partially Effective', 'Controls with minor design issues', '=COUNTIF(\'Test of Design (ToD)\'!E3:E20,"Partially Effective") & " controls to enhance"'],
    ['MEDIUM', 'Operating Partially', 'Controls with exceptions in testing', '=COUNTIF(\'Test of Effectiveness (ToE)\'!I3:I20,"Operating Partially") & " controls to monitor"'],
    ['LOW', 'Not Tested', 'Controls pending testing', '=(COUNTIF(\'Test of Design (ToD)\'!E3:E20,"Not Tested")+COUNTIF(\'Test of Effectiveness (ToE)\'!I3:I20,"Not Tested")) & " tests pending"']
  ];
  
  sheet.getRange(findingsRow + 1, 1, findingsData.length, 4).setValues(findingsData);
  
  // Format findings section
  sheet.getRange(findingsRow + 1, 1, 1, 4)
    .setBackground('#3f51b5')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.getRange(findingsRow + 1, 1, findingsData.length, 4)
    .setBorder(true, true, true, true, true, true);
  
  // Set formulas for action items
  for (let i = 2; i <= findingsData.length; i++) {
    sheet.getRange(findingsRow + i, 4).setFormula(findingsData[i-1][3]);
  }
  
  // Priority color coding
  const highPriorityRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('HIGH')
    .setBackground('#ffcdd2')
    .setFontColor('#c62828')
    .setRanges([sheet.getRange(findingsRow + 2, 1, 2, 1)])
    .build();
  
  const mediumPriorityRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('MEDIUM')
    .setBackground('#fff9c4')
    .setFontColor('#f57c00')
    .setRanges([sheet.getRange(findingsRow + 4, 1, 2, 1)])
    .build();
  
  const lowPriorityRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('LOW')
    .setBackground('#c8e6c9')
    .setFontColor('#2e7d32')
    .setRanges([sheet.getRange(findingsRow + 6, 1, 1, 1)])
    .build();
  
  const priorityRules = sheet.getConditionalFormatRules();
  priorityRules.push(highPriorityRule, mediumPriorityRule, lowPriorityRule);
  sheet.setConditionalFormatRules(priorityRules);
  
  // Footer
  const footerRow = findingsRow + findingsData.length + 2;
  sheet.getRange('A' + footerRow + ':H' + footerRow)
    .merge()
    .setValue('Dashboard auto-updates from RCM, ToD, and ToE sheets. Refresh to see latest status.')
    .setFontSize(9)
    .setFontStyle('italic')
    .setHorizontalAlignment('center')
    .setBackground('#e3f2fd');
  
  // Freeze header
  sheet.setFrozenRows(1);
}

// ==================== HELPER FUNCTIONS ====================

/**
 * Creates a custom menu when the spreadsheet opens
 */
// onOpen() is handled by common/utilities.gs - auto-detects workbook type

/**
 * Shows information about the workpaper
 */
// showAbout() is handled by common/utilities.gs

/**
 * Run this function to set up the workpaper
 */
function setupWorkpaper() {
  createP2PWorkpaper();
}