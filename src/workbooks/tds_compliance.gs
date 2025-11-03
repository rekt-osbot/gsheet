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
