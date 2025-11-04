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
