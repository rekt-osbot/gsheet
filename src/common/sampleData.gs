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
  'ICFR_P2P': getSampleDataICFRP2P,
  'IA_MASTER': getSampleDataIAMaster
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

function getSampleDataIAMaster() {
  return {
    'Workpaper Index': {
      'A4:G9': [
        ['H1-REV-01', 'H1', 'Revenue/OTC', 'Revenue recognition walkthrough', 'TM1', 'Complete', "=DATE(2025,10,5)"],
        ['H1-P2P-02', 'H1', 'P2P', 'Three-way match testing', 'TM2', 'In Progress', "=DATE(2025,10,7)"],
        ['H1-TAX-03', 'H1', 'Taxation', 'Indirect tax reconciliations', 'TM3', 'In Progress', "=DATE(2025,10,9)"],
        ['H1-TRE-04', 'H1', 'Treasury', 'Cash forecasting controls review', 'TM1', 'Complete', "=DATE(2025,10,12)"],
        ['Q3-PAY-01', 'Q3', 'Payroll/HR', 'Payroll master data audit', 'IA Manager', 'Complete', "=DATE(2026,1,18)"],
        ['Q4-FAR-01', 'Q4', 'Fixed Assets', 'Asset verification procedures', 'TM2', 'Complete', "=DATE(2026,4,12)"],
        ['Q4-IFC-02', 'Q4', 'IFC', 'Control design walkthroughs', 'TM3', 'In Progress', "=DATE(2026,4,20)"]
      ]
    },
    'Findings Tracker': {
      'A4:I8': [
        ['IA-001', 'H1', 'Revenue/OTC', 'Segregation of duties gap in billing', 'High', 'Introduce automated approval workflow', 'To be discussed with ERP team', 'Revenue Head', 'Open'],
        ['IA-002', 'H1', 'P2P', 'Delayed vendor reconciliations', 'Medium', 'Implement monthly reconciliation tracker', 'Process roll-out underway', 'Procurement Lead', 'In Progress'],
        ['IA-003', 'Q4', 'Fixed Assets', 'Physical verification gaps noted', 'Critical', 'Schedule quarterly surprise checks', 'Awaiting management confirmation', 'IA Manager', 'Remediated'],
        ['IA-004', 'Q4', 'IFC', 'User access reviews not documented', 'Low', 'Implement quarterly review checklist', 'Checklist drafted for approval', 'TM2', 'Closed']
      ]
    }
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
    'This will clear all data from input sections linked to named ranges. Are you sure?',
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) {
    return;
  }

  const prefix = '_INPUT_';
  const namedRanges = ss.getNamedRanges();
  let clearedCells = 0;

  namedRanges.forEach(namedRange => {
    const name = namedRange.getName();
    if (name && name.startsWith(prefix)) {
      const range = namedRange.getRange();
      clearedCells += range.getNumRows() * range.getNumColumns();
      range.clearContent();
    }
  });

  ui.alert(
    'Sample Data Cleared',
    clearedCells > 0
      ? `Cleared input data across ${clearedCells} cells linked to named ranges.`
      : 'No named ranges with the _INPUT_ prefix were found.',
    ui.ButtonSet.OK
  );
}
