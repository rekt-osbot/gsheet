/**
 * @OnlyCurrentDoc
 * IND AS 116 LEASE ACCOUNTING WORKPAPER
 *
 * This script creates a comprehensive, audit-ready workpaper for lease accounting
 * under Ind AS 116 standards (Indian Accounting Standards). The workpaper includes:
 * - Summary dashboard with key metrics
 * - Detailed amortization schedule (lease liability)
 * - Detailed depreciation schedule (ROU asset)
 * - Journal entry summaries for each reporting period
 * - Comprehensive formatting for professional presentation
 *
 * Version: 2.0
 * Last Updated: November 2025
 */

/**
 * Creates a custom menu in the spreadsheet UI to run the script.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Lease Accounting (Ind AS 116)')
    .addItem('Generate Workpaper', 'createIndAS116Workpaper')
    .addSeparator()
    .addItem('Refresh Formatting', 'applyFormattingOnly')
    .addToUi();
}

/**
 * Main function to generate the entire Ind AS 116 workpaper.
 */
function createIndAS116Workpaper() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Ind AS 116 Workpaper");

  if (sheet) {
    // Clear existing content but keep the sheet
    sheet.clear();
    sheet.clearFormats();
  } else {
    sheet = ss.insertSheet("Ind AS 116 Workpaper");
  }
  ss.setActiveSheet(sheet);

  // --- STEP 1: SETUP HEADER & INSTRUCTIONS ---
  setupHeader(sheet);
  
  // --- STEP 2: SETUP INPUT SECTION ---
  setupInputs(sheet);
  
  // --- STEP 3: READ INPUTS AND PERFORM CALCULATIONS ---
  const inputs = readInputs(sheet);
  const calculations = performCalculations(inputs);
  
  // --- STEP 4: CREATE SUMMARY DASHBOARD ---
  createSummaryDashboard(sheet, calculations, inputs);
  
  // --- STEP 5: GENERATE DETAILED SCHEDULES ---
  const amortizationData = generateAmortizationSchedule(calculations, inputs);
  const rouAssetData = generateROUAssetSchedule(calculations, inputs);

  // Safety check: if no data was generated, show error and exit
  if (!amortizationData || amortizationData.length === 0) {
    SpreadsheetApp.getUi().alert('❌ Error: Invalid inputs detected.\n\nPlease check:\n- Lease term is a reasonable number (e.g., 5)\n- Payment frequency is "Annually", "Quarterly", or "Monthly"\n- All required fields are filled');
    return;
  }

  // --- STEP 6: WRITE SCHEDULES TO SHEET ---
  writeSchedules(sheet, amortizationData, rouAssetData, calculations);

  // --- STEP 7: GENERATE COMPREHENSIVE JOURNAL ENTRIES ---
  generateComprehensiveJournalEntries(sheet, amortizationData, rouAssetData, calculations, inputs);

  // --- STEP 8: ADD NOTES AND ASSUMPTIONS ---
  addNotesSection(sheet, inputs, calculations);

  // --- STEP 9: APPLY COMPREHENSIVE FORMATTING ---
  applyFormatting(sheet, calculations.totalPeriods);
  
  // --- STEP 10: PROTECT FORMULA CELLS ---
  protectFormulaCells(sheet);

  SpreadsheetApp.getUi().alert('✅ Ind AS 116 Workpaper successfully created!\n\nThe workpaper is now ready for review by management and auditors.');
}

/**
 * Sets up the header section with title and document information.
 */
function setupHeader(sheet) {
  const headerData = [
    ['LEASE ACCOUNTING WORKPAPER', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['Ind AS 116 Compliance (Indian Accounting Standards)', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', '', '', '', ''],
  ];
  sheet.getRange("A1:M3").setValues(headerData);

  // Add document metadata
  const metadata = [
    ['Prepared by:', '', ''],
    ['Date prepared:', '', new Date()],
    ['Purpose:', '', 'Lease liability and ROU asset calculation and tracking']
  ];
  sheet.getRange("K1:M3").setValues(metadata);
}

/**
 * Sets up the input area with clear labels and instructions.
 */
function setupInputs(sheet) {
  const inputData = [
    ['', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['LEASE INFORMATION', '', '', '(Complete all yellow fields)', '', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['Lease Commencement Date:', '', new Date(2024, 0, 1), '', 'The date when the lessee can begin to use the leased asset', '', '', '', '', '', '', '', ''],
    ['Lease Term:', '', 5, 'years', 'Total duration of the lease agreement', '', '', '', '', '', '', '', ''],
    ['Lease Payment Amount:', '', 10000, '', 'Payment amount per period (excluding variable payments)', '', '', '', '', '', '', '', ''],
    ['Payment Frequency:', '', 'Annually', '', 'Options: Annually, Quarterly, or Monthly', '', '', '', '', '', '', '', ''],
    ['Incremental Borrowing Rate (IBR):', '', 0.05, '', 'Annual discount rate used to calculate present value', '', '', '', '', '', '', '', ''],
    ['First Reporting Period End:', '', new Date(2024, 11, 31), '', 'End date of your first reporting period (e.g., FY end)', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['Initial Direct Costs (optional):', '', 0, '', 'Costs directly attributable to negotiating and arranging the lease', '', '', '', '', '', '', '', ''],
    ['Prepaid Lease Payments (optional):', '', 0, '', 'Any payments made at or before commencement date', '', '', '', '', '', '', '', ''],
  ];
  sheet.getRange("A5:M16").setValues(inputData);
}

/**
 * Reads and validates input values from the sheet.
 */
function readInputs(sheet) {
  return {
    leaseStartDate: new Date(sheet.getRange("C7").getValue()),
    leaseTermYears: sheet.getRange("C8").getValue(),
    paymentAmount: sheet.getRange("C9").getValue(),
    paymentFrequency: String(sheet.getRange("C10").getValue() || 'Annually'),
    discountRate: sheet.getRange("C11").getValue(),
    reportingEndDate: new Date(sheet.getRange("C12").getValue()),
    initialDirectCosts: sheet.getRange("C14").getValue() || 0,
    prepaidPayments: sheet.getRange("C15").getValue() || 0
  };
}

/**
 * Performs all necessary calculations based on inputs.
 */
function performCalculations(inputs) {
  // Determine periods based on frequency
  let periodsPerYear;
  switch (inputs.paymentFrequency.toLowerCase().trim()) {
    case 'monthly':
      periodsPerYear = 12;
      break;
    case 'quarterly':
      periodsPerYear = 4;
      break;
    case 'annually':
    default:
      periodsPerYear = 1;
      break;
  }

  const totalPeriods = Math.floor(inputs.leaseTermYears * periodsPerYear);
  const periodicRate = inputs.discountRate / periodsPerYear;

  // Calculate Present Value of Lease Payments (Initial Lease Liability)
  // Only calculate if we have valid inputs
  let pvFactor = 0;
  let pvOfLeasePayments = 0;

  if (periodicRate > 0 && totalPeriods > 0) {
    pvFactor = (1 - Math.pow(1 + periodicRate, -totalPeriods)) / periodicRate;
    pvOfLeasePayments = inputs.paymentAmount * pvFactor;
  } else if (periodicRate === 0 && totalPeriods > 0) {
    // If rate is 0, PV is just sum of payments
    pvOfLeasePayments = inputs.paymentAmount * totalPeriods;
    pvFactor = totalPeriods;
  }

  // Initial Lease Liability = PV of lease payments
  const initialLiability = pvOfLeasePayments;

  // Initial ROU Asset = Lease Liability + Initial Direct Costs + Prepayments - Lease Incentives
  const initialROUAsset = initialLiability + inputs.initialDirectCosts + inputs.prepaidPayments;

  return {
    periodsPerYear: periodsPerYear,
    totalPeriods: totalPeriods,
    periodicRate: periodicRate,
    pvFactor: pvFactor,
    initialLiability: initialLiability,
    initialROUAsset: initialROUAsset,
    totalCashOutflow: inputs.paymentAmount * totalPeriods
  };
}

/**
 * Creates a summary dashboard showing key metrics and balances.
 */
function createSummaryDashboard(sheet, calculations, inputs) {
  // Set up the dashboard structure with labels only
  const dashboardData = [
    ['', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['SUMMARY DASHBOARD', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['Key Metrics', '', '', 'Initial Recognition', '', '', 'Lease Economics', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['Total Lease Payments:', '', '', 'Initial Lease Liability:', '', '', 'Implied Interest Cost:', '', '', '', '', '', ''],
    ['Number of Payments:', '', '', 'Initial ROU Asset:', '', '', 'Effective Interest Rate (periodic):', '', '', '', '', '', ''],
    ['Payment per Period:', '', '', 'Initial Direct Costs:', '', '', 'Effective Interest Rate (annual):', '', '', '', '', '', ''],
    ['Payment Frequency:', '', '', 'Prepaid Amounts:', '', '', '', '', '', '', '', '', ''],
  ];
  sheet.getRange("A18:M26").setValues(dashboardData);

  // Now set formulas for calculated values
  // Key Metrics section
  sheet.getRange("C23").setFormula('=C9*C8*IF(C10="Monthly",12,IF(C10="Quarterly",4,1))'); // Total Lease Payments
  sheet.getRange("C24").setFormula('=C8*IF(C10="Monthly",12,IF(C10="Quarterly",4,1))'); // Number of Payments
  sheet.getRange("C25").setFormula('=C9'); // Payment per Period
  sheet.getRange("C26").setFormula('=C10'); // Payment Frequency

  // Initial Recognition section
  // PV of lease payments calculation
  const pvFormula = '=C9*((1-POWER(1+C11/IF(C10="Monthly",12,IF(C10="Quarterly",4,1)),-C8*IF(C10="Monthly",12,IF(C10="Quarterly",4,1))))/(C11/IF(C10="Monthly",12,IF(C10="Quarterly",4,1))))';
  sheet.getRange("F23").setFormula(pvFormula); // Initial Lease Liability
  sheet.getRange("F24").setFormula('=F23+C14+C15'); // Initial ROU Asset
  sheet.getRange("F25").setFormula('=C14'); // Initial Direct Costs
  sheet.getRange("F26").setFormula('=C15'); // Prepaid Amounts

  // Lease Economics section
  sheet.getRange("I23").setFormula('=C23-F23'); // Implied Interest Cost
  sheet.getRange("I24").setFormula('=C11/IF(C10="Monthly",12,IF(C10="Quarterly",4,1))'); // Effective Interest Rate (periodic)
  sheet.getRange("I25").setFormula('=C11'); // Effective Interest Rate (annual)
}

/**
 * Generates the lease liability amortization schedule.
 */
function generateAmortizationSchedule(calculations, inputs) {
  const data = [];

  // Safety check: prevent infinite loops or excessive iterations
  if (!calculations.totalPeriods || calculations.totalPeriods <= 0 || calculations.totalPeriods > 1000) {
    Logger.log('Invalid total periods: ' + calculations.totalPeriods);
    return data; // Return empty array
  }

  let openingBalance = calculations.initialLiability;
  const periodicRate = calculations.periodicRate;
  const periodsPerYear = calculations.periodsPerYear;

  for (let i = 1; i <= calculations.totalPeriods; i++) {
    const interestExpense = openingBalance * periodicRate;
    const principalRepayment = inputs.paymentAmount - interestExpense;
    const closingBalance = openingBalance - principalRepayment;

    // Calculate payment date based on frequency
    let paymentDate = new Date(inputs.leaseStartDate);
    if (periodsPerYear === 12) {
      paymentDate.setMonth(paymentDate.getMonth() + i);
    } else if (periodsPerYear === 4) {
      paymentDate.setMonth(paymentDate.getMonth() + i * 3);
    } else {
      paymentDate.setFullYear(paymentDate.getFullYear() + i);
    }

    // Determine fiscal year
    const fiscalYear = paymentDate.getFullYear();

    data.push([
      i,
      paymentDate,
      fiscalYear,
      openingBalance,
      inputs.paymentAmount,
      interestExpense,
      principalRepayment,
      closingBalance
    ]);

    openingBalance = closingBalance;
  }
  return data;
}

/**
 * Generates the Right-of-Use (ROU) Asset depreciation schedule.
 */
function generateROUAssetSchedule(calculations, inputs) {
  const data = [];

  // Safety check: prevent infinite loops or excessive iterations
  if (!calculations.totalPeriods || calculations.totalPeriods <= 0 || calculations.totalPeriods > 1000) {
    Logger.log('Invalid total periods: ' + calculations.totalPeriods);
    return data; // Return empty array
  }

  let openingBalance = calculations.initialROUAsset;
  const depreciationPerPeriod = calculations.initialROUAsset / calculations.totalPeriods;
  const periodsPerYear = calculations.periodsPerYear;

  for (let i = 1; i <= calculations.totalPeriods; i++) {
    const closingBalance = openingBalance - depreciationPerPeriod;

    // Calculate period end date
    let periodEndDate = new Date(inputs.leaseStartDate);
    if (periodsPerYear === 12) {
      periodEndDate.setMonth(periodEndDate.getMonth() + i);
    } else if (periodsPerYear === 4) {
      periodEndDate.setMonth(periodEndDate.getMonth() + i * 3);
    } else {
      periodEndDate.setFullYear(periodEndDate.getFullYear() + i);
    }

    // Determine fiscal year
    const fiscalYear = periodEndDate.getFullYear();

    data.push([
      i,
      periodEndDate,
      fiscalYear,
      openingBalance,
      depreciationPerPeriod,
      closingBalance
    ]);

    openingBalance = closingBalance;
  }
  return data;
}

/**
 * Writes the calculated schedule data to the sheet with proper headers and FORMULAS.
 */
function writeSchedules(sheet, amortizationData, rouAssetData, calculations) {
  const startRow = 28;

  // Section header for schedules
  sheet.getRange(startRow, 1, 1, 13).merge()
    .setValue('DETAILED SCHEDULES')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground('#1c4587')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  // Lease Liability Amortization Schedule
  const amortStartRow = startRow + 2;
  sheet.getRange(amortStartRow, 1, 1, 8).merge()
    .setValue('A. LEASE LIABILITY AMORTIZATION SCHEDULE')
    .setFontWeight('bold')
    .setBackground('#3c78d8')
    .setFontColor('#ffffff');

  const amortHeaders = [[
    'Period #',
    'Payment Date',
    'Fiscal Year',
    'Opening Balance',
    'Cash Payment',
    'Interest Expense',
    'Principal Reduction',
    'Closing Balance'
  ]];

  const amortHeaderRow = amortStartRow + 1;
  sheet.getRange(amortHeaderRow, 1, 1, 8).setValues(amortHeaders);

  // Write formulas for amortization schedule
  const totalPeriods = calculations.totalPeriods;
  const firstDataRow = amortHeaderRow + 1;

  for (let i = 0; i < totalPeriods; i++) {
    const row = firstDataRow + i;
    const periodNum = i + 1;

    // Period # (static value)
    sheet.getRange(row, 1).setValue(periodNum);

    // Payment Date (formula based on frequency)
    const dateFormula = `=IF($C$10="Monthly",EDATE($C$7,${periodNum}),IF($C$10="Quarterly",EDATE($C$7,${periodNum}*3),DATE(YEAR($C$7)+${periodNum},MONTH($C$7),DAY($C$7))))`;
    sheet.getRange(row, 2).setFormula(dateFormula);

    // Fiscal Year (extract year from payment date)
    sheet.getRange(row, 3).setFormula(`=YEAR(B${row})`);

    // Opening Balance (first row = initial liability, subsequent = previous closing)
    if (i === 0) {
      sheet.getRange(row, 4).setFormula('=$F$23'); // Reference to Initial Lease Liability
    } else {
      sheet.getRange(row, 4).setFormula(`=H${row - 1}`); // Previous closing balance
    }

    // Cash Payment
    sheet.getRange(row, 5).setFormula('=$C$9');

    // Interest Expense = Opening Balance * Periodic Rate
    sheet.getRange(row, 6).setFormula(`=D${row}*$I$24`);

    // Principal Reduction = Payment - Interest
    sheet.getRange(row, 7).setFormula(`=E${row}-F${row}`);

    // Closing Balance = Opening - Principal
    sheet.getRange(row, 8).setFormula(`=D${row}-G${row}`);
  }

  // ROU Asset Depreciation Schedule
  const rouStartRow = amortStartRow;
  sheet.getRange(rouStartRow, 10, 1, 6).merge()
    .setValue('B. RIGHT-OF-USE ASSET DEPRECIATION SCHEDULE')
    .setFontWeight('bold')
    .setBackground('#3c78d8')
    .setFontColor('#ffffff');

  const rouHeaders = [[
    'Period #',
    'Period End',
    'Fiscal Year',
    'Opening Balance',
    'Depreciation',
    'Closing Balance'
  ]];

  const rouHeaderRow = rouStartRow + 1;
  sheet.getRange(rouHeaderRow, 10, 1, 6).setValues(rouHeaders);

  // Write formulas for ROU Asset schedule
  for (let i = 0; i < totalPeriods; i++) {
    const row = firstDataRow + i;
    const periodNum = i + 1;

    // Period # (static value)
    sheet.getRange(row, 10).setValue(periodNum);

    // Period End Date (same as payment date)
    sheet.getRange(row, 11).setFormula(`=B${row}`);

    // Fiscal Year (same as amortization)
    sheet.getRange(row, 12).setFormula(`=C${row}`);

    // Opening Balance (first row = initial ROU asset, subsequent = previous closing)
    if (i === 0) {
      sheet.getRange(row, 13).setFormula('=$F$24'); // Reference to Initial ROU Asset
    } else {
      sheet.getRange(row, 13).setFormula(`=O${row - 1}`); // Previous closing balance
    }

    // Depreciation (straight-line over lease term)
    sheet.getRange(row, 14).setFormula('=$F$24/$C$24');

    // Closing Balance = Opening - Depreciation
    sheet.getRange(row, 15).setFormula(`=M${row}-N${row}`);
  }
}

/**
 * Generates comprehensive journal entries for all reporting periods using FORMULAS.
 */
function generateComprehensiveJournalEntries(sheet, amortizationData, rouAssetData, calculations, inputs) {
  const startRow = 32 + calculations.totalPeriods;
  const scheduleDataRange = `$C$32:$O$${31 + calculations.totalPeriods}`; // Range of schedule data

  // Section header
  sheet.getRange(startRow, 1, 1, 13).merge()
    .setValue('JOURNAL ENTRIES')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground('#1c4587')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  let currentRow = startRow + 2;

  // 1. INITIAL RECOGNITION - Use formulas
  const initialJournalLabels = [
    ['1. INITIAL RECOGNITION OF LEASE (At Commencement)', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['Date:', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['Account', 'Description', 'Debit', 'Credit', 'Reference', '', '', '', '', '', '', '', ''],
    ['Right-of-Use Asset', 'Initial recognition of leased asset', '', '', 'Ind AS 116.23', '', '', '', '', '', '', '', ''],
    ['  Lease Liability', 'Present value of lease payments', '', '', 'Ind AS 116.26', '', '', '', '', '', '', '', ''],
    ['  Cash/Payables', 'Initial direct costs (if any)', '', '', 'Ind AS 116.24(a)', '', '', '', '', '', '', '', ''],
    ['  Prepaid Rent', 'Prepaid lease payments (if any)', '', '', 'Ind AS 116.24(b)', '', '', '', '', '', '', '', ''],
    ['', 'To recognize lease on balance sheet', '', '', '', '', '', '', '', '', '', '', '']
  ];

  sheet.getRange(currentRow, 1, initialJournalLabels.length, 13).setValues(initialJournalLabels);

  // Add formulas for initial journal entry
  sheet.getRange(currentRow + 1, 2).setFormula('=TEXT($C$7,"dd-mmm-yyyy")'); // Date
  sheet.getRange(currentRow + 4, 3).setFormula('=$F$24'); // ROU Asset Debit
  sheet.getRange(currentRow + 5, 4).setFormula('=$F$23'); // Lease Liability Credit
  sheet.getRange(currentRow + 6, 4).setFormula('=IF($C$14>0,$C$14,"")'); // Initial Direct Costs
  sheet.getRange(currentRow + 7, 4).setFormula('=IF($C$15>0,$C$15,"")'); // Prepaid Payments

  currentRow += initialJournalLabels.length + 2;

  // 2. SUBSEQUENT MEASUREMENT - Create year summaries using SUMIF
  // First, get unique years from the schedule
  if (!amortizationData || amortizationData.length === 0) {
    Logger.log('No amortization data available for journal entries');
    return; // Exit early if no data
  }

  const uniqueYears = [...new Set(amortizationData.map(row => row[2]))].sort();
  let yearNumber = 1;

  for (const year of uniqueYears) {
    const yearJournalLabels = [
      [`${yearNumber + 1}. YEAR ${yearNumber} ENTRIES (Fiscal Year ${year})`, '', '', '', '', '', '', '', '', '', '', '', ''],
      ['', '', '', '', '', '', '', '', '', '', '', '', ''],
      ['Account', 'Description', 'Debit', 'Credit', 'Reference', '', '', '', '', '', '', '', ''],
      ['Interest Expense', 'Interest on lease liability', '', '', 'Ind AS 116.29', '', '', '', '', '', '', '', ''],
      ['Depreciation Expense', 'Depreciation of ROU asset', '', '', 'Ind AS 116.31', '', '', '', '', '', '', '', ''],
      ['Lease Liability', 'Principal portion of payments', '', '', '', '', '', '', '', '', '', '', ''],
      ['  Cash', 'Cash payments made during the year', '', '', '', '', '', '', '', '', '', '', ''],
      ['', 'To record payment(s) and related expenses', '', '', '', '', '', '', '', '', '', '', ''],
      ['', '', '', '', '', '', '', '', '', '', '', '', ''],
      ['', 'Year-End Balance Sheet Position:', '', '', '', '', '', '', '', '', '', '', ''],
      ['', '  Lease Liability (Closing)', '', '', '', '', '', '', '', '', '', '', ''],
      ['', '  ROU Asset (Closing)', '', '', '', '', '', '', '', '', '', '', '']
    ];

    sheet.getRange(currentRow, 1, yearJournalLabels.length, 13).setValues(yearJournalLabels);

    // Add SUMIF formulas for yearly totals
    const scheduleStart = 32;
    const scheduleEnd = 31 + calculations.totalPeriods;

    // Interest Expense (sum of column F where fiscal year = this year)
    sheet.getRange(currentRow + 3, 3).setFormula(`=SUMIF($C$${scheduleStart}:$C$${scheduleEnd},${year},$F$${scheduleStart}:$F$${scheduleEnd})`);

    // Depreciation (sum of column N where fiscal year = this year)
    sheet.getRange(currentRow + 4, 3).setFormula(`=SUMIF($L$${scheduleStart}:$L$${scheduleEnd},${year},$N$${scheduleStart}:$N$${scheduleEnd})`);

    // Principal Reduction (sum of column G where fiscal year = this year)
    sheet.getRange(currentRow + 5, 3).setFormula(`=SUMIF($C$${scheduleStart}:$C$${scheduleEnd},${year},$G$${scheduleStart}:$G$${scheduleEnd})`);

    // Cash payments (sum of column E where fiscal year = this year)
    sheet.getRange(currentRow + 6, 4).setFormula(`=SUMIF($C$${scheduleStart}:$C$${scheduleEnd},${year},$E$${scheduleStart}:$E$${scheduleEnd})`);

    // Closing Lease Liability (last closing balance for this year)
    sheet.getRange(currentRow + 10, 3).setFormula(`=IFERROR(INDEX($H$${scheduleStart}:$H$${scheduleEnd},MAX(IF($C$${scheduleStart}:$C$${scheduleEnd}=${year},ROW($C$${scheduleStart}:$C$${scheduleEnd})-${scheduleStart - 1}))),0)`);

    // Closing ROU Asset (last closing balance for this year)
    sheet.getRange(currentRow + 11, 3).setFormula(`=IFERROR(INDEX($O$${scheduleStart}:$O$${scheduleEnd},MAX(IF($L$${scheduleStart}:$L$${scheduleEnd}=${year},ROW($L$${scheduleStart}:$L$${scheduleEnd})-${scheduleStart - 1}))),0)`);

    currentRow += yearJournalLabels.length + 2;
    yearNumber++;
  }
}

/**
 * Groups amortization and ROU data by fiscal year for journal entries.
 */
function groupByFiscalYear(amortizationData, rouAssetData) {
  const yearlyData = {};
  
  amortizationData.forEach((row, index) => {
    const year = row[2]; // Fiscal year column
    
    if (!yearlyData[year]) {
      yearlyData[year] = {
        totalInterest: 0,
        totalPrincipal: 0,
        totalPayments: 0,
        totalDepreciation: 0,
        periodCount: 0,
        closingLiability: 0,
        closingROUAsset: 0
      };
    }
    
    yearlyData[year].totalInterest += row[5]; // Interest Expense
    yearlyData[year].totalPrincipal += row[6]; // Principal Repayment
    yearlyData[year].totalPayments += row[4]; // Payment
    yearlyData[year].totalDepreciation += rouAssetData[index][4]; // Depreciation
    yearlyData[year].periodCount++;
    yearlyData[year].closingLiability = row[7]; // Closing Balance
    yearlyData[year].closingROUAsset = rouAssetData[index][5]; // Closing ROU Asset
  });
  
  return yearlyData;
}

/**
 * Adds a notes and assumptions section for audit trail.
 */
function addNotesSection(sheet, inputs, calculations) {
  const lastRow = sheet.getLastRow();
  const notesStartRow = lastRow + 3;
  
  const notesData = [
    ['NOTES AND ASSUMPTIONS', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['Accounting Standard:', 'Ind AS 116 Leases (Indian Accounting Standards)', '', '', '', '', '', '', '', '', '', '', ''],
    ['Applicability:', 'Indian companies and entities following Ind AS', '', '', '', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['Discount Rate Basis:', 'Incremental Borrowing Rate (IBR)', '', '', '', '', '', '', '', '', '', '', ''],
    ['', 'The rate of interest that a lessee would have to pay to borrow over a similar term,', '', '', '', '', '', '', '', '', '', '', ''],
    ['', 'and with similar security, the funds necessary to obtain an asset of similar value', '', '', '', '', '', '', '', '', '', '', ''],
    ['', 'to the right-of-use asset in a similar economic environment.', '', '', '', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['Depreciation Method:', 'Straight-line over lease term', '', '', '', '', '', '', '', '', '', '', ''],
    ['', 'The ROU asset is depreciated evenly over the lease term.', '', '', '', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['Variable Payments:', 'Not included in measurement', '', '', '', '', '', '', '', '', '', '', ''],
    ['', 'Only fixed payments are included in the initial measurement.', '', '', '', '', '', '', '', '', '', '', ''],
    ['', 'Variable lease payments are expensed as incurred.', '', '', '', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['Currency:', 'Indian Rupees (₹)', '', '', '', '', '', '', '', '', '', '', ''],
    ['Calculation Date:', new Date(), '', '', '', '', '', '', '', '', '', '', ''],
    ['Calculation Method:', 'Present Value using Annuity Formula: PV = PMT × [(1 - (1 + r)^-n) / r]', '', '', '', '', '', '', '', '', '', '', ''],
  ];
  
  sheet.getRange(notesStartRow, 1, notesData.length, 13).setValues(notesData);
}

/**
 * Applies comprehensive formatting to make the workpaper professional and audit-ready.
 */
function applyFormatting(sheet, scheduleRows) {
  // Set column widths for optimal readability
  sheet.setColumnWidths(1, 1, 80);   // Period #
  sheet.setColumnWidths(2, 1, 120);  // Dates
  sheet.setColumnWidths(3, 1, 90);   // Fiscal Year / Labels
  sheet.setColumnWidths(4, 5, 130);  // Monetary columns
  sheet.setColumnWidths(10, 1, 80);  // Period # (ROU)
  sheet.setColumnWidths(11, 1, 120); // Dates (ROU)
  sheet.setColumnWidths(12, 4, 130); // ROU monetary columns
  
  // === HEADER SECTION ===
  sheet.getRange("A1:I2").merge();
  sheet.getRange("A1").setFontSize(18).setFontWeight('bold')
    .setFontColor('#1c4587').setHorizontalAlignment('left');
  
  sheet.getRange("K1:K3").setFontWeight('bold').setFontSize(9);
  sheet.getRange("M2").setNumberFormat('dd-mmm-yyyy hh:mm');
  
  // === INPUT SECTION ===
  sheet.getRange("A6:M6").merge();
  sheet.getRange("A6").setBackground('#1c4587').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(12);
  
  // Input labels
  sheet.getRange("A7:A15").setFontWeight('bold').setBackground('#e8eaf6');
  
  // Input fields (yellow highlight)
  const inputRanges = ["C7", "C8", "C9", "C10", "C11", "C12", "C14", "C15"];
  inputRanges.forEach(range => {
    sheet.getRange(range).setBackground('#fff9c4').setBorder(true, true, true, true, null, null, '#f57c00', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });
  
  // Help text
  sheet.getRange("E7:E15").setFontStyle('italic').setFontSize(9).setFontColor('#666666');
  
  // Number formatting for inputs (Indian Rupees)
  sheet.getRange("C7").setNumberFormat('dd-mmm-yyyy');
  sheet.getRange("C8").setNumberFormat('0');  // Lease term - just a number
  sheet.getRange("C9").setNumberFormat('₹#,##,##0.00');  // Indian number format
  sheet.getRange("C10").setNumberFormat('@');  // Payment frequency - text
  sheet.getRange("C11").setNumberFormat('0.00%');
  sheet.getRange("C12").setNumberFormat('dd-mmm-yyyy');
  sheet.getRange("C14:C15").setNumberFormat('₹#,##,##0.00');  // Indian number format
  
  // Input section border
  sheet.getRange("A6:M16").setBorder(true, true, true, true, null, null, '#1c4587', SpreadsheetApp.BorderStyle.SOLID_THICK);
  
  // === DASHBOARD SECTION ===
  sheet.getRange("A19").setBackground('#1c4587').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(12);
  sheet.getRange("A19:M19").merge();
  
  // Dashboard subsection headers
  sheet.getRange("A21").setValue('Key Metrics').setFontWeight('bold').setBackground('#6d9eeb').setFontColor('#ffffff');
  sheet.getRange("A21:C21").merge();
  sheet.getRange("D21").setValue('Initial Recognition').setFontWeight('bold').setBackground('#6d9eeb').setFontColor('#ffffff');
  sheet.getRange("D21:F21").merge();
  sheet.getRange("G21").setValue('Lease Economics').setFontWeight('bold').setBackground('#6d9eeb').setFontColor('#ffffff');
  sheet.getRange("G21:I21").merge();
  
  // Dashboard labels and values
  sheet.getRange("A23:A26").setFontWeight('bold').setBackground('#e8eaf6');
  sheet.getRange("D23:D26").setFontWeight('bold').setBackground('#e8eaf6');
  sheet.getRange("G23:G26").setFontWeight('bold').setBackground('#e8eaf6');
  
  // Dashboard value formatting (Indian Rupees)
  const dashboardMoneyRanges = ["C23", "C25", "F23:F26", "I23"];
  dashboardMoneyRanges.forEach(range => {
    sheet.getRange(range).setNumberFormat('₹#,##,##0.00').setFontWeight('bold').setFontColor('#1c4587');
  });

  sheet.getRange("C24").setNumberFormat('#,##,##0');  // Indian number format
  sheet.getRange("I24:I25").setNumberFormat('0.00%').setFontWeight('bold').setFontColor('#1c4587');
  
  // Dashboard border
  sheet.getRange("A19:M26").setBorder(true, true, true, true, true, true, '#1c4587', SpreadsheetApp.BorderStyle.SOLID_THICK);
  
  // === SCHEDULE SECTION ===
  const scheduleStartRow = 30;
  
  // Amortization schedule header
  sheet.getRange(scheduleStartRow, 1, 1, 8)
    .setBackground('#3c78d8').setFontColor('#ffffff').setFontWeight('bold')
    .setBorder(true, true, true, true, null, null);
  
  // Amortization schedule column headers
  sheet.getRange(scheduleStartRow + 1, 1, 1, 8)
    .setBackground('#9fc5e8').setFontWeight('bold').setHorizontalAlignment('center')
    .setWrap(true).setBorder(true, true, true, true, null, null);
  
  // Amortization data formatting
  if (scheduleRows > 0) {
    const dataStartRow = scheduleStartRow + 2;
    
    // Dates
    sheet.getRange(dataStartRow, 2, scheduleRows, 1).setNumberFormat('dd-mmm-yyyy');

    // Monetary columns (Indian Rupees)
    sheet.getRange(dataStartRow, 4, scheduleRows, 5).setNumberFormat('₹#,##,##0.00');
    
    // Alternating row colors for readability
    for (let i = 0; i < scheduleRows; i++) {
      const rowColor = (i % 2 === 0) ? '#f3f3f3' : '#ffffff';
      sheet.getRange(dataStartRow + i, 1, 1, 8).setBackground(rowColor);
    }
    
    // Add borders to data
    sheet.getRange(dataStartRow, 1, scheduleRows, 8)
      .setBorder(null, null, null, null, true, null, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
    
    // Highlight closing balance
    sheet.getRange(dataStartRow, 8, scheduleRows, 1).setFontWeight('bold');
  }
  
  // ROU Asset schedule header
  sheet.getRange(scheduleStartRow, 10, 1, 6)
    .setBackground('#3c78d8').setFontColor('#ffffff').setFontWeight('bold')
    .setBorder(true, true, true, true, null, null);
  
  // ROU schedule column headers
  sheet.getRange(scheduleStartRow + 1, 10, 1, 6)
    .setBackground('#9fc5e8').setFontWeight('bold').setHorizontalAlignment('center')
    .setWrap(true).setBorder(true, true, true, true, null, null);
  
  // ROU data formatting
  if (scheduleRows > 0) {
    const dataStartRow = scheduleStartRow + 2;
    
    // Dates
    sheet.getRange(dataStartRow, 11, scheduleRows, 1).setNumberFormat('dd-mmm-yyyy');

    // Monetary columns (Indian Rupees)
    sheet.getRange(dataStartRow, 13, scheduleRows, 3).setNumberFormat('₹#,##,##0.00');
    
    // Alternating row colors
    for (let i = 0; i < scheduleRows; i++) {
      const rowColor = (i % 2 === 0) ? '#f3f3f3' : '#ffffff';
      sheet.getRange(dataStartRow + i, 10, 1, 6).setBackground(rowColor);
    }
    
    // Add borders
    sheet.getRange(dataStartRow, 10, scheduleRows, 6)
      .setBorder(null, null, null, null, true, null, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
    
    // Highlight closing balance
    sheet.getRange(dataStartRow, 15, scheduleRows, 1).setFontWeight('bold');
  }
  
  // === JOURNAL ENTRIES SECTION ===
  formatJournalEntries(sheet);
  
  // === NOTES SECTION ===
  formatNotesSection(sheet);
  
  // === GENERAL FORMATTING ===
  sheet.setFrozenRows(2);
  
  // Apply conditional formatting for negative values (if any)
  applyConditionalFormatting(sheet);
}

/**
 * Formats the journal entries section.
 */
function formatJournalEntries(sheet) {
  const journalStartRow = findRowByText(sheet, 'JOURNAL ENTRIES');
  
  if (journalStartRow > 0) {
    // Main section header
    sheet.getRange(journalStartRow, 1, 1, 13)
      .setBackground('#1c4587').setFontColor('#ffffff');
    
    // Find and format each journal entry
    let currentRow = journalStartRow + 2;
    const lastRow = sheet.getLastRow();
    
    while (currentRow < lastRow) {
      const cellValue = sheet.getRange(currentRow, 1).getValue();
      
      if (cellValue && cellValue.toString().match(/^\d+\./)) {
        // Journal entry title
        sheet.getRange(currentRow, 1, 1, 13).merge()
          .setBackground('#6d9eeb').setFontColor('#ffffff').setFontWeight('bold');
        
        // Column headers (3 rows down from title)
        const headerRow = currentRow + 3;
        sheet.getRange(headerRow, 1, 1, 13)
          .setBackground('#cfe2f3').setFontWeight('bold')
          .setBorder(true, true, true, true, null, null);
        
        // Format monetary columns in this journal entry
        let entryRow = headerRow + 1;
        while (entryRow < lastRow && sheet.getRange(entryRow, 1).getValue()) {
          // Apply number format to debit and credit columns (Indian Rupees)
          sheet.getRange(entryRow, 3, 1, 2).setNumberFormat('₹#,##,##0.00');
          
          // Indent sub-accounts
          const account = sheet.getRange(entryRow, 1).getValue();
          if (account && account.toString().trim().startsWith('')) {
            sheet.getRange(entryRow, 1).setFontStyle('italic');
          }
          
          // Format reference column
          sheet.getRange(entryRow, 5).setFontSize(8).setFontColor('#666666');
          
          entryRow++;
          if (!sheet.getRange(entryRow, 1).getValue() || 
              sheet.getRange(entryRow, 1).getValue().toString().match(/^\d+\./)) {
            break;
          }
        }
        
        // Add border around this journal entry
        const entryHeight = entryRow - currentRow;
        sheet.getRange(currentRow, 1, entryHeight, 5)
          .setBorder(true, true, true, true, true, null, '#3c78d8', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        
        currentRow = entryRow + 1;
      } else {
        currentRow++;
      }
    }
  }
}

/**
 * Formats the notes section at the bottom.
 */
function formatNotesSection(sheet) {
  const notesStartRow = findRowByText(sheet, 'NOTES AND ASSUMPTIONS');
  
  if (notesStartRow > 0) {
    // Section header
    sheet.getRange(notesStartRow, 1, 1, 5).merge()
      .setBackground('#1c4587').setFontColor('#ffffff')
      .setFontWeight('bold').setFontSize(12);
    
    // Format labels
    const lastRow = sheet.getLastRow();
    for (let row = notesStartRow + 2; row <= lastRow; row++) {
      const value = sheet.getRange(row, 1).getValue();
      if (value && value.toString().endsWith(':')) {
        sheet.getRange(row, 1).setFontWeight('bold').setBackground('#e8eaf6');
      }
    }
    
    // Format date
    sheet.getRange(notesStartRow + 17, 2).setNumberFormat('dd-mmm-yyyy hh:mm');
    
    // Add border
    sheet.getRange(notesStartRow, 1, lastRow - notesStartRow + 1, 5)
      .setBorder(true, true, true, true, null, null, '#1c4587', SpreadsheetApp.BorderStyle.SOLID_THICK);
  }
}

/**
 * Applies conditional formatting rules.
 */
function applyConditionalFormatting(sheet) {
  // Highlight any negative balances in red (for error detection)
  const lastRow = sheet.getLastRow();
  
  // For closing liability column
  const rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground('#f4cccc')
    .setFontColor('#cc0000')
    .setRanges([sheet.getRange('H32:H' + lastRow)])
    .build();
  
  // For closing ROU asset column
  const rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground('#f4cccc')
    .setFontColor('#cc0000')
    .setRanges([sheet.getRange('O32:O' + lastRow)])
    .build();
  
  const rules = sheet.getConditionalFormatRules();
  rules.push(rule1);
  rules.push(rule2);
  sheet.setConditionalFormatRules(rules);
}

/**
 * Protects formula cells while allowing input cells to be edited.
 */
function protectFormulaCells(sheet) {
  // Protect the entire sheet
  const protection = sheet.protect().setDescription('Protected Workpaper');
  
  // Allow editing only for the input cells
  const unprotectedRanges = [
    sheet.getRange('C7'),  // Lease Start Date
    sheet.getRange('C8'),  // Lease Term
    sheet.getRange('C9'),  // Payment Amount
    sheet.getRange('C10'), // Payment Frequency
    sheet.getRange('C11'), // Discount Rate
    sheet.getRange('C12'), // Reporting End Date
    sheet.getRange('C14'), // Initial Direct Costs
    sheet.getRange('C15')  // Prepaid Payments
  ];
  
  protection.setUnprotectedRanges(unprotectedRanges);
  
  // Allow all users to edit unprotected ranges
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}

/**
 * Utility function to find a row containing specific text.
 */
function findRowByText(sheet, searchText) {
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().includes(searchText)) {
      return i + 1; // Return 1-indexed row number
    }
  }
  return -1;
}

/**
 * Standalone function to refresh formatting without recalculating.
 */
function applyFormattingOnly() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const scheduleRows = sheet.getLastRow() - 32; // Approximate schedule rows
  applyFormatting(sheet, scheduleRows);
  SpreadsheetApp.getUi().alert('✅ Formatting refreshed!');
}