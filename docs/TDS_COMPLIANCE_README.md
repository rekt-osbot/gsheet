# TDS Compliance Tracker for Google Sheets

> Comprehensive Tax Deducted at Source (TDS) compliance management system for Indian businesses

## 📋 Table of Contents
- [What is This?](#what-is-this)
- [Who Should Use This?](#who-should-use-this)
- [Key Features](#key-features)
- [Quick Start Guide](#quick-start-guide)
- [Sample Data Demo](#sample-data-demo)
- [Sheet-by-Sheet Guide](#sheet-by-sheet-guide)
- [TDS Sections Covered](#tds-sections-covered)
- [Compliance Checklist](#compliance-checklist)
- [Troubleshooting](#troubleshooting)

---

## What is This?

A complete TDS compliance tracking system built for Google Sheets that automates:
- TDS rate calculations across 30+ sections
- Vendor master with PAN validation
- Monthly liability tracking
- Form 26AS reconciliation
- Interest calculation on late payments
- Quarterly return preparation (24Q/26Q)
- Lower deduction certificate management

**Think of it as:** Your complete TDS compliance department in a spreadsheet.

---

## Who Should Use This?

- **Chartered Accountants** - Managing multiple client TDS compliance
- **Finance Teams** - In-house TDS management for companies
- **Tax Consultants** - TDS advisory and compliance services
- **Audit Firms** - TDS audit and verification
- **Small Businesses** - Self-managed TDS compliance
- **Startups** - Cost-effective TDS tracking solution

---

## Key Features

### ✅ Automated Calculations
- **Smart Rate Lookup** - Automatically applies correct TDS rate based on section + entity type
- **Threshold Management** - Only deducts TDS when payment exceeds threshold limits
- **Interest Calculator** - Auto-calculates Section 201 interest for late payments
- **Quarter Assignment** - Automatically tags transactions to Q1/Q2/Q3/Q4

### 🔍 Compliance Tracking
- **PAN Validation** - Regex-based validation for valid PAN format
- **Payment Status** - Color-coded tracking (Paid/Pending/Late)
- **26AS Reconciliation** - Variance analysis between books and Form 26AS
- **Certificate Tracking** - Monitors Section 197 certificate validity with expiry alerts

### 📊 Reporting & Analytics
- **Real-time Dashboard** - Key metrics and compliance status at a glance
- **Section-wise Summary** - TDS breakdown by section
- **Vendor-wise Analysis** - Track TDS per vendor
- **Monthly Ledger** - Month-wise liability and payment tracking

### 🎯 User-Friendly Design
- **Input Protection** - Only editable cells are highlighted (light orange)
- **Auto-population** - Vendor details auto-fill from master
- **Data Validation** - Dropdowns for entity types, sections, quarters
- **Professional Formatting** - Print-ready working papers

---

## Quick Start Guide

### Step 1: Create the Workbook

1. Open a new [Google Sheet](https://sheets.google.com)
2. Go to **Extensions > Apps Script**
3. Delete any existing code
4. Copy the entire `tds_compliance.gs` file content
5. Paste into the Apps Script editor
6. Click **Save** (💾 icon)
7. Name your project: "TDS Compliance Tracker"

### Step 2: Run the Script

1. In the function dropdown, select `createTDSComplianceWorkbook`
2. Click **▶ Run**
3. **First time only:** Authorize the script
   - Click "Review Permissions"
   - Choose your Google account
   - Click "Advanced" → "Go to [Project Name] (unsafe)"
   - Click "Allow"
4. Wait 10-15 seconds for workbook creation
5. You'll see 12 sheets created automatically

### Step 3: Load Sample Data (Recommended)

1. Go to **Extensions > Apps Script** again
2. In the function dropdown, select `populateSampleData`
3. Click **▶ Run**
4. Sample data will populate all sheets
5. Review the Dashboard to see the system in action

### Step 4: Customize for Your Entity

1. Go to **Assumptions** sheet
2. Update entity details (Name, PAN, TAN, etc.)
3. Go to **Vendor_Master** sheet
4. Replace sample vendors with your actual vendors
5. Go to **TDS_Register** sheet
6. Start entering your actual transactions

---

## Sample Data Demo

The `populateSampleData()` function creates a realistic FY 2024-25 scenario:

### Sample Entity
- **Company:** ABC Manufacturing Pvt Ltd
- **PAN:** AAAPL1234C
- **TAN:** MUMM12345D
- **Location:** Mumbai, Maharashtra

### Sample Vendors (10 vendors)
- XYZ Contractors (194C - Construction)
- Ramesh Consultants (194J - Technical, with lower deduction cert)
- Global Tech Solutions (194J - Software)
- Priya Advertising (194H - Commission)
- Sharma & Associates (194J - Legal)
- Metro Property Rentals (194I - Rent)
- Suresh Transport (194C - Logistics)
- ICICI Bank (194A - Interest)
- Anita Design Studio (194J - Design)
- BuildRight Engineers (194C - Construction)

### Sample Transactions (30 transactions)
- Spread across all 4 quarters
- Multiple TDS sections (194A, 194C, 194H, 194I, 194J)
- Includes late payments to demonstrate interest calculation
- Total TDS: ₹1,05,720
- Demonstrates threshold application
- Shows lower deduction certificate usage

### Key Scenarios Demonstrated
1. **Regular payments** - Most transactions paid on time
2. **Late payments** - Oct and Feb payments delayed (interest calculated)
3. **Lower deduction cert** - Ramesh Consultants has 2% rate instead of 10%
4. **Threshold application** - Some payments below threshold (no TDS)
5. **26AS variance** - Sample shows reconciliation differences
6. **Pending payment** - March TDS not yet deposited

---

## Sheet-by-Sheet Guide

### 1. Dashboard
**Purpose:** Real-time compliance overview

**Key Metrics:**
- Total Vendors: 10
- Total Transactions: 30
- Total TDS Deducted: ₹1,05,720
- TDS Payable: Shows outstanding balance
- Interest on Late Payment: Auto-calculated
- Active Lower Deduction Certs: 1

**Status Table:** Month-wise compliance status with color coding
- 🟢 Green = Paid on time
- 🟡 Yellow = Pending
- 🔴 Red = Late payment

### 2. Assumptions
**Purpose:** Entity configuration and settings

**What to Fill:**
- Entity name, PAN, TAN
- Address details
- Contact information
- Financial year
- Bank details for TDS payment

### 3. Vendor_Master
**Purpose:** Complete vendor/deductee database

**Columns:**
- Vendor Code (unique ID)
- Vendor Name
- PAN (auto-validates format)
- Entity Type (Company/Individual/HUF/Firm/etc.)
- Contact details
- Lower Deduction Certificate flag

**PAN Validation:** Automatically checks if PAN matches format: AAAAA9999A

### 4. TDS_Register
**Purpose:** Transaction-wise TDS deduction log

**Input Columns (orange background):**
- Date
- Vendor Code
- TDS Section
- Nature of Payment
- Gross Amount

**Auto-Calculated Columns:**
- Vendor Name (lookup from master)
- PAN (lookup from master)
- Entity Type (lookup from master)
- Threshold Limit (lookup from rates)
- Applicable Rate % (lookup based on section + entity type)
- TDS Amount (calculated if above threshold)
- Net Payment (Gross - TDS)
- Quarter (auto-assigned based on date)

**Formula Logic:**
```
TDS Amount = IF(Gross Amount > Threshold, ROUND(Gross Amount × Rate%, 0), 0)
```

### 5. Section_Rates
**Purpose:** TDS rate master for all sections

**Sections Covered:** 192, 192A, 193, 194, 194A, 194B, 194C, 194D, 194DA, 194EE, 194F, 194G, 194H, 194I, 194IA, 194IB, 194IC, 194J, 194K, 194LA, 194LB, 194LBA, 194LBB, 194LBC, 194M, 194N, 194O, 194Q, 195

**Rate Variations:** Different rates for Company, Individual, HUF, Firm, Non-Resident

**Lookup Key:** Combines Section + Entity Type (e.g., "194CCompany", "194JIndividual")

### 6. Lower_Deduction_Cert
**Purpose:** Track Section 197 certificates

**Features:**
- Certificate number tracking
- Validity period monitoring
- Reduced rate specification
- Maximum amount limit
- Auto-status: Active/Expired/Not Yet Valid

**Color Coding:**
- 🟢 Green = Active
- 🔴 Red = Expired

### 7. TDS_Payable_Ledger
**Purpose:** Month-wise TDS liability tracking

**Auto-Calculated:**
- TDS Deducted (sum from TDS_Register)
- Due Date (7th of next month)
- Balance (TDS Deducted - Amount Paid)
- Status (Paid/Pending/Late Payment)

**Input Fields:**
- Challan Number
- Payment Date
- Amount Paid

### 8. 26AS_Reconciliation
**Purpose:** Reconcile books with Form 26AS

**Three Sections:**
1. **As Per Books** - Auto-calculated from TDS_Register
2. **As Per Form 26AS** - Manual input from TRACES portal
3. **Variance Analysis** - Auto-calculated differences

**Status:**
- 🟢 Matched = No variance
- 🔴 Variance = Differences found (investigate)

### 9. Quarterly_Return
**Purpose:** Prepare data for 24Q/26Q filing

**Features:**
- Quarter selection dropdown
- Auto-summary of selected quarter
- Deductee-wise details
- Challan mapping
- BSR code and serial number fields

### 10. Interest_Calculator
**Purpose:** Calculate Section 201 interest

**Two Sections:**
1. **Late Payment Analysis** - Auto-calculates interest for all months
2. **Manual Calculator** - Calculate interest for specific scenarios

**Formula:**
```
Interest = TDS Amount × 1% × (Delay Days / 30)
```

**Example:** ₹10,000 TDS paid 15 days late = ₹10,000 × 1% × (15/30) = ₹50

### 11. Audit_Notes
**Purpose:** Documentation and reference

**Contents:**
- Compliance requirements
- Important sections explained
- Interest & penalty provisions
- Quarterly return due dates
- Useful links (TRACES, Income Tax portal)
- Best practices
- Audit checklist

---

## TDS Sections Covered

### Salary & Employment
- **192** - Salary (as per slab rates)
- **192A** - Premature EPF withdrawal (10%)

### Interest & Dividends
- **193** - Interest on securities (10%)
- **194** - Dividend (10%)
- **194A** - Interest other than securities (10%)
- **194K** - Income from mutual fund units (10%)

### Contractors & Services
- **194C** - Payments to contractors (1% Company, 2% Individual/HUF)
- **194H** - Commission/Brokerage (5%)
- **194J** - Professional/Technical fees (10%)

### Rent
- **194I** - Rent on Plant & Machinery (2%)
- **194I** - Rent on Land/Building (10%)
- **194IB** - Rent by Individual/HUF (5%)

### Property
- **194IA** - Transfer of immovable property (1%)
- **194IC** - Joint Development Agreement (10%)
- **194LA** - Compensation on land acquisition (10%)

### Winnings & Insurance
- **194B** - Lottery/Game show winnings (30%)
- **194D** - Insurance commission (5%)
- **194DA** - Life insurance maturity (5%)
- **194G** - Commission on lottery tickets (5%)

### Investments & Trusts
- **194EE** - NSS deposits (10%)
- **194F** - Mutual fund repurchase (20%)
- **194LB** - Infrastructure debt fund interest (5%)
- **194LBA** - Business trust income (10%)
- **194LBB** - Investment fund income (10%)
- **194LBC** - Securitization trust income (25-30%)

### Recent Additions
- **194M** - Payment by individuals/HUF (5%)
- **194N** - Cash withdrawal (2%)
- **194O** - E-commerce participants (1%)
- **194Q** - Purchase of goods (0.1%)

### Non-Residents
- **195** - Non-resident payments (as per DTAA/Act)

---

## Compliance Checklist

### Monthly Tasks
- [ ] Enter all payment transactions in TDS_Register
- [ ] Verify TDS calculations are correct
- [ ] Check for lower deduction certificates
- [ ] Deposit TDS by 7th of next month
- [ ] Update challan details in TDS_Payable_Ledger
- [ ] Review Dashboard for pending items

### Quarterly Tasks
- [ ] Download Form 26AS from TRACES
- [ ] Perform 26AS reconciliation
- [ ] Investigate and resolve variances
- [ ] Prepare quarterly return (24Q/26Q)
- [ ] File return before due date
- [ ] Issue TDS certificates (Form 16/16A)

### Annual Tasks
- [ ] Review all vendor PANs for validity
- [ ] Update TDS rates for new financial year
- [ ] Renew lower deduction certificates
- [ ] Archive previous year data
- [ ] Conduct internal TDS audit
- [ ] Prepare for external audit

---

## Troubleshooting

### Common Issues

**Q: PAN validation shows ✗ for valid PAN**
- Ensure PAN is in uppercase
- Check for extra spaces
- Format must be: AAAAA9999A (5 letters, 4 digits, 1 letter)

**Q: TDS not calculating despite amount above threshold**
- Check if TDS Section is entered correctly
- Verify Entity Type is selected in Vendor_Master
- Ensure Section_Rates sheet has matching lookup key

**Q: Vendor details not auto-filling in TDS_Register**
- Verify Vendor Code matches exactly with Vendor_Master
- Check for leading/trailing spaces
- Ensure Vendor_Master has data in correct columns

**Q: Interest calculation showing zero**
- Verify Payment Date is after Due Date
- Check if TDS Amount is populated
- Ensure dates are in proper date format (not text)

**Q: Dashboard metrics not updating**
- Click "TDS Compliance" menu → "Refresh Dashboard"
- Or press Ctrl+R (Cmd+R on Mac) to recalculate

**Q: 26AS reconciliation showing variance**
- Check if all transactions are entered
- Verify challan numbers match
- Ensure dates are in correct quarter
- Check for TDS deducted by others (not in your books)

### Performance Tips

**For Large Datasets (500+ transactions):**
1. Use filters instead of scrolling
2. Archive old financial year data to separate sheet
3. Consider splitting by quarter
4. Use named ranges for faster lookups

**Formula Optimization:**
- Avoid volatile functions (NOW, TODAY) in large ranges
- Use IFERROR sparingly
- Consider replacing VLOOKUP with INDEX-MATCH for speed

---

## Best Practices

### Data Entry
1. **Enter transactions promptly** - Don't wait until month-end
2. **Use consistent vendor codes** - Maintain proper vendor master
3. **Verify PANs immediately** - Incorrect PAN = 20% TDS rate
4. **Document special cases** - Use Remarks column liberally
5. **Backup regularly** - Download Excel copy monthly

### Compliance
1. **Pay before due date** - Avoid interest and penalties
2. **Reconcile quarterly** - Don't wait for year-end
3. **Track certificates** - Monitor lower deduction cert expiry
4. **Maintain documentation** - Keep invoices, contracts, certificates
5. **Review rates annually** - Update Section_Rates for budget changes

### Audit Readiness
1. **Complete all fields** - Don't leave blanks in critical columns
2. **Cross-reference** - Ensure challan numbers match bank statements
3. **Explain variances** - Document reasons for 26AS differences
4. **Organize certificates** - Keep TDS certificates in order
5. **Prepare summaries** - Use Dashboard for audit discussions

---

## Advanced Features

### Custom Menu
After running the script, you'll see a "TDS Compliance" menu with:
- Create New Workbook
- Populate Sample Data
- Refresh Dashboard
- Export for Return Filing

### Named Ranges
The workbook uses named ranges for easy formula reference:
- `VendorCodes` - All vendor codes
- `VendorNames` - All vendor names
- `VendorPANs` - All vendor PANs
- `TDSTransactions` - Complete TDS register
- `SectionRates` - Rate master table

### Conditional Formatting
- Status columns use color coding
- Late payments highlighted in red
- Pending items in yellow
- Completed items in green

---

## Updates & Maintenance

### When to Update Rates
- Budget announcements (usually February)
- Finance Act amendments
- CBDT circulars and notifications

### How to Update Rates
1. Go to Section_Rates sheet
2. Update the rate columns (E-I)
3. Add new sections if introduced
4. Update Remarks column with effective date

### Version Control
- Save dated copies before major changes
- Use "File > Version History" in Google Sheets
- Document changes in Audit_Notes sheet

---

## Support & Resources

### Official Resources
- **TRACES Portal:** https://www.tdscpc.gov.in/
- **Income Tax Department:** https://www.incometax.gov.in/
- **TDS Rates:** https://www.incometax.gov.in/iec/foportal/help/individual/return-applicable-1/tds-rates

### Learning Resources
- Income Tax Act, 1961 - Chapter XVII-B (TDS provisions)
- CBDT Circulars on TDS
- Form 24Q/26Q filing guides on TRACES

---

## Disclaimer

This workbook is a tool for TDS compliance tracking. Users should:
- Verify rates as per latest Income Tax Act amendments
- Consult a tax professional for specific situations
- Validate calculations before filing returns
- Keep updated with CBDT notifications

The creator assumes no liability for errors, omissions, or compliance failures arising from use of this tool.

---

## License

Open source - Free to use, modify, and distribute.

---

## Changelog

### Version 1.0 (November 2025)
- Initial release
- 12 interconnected sheets
- 30+ TDS sections covered
- Automated calculations and validations
- Sample data for demonstration
- Comprehensive documentation

---

**Built with ❤️ for the Indian accounting community**

*Simplifying TDS compliance, one transaction at a time.*
