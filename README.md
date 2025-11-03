# IGAAP-Ind AS Audit Builder for Google Sheets

> Automated audit working papers generator for Indian Accounting Standards (Ind AS) and IGAAP compliance

## Table of Contents
- [What is This Project?](#what-is-this-project)
- [Who Should Use This?](#who-should-use-this)
- [What You'll Need](#what-youll-need)
- [Getting Started (Step-by-Step for Beginners)](#getting-started-step-by-step-for-beginners)
- [How to Use Each Script](#how-to-use-each-script)
- [Features](#features)
- [Troubleshooting](#troubleshooting)
- [Understanding the Code](#understanding-the-code)
- [Contributing](#contributing)

---

## What is This Project?

This project contains **Google Apps Script** files that automatically create professional audit working papers in Google Sheets. These working papers help accountants and auditors comply with Indian Accounting Standards.

**Think of it as:** A template generator that creates pre-formatted, formula-filled spreadsheets for accounting compliance work.

### What's Included:

1. **indas109.gs** - Financial Instruments Audit Workings (Ind AS 109)
2. **indas116.gs** - Lease Accounting Workings (Ind AS 116)
3. **deferredtax.gs** - Deferred Taxation Workings (IGAAP/Ind AS 12)

---

## Who Should Use This?

- Chartered Accountants (CAs)
- Audit professionals
- Finance teams implementing Ind AS
- Accounting students learning Ind AS compliance
- Anyone preparing period-end financial statements under Indian GAAP or Ind AS

---

## What You'll Need

### Prerequisites

1. **A Google Account** - Free Gmail account works fine
2. **Basic Spreadsheet Knowledge** - You should know how to open and navigate Google Sheets
3. **No Coding Experience Required** - The scripts are ready to run!

### Technical Requirements

- Internet connection
- Web browser (Chrome, Firefox, Safari, or Edge)
- Access to Google Drive

---

## Getting Started (Step-by-Step for Beginners)

### Step 1: Create a New Google Spreadsheet

1. Go to [Google Sheets](https://sheets.google.com)
2. Click the **+ Blank** button to create a new spreadsheet
3. Give it a meaningful name (e.g., "Ind AS 109 Workings - XYZ Company")

### Step 2: Open the Script Editor

1. In your Google Sheet, click **Extensions** in the menu bar
2. Select **Apps Script**
3. A new tab will open with the Apps Script editor

### Step 3: Add the Script Code

**For Ind AS 109 (Financial Instruments):**

1. Delete any code in the editor (it usually has a placeholder function)
2. Open the `indas109.gs` file from this repository
3. Copy ALL the code
4. Paste it into the Apps Script editor
5. Click the **Save** icon (💾) or press `Ctrl+S` (Windows) / `Cmd+S` (Mac)
6. Give your project a name (e.g., "Ind AS 109 Builder")

**For Ind AS 116 (Leases):**
- Follow the same steps but use code from `indas116.gs`

**For Deferred Tax:**
- Follow the same steps but use code from `deferredtax.gs`

### Step 4: Run the Script

1. In the Apps Script editor, find the function dropdown (near the top)
2. Select the main function:
   - For Ind AS 109: `createIndAS109WorkingPapers`
   - For Ind AS 116: `createIndAS116Workbook`
   - For Deferred Tax: `createDeferredTaxWorkbook`
3. Click the **▶ Run** button
4. **First Time Only:** You'll see an authorization screen:
   - Click **Review Permissions**
   - Choose your Google account
   - Click **Advanced** (if you see a warning)
   - Click **Go to [Your Project Name] (unsafe)**
   - Click **Allow**

### Step 5: See Your Working Papers!

1. Go back to your Google Sheet tab
2. You'll see multiple sheets created automatically
3. An alert will appear confirming successful creation
4. Start with the **Cover** or **Assumptions** sheet

---

## How to Use Each Script

### 1. Ind AS 109 - Financial Instruments (`indas109.gs`)

**Purpose:** Create audit workings for financial instruments including fair value adjustments, expected credit loss, and amortization calculations.

**Steps After Running:**

1. Navigate to **Input_Variables** sheet
2. Fill in the highlighted cells (light blue background):
   - Company name
   - Reporting period
   - Currency
   - Tax rates
3. Go to **Instruments_Register** sheet
4. Enter your financial instruments details
5. The following sheets will auto-calculate:
   - Fair Value Workings
   - ECL Impairment
   - Amortization Schedule
   - Period End Entries
   - Reconciliation

**What You Get:**
- 11 interconnected sheets
- Automatic formula calculations
- Journal entries for book closure
- Audit trail and references

---

### 2. Ind AS 116 - Lease Accounting (`indas116.gs`)

**Purpose:** Automate Right-of-Use (ROU) asset calculations, lease liability schedules, and depreciation workings.

**Steps After Running:**

1. Start at the **Assumptions** sheet
2. Enter:
   - Company details
   - Reporting date
   - Base currency
   - Default incremental borrowing rate (IBR)
3. Go to **Lease_Register** sheet
4. Enter each lease contract details
5. Review auto-generated schedules:
   - ROU Asset Schedule
   - Lease Liability Schedule
   - Payment Schedule
   - Journal Entries

**What You Get:**
- Complete lease accounting workings
- Automatic interest calculations
- Period movement analysis
- Reconciliation with IGAAP

---

### 3. Deferred Taxation (`deferredtax.gs`)

**Purpose:** Generate deferred tax asset/liability calculations with temporary differences tracking.

**Steps After Running:**

1. Open **Assumptions** sheet
2. Select framework: IGAAP or Ind AS
3. Enter tax rates and entity details
4. Go to **Temp_Differences** sheet
5. Enter temporary differences between:
   - Book values (financial statements)
   - Tax base values (tax returns)
6. Review calculations in:
   - DT_Schedule (Deferred Tax Schedule)
   - Movement_Analysis
   - P&L Reconciliation
   - Balance Sheet Reconciliation

**What You Get:**
- Dynamic deferred tax calculations
- Automatic DTA/DTL classification
- Movement schedules
- Full reconciliation with financial statements

---

## Features

### What Makes These Scripts Useful?

- **Zero Manual Formulas:** All calculations are pre-built
- **Input Protection:** Only editable cells are highlighted
- **Audit Trail:** Every calculation is traceable
- **Professional Formatting:** Print-ready working papers
- **Standards Compliant:** Based on official Ind AS requirements
- **Cross-Referenced:** Sheets link to each other automatically
- **Control Totals:** Built-in reconciliation checks

### Visual Indicators

- **Light Blue Cells** = Input cells (you fill these)
- **White/Gray Cells** = Calculated cells (auto-filled)
- **Dark Blue Headers** = Section titles
- **Green/Yellow/Red** = Status indicators (where applicable)

---

## Troubleshooting

### Common Issues for Beginners

**Problem:** "Authorization required" error
- **Solution:** Follow Step 4 authorization process carefully. This is normal for first-time use.

**Problem:** "Script function not found"
- **Solution:** Make sure you selected the correct function name from the dropdown before clicking Run.

**Problem:** Nothing happens when I click Run
- **Solution:**
  1. Check if you saved the script (💾 icon)
  2. Refresh your Google Sheet tab
  3. Try running again

**Problem:** Formulas show #REF! or #NAME? errors
- **Solution:** Don't manually delete sheets. If errors appear, re-run the main function to rebuild.

**Problem:** Can't find the function dropdown
- **Solution:** It's right above your code, next to the Run button. It might say "Select function" initially.

**Problem:** Script runs but no sheets are created
- **Solution:** Check the Execution log (View > Logs in Apps Script). Look for error messages.

### Getting Help

If you encounter issues:
1. Check the **Audit_Notes** or **References** sheet in your workbook
2. Review the Execution log in Apps Script editor
3. Make sure you're using a desktop browser (mobile not recommended)

---

## Understanding the Code

### For Those Who Want to Learn

Even if you don't code, here's what's happening:

1. **Main Function:** Each script has one main function (e.g., `createIndAS109WorkingPapers`)
2. **Sheet Creation:** Functions like `createCoverSheet()` build individual sheets
3. **Formatting Functions:** Functions like `formatHeader()` apply colors and styles
4. **Formula Functions:** Functions insert Excel-like formulas into cells
5. **Named Ranges:** Important cells are named for easy reference

### Code Structure

```
Main Function
├── Clear existing sheets (optional)
├── Create Cover sheet
├── Create Input/Assumptions sheet
├── Create calculation sheets
├── Create reconciliation sheets
├── Setup named ranges
├── Apply formatting
└── Show success message
```

### Key Functions You Might See

- `SpreadsheetApp.getActiveSpreadsheet()` - Gets the current spreadsheet
- `sheet.getRange()` - Selects cells
- `range.setValue()` - Puts data in cells
- `range.setFormula()` - Puts formulas in cells
- `range.setBackground()` - Changes cell colors

---

## Contributing

### How to Improve This Project

If you'd like to contribute:

1. **Fork** this repository
2. **Make improvements** (fix bugs, add features, improve documentation)
3. **Test** your changes thoroughly
4. **Submit a Pull Request** with clear description

### Ideas for Contribution

- Add more Ind AS standards (Ind AS 115, 19, etc.)
- Improve error handling
- Add data validation
- Create video tutorials
- Translate documentation
- Add example datasets

---

## License

This project is open source. Feel free to use, modify, and distribute.

---

## Support

For questions or issues:
- Check the [Troubleshooting](#troubleshooting) section
- Review script comments for detailed explanations
- Consult official Ind AS documentation for accounting guidance

---

## About Ind AS

**Ind AS** (Indian Accounting Standards) are accounting standards adopted by companies in India and are based on IFRS (International Financial Reporting Standards).

**Key Standards Covered:**
- **Ind AS 109:** Financial Instruments (replaces AS 30 under IGAAP)
- **Ind AS 116:** Leases (replaces AS 19 under IGAAP)
- **Ind AS 12:** Income Taxes / AS 22 (IGAAP)

---

## Changelog

### Version 1.0 (November 2025)
- Initial release
- Three core scripts: Ind AS 109, 116, and Deferred Tax
- Full automation with formula-driven calculations
- Professional formatting and audit trail

---

**Happy Auditing! 📊**

*This project is designed to save time and improve accuracy in Ind AS compliance work. Remember: always review and verify calculations according to your specific circumstances.*
