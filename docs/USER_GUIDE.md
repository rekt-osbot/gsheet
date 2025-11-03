# User Guide - IGAAP-Ind AS Audit Workpaper Builder

## For End Users (Auditors, Accountants)

This guide is for users who want to use the pre-built audit workpapers in Google Sheets.

## Quick Start (5 Minutes)

### Step 1: Choose Your Workbook

Navigate to the `dist/` folder and select the workbook you need:

| Workbook | File Name | Use Case |
|----------|-----------|----------|
| **Deferred Tax** | `deferredtax_standalone.gs` | IGAAP (AS 22) & Ind AS 12 deferred tax calculations |
| **Fixed Assets** | `far_wp_standalone.gs` | Fixed assets register and depreciation schedules |
| **ICFR P2P** | `ifc_p2p_standalone.gs` | Internal controls testing for Procure-to-Pay |
| **Ind AS 109** | `indas109_standalone.gs` | Financial instruments (classification, ECL, fair value) |
| **Ind AS 115** | `indas115_standalone.gs` | Revenue recognition (5-step model) |
| **Ind AS 116** | `indas116_standalone.gs` | Lease accounting workings |
| **TDS Compliance** | `tds_compliance_standalone.gs` | TDS compliance and reconciliation |

### Step 2: Open Google Sheets

1. Go to [Google Sheets](https://sheets.google.com)
2. Create a new blank spreadsheet
3. Give it a meaningful name (e.g., "ABC Ltd - Deferred Tax FY25")

### Step 3: Open Apps Script Editor

1. In your Google Sheet, click **Extensions** → **Apps Script**
2. You'll see a code editor with some default code
3. **Delete all the existing code** in the editor

### Step 4: Copy the Script

1. Open the standalone file you chose (e.g., `deferredtax_standalone.gs`)
2. **Select all the code** (Ctrl+A or Cmd+A)
3. **Copy it** (Ctrl+C or Cmd+C)

### Step 5: Paste and Save

1. Go back to the Apps Script editor
2. **Paste the code** (Ctrl+V or Cmd+V)
3. Click the **Save** icon (💾) or press Ctrl+S
4. Give your project a name (e.g., "Deferred Tax Automation")

### Step 6: Authorize the Script

1. Close the Apps Script editor tab
2. Go back to your Google Sheet
3. **Refresh the page** (F5 or Cmd+R)
4. You'll see a new menu appear (e.g., "Deferred Tax Tools")
5. Click the menu → **Create/Refresh Workbook**
6. Google will ask you to authorize the script:
   - Click **Continue**
   - Select your Google account
   - Click **Advanced** → **Go to [Project Name] (unsafe)**
   - Click **Allow**

### Step 7: Generate Your Workbook

1. After authorization, click the menu again
2. Click **Create/Refresh Workbook**
3. Wait 30-60 seconds (depending on workbook complexity)
4. Your workbook will be automatically created with all sheets!

## Using the Workbooks

### General Structure

All workbooks follow a similar structure:

1. **Cover/Index Sheet** - Overview and navigation
2. **Assumptions/Input Sheet** - Enter your parameters here (highlighted in yellow/blue)
3. **Data Input Sheets** - Enter transaction data
4. **Calculation Sheets** - Auto-calculated (formulas)
5. **Summary/Reconciliation Sheets** - Final outputs
6. **Reference Sheets** - Accounting standards guidance
7. **Audit Notes** - Document your review

### Input Cells

- **Yellow/Light Blue cells** = Input required
- **White/Grey cells** = Auto-calculated (don't edit)
- **Cells with notes** = Hover to see instructions

### Navigation

- Use the **Cover/Index sheet** for quick navigation
- Sheet tabs are color-coded by function
- Cross-references link automatically

## Workbook-Specific Guides

### Deferred Tax Workbook

**Purpose:** Calculate deferred tax assets (DTA) and liabilities (DTL) per IGAAP AS 22 or Ind AS 12.

**Key Inputs:**
1. **Assumptions Sheet:**
   - Entity name, financial year
   - Framework selection (IGAAP or Ind AS)
   - Tax rates (current and deferred)
   - Opening balances

2. **Temp_Differences Sheet:**
   - List all temporary differences
   - Enter tax base and book base
   - System auto-calculates DTA/DTL

**Outputs:**
- DT Schedule with full calculations
- Movement analysis (opening to closing)
- P&L reconciliation
- Balance sheet presentation

**Tips:**
- Framework selection affects recognition criteria
- IGAAP requires "virtual certainty" for loss-related DTA
- Ind AS allows DTA if "probable" (>50% likelihood)

### Ind AS 109 Workbook

**Purpose:** Financial instruments classification, measurement, and ECL impairment.

**Key Inputs:**
1. **Input_Variables Sheet:**
   - Reporting dates
   - Risk-free rate, PD/LGD parameters
   - Fair value parameters

2. **Instruments_Register Sheet:**
   - List all financial instruments
   - SPPI test results
   - Business model classification
   - DPD (days past due)

**Outputs:**
- Classification matrix (Amortized Cost/FVOCI/FVTPL)
- Fair value workings
- ECL impairment (3-stage approach)
- Period-end journal entries

**Tips:**
- SPPI test: "Pass" if cash flows are solely principal + interest
- Business model: Hold to Collect / Hold to Collect & Sell / Other
- Stage 1 = 12-month ECL, Stage 2/3 = Lifetime ECL

### Ind AS 115 Workbook

**Purpose:** Revenue recognition per 5-step model.

**Key Inputs:**
1. **Assumptions Sheet:**
   - Reporting period dates
   - Revenue recognition policies
   - Materiality thresholds

2. **Contract Register:**
   - All customer contracts
   - Contract dates, values
   - Performance obligations count

3. **Revenue Recognition Sheet:**
   - Revenue recognized per period
   - Progress % for over-time contracts
   - Billed amounts

**Outputs:**
- 5-step model application
- Contract assets/liabilities
- Period-end adjustments
- IGAAP reconciliation

**Tips:**
- Step 1: Identify contract
- Step 2: Identify performance obligations
- Step 3: Determine transaction price
- Step 4: Allocate price to POs
- Step 5: Recognize revenue when PO satisfied

### Ind AS 116 Workbook

**Purpose:** Lease accounting (lessee and lessor).

**Key Inputs:**
1. **Assumptions Sheet:**
   - Entity details
   - Incremental borrowing rate (IBR)
   - Short-term/low-value thresholds

2. **Lease Register:**
   - All lease contracts
   - Lease terms, payments
   - Lease classification

**Outputs:**
- Right-of-use asset schedule
- Lease liability amortization
- Depreciation and interest expense
- Journal entries
- Disclosure schedules

**Tips:**
- Short-term leases (≤12 months) can use simplified approach
- Low-value assets (typically <$5,000) can be expensed
- IBR = rate lessee would pay to borrow for similar term

### Fixed Assets Workbook

**Purpose:** Fixed assets register and audit testing.

**Key Inputs:**
1. **Summary Sheet:**
   - Opening balances by category
   - Additions, disposals, depreciation

2. **Roll Forward Sheet:**
   - Detailed asset movements

**Outputs:**
- Depreciation schedules
- Additions testing
- Disposals testing
- Physical verification
- Capitalization testing
- Disclosure checklist

**Tips:**
- Use consistent depreciation methods
- Test high-value additions
- Verify physical existence of assets

### TDS Compliance Workbook

**Purpose:** TDS compliance tracking and reconciliation.

**Key Inputs:**
1. **Assumptions Sheet:**
   - Financial year
   - Entity details
   - TDS rates

2. **Transactions Sheet:**
   - All TDS-applicable transactions
   - Payment dates, amounts
   - TDS deducted

**Outputs:**
- TDS summary by section
- Quarterly returns (26Q/27Q)
- Reconciliation with Form 26AS
- Late payment interest calculation
- Compliance checklist

**Tips:**
- Ensure timely deposit (by 7th of next month)
- File returns quarterly
- Reconcile with Form 26AS regularly

### ICFR P2P Workbook

**Purpose:** Internal controls testing for Procure-to-Pay process.

**Key Inputs:**
1. **Control Matrix:**
   - List of controls
   - Control owners
   - Testing frequency

2. **Testing Results:**
   - Sample selections
   - Test results (Pass/Fail)
   - Exceptions noted

**Outputs:**
- Control effectiveness summary
- Exception tracking
- Management letter points
- Remediation plan

**Tips:**
- Test key controls quarterly
- Document all exceptions
- Follow up on remediation

## Troubleshooting

### Script Authorization Issues

**Problem:** "This app isn't verified" warning

**Solution:**
1. Click **Advanced**
2. Click **Go to [Project Name] (unsafe)**
3. Click **Allow**

This is normal for custom scripts. Google shows this warning because the script isn't published in the Google Workspace Marketplace.

### Menu Not Appearing

**Problem:** Custom menu doesn't show after pasting script

**Solution:**
1. Make sure you saved the script (💾 icon)
2. Refresh your Google Sheet (F5)
3. Wait a few seconds for the menu to load
4. If still not appearing, close and reopen the spreadsheet

### Formulas Showing #REF! Errors

**Problem:** Formulas show #REF! errors

**Solution:**
1. This usually means a sheet name was changed
2. Re-run the workbook creation (menu → Create/Refresh Workbook)
3. Don't rename sheets after creation

### Slow Performance

**Problem:** Workbook is slow to calculate

**Solution:**
1. Reduce the number of input rows (delete unused rows)
2. Clear browser cache
3. Use Google Chrome for best performance
4. Avoid having too many sheets open simultaneously

### Data Lost After Refresh

**Problem:** Data disappeared after refreshing workbook

**Solution:**
1. The "Create/Refresh Workbook" function recreates all sheets
2. Always **save a copy** before refreshing
3. Or manually copy your input data to a separate sheet first

## Best Practices

### Data Entry

1. ✅ **Fill yellow/blue cells first** - these are required inputs
2. ✅ **Use consistent date formats** - DD-MMM-YYYY
3. ✅ **Enter amounts without currency symbols** - just numbers
4. ✅ **Use dropdowns where provided** - ensures data consistency
5. ❌ **Don't edit formula cells** - they auto-calculate

### Version Control

1. **Save versions regularly:**
   - File → Make a copy
   - Name with date: "ABC Ltd DT FY25 - 2025-11-03"

2. **Before major changes:**
   - Make a backup copy
   - Test in the copy first

3. **Final version:**
   - Lock input sheets (Data → Protect sheets)
   - Export to PDF for records

### Collaboration

1. **Share with team:**
   - Click Share button
   - Add team members with "Editor" access

2. **Track changes:**
   - Use version history (File → Version history)
   - Add comments (Insert → Comment)

3. **Review workflow:**
   - Preparer fills inputs
   - Reviewer checks calculations
   - Manager approves final version

## Support

### Getting Help

1. **Check the documentation:**
   - Review this guide
   - Check workbook-specific README files in `docs/` folder

2. **Review the code:**
   - Apps Script editor has comments explaining logic
   - Look for `// NOTE:` comments for important info

3. **Common issues:**
   - Most issues are due to missing inputs
   - Check all yellow/blue cells are filled
   - Verify date formats are correct

### Reporting Issues

If you find a bug:
1. Note the workbook name and version
2. Describe the steps to reproduce
3. Include screenshots if possible
4. Check if issue persists after refreshing workbook

## Updates

### Getting Latest Version

1. Check the `dist/` folder for updated files
2. Compare version numbers in file headers
3. Copy new version to Apps Script editor
4. Re-run workbook creation

### Migrating Data

To migrate data to a new version:
1. Export current data to CSV (File → Download → CSV)
2. Create new workbook with latest script
3. Import data back (File → Import)
4. Verify calculations

## Appendix: Keyboard Shortcuts

| Action | Windows | Mac |
|--------|---------|-----|
| Save | Ctrl+S | Cmd+S |
| Copy | Ctrl+C | Cmd+C |
| Paste | Ctrl+V | Cmd+V |
| Undo | Ctrl+Z | Cmd+Z |
| Find | Ctrl+F | Cmd+F |
| Refresh | F5 | Cmd+R |

---

**Version:** 1.0  
**Last Updated:** November 2025  
**For technical documentation, see:** `BUILD_SYSTEM.md`
