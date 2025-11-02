# IND AS 109 AUDIT BUILDER - IMPLEMENTATION GUIDE

## 📋 OVERVIEW

This Google Apps Script creates a complete, audit-ready working paper system for **Ind AS 109 - Financial Instruments** period book closure entries. The script automatically generates 11 interconnected sheets with dynamic formulas, audit trails, and compliance checks.

---

## 🚀 INSTALLATION INSTRUCTIONS

### Step 1: Create New Google Sheet
1. Open Google Sheets (sheets.google.com)
2. Create a new blank spreadsheet
3. Name it: "Ind AS 109 - Period Book Closure Workings [FY 2024-25]"

### Step 2: Open Apps Script Editor
1. Click **Extensions** → **Apps Script**
2. Delete any existing code in the editor
3. Copy the entire contents of `IndAS109_AuditBuilder.gs`
4. Paste into the Apps Script editor
5. Click **Save** (disk icon) and name the project "IndAS109_Builder"

### Step 3: Run the Script
1. Ensure the function dropdown shows `createIndAS109WorkingPapers`
2. Click **Run** (▶️ play button)
3. **First time only**: You'll need to authorize the script:
   - Click "Review Permissions"
   - Choose your Google account
   - Click "Advanced" → "Go to IndAS109_Builder (unsafe)"
   - Click "Allow"
4. Wait 10-20 seconds for all sheets to be created
5. Success message will appear!

### Step 4: Return to Spreadsheet
1. Close the Apps Script tab
2. Return to your Google Sheet
3. You'll see all 11 sheets created with the "Cover" sheet active

---

## 📊 WORKBOOK STRUCTURE

### Sheet 1: **Cover**
- **Purpose**: Executive dashboard with key metrics
- **Features**:
  - Company details input section
  - Real-time summary of financial asset classifications
  - ECL provision totals by stage
  - Net financial assets position
  - Navigation guide to other sheets
  - Compliance note

**What to do**: Fill in company name, financial year, period end date, and currency.

---

### Sheet 2: **Input_Variables** ⭐ START HERE
- **Purpose**: Master control panel for all assumptions
- **Input Cells** (Light Green):
  - Reporting dates (current and previous)
  - Risk-free rate (G-Sec 10Y)
  - ECL parameters (PD, LGD, DPD thresholds)
  - Fair value parameters (credit spreads, risk premium)
  - Materiality thresholds
  - Hedge effectiveness bounds

**Validation**: Built-in data validation for dates and ranges
**Impact**: Changes here cascade through all calculations

---

### Sheet 3: **Instruments_Register** 🔢 CRITICAL
- **Purpose**: Master list of ALL financial instruments
- **Input Cells** (Orange):
  - Instrument details (name, type, counterparty)
  - Dates (issue, maturity)
  - Financial terms (face value, coupon, EIR)
  - Opening balances
  - Security type and credit rating
  - DPD (Days Past Due)
  - SPPI test result (Pass/Fail)
  - Business model classification

**Sample Data**: 7 pre-populated instruments for demonstration
**Add More**: Simply copy the formula row and paste below

**Dropdowns Available**:
- Type: Loan, Bond, Equity, Mutual Fund, etc.
- Security: Secured, Unsecured, Equity, Sovereign
- Rating: AAA to D, Not Rated
- SPPI: Pass, Fail, Not Applicable
- Business Model: Hold to Collect, Hold to Collect & Sell, FVTPL

---

### Sheet 4: **Classification_Matrix**
- **Purpose**: Auto-classify per Ind AS 109 decision tree
- **Logic**:
  ```
  IF SPPI = Fail → FVTPL
  ELSE IF Business Model = FVTPL → FVTPL
  ELSE IF SPPI = Pass AND Business Model = Hold to Collect → Amortized Cost
  ELSE IF SPPI = Pass AND Business Model = Hold to Collect & Sell → FVOCI
  ELSE → FVTPL (default)
  ```
- **Features**:
  - Automatic classification for each instrument
  - Color coding (Green=AC, Blue=FVOCI, Orange=FVTPL)
  - Summary count and balance by classification
  - Ind AS 109 reference paragraphs

**What to Review**: Ensure classification matches intended accounting treatment

---

### Sheet 5: **Fair_Value_Workings**
- **Purpose**: Calculate fair value adjustments for FVTPL & FVOCI
- **Auto-Calculations**:
  - Period-end fair value (simplified model - uses +/-15% variance)
  - Fair value gain/loss
  - P&L impact (FVTPL instruments)
  - OCI impact (FVOCI instruments)
  - Fair value hierarchy level

- **Manual Override Section**:
  - Input actual market quotes/valuations
  - Document valuation source
  - Record valuation date

**Color Coding**: Green = Gains, Red = Losses

---

### Sheet 6: **ECL_Impairment** 🎯 CRITICAL
- **Purpose**: Three-stage Expected Credit Loss calculations
- **Auto-Calculations**:
  - Stage determination (based on DPD thresholds)
  - PD (Probability of Default) by stage
  - LGD (Loss Given Default) by security type
  - EAD (Exposure at Default) = Gross carrying amount
  - ECL = EAD × PD × LGD
  - Opening, movement, and closing provisions

**Stage Logic**:
- **Stage 1** (Green): DPD < 30 days → 12-month ECL
- **Stage 2** (Yellow): DPD 30-89 days → Lifetime ECL
- **Stage 3** (Red): DPD ≥ 90 days → Lifetime ECL (NPA)

**Summary Section**:
- Count and balance by stage
- ECL rate and coverage ratio
- Total provision movement

---

### Sheet 7: **Amortization_Schedule**
- **Purpose**: EIR-based amortization for Amortized Cost instruments
- **Calculations**:
  - Interest income = Opening Balance × EIR × (Days/365)
  - Cash received = Face Value × Coupon Rate
  - Amortization = Interest Income - Cash Received
  - Closing balance = Opening + Interest - Cash - Impairment + Adjustments

- **Input Cell**: "Other Adjustments" (light green) for manual entries

**Verification**: Summary section includes zero-balance check

---

### Sheet 8: **Period_End_Entries** 📝 EXTRACT FOR POSTING
- **Purpose**: Ready-to-post journal entries
- **Entries Generated**:
  1. **JE001**: FVTPL Fair Value Adjustments
  2. **JE002**: FVOCI Fair Value Adjustments (to OCI)
  3. **JE003**: Interest Income (EIR method)
  4. **JE004**: ECL Provision Movement
  5. **JE005**: Premium/Discount Amortization

**Features**:
- Debit/Credit columns with auto-balancing
- Narration for each entry
- Ind AS 109 reference paragraphs
- Summary table with P&L impact
- **Balancing check** (should be zero)

**Action Required**: Copy these entries to your general ledger

---

### Sheet 9: **Reconciliation**
- **Purpose**: Complete opening-to-closing trail
- **Reconciles**:
  - Financial assets by classification
  - ECL provisions by stage
  - Net financial assets position
  - P&L impact summary
  - OCI movements

**Control Total**: Opening + Movements - Closing = 0 (with conditional formatting)

---

### Sheet 10: **References**
- **Purpose**: Quick reference to Ind AS 109 key provisions
- **Sections**:
  - Classification & Measurement
  - Business Model Assessment
  - SPPI Test
  - Effective Interest Rate
  - ECL Impairment Model
  - Fair Value Measurement
  - Hedge Accounting
  - Derecognition
  - Disclosure Requirements (Ind AS 107)

**Use Case**: Support for audit queries and documentation

---

### Sheet 11: **Audit_Notes**
- **Purpose**: Control checks, assertions, and sign-off
- **Features**:
  - **5 Control Totals** with Pass/Fail status
  - **8 Audit Assertions** (Completeness, Accuracy, Valuation, etc.)
  - **7 Risk Areas** with priority coding
  - **Materiality Assessment** section
  - **Audit Conclusion & Sign-Off** area

**Color Coding**:
- ✓ Green = Pass
- ✗ Red = Fail
- ⚠ Yellow = Review Required

---

## 🎨 COLOR CODING SYSTEM

| Color | Meaning | Usage |
|-------|---------|-------|
| **Light Blue (#e3f2fd)** | Primary Input Cell | General user entries (Cover, Input_Variables) |
| **Light Green (#e1f5e1, #e8f5e9)** | Input Cell | Specific adjustments (Amortization) |
| **Orange (#fff3e0)** | Critical Input | Instruments Register (all instrument data) |
| **Dark Blue (#1a237e)** | Main Header | Sheet titles |
| **Medium Blue (#283593)** | Sub Header | Column headers |
| **Green (#c8e6c9)** | Positive/Stage 1/AC | Gains, performing assets |
| **Yellow (#fff9c4)** | Caution/Stage 2 | Underperforming assets, review items |
| **Red (#ffcdd2)** | Negative/Stage 3/Alert | Losses, NPAs, errors |

---

## 🔧 MAINTENANCE & UPDATES

### Adding New Instruments
1. Go to **Instruments_Register**
2. Copy row 9 (last sample instrument)
3. Paste in next empty row
4. Update all values
5. All other sheets auto-update via formulas

### Modifying Assumptions
1. Go to **Input_Variables**
2. Change relevant parameters
3. All calculations refresh automatically

### Quarterly/Annual Updates
1. Update reporting date in **Cover** and **Input_Variables**
2. Update opening balances in **Instruments_Register**
3. Refresh DPD values
4. Review credit ratings
5. Update fair values (if manual overrides used)

### Formula Protection (Optional)
To prevent accidental formula deletion:
1. Select all formula cells (non-input cells)
2. Right-click → **Protect range**
3. Set permissions to "Only you"
4. Leave input cells unprotected

---

## ✅ VALIDATION CHECKLIST

Before finalizing:

- [ ] All input cells (blue/green/orange) are filled
- [ ] Control totals in **Audit_Notes** show "✓ Pass"
- [ ] Balancing check in **Period_End_Entries** = 0
- [ ] Reconciliation verification = 0
- [ ] ECL provisions reviewed for reasonableness
- [ ] Fair values supported by external sources
- [ ] SPPI and Business Model assessments documented
- [ ] All instruments classified correctly
- [ ] Journal entries extracted and ready for posting
- [ ] Audit sign-off completed

---

## 🔍 COMMON ISSUES & TROUBLESHOOTING

### Issue: #REF! Errors Appearing
**Cause**: Sheet referenced in formula was renamed or deleted
**Solution**: Do not rename or delete any sheets. If needed, re-run the script.

### Issue: Control Total Fails
**Cause**: Manual edits broke formula linkages
**Solution**: Check formulas in failing control. Compare with original structure.

### Issue: Circular Reference Warning
**Cause**: Should not occur with this script
**Solution**: If appears, check for accidental manual formula edits

### Issue: Classification Shows "Review Required"
**Cause**: Invalid combination of SPPI/Business Model
**Solution**: Review logic in **Classification_Matrix** notes and correct source data

### Issue: ECL Provision Seems Too High/Low
**Cause**: PD/LGD parameters may need adjustment
**Solution**: Review **Input_Variables** B10:B14 and adjust based on historical data

---

## 📈 ADVANCED FEATURES

### Named Ranges
The script creates named ranges for key inputs:
- `ReportingDate`
- `RiskFreeRate`
- `PD_Stage1`, `PD_Stage2`, `PD_Stage3`
- `InstrumentsList`

**Use in formulas**: =ReportingDate instead of =Input_Variables!B4

### Custom Menu
After running the script once, you'll see a custom menu:
**📊 Ind AS 109 Tools**
- 🚀 Create Working Papers
- 🔄 Refresh All Formulas
- 📋 Export Journal Entries
- 📖 Help & Documentation

### Data Validation
Dropdowns are automatically configured for:
- Dates (with date picker)
- Instrument types
- Security types
- Credit ratings
- SPPI test results
- Business models

---

## 📊 SAMPLE DATA OVERVIEW

The script includes 7 sample instruments for demonstration:

1. **FI001**: Term Loan (Amortized Cost, Stage 1)
2. **FI002**: Corporate Bond (Amortized Cost, Stage 1)
3. **FI003**: Equity Investment (FVTPL)
4. **FI004**: Trade Receivable (Stage 2 - DPD 15 days)
5. **FI005**: Government Security (FVOCI)
6. **FI006**: Mutual Fund (FVTPL)
7. **FI007**: Stressed Loan (Stage 3 - DPD 120 days)

**Action**: Delete sample data and replace with your actual instruments.

---

## 🎯 IND AS 109 COMPLIANCE NOTES

### Classification Requirements
Per **Para 4.1.1**, this workbook classifies instruments as:
- **Amortized Cost**: SPPI Pass + Hold to Collect
- **FVOCI**: SPPI Pass + Hold to Collect & Sell
- **FVTPL**: Default category, SPPI Fail, or irrevocable election

### Impairment (ECL) Methodology
Per **Para 5.5**, the three-stage approach:
- **Stage 1**: No significant increase in credit risk → 12-month ECL
- **Stage 2**: Significant increase but not impaired → Lifetime ECL
- **Stage 3**: Credit-impaired → Lifetime ECL on net carrying amount

**Rebuttable Presumption** (Para 5.5.11): DPD > 30 days = Stage 2

### EIR Calculation
Per **Para 5.4.1 & B5.4.1**, EIR includes:
- All fees paid/received between parties
- Transaction costs
- Premiums or discounts
- Excludes expected credit losses

### Fair Value Hierarchy
Per **Ind AS 113**:
- **Level 1**: Quoted prices in active markets
- **Level 2**: Observable inputs (used in this workbook for demo)
- **Level 3**: Unobservable inputs

---

## 📋 DISCLOSURE REQUIREMENTS

Users must prepare separate disclosures per **Ind AS 107**:

1. **Significance of Financial Instruments**
   - Carrying amounts by classification
   - Income, expenses, gains, losses

2. **Nature and Extent of Risks**
   - Credit risk, liquidity risk, market risk
   - Credit risk concentrations

3. **ECL Disclosures**
   - Reconciliation of ECL provision
   - Staging analysis
   - Significant judgments

4. **Fair Value**
   - Fair value hierarchy levels
   - Valuation techniques
   - Significant unobservable inputs

5. **Hedge Accounting** (if applicable)
   - Types of hedges
   - Risk management objectives
   - Effectiveness testing

---

## 🛡️ AUDIT ASSERTIONS COVERAGE

| Assertion | How This Workbook Addresses It |
|-----------|-------------------------------|
| **Existence** | All instruments traced to source documents (register) |
| **Completeness** | Control total ensures all instruments included |
| **Accuracy** | Mathematical checks at multiple levels |
| **Valuation** | Fair value sources documented, ECL per approved methodology |
| **Classification** | Systematic SPPI and business model tests |
| **Presentation** | Assets shown net of ECL, proper P&L vs OCI allocation |
| **Disclosure** | References sheet guides note preparation |

---

## 💡 BEST PRACTICES

1. **Monthly Review**: Update DPD and staging monthly, not just year-end
2. **Documentation**: Maintain separate files for:
   - SPPI test conclusions
   - Business model assessment
   - Fair value sources
   - PD/LGD derivation
3. **Version Control**: Save dated copies (e.g., "Ind AS 109 - Q1 FY25")
4. **Backup**: Keep PDF exports of final working papers
5. **Audit Trail**: Use cell comments to explain significant judgments
6. **Review**: Have a second person review all inputs and classifications
7. **Reconciliation**: Tie to general ledger before finalizing

---

## 🔗 INTEGRATION WITH OTHER SYSTEMS

### From General Ledger
**Import to Instruments_Register**:
- Opening balances (Column J)
- Instrument details
- Current credit ratings

### To General Ledger
**Export from Period_End_Entries**:
- Copy journal entries (JE001 to JE005)
- Post with appropriate dates
- Reference this workbook as support

### To Financial Statements
**Use Cover Sheet**:
- Financial asset totals
- ECL provision totals
- Net financial assets

### To Disclosures (Notes to Accounts)
**Use References & Audit_Notes**:
- Staging analysis
- Fair value levels
- Risk concentrations
- Accounting policies

---

## 📞 SUPPORT & CUSTOMIZATION

### Need More Features?
This script can be extended to include:
- Hedge accounting effectiveness tests
- Foreign currency revaluation
- Complex derivatives valuation
- Integration with external data sources
- Automated PDF report generation
- Email notifications

### Customization Areas
Edit the script to:
- Change color schemes
- Modify materiality thresholds
- Add custom risk models
- Include additional sheets
- Modify journal entry templates

### Professional Consultation
For complex instruments or significant portfolios, consult:
- Chartered Accountants specializing in Ind AS
- Financial instrument valuation experts
- External auditors

---

## 📚 LEARNING RESOURCES

### Ind AS 109 - Full Text
Available at: ICAI website → Standards → Indian Accounting Standards

### Key References
- **Ind AS 109**: Financial Instruments
- **Ind AS 107**: Financial Instruments: Disclosures
- **Ind AS 113**: Fair Value Measurement
- **Ind AS 32**: Financial Instruments: Presentation

### Related Guidance
- ICAI Guidance Note on Accounting for Derivatives
- RBI Master Circular on NPA Recognition
- ICAI Educational Material on ECL

---

## ✨ VERSION HISTORY

**Version 1.0** (Current)
- Complete 11-sheet working paper system
- Auto-classification per Ind AS 109
- Three-stage ECL impairment
- Fair value workings (FVTPL/FVOCI)
- EIR-based amortization
- Period-end journal entries
- Full audit trail and controls
- Sample data for 7 instruments

---

## 📄 LICENSE & DISCLAIMER

**Usage Rights**: This script is provided for legitimate accounting and audit purposes.

**Disclaimer**:
- This tool is for reference and should not replace professional judgment
- Users are responsible for ensuring accuracy of inputs and assumptions
- Complex instruments may require additional analysis beyond this workbook
- Always consult with qualified auditors and accountants
- This tool does not constitute professional advice

**Professional Responsibility**:
- The preparer and reviewer are responsible for the accuracy of financial reporting
- Ind AS 109 requires significant professional judgment
- External audit may require additional procedures

---

## 🎓 TRAINING RECOMMENDATIONS

### For Finance Teams
1. Ind AS 109 fundamentals (8 hours)
2. Hands-on with this workbook (4 hours)
3. Case studies on classification (4 hours)
4. ECL modeling workshop (8 hours)

### For Auditors
1. Ind AS 109 audit approach (4 hours)
2. Testing ECL models (4 hours)
3. Fair value audit procedures (4 hours)
4. Using this workbook for audit (2 hours)

---

## ⚡ QUICK START GUIDE (5 Minutes)

1. **Run Script** → Creates all sheets
2. **Fill Input_Variables** → Basic parameters
3. **Enter Instruments_Register** → Your actual instruments
4. **Review Classification_Matrix** → Verify auto-classification
5. **Check Period_End_Entries** → Extract journal entries

**That's it!** All other sheets auto-populate.

---

## 📧 FEEDBACK & IMPROVEMENTS

This is a living tool. Suggested improvements welcome:
- Additional automation features
- Integration with accounting systems
- Enhanced risk models
- More sophisticated fair value calculations
- Additional control checks

---

**END OF IMPLEMENTATION GUIDE**

For questions or issues, refer to the "References" sheet in the workbook or consult your external auditors.

**Happy Auditing! 🎯**
