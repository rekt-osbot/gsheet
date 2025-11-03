# Deferred Tax Workbook (Ind AS 12 / AS 22)

> Comprehensive audit working papers for Deferred Taxation under Ind AS 12 or IGAAP AS 22

## 📋 Overview

This Google Sheets workbook automates the creation of audit working papers for deferred tax calculations, covering temporary differences, deferred tax assets (DTA), deferred tax liabilities (DTL), and reconciliations.

**Applicable Standards:**
- **Ind AS 12** - Income Taxes (converged with IAS 12)
- **AS 22** - Accounting for Taxes on Income (IGAAP)

**Key Difference:** Ind AS 12 uses balance sheet approach; AS 22 uses timing difference approach

---

## 🎯 What Does This Workbook Do?

### Core Functions
1. **Framework Selection** - Choose between Ind AS 12 or AS 22
2. **Temporary Differences** - Track differences between book and tax base
3. **DTA/DTL Calculation** - Compute deferred tax assets and liabilities
4. **Movement Analysis** - Track opening, additions, reversals, closing
5. **P&L Reconciliation** - Reconcile accounting profit to tax expense
6. **Balance Sheet Reconciliation** - Reconcile DTA/DTL balances
7. **Tax Rate Changes** - Handle changes in enacted tax rates
8. **Disclosure Schedules** - Prepare required disclosures

---

## 📊 Sheets Included

### 1. Cover Sheet
- Workbook overview
- Entity details
- Financial year
- Framework selection (Ind AS/IGAAP)

### 2. Assumptions
**Purpose:** Central configuration sheet

**Key Inputs:**
- Entity name and PAN
- Financial year
- Framework: Ind AS 12 or AS 22
- Current tax rate (%)
- Future tax rate (if enacted change)
- MAT rate (if applicable)
- Minimum Alternate Tax (MAT) credit tracking

**Tax Rates:**
- Domestic company: 25.17% (with surcharge & cess)
- Foreign company: 43.68%
- MAT: 18.5%

### 3. Temp_Differences
**Purpose:** Master list of all temporary differences

**Columns:**
- Item description
- Category (Property/Depreciation/Provisions/etc.)
- Book value (as per financial statements)
- Tax base (as per tax return)
- Temporary difference (Book - Tax)
- Nature (Deductible/Taxable)
- DTA/DTL amount
- Recognition (Yes/No with reason)

**Common Temporary Differences:**

**Deductible (DTA):**
- Provisions not allowed for tax (warranty, doubtful debts)
- Expenses disallowed u/s 43B (paid basis)
- Unabsorbed depreciation
- Carried forward losses
- Employee benefits (gratuity, leave)

**Taxable (DTL):**
- Depreciation (book < tax)
- Prepaid expenses
- Revenue recognition differences
- Fair value gains (not taxable yet)

### 4. DT_Schedule
**Purpose:** Detailed deferred tax calculation

**Structure:**
- Opening DTA/DTL
- Current year origination
- Current year reversal
- Tax rate changes
- Closing DTA/DTL

**Formula:**
```
DTA = Deductible Temporary Difference × Tax Rate
DTL = Taxable Temporary Difference × Tax Rate
Net DT = DTA - DTL
```

**Recognition Criteria (Ind AS 12):**
- DTA: Recognize if probable future taxable profit
- DTL: Recognize for all taxable temporary differences

**Recognition Criteria (AS 22):**
- Timing differences only (not all temporary differences)
- Virtual certainty for DTA (stricter than Ind AS)

### 5. Movement_Analysis
**Purpose:** Track movements in DTA/DTL

**By Category:**
- Property, plant & equipment
- Intangible assets
- Provisions
- Employee benefits
- Carried forward losses
- Others

**Movements:**
- Opening balance
- Additions (new temporary differences)
- Reversals (temporary differences reversed)
- Tax rate changes
- Reclassifications
- Closing balance

### 6. MAT_Credit
**Purpose:** Track Minimum Alternate Tax credit (India-specific)

**MAT Credit:**
- Arises when MAT > Normal tax
- Can be carried forward 15 years
- Utilized when Normal tax > MAT

**Columns:**
- Year of origin
- MAT credit amount
- Utilized (year-wise)
- Balance available
- Expiry year

**Recognition:**
- Recognize as DTA if probable utilization
- Review annually for impairment

### 7. Unrecognized_DTA
**Purpose:** Track DTA not recognized due to uncertainty

**Reasons for Non-Recognition:**
- Insufficient future taxable profit
- Expiry of carry forward period
- Change in business model
- Regulatory restrictions

**Columns:**
- Item description
- Temporary difference amount
- Potential DTA
- Reason for non-recognition
- Reassessment date

**Disclosure:** Required to disclose unrecognized DTA

### 8. Tax_Rate_Changes
**Purpose:** Track impact of enacted tax rate changes

**Accounting:**
- Remeasure all DTA/DTL at new rate
- Recognize impact in P&L (or OCI if related to OCI item)

**Example:**
```
Old rate: 30%
New rate: 25%
Temporary difference: ₹1,00,000
Old DTL: ₹30,000
New DTL: ₹25,000
P&L credit: ₹5,000
```

### 9. P&L_Reconciliation
**Purpose:** Reconcile accounting profit to tax expense

**Format:**
```
Accounting Profit before Tax         ₹ XXX
Add: Permanent differences (expenses)    XXX
Less: Permanent differences (income)    (XXX)
                                     -------
Taxable Income                       ₹ XXX
                                     =======

Current Tax @ XX%                    ₹ XXX
Deferred Tax (Net)                       XXX
                                     -------
Total Tax Expense                    ₹ XXX
                                     =======

Effective Tax Rate                      XX%
```

**Permanent Differences:**
- Non-deductible expenses (penalties, CSR)
- Exempt income (dividends, capital gains)
- Disallowances (80% of certain expenses)

### 10. BS_Reconciliation
**Purpose:** Reconcile DTA/DTL to balance sheet

**Format:**
```
Deferred Tax Assets:
- Property, plant & equipment         ₹ XXX
- Provisions                              XXX
- Employee benefits                       XXX
- Carried forward losses                  XXX
                                     -------
Total DTA                            ₹ XXX

Deferred Tax Liabilities:
- Property, plant & equipment         ₹ XXX
- Intangible assets                       XXX
- Others                                  XXX
                                     -------
Total DTL                            ₹ XXX

Net DTA/(DTL)                        ₹ XXX
                                     =======
```

**Presentation:**
- Offset DTA and DTL if legally enforceable right
- Present net amount in balance sheet

### 11. Disclosure_Schedule
**Purpose:** Prepare required disclosures

**Disclosures Required:**

**Ind AS 12:**
- Major components of tax expense
- Reconciliation of effective tax rate
- DTA/DTL by nature
- Unrecognized DTA
- Unrecognized DTL (subsidiaries)
- Tax consequences of dividends

**AS 22:**
- Current tax and deferred tax
- Timing differences
- Unabsorbed depreciation/losses
- Reasons for DTA non-recognition

### 12. Audit_Notes
**Purpose:** Documentation and references

**Contents:**
- Ind AS 12 vs AS 22 key differences
- Temporary vs timing differences
- Recognition criteria
- Measurement principles
- Significant judgments
- Tax planning strategies
- Audit procedures performed

---

## 🚀 Quick Start Guide

### Step 1: Create Workbook
1. Open new Google Sheet
2. Extensions > Apps Script
3. Copy `deferredtax.gs` code
4. Run `createDeferredTaxWorkbook()`
5. Authorize when prompted

### Step 2: Configure
1. Go to **Assumptions** sheet
2. Select framework: Ind AS 12 or AS 22
3. Enter tax rates
4. Specify entity details

### Step 3: Enter Temporary Differences
1. Go to **Temp_Differences** sheet
2. Enter each item:
   - Description
   - Book value
   - Tax base
   - Nature (Deductible/Taxable)

**Common Items:**
- Fixed assets (depreciation difference)
- Provisions (warranty, doubtful debts)
- Employee benefits (gratuity, leave)
- Expenses disallowed u/s 43B
- Carried forward losses

### Step 4: Review Calculations
1. **DT_Schedule** - Check DTA/DTL calculations
2. **Movement_Analysis** - Review movements
3. **P&L_Reconciliation** - Verify tax expense
4. **BS_Reconciliation** - Confirm balance sheet amounts

### Step 5: Assess Recognition
1. Review **Unrecognized_DTA** sheet
2. Assess probability of future taxable profit
3. Document reasons for non-recognition
4. Update recognition decisions

---

## 📐 Key Formulas & Logic

### Temporary Difference Calculation

```
Temporary Difference = Book Value - Tax Base

If positive (Book > Tax) → Taxable difference → DTL
If negative (Book < Tax) → Deductible difference → DTA
```

### DTA/DTL Calculation

```
DTA = Deductible Temporary Difference × Tax Rate
DTL = Taxable Temporary Difference × Tax Rate
```

### Recognition Decision Tree (Ind AS 12)

```
Is it a temporary difference?
├─ Taxable difference → Recognize DTL (always)
└─ Deductible difference → Probable future profit?
   ├─ YES → Recognize DTA
   └─ NO → Don't recognize (disclose)
```

### Recognition Decision Tree (AS 22)

```
Is it a timing difference?
├─ YES → Will it reverse?
│  ├─ YES → Recognize
│  └─ NO → Don't recognize
└─ NO → Not in scope (permanent difference)
```

### Effective Tax Rate Reconciliation

```
Statutory Tax Rate                    XX%
Add/Less:
- Permanent differences               X%
- Tax rate changes                    X%
- Unrecognized DTA                    X%
- Prior year adjustments              X%
                                    -----
Effective Tax Rate                    XX%
                                    =====
```

---

## ⚠️ Known Limitations (See todo.md)

### 1. Movement Analysis Flaws
**Issue:** Uses hardcoded percentages to distribute opening balances  
**Impact:** Unreliable movement analysis  
**Workaround:** Enter opening temporary differences directly

### 2. Arbitrary Additions/Reversals
**Issue:** Fixed proportions assumed for movements  
**Impact:** Doesn't reflect actual transactions  
**Workaround:** Link to actual transaction data

**Status:** HIGH PRIORITY FIX REQUIRED

---

## 🎓 Key Concepts

### Ind AS 12 vs AS 22

| Aspect | Ind AS 12 | AS 22 |
|--------|-----------|-------|
| Approach | Balance sheet (temporary differences) | P&L (timing differences) |
| Scope | All temporary differences | Only timing differences |
| DTA Recognition | Probable future profit | Virtual certainty |
| Initial Recognition | Exceptions for goodwill, certain transactions | Similar |
| Rate Changes | Remeasure immediately | Similar |
| Presentation | Net off if legal right | Similar |

### Temporary vs Timing Differences

**Temporary Differences (Ind AS 12):**
- Difference between book value and tax base
- Broader scope
- Includes items never in P&L (e.g., revaluation)

**Timing Differences (AS 22):**
- Difference in recognition timing in P&L
- Narrower scope
- Only items that go through P&L

**Example:**
- Revaluation of property: Temporary difference (Ind AS 12) but NOT timing difference (AS 22)

### Deferred Tax Asset Recognition

**Ind AS 12 - Probable:**
- More than 50% likelihood
- Based on business plans
- Taxable profit forecasts
- Tax planning strategies

**AS 22 - Virtual Certainty:**
- Very high degree of certainty
- Stricter than Ind AS
- Usually only for timing differences that will reverse

### Tax Base

**Definition:** Amount attributed to asset/liability for tax purposes

**Examples:**
```
Asset:
- Book value: ₹1,00,000
- Tax WDV: ₹80,000
- Tax base: ₹80,000
- Temporary difference: ₹20,000 (Taxable → DTL)

Liability (Provision):
- Book value: ₹50,000
- Tax deduction: When paid
- Tax base: ₹0
- Temporary difference: ₹50,000 (Deductible → DTA)
```

---

## 📋 Compliance Checklist

### Calculation
- [ ] All temporary differences identified
- [ ] Book values agree to financial statements
- [ ] Tax bases agree to tax computation
- [ ] Correct tax rate applied
- [ ] DTA/DTL calculated correctly

### Recognition
- [ ] DTA recognition assessed (probability/virtual certainty)
- [ ] Unrecognized DTA documented
- [ ] MAT credit recognized appropriately
- [ ] Offsetting applied correctly

### Measurement
- [ ] Enacted tax rates used
- [ ] Rate changes accounted for
- [ ] Discounting not applied (unless required)

### Presentation
- [ ] Current vs deferred tax split
- [ ] DTA and DTL offset if permitted
- [ ] Separate line items in balance sheet

### Disclosure
- [ ] Major components disclosed
- [ ] Effective tax rate reconciliation
- [ ] Unrecognized DTA disclosed
- [ ] Expiry dates disclosed

---

## 🔍 Audit Procedures

### Temporary Differences
1. Obtain schedule of temporary differences
2. Agree book values to financial statements
3. Agree tax bases to tax computation
4. Recalculate temporary differences
5. Verify classification (deductible/taxable)

### DTA/DTL Calculation
1. Verify tax rates used
2. Recalculate DTA/DTL
3. Check for rate changes
4. Test mathematical accuracy

### Recognition Assessment
1. Review business plans and forecasts
2. Assess probability of future taxable profit
3. Review tax planning strategies
4. Challenge management assumptions
5. Test sensitivity analysis

### Movement Analysis
1. Agree opening balances to prior year
2. Test additions (new temporary differences)
3. Test reversals (actual vs expected)
4. Verify rate change impacts
5. Agree closing balances to current year

### Reconciliations
1. Reconcile tax expense to accounting profit
2. Explain permanent differences
3. Verify effective tax rate
4. Reconcile DTA/DTL to balance sheet

---

## 💡 Best Practices

### Data Management
1. Maintain detailed temporary difference schedule
2. Update quarterly (minimum)
3. Document all assumptions
4. Keep audit trail of changes

### Recognition Assessment
1. Prepare detailed profit forecasts
2. Consider tax planning strategies
3. Review annually (minimum)
4. Document judgment and rationale
5. Obtain management representation

### Tax Rate Management
1. Monitor for enacted rate changes
2. Update immediately when enacted
3. Remeasure all DTA/DTL
4. Disclose impact separately

### Controls
1. Reconcile to tax computation
2. Independent review of calculations
3. Management approval of recognition
4. Quarterly monitoring
5. Year-end detailed review

---

## 📊 Common Scenarios

### Depreciation Difference
```
Fixed Asset Cost: ₹10,00,000
Book Depreciation (10 years): ₹1,00,000/year
Tax Depreciation (15%): ₹1,50,000/year

Year 1:
Book WDV: ₹9,00,000
Tax WDV: ₹8,50,000
Temporary Difference: ₹50,000 (Taxable)
DTL @ 25%: ₹12,500
```

### Provision for Warranty
```
Provision created: ₹2,00,000
Tax deduction: When paid

Book value: ₹2,00,000 (liability)
Tax base: ₹0 (no deduction yet)
Temporary Difference: ₹2,00,000 (Deductible)
DTA @ 25%: ₹50,000
```

### Carried Forward Loss
```
Tax loss: ₹50,00,000
Carry forward period: 8 years
Expected utilization: 5 years

If probable future profit:
DTA @ 25%: ₹12,50,000

If not probable:
DTA: ₹0 (disclose unrecognized)
```

---

## 📚 References

### Standards
- Ind AS 12 - Income Taxes
- AS 22 - Accounting for Taxes on Income
- Income Tax Act, 1961

### Guidance
- ICAI Implementation Guide on Ind AS 12
- ICAI Guidance Note on AS 22
- IASB Educational Material on IAS 12

### Useful Links
- [MCA Ind AS Portal](https://www.mca.gov.in/)
- [ICAI Resources](https://www.icai.org/)
- [Income Tax Department](https://www.incometax.gov.in/)

---

## 🔄 Updates & Maintenance

### Quarterly Tasks
- Update temporary differences
- Recalculate DTA/DTL
- Review recognition assessment
- Prepare journal entries
- Reconcile to GL

### Annual Tasks
- Comprehensive review of all temporary differences
- Reassess DTA recognition
- Update tax rates if changed
- Prepare disclosure schedules
- Archive working papers

### When Tax Rates Change
- Identify effective date
- Remeasure all DTA/DTL
- Calculate P&L impact
- Update disclosure
- Communicate to stakeholders

---

## ⚙️ Customization Tips

### Adding New Categories
1. Add row in Temp_Differences
2. Classify as deductible/taxable
3. Calculations auto-populate
4. Update Movement_Analysis grouping

### Changing Tax Rates
1. Update in Assumptions sheet
2. All calculations auto-update
3. Review Tax_Rate_Changes sheet
4. Document in Audit_Notes

### Custom Reports
1. Pivot table by category
2. Chart DTA/DTL trends
3. Analyze effective tax rate
4. Compare budget vs actual

---

## 🐛 Troubleshooting

**Issue:** DTA/DTL not calculating  
**Fix:** Check if temporary difference and tax rate are entered

**Issue:** Movement analysis incorrect  
**Fix:** Known issue - enter opening balances directly in Temp_Differences

**Issue:** Effective tax rate doesn't reconcile  
**Fix:** Check permanent differences, verify tax rate, review prior year adjustments

**Issue:** MAT credit not showing  
**Fix:** Ensure MAT > Normal tax, enter in MAT_Credit sheet

**Issue:** Reconciliation not balancing  
**Fix:** Check for missing entries, verify all movements captured

---

## 📄 License

Open source - Free to use, modify, and distribute for audit and compliance purposes.

---

## 📝 Version History

**Version 1.0 (November 2024)**
- Initial release
- Supports both Ind AS 12 and AS 22
- 12 interconnected sheets
- Comprehensive reconciliations
- MAT credit tracking
- Known issues documented

**Version 1.1 (Planned)**
- Fix movement analysis flaws
- Add transaction-level tracking
- Enhanced forecasting tools

---

**For complex tax situations, always consult with tax experts and auditors.**

*Simplifying deferred tax compliance, one temporary difference at a time.*
