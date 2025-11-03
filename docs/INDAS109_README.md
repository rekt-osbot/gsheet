# Ind AS 109 - Financial Instruments Audit Workbook

> Comprehensive audit working papers for Financial Instruments under Indian Accounting Standard 109

## 📋 Overview

This Google Sheets workbook automates the creation of audit working papers for Ind AS 109 (Financial Instruments), covering classification, measurement, impairment (Expected Credit Loss), and hedge accounting.

**Replaces:** AS 30 (IGAAP)  
**Effective:** For periods beginning on or after April 1, 2018  
**Complexity:** High - Requires understanding of fair value, ECL, and effective interest rate

---

## 🎯 What Does This Workbook Do?

### Core Functions
1. **Classification & Measurement** - Categorizes financial instruments (Amortized Cost, FVOCI, FVTPL)
2. **Fair Value Calculations** - Tracks fair value changes and adjustments
3. **ECL Impairment** - Calculates Expected Credit Loss using simplified/general approach
4. **Amortization Schedule** - Effective Interest Rate (EIR) method calculations
5. **Hedge Accounting** - Tracks hedge effectiveness and adjustments
6. **Journal Entries** - Auto-generates period-end accounting entries
7. **Reconciliation** - Opening to closing balance reconciliation

---

## 📊 Sheets Included

### 1. Cover Sheet
- Workbook overview
- Entity details
- Reporting period
- Preparer information

### 2. Input_Variables
**Purpose:** Central configuration sheet

**Key Inputs:**
- Entity name and reporting period
- Base currency
- Tax rates (for deferred tax on FVOCI)
- Discount rates for ECL
- Risk-free rates for fair value

**Color Coding:**
- 🟦 Light blue = Input cells (fill these)
- ⬜ White = Calculated cells (auto-filled)

### 3. Instruments_Register
**Purpose:** Master list of all financial instruments

**Columns:**
- Instrument ID
- Instrument Type (Debt/Equity/Derivative)
- Classification (AC/FVOCI/FVTPL)
- Counterparty details
- Original terms (amount, rate, maturity)
- Current carrying amount
- Fair value

**Classification Logic:**
- **Amortized Cost (AC):** Hold to collect, SPPI test passed
- **FVOCI:** Hold to collect & sell, SPPI test passed
- **FVTPL:** Trading, failed SPPI test, or designated

### 4. Fair_Value_Workings
**Purpose:** Fair value hierarchy and valuation

**Features:**
- Level 1: Quoted prices (mark-to-market)
- Level 2: Observable inputs (yield curves, credit spreads)
- Level 3: Unobservable inputs (DCF models)
- Valuation technique documentation
- Fair value adjustments tracking

**Formulas:**
```
Fair Value (Level 2) = PV of future cash flows at market rate
Fair Value Gain/Loss = Current FV - Previous FV
```

### 5. ECL_Impairment
**Purpose:** Expected Credit Loss calculation

**Approaches:**
1. **Simplified Approach** - Trade receivables, contract assets
2. **General Approach** - All other financial assets

**Three Stages:**
- **Stage 1:** 12-month ECL (no significant increase in credit risk)
- **Stage 2:** Lifetime ECL (significant increase in credit risk)
- **Stage 3:** Lifetime ECL (credit-impaired)

**ECL Formula:**
```
ECL = EAD × PD × LGD
Where:
- EAD = Exposure at Default
- PD = Probability of Default
- LGD = Loss Given Default
```

**Columns:**
- Instrument ID
- Stage (1/2/3)
- Gross carrying amount
- EAD (Exposure at Default)
- PD (Probability of Default %)
- LGD (Loss Given Default %)
- ECL Amount
- Allowance account

### 6. Amortization_Schedule
**Purpose:** Effective Interest Rate (EIR) calculations

**For:** Instruments measured at Amortized Cost or FVOCI

**Columns:**
- Period
- Opening balance
- Interest income (at EIR)
- Cash flows (principal + interest)
- Closing balance

**EIR Method:**
```
Interest Income = Opening Carrying Amount × EIR
Carrying Amount = Opening + Interest - Cash Received
```

**Handles:**
- Transaction costs
- Premiums/discounts
- Modification gains/losses

### 7. FVOCI_Reserve
**Purpose:** Track Other Comprehensive Income for FVOCI instruments

**Movements:**
- Fair value gains/losses (OCI)
- Reclassification to P&L (on derecognition)
- Tax effects on OCI
- Recycling adjustments

**Formula:**
```
FVOCI Reserve = Cumulative FV changes - Tax effect
```

### 8. Hedge_Accounting
**Purpose:** Document hedge relationships and effectiveness

**Hedge Types:**
- Fair value hedge
- Cash flow hedge
- Net investment hedge

**Effectiveness Testing:**
- Prospective test (forward-looking)
- Retrospective test (80-125% range)
- Ineffectiveness calculation

**Columns:**
- Hedge relationship ID
- Hedged item
- Hedging instrument
- Hedge ratio
- Effectiveness %
- Ineffective portion

### 9. Derecognition_Log
**Purpose:** Track disposal/settlement of instruments

**Triggers:**
- Sale of instrument
- Maturity/settlement
- Substantial modification
- Transfer of risks & rewards

**Gain/Loss Calculation:**
```
Gain/Loss = Proceeds - Carrying Amount
```

### 10. Period_End_Entries
**Purpose:** Auto-generated journal entries

**Entry Types:**
- Fair value adjustments (FVTPL → P&L)
- ECL provision movements
- Interest income accrual
- FVOCI movements (→ OCI)
- Hedge accounting adjustments
- Reclassifications

**Format:**
- Date
- Account
- Debit
- Credit
- Narration
- Reference to working

### 11. Reconciliation
**Purpose:** Opening to closing balance reconciliation

**By Classification:**
- Amortized Cost
- FVOCI
- FVTPL

**Movements:**
- Opening balance
- Additions
- Disposals
- Fair value changes
- ECL movements
- Interest accrued
- Closing balance

### 12. Audit_Notes
**Purpose:** Documentation and references

**Contents:**
- Ind AS 109 key requirements
- Classification decision tree
- ECL methodology
- Significant judgments
- Audit procedures performed
- References to supporting documents

---

## 🚀 Quick Start Guide

### Step 1: Create Workbook
1. Open new Google Sheet
2. Extensions > Apps Script
3. Copy `indas109.gs` code
4. Run `createIndAS109WorkingPapers()`
5. Authorize when prompted

### Step 2: Configure
1. Go to **Input_Variables** sheet
2. Fill in:
   - Entity name
   - Reporting period
   - Currency
   - Tax rates
   - Discount rates

### Step 3: Enter Instruments
1. Go to **Instruments_Register**
2. Enter each financial instrument:
   - Debt securities
   - Equity investments
   - Loans & receivables
   - Derivatives
   - Trade receivables

### Step 4: Review Calculations
1. **Fair_Value_Workings** - Check FV calculations
2. **ECL_Impairment** - Review ECL provisions
3. **Amortization_Schedule** - Verify EIR calculations
4. **Period_End_Entries** - Review journal entries

### Step 5: Reconcile
1. Go to **Reconciliation** sheet
2. Verify opening balances
3. Check all movements
4. Confirm closing balances tie to financial statements

---

## 📐 Key Formulas & Logic

### Classification Decision Tree

```
Is it a financial asset?
├─ YES → Does it pass SPPI test?
│  ├─ YES → Business model?
│  │  ├─ Hold to collect → Amortized Cost
│  │  ├─ Hold to collect & sell → FVOCI
│  │  └─ Other → FVTPL
│  └─ NO → FVTPL
└─ NO → Not in scope
```

### SPPI Test (Solely Payments of Principal and Interest)
- Principal = Amount advanced
- Interest = Consideration for time value of money + credit risk
- Fails if: Leverage, inverse relationship, non-recourse

### ECL Staging

```
Stage 1: No significant increase in credit risk
- 12-month ECL
- Interest on gross carrying amount

Stage 2: Significant increase in credit risk
- Lifetime ECL
- Interest on gross carrying amount

Stage 3: Credit-impaired
- Lifetime ECL
- Interest on net carrying amount (gross - ECL)
```

### Effective Interest Rate

```
EIR = Rate that discounts future cash flows to initial carrying amount

Initial Carrying Amount = Fair Value + Transaction Costs

Interest Income = Carrying Amount × EIR × Time Period
```

---

## ✅ Recent Improvements (v1.0.1)

### ECL Discounting Enhancement
**Fixed in v1.0.1:** ECL calculations now include present value discounting  
**Benefit:** Accurate ECL measurement for long-term exposures  
**Status:** ✅ Resolved

### EIR Method Improvement
**Enhanced in v1.0.1:** Improved effective interest rate calculations  
**Benefit:** More accurate interest income recognition  
**Status:** ✅ Enhanced

All previously documented limitations have been addressed. The workbook is production-ready for professional use.

---

## 🎓 Ind AS 109 Key Concepts

### Classification Categories

**Amortized Cost (AC)**
- Business model: Hold to collect contractual cash flows
- Cash flows: SPPI test passed
- Measurement: EIR method, less ECL
- Example: Loans, bonds held to maturity

**Fair Value through OCI (FVOCI)**
- Business model: Hold to collect AND sell
- Cash flows: SPPI test passed
- Measurement: Fair value, changes in OCI
- Example: Debt securities (available for sale)

**Fair Value through P&L (FVTPL)**
- Default category (residual)
- Measurement: Fair value, changes in P&L
- Example: Trading securities, derivatives

### Impairment - Expected Credit Loss

**Key Principles:**
- Forward-looking (not incurred loss)
- Probability-weighted
- Time value of money
- Reasonable & supportable information

**Simplified Approach:**
- Trade receivables without financing component
- Lifetime ECL from day 1
- Provision matrix based on aging

**General Approach:**
- Three-stage model
- 12-month vs lifetime ECL
- Significant increase in credit risk trigger

### Hedge Accounting

**Qualifying Criteria:**
- Formal designation & documentation
- Hedge effectiveness (80-125%)
- Economic relationship exists
- Credit risk doesn't dominate

**Fair Value Hedge:**
- Hedged item: Recognized asset/liability
- Gain/loss: Both in P&L
- Example: Fixed-rate bond hedged with IRS

**Cash Flow Hedge:**
- Hedged item: Forecast transaction
- Gain/loss: Effective portion in OCI
- Example: Foreign currency forecast sale

---

## 📋 Compliance Checklist

### Classification
- [ ] All financial instruments identified
- [ ] SPPI test performed and documented
- [ ] Business model assessment documented
- [ ] Classification appropriate and consistent

### Measurement
- [ ] Fair values determined using appropriate techniques
- [ ] Fair value hierarchy disclosed
- [ ] EIR calculated correctly including transaction costs
- [ ] Amortization schedules prepared

### Impairment
- [ ] ECL model selected (simplified vs general)
- [ ] Staging assessment performed
- [ ] PD, LGD, EAD determined
- [ ] Forward-looking information incorporated
- [ ] Significant increase in credit risk defined

### Hedge Accounting
- [ ] Hedge relationships formally designated
- [ ] Hedge documentation prepared
- [ ] Effectiveness testing performed
- [ ] Ineffectiveness quantified

### Disclosure
- [ ] Accounting policies disclosed
- [ ] Significant judgments explained
- [ ] Fair value hierarchy disclosed
- [ ] ECL methodology disclosed
- [ ] Credit risk disclosures complete

---

## 🔍 Audit Procedures

### Classification Testing
1. Review business model documentation
2. Test SPPI cash flows
3. Verify classification consistency
4. Check for reclassifications

### Fair Value Testing
1. Obtain independent valuations
2. Test Level 1 prices to market data
3. Review Level 2 inputs for reasonableness
4. Challenge Level 3 assumptions
5. Recalculate fair values

### ECL Testing
1. Review ECL methodology
2. Test PD, LGD, EAD inputs
3. Verify staging assessment
4. Check forward-looking adjustments
5. Recalculate ECL provision
6. Test write-offs and recoveries

### EIR Testing
1. Recalculate EIR
2. Verify transaction costs included
3. Test amortization calculations
4. Check for modifications

---

## 💡 Best Practices

### Data Management
1. Maintain complete instrument register
2. Update fair values regularly
3. Document all assumptions
4. Keep audit trail of changes

### ECL Modeling
1. Use historical loss rates as starting point
2. Adjust for forward-looking information
3. Segment portfolio by risk characteristics
4. Review staging quarterly
5. Document significant judgments

### Documentation
1. Prepare classification memos
2. Document business model
3. Maintain valuation reports
4. Keep hedge documentation current
5. Update accounting policies

### Controls
1. Segregate duties (recording vs valuation)
2. Independent price verification
3. Management review of ECL
4. Hedge effectiveness monitoring
5. Reconciliation to GL

---

## 📚 References

### Standards
- Ind AS 109 - Financial Instruments
- Ind AS 107 - Financial Instruments: Disclosures
- Ind AS 32 - Financial Instruments: Presentation

### Guidance
- ICAI Implementation Guide on Ind AS 109
- IASB Educational Material on IFRS 9
- RBI Guidelines on ECL (for banks)

### Useful Links
- [MCA Ind AS Portal](https://www.mca.gov.in/)
- [ICAI Resources](https://www.icai.org/)
- [IFRS Foundation](https://www.ifrs.org/)

---

## 🔄 Updates & Maintenance

### When to Update
- New financial instruments acquired
- Fair value changes (quarterly minimum)
- Credit risk changes
- Modifications to instruments
- Hedge relationship changes

### Quarterly Tasks
- Update fair values
- Reassess ECL staging
- Test hedge effectiveness
- Prepare journal entries
- Reconcile to GL

### Year-End Tasks
- Complete impairment assessment
- Finalize fair value measurements
- Prepare disclosure schedules
- Document significant judgments
- Archive working papers

---

## ⚙️ Customization Tips

### Adding New Instrument Types
1. Add row in Instruments_Register
2. Update classification logic if needed
3. Add to relevant calculation sheets
4. Update reconciliation

### Modifying ECL Model
1. Adjust PD/LGD/EAD inputs in ECL_Impairment
2. Update staging criteria
3. Modify formulas as needed
4. Document changes in Audit_Notes

### Custom Reports
1. Use pivot tables on Instruments_Register
2. Create charts for fair value trends
3. Build ECL movement analysis
4. Design custom dashboards

---

## 🐛 Troubleshooting

**Issue:** Fair value not calculating  
**Fix:** Check if market data is entered, verify formula references

**Issue:** ECL showing zero  
**Fix:** Ensure PD, LGD, EAD are populated, check staging

**Issue:** EIR amortization incorrect  
**Fix:** Verify initial carrying amount includes transaction costs

**Issue:** Reconciliation not balancing  
**Fix:** Check for missing entries, verify all movements captured

**Issue:** Journal entries not generating  
**Fix:** Ensure all input sheets are complete, check formula errors

---

## 📄 License

Open source - Free to use, modify, and distribute for audit and compliance purposes.

---

## 📝 Version History

**Version 1.0.1 (November 2025)**
- Added present value discounting to ECL calculations
- Improved EIR method accuracy
- Enhanced stage determination logic
- All known issues resolved

**Version 1.0 (November 2025)**
- Initial release
- 12 interconnected sheets
- Classification, measurement, ECL, hedge accounting
- Auto-generated journal entries
- Comprehensive reconciliation

---

**For complex financial instruments, always consult with accounting experts and auditors.**

*Simplifying Ind AS 109 compliance, one instrument at a time.*
