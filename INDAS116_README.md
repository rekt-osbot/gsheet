# Ind AS 116 - Lease Accounting Workbook

> Comprehensive audit working papers for Lease Accounting under Indian Accounting Standard 116

## 📋 Overview

This Google Sheets workbook automates the creation of audit working papers for Ind AS 116 (Leases), covering lease identification, measurement of Right-of-Use (ROU) assets and lease liabilities, and subsequent accounting.

**Replaces:** AS 19 (IGAAP) - Operating vs Finance Lease distinction  
**Effective:** For periods beginning on or after April 1, 2019  
**Key Change:** Single lessee accounting model - all leases on balance sheet

---

## 🎯 What Does This Workbook Do?

### Core Functions
1. **Lease Identification** - Determines if contract contains a lease
2. **Initial Measurement** - Calculates ROU asset and lease liability at commencement
3. **ROU Asset Schedule** - Tracks depreciation and impairment
4. **Lease Liability Schedule** - Amortizes liability using effective interest method
5. **Payment Schedule** - Tracks actual vs expected payments
6. **Journal Entries** - Auto-generates accounting entries
7. **Reconciliation** - Opening to closing balance movements
8. **IGAAP Comparison** - Shows impact of transition from AS 19

---

## 📊 Sheets Included

### 1. Cover Sheet
- Workbook overview
- Entity details
- Reporting period
- Transition date (if applicable)

### 2. Assumptions
**Purpose:** Central configuration sheet

**Key Inputs:**
- Entity name and reporting date
- Base currency
- Default incremental borrowing rate (IBR)
- Depreciation policy
- Materiality thresholds
- Short-term lease threshold (12 months)
- Low-value asset threshold (₹5 lakhs)

**Practical Expedients:**
- Short-term leases (≤12 months)
- Low-value assets
- Non-lease components separation

### 3. Lease_Register
**Purpose:** Master list of all lease contracts

**Columns:**
- Lease ID
- Lease description
- Lessor name
- Asset type (Property/Vehicle/Equipment)
- Commencement date
- Lease term (months)
- Payment frequency
- Monthly/Annual payment
- Incremental Borrowing Rate (IBR)
- Initial direct costs
- Lease incentives received
- Variable payment terms
- Extension/termination options

**Lease Identification Criteria:**
- Identified asset
- Right to obtain substantially all economic benefits
- Right to direct use of asset

**Exemptions Applied:**
- Short-term leases (≤12 months)
- Low-value assets (≤₹5 lakhs)

### 4. ROU_Asset_Schedule
**Purpose:** Track Right-of-Use asset from commencement to end

**Initial Measurement:**
```
ROU Asset = Lease Liability (initial)
          + Initial direct costs
          + Restoration costs (if any)
          - Lease incentives received
```

**Columns:**
- Lease ID
- Opening balance
- Additions (new leases)
- Depreciation charge
- Impairment loss (if any)
- Disposals/terminations
- Closing balance

**Depreciation:**
- Method: Straight-line (or pattern of benefit)
- Period: Shorter of lease term or useful life
- Formula: `ROU Asset / Lease Term (months)`

### 5. Lease_Liability_Schedule
**Purpose:** Amortize lease liability using effective interest method

**Initial Measurement:**
```
Lease Liability = Present Value of:
                - Fixed payments
                - Variable payments (based on index/rate)
                - Residual value guarantees
                - Purchase option (if reasonably certain)
                - Termination penalties (if applicable)
```

**Columns:**
- Period (month/year)
- Opening liability
- Interest expense (Opening × IBR × Time)
- Lease payment
- Closing liability
- Current portion
- Non-current portion

**Interest Calculation:**
```
Interest Expense = Opening Liability × IBR × (Days/365)
```

**Current/Non-Current Split:**
- Current: Payments due within 12 months
- Non-current: Payments due after 12 months

### 6. Payment_Schedule
**Purpose:** Track actual lease payments vs schedule

**Columns:**
- Payment date
- Scheduled amount
- Actual amount paid
- Variance
- Cumulative payments
- Payment method
- Reference (invoice/receipt)

**Variance Analysis:**
- Timing differences
- Amount differences
- Reasons for variance

### 7. Modifications_Log
**Purpose:** Track lease modifications and remeasurements

**Modification Triggers:**
- Change in lease term
- Change in assessment of purchase option
- Change in amounts payable (index/rate change)
- Change in scope (add/remove asset)

**Accounting Treatment:**
- Separate lease: New lease accounting
- Not separate: Remeasure liability, adjust ROU asset

**Columns:**
- Modification date
- Type of modification
- Original terms
- Modified terms
- Remeasurement amount
- Accounting treatment

### 8. Variable_Payments
**Purpose:** Track variable lease payments not in liability

**Types:**
- Performance-based (e.g., % of sales)
- Usage-based (e.g., per km, per hour)
- Index-linked (after initial measurement)

**Accounting:**
- Not included in initial liability
- Expensed as incurred
- Disclosed separately

### 9. Sublease_Register
**Purpose:** Track subleases (if lessee subleases to third party)

**Accounting:**
- Lessee becomes intermediate lessor
- Classify sublease as finance or operating
- Based on ROU asset (not underlying asset)

**Columns:**
- Sublease ID
- Sublessee name
- Related head lease
- Sublease classification
- Sublease income
- Net position

### 10. Period_End_Entries
**Purpose:** Auto-generated journal entries

**Entry Types:**

**At Commencement:**
```
Dr. ROU Asset
Cr. Lease Liability
```

**Monthly/Periodic:**
```
Dr. Depreciation Expense
Cr. Accumulated Depreciation - ROU Asset

Dr. Interest Expense
Dr. Lease Liability
Cr. Cash/Bank
```

**Variable Payments:**
```
Dr. Lease Expense (Variable)
Cr. Cash/Bank
```

### 11. Reconciliation
**Purpose:** Opening to closing balance reconciliation

**ROU Asset Reconciliation:**
- Opening balance
- Additions (new leases)
- Depreciation
- Impairment
- Disposals
- Closing balance

**Lease Liability Reconciliation:**
- Opening balance
- Additions (new leases)
- Interest accrued
- Payments made
- Modifications
- Closing balance

### 12. IGAAP_Comparison
**Purpose:** Show impact of transition from AS 19

**Comparison:**
- AS 19: Operating lease (off balance sheet)
- Ind AS 116: On balance sheet

**Impact Analysis:**
- Balance sheet impact (ROU asset + Liability)
- P&L impact (Depreciation + Interest vs Rent)
- Cash flow impact (Operating vs Financing)
- Key ratios impact (Debt/Equity, EBITDA, etc.)

### 13. Disclosure_Schedules
**Purpose:** Prepare disclosure requirements

**Disclosures:**
- Maturity analysis of lease liabilities
- Lease expense breakdown
- Cash flow information
- Short-term and low-value lease expense
- Variable lease payments
- Income from subleasing
- Gains/losses on sale and leaseback

### 14. Audit_Notes
**Purpose:** Documentation and references

**Contents:**
- Ind AS 116 key requirements
- Lease identification assessment
- IBR determination methodology
- Lease term judgments
- Discount rate assumptions
- Significant judgments made

---

## 🚀 Quick Start Guide

### Step 1: Create Workbook
1. Open new Google Sheet
2. Extensions > Apps Script
3. Copy `indas116.gs` code
4. Run `createIndAS116Workbook()`
5. Authorize when prompted

### Step 2: Configure Assumptions
1. Go to **Assumptions** sheet
2. Fill in:
   - Entity details
   - Reporting date
   - Default IBR (e.g., 8-10%)
   - Depreciation policy
   - Exemption thresholds

### Step 3: Enter Lease Contracts
1. Go to **Lease_Register**
2. Enter each lease:
   - Office space
   - Vehicles
   - Equipment
   - Warehouses
3. Specify:
   - Lease term
   - Payment amount
   - IBR (if different from default)

### Step 4: Review Calculations
1. **ROU_Asset_Schedule** - Check initial measurement and depreciation
2. **Lease_Liability_Schedule** - Verify interest calculations
3. **Period_End_Entries** - Review journal entries
4. **Reconciliation** - Confirm balances

### Step 5: Analyze Impact
1. Go to **IGAAP_Comparison**
2. Review balance sheet impact
3. Analyze P&L impact
4. Assess ratio changes

---

## 📐 Key Formulas & Logic

### Lease Identification

```
Is there a lease?
├─ Identified asset? (specific asset, not substitutable)
│  └─ YES → Right to obtain economic benefits?
│     └─ YES → Right to direct use?
│        └─ YES → LEASE
│        └─ NO → Service contract
└─ NO → Service contract
```

### Initial Measurement

**Lease Liability:**
```
PV = PMT × [(1 - (1 + r)^-n) / r]

Where:
- PMT = Periodic payment
- r = IBR per period
- n = Number of periods
```

**ROU Asset:**
```
ROU Asset = Lease Liability
          + Initial direct costs
          + Restoration costs
          - Lease incentives
```

### Subsequent Measurement

**Depreciation:**
```
Annual Depreciation = ROU Asset / Lease Term (years)
Monthly Depreciation = ROU Asset / Lease Term (months)
```

**Interest:**
```
Interest Expense = Opening Liability × IBR × Time Period
```

**Liability Reduction:**
```
Principal Repayment = Payment - Interest Expense
Closing Liability = Opening Liability - Principal Repayment
```

### Current vs Non-Current

```
Current Portion = Sum of principal repayments in next 12 months
Non-Current Portion = Total Liability - Current Portion
```

---

## ⚠️ Known Limitations (See todo.md)

### 1. Simplified Interest Calculation
**Issue:** Uses average balance method instead of true EIR  
**Impact:** Minor variance in interest expense  
**Workaround:** Acceptable for monthly/quarterly reporting

### 2. Current Portion Approximation
**Issue:** Uses simplified formula instead of forward-looking schedule  
**Impact:** May not be exact for irregular payment patterns  
**Workaround:** Manual adjustment for material leases

### 3. Variable Payment Tracking
**Issue:** Basic tracking, doesn't handle complex formulas  
**Impact:** Manual calculation needed for complex variable terms  
**Workaround:** Use Variable_Payments sheet for detailed tracking

---

## 🎓 Ind AS 116 Key Concepts

### Lessee Accounting

**Recognition:**
- ROU asset (asset)
- Lease liability (liability)
- At commencement date

**Measurement:**
- Initial: PV of lease payments
- Subsequent: Cost model (depreciation + impairment)

**Exemptions:**
- Short-term leases (≤12 months)
- Low-value assets (≤$5,000 or ₹5 lakhs)

### Lease Term

**Includes:**
- Non-cancellable period
- Extension options (reasonably certain)
- Termination penalties (if applicable)

**Reassessment Triggers:**
- Significant event
- Change in circumstances
- Lessee controls exercise of option

### Incremental Borrowing Rate (IBR)

**Definition:** Rate lessee would pay to borrow over similar term, with similar security, to obtain asset of similar value

**Determination:**
- Start with risk-free rate
- Add credit spread
- Adjust for lease-specific factors

**Hierarchy:**
1. Rate implicit in lease (if determinable)
2. IBR (if implicit rate not available)

### Lease Modifications

**Types:**
1. **Separate lease** - Additional ROU asset at standalone price
2. **Not separate lease** - Remeasure liability, adjust ROU asset

**Remeasurement:**
- Change in lease term
- Change in purchase option assessment
- Change in residual value guarantee
- Change in index/rate

---

## 📋 Compliance Checklist

### Lease Identification
- [ ] All contracts reviewed for embedded leases
- [ ] Identified asset assessment documented
- [ ] Right to control use assessed
- [ ] Service components separated (if practical expedient not used)

### Initial Measurement
- [ ] Lease term determined (including options)
- [ ] Lease payments identified
- [ ] IBR determined and documented
- [ ] Initial direct costs identified
- [ ] Lease incentives accounted for

### Subsequent Measurement
- [ ] Depreciation calculated correctly
- [ ] Interest expense calculated using EIR
- [ ] Current/non-current split correct
- [ ] Modifications identified and accounted for

### Exemptions
- [ ] Short-term leases identified
- [ ] Low-value assets identified
- [ ] Exemption policy documented
- [ ] Exemption expense tracked

### Disclosure
- [ ] Maturity analysis prepared
- [ ] Expense breakdown complete
- [ ] Cash flow information disclosed
- [ ] Significant judgments explained

---

## 🔍 Audit Procedures

### Lease Identification
1. Obtain lease register
2. Review contracts for embedded leases
3. Test lease identification criteria
4. Verify exemptions applied correctly

### Initial Measurement
1. Recalculate lease liability (PV)
2. Verify IBR used
3. Check lease term assessment
4. Verify ROU asset calculation
5. Test initial direct costs

### Subsequent Measurement
1. Recalculate depreciation
2. Test interest expense calculation
3. Verify current/non-current split
4. Check for impairment indicators
5. Test modifications accounting

### Reconciliation
1. Agree opening balances to prior year
2. Test additions (new leases)
3. Verify depreciation and interest
4. Check payments to bank statements
5. Agree closing balances to GL

---

## 💡 Best Practices

### Lease Management
1. Maintain centralized lease register
2. Set up lease renewal reminders
3. Track critical dates (commencement, expiry, options)
4. Document all lease modifications
5. Review lease terms annually

### IBR Determination
1. Document IBR methodology
2. Update IBR periodically
3. Use consistent approach across leases
4. Consider lease-specific factors
5. Obtain treasury/finance input

### System & Controls
1. Implement lease management system
2. Segregate duties (contract vs accounting)
3. Regular reconciliation to GL
4. Management review of assumptions
5. Audit trail for all changes

### Transition Planning
1. Identify all leases early
2. Gather lease documentation
3. Determine IBR for each lease
4. Calculate transition impact
5. Update systems and processes

---

## 📊 Common Lease Scenarios

### Office Space Lease
- **Term:** 3-5 years
- **Payments:** Monthly rent
- **Variables:** Maintenance, utilities (separate)
- **Options:** Renewal option common
- **IBR:** 8-10%

### Vehicle Lease
- **Term:** 3-4 years
- **Payments:** Monthly EMI
- **Variables:** Fuel, maintenance (separate)
- **Options:** Purchase option at end
- **IBR:** 9-12%

### Equipment Lease
- **Term:** 5-7 years
- **Payments:** Quarterly/Annual
- **Variables:** Usage-based charges
- **Options:** Return or purchase
- **IBR:** 10-14%

### Warehouse Lease
- **Term:** 5-10 years
- **Payments:** Monthly/Quarterly
- **Variables:** Property tax, insurance
- **Options:** Extension options
- **IBR:** 8-11%

---

## 📚 References

### Standards
- Ind AS 116 - Leases
- Ind AS 36 - Impairment of Assets (for ROU assets)
- Ind AS 8 - Accounting Policies, Changes in Estimates

### Guidance
- ICAI Implementation Guide on Ind AS 116
- IASB Educational Material on IFRS 16
- Transition Resource Group discussions

### Useful Links
- [MCA Ind AS Portal](https://www.mca.gov.in/)
- [ICAI Resources](https://www.icai.org/)
- [IFRS Foundation](https://www.ifrs.org/)

---

## 🔄 Updates & Maintenance

### Monthly Tasks
- Enter new lease contracts
- Record lease payments
- Accrue interest expense
- Calculate depreciation
- Update payment schedule

### Quarterly Tasks
- Review lease modifications
- Reassess lease terms
- Update IBR if needed
- Prepare journal entries
- Reconcile to GL

### Annual Tasks
- Review all lease terms
- Assess extension/termination options
- Test ROU assets for impairment
- Update disclosure schedules
- Archive working papers

---

## ⚙️ Customization Tips

### Adding New Lease Types
1. Add row in Lease_Register
2. Specify asset type and terms
3. Calculations auto-populate
4. Review and adjust if needed

### Modifying IBR
1. Update in Assumptions (default)
2. Or override in Lease_Register (specific lease)
3. Remeasure liability if mid-term change
4. Document reason for change

### Custom Reports
1. Pivot table on Lease_Register by asset type
2. Chart lease liability maturity profile
3. Analyze lease expense trends
4. Compare actual vs budget

---

## 🐛 Troubleshooting

**Issue:** ROU asset not calculating  
**Fix:** Check lease liability is calculated first, verify initial costs entered

**Issue:** Interest expense seems high  
**Fix:** Verify IBR is entered as decimal (e.g., 0.10 for 10%), check opening liability

**Issue:** Current portion incorrect  
**Fix:** Ensure payment schedule is complete, check date calculations

**Issue:** Depreciation not matching  
**Fix:** Verify lease term in months, check ROU asset opening balance

**Issue:** Reconciliation not balancing  
**Fix:** Check for missing entries, verify all movements captured

---

## 📄 License

Open source - Free to use, modify, and distribute for audit and compliance purposes.

---

## 📝 Version History

**Version 1.0 (November 2024)**
- Initial release
- 14 interconnected sheets
- ROU asset and lease liability schedules
- Auto-generated journal entries
- IGAAP comparison analysis
- Comprehensive reconciliation

---

**For complex lease arrangements, always consult with accounting experts and auditors.**

*Simplifying Ind AS 116 compliance, one lease at a time.*
