# Ind AS 115 - Revenue Recognition Workbook

> Comprehensive audit working papers for Revenue from Contracts with Customers under Indian Accounting Standard 115

## 📋 Overview

This Google Sheets workbook automates the creation of audit working papers for Ind AS 115 (Revenue from Contracts with Customers), implementing the 5-step revenue recognition model.

**Replaces:** AS 9 (Sale of Goods), AS 7 (Construction Contracts), AS 28 (Construction Contracts - revised)  
**Effective:** For periods beginning on or after April 1, 2018  
**Key Change:** Principle-based, 5-step model for all revenue contracts

---

## 🎯 What Does This Workbook Do?

### Core Functions
1. **5-Step Model Implementation** - Structured approach to revenue recognition
2. **Contract Identification** - Assess if contract exists per Ind AS 115
3. **Performance Obligations** - Identify and separate distinct goods/services
4. **Transaction Price** - Determine and allocate consideration
5. **Revenue Recognition** - Recognize revenue when/as obligations satisfied
6. **Contract Assets/Liabilities** - Track unbilled revenue and deferred revenue
7. **Disclosure Schedules** - Prepare required disclosures

---

## 📊 The 5-Step Model

### Step 1: Identify the Contract
- Agreement between parties
- Commercial substance
- Rights and obligations identifiable
- Payment terms clear
- Probable collection

### Step 2: Identify Performance Obligations
- Promises to transfer goods/services
- Distinct (capable of being distinct + distinct in context)
- Separate or bundle

### Step 3: Determine Transaction Price
- Fixed consideration
- Variable consideration (estimate)
- Significant financing component
- Non-cash consideration
- Consideration payable to customer

### Step 4: Allocate Transaction Price
- Standalone selling price (SSP)
- Relative SSP method
- Residual approach (if applicable)

### Step 5: Recognize Revenue
- Over time (continuous transfer)
- Point in time (control transfers)
- Measure progress (output/input method)

---

## 📊 Sheets Included

### 1. Cover Sheet
- Workbook overview
- Entity details
- Reporting period
- Revenue streams covered

### 2. Assumptions
**Purpose:** Central configuration sheet

**Key Inputs:**
- Entity name and industry
- Reporting period
- Revenue recognition policies
- Significant judgments
- Practical expedients applied

**Practical Expedients:**
- Significant financing component (<12 months)
- Incremental costs of obtaining contract
- Right to invoice (output method)

### 3. Contract_Register
**Purpose:** Master list of all customer contracts

**Columns:**
- Contract ID
- Customer name
- Contract date
- Contract value
- Payment terms
- Contract period
- Status (Active/Completed/Terminated)
- Revenue recognized to date
- Remaining performance obligations

**Contract Modifications:**
- Separate contract
- Termination + new contract
- Cumulative catch-up adjustment

### 4. Step1_Contract_Assessment
**Purpose:** Assess if contract exists per Ind AS 115

**Criteria Checklist:**
- [ ] Parties committed (approved & enforceable)
- [ ] Rights identifiable
- [ ] Payment terms identifiable
- [ ] Commercial substance
- [ ] Collection probable

**Conclusion:** Contract exists (Yes/No)

**If No:** Account as received (liability) until criteria met

### 5. Step2_Performance_Obligations
**Purpose:** Identify distinct performance obligations

**Columns:**
- Contract ID
- Performance obligation description
- Distinct? (Yes/No)
- Reasoning
- Timing (Point in time / Over time)
- Transfer of control indicators

**Distinct Criteria:**
- Customer can benefit from good/service
- Separately identifiable from other promises

**Common Scenarios:**
- Product + installation (separate or combined?)
- Product + warranty (standard vs service-type)
- License + implementation (separate or combined?)

### 6. Step3_Transaction_Price
**Purpose:** Determine total transaction price

**Components:**
- Fixed consideration
- Variable consideration (estimate)
- Constraint on variable consideration
- Significant financing component
- Non-cash consideration (fair value)
- Consideration payable to customer (reduce price)

**Variable Consideration Methods:**
- Expected value (probability-weighted)
- Most likely amount

**Constraint:** Include only to extent highly probable no reversal

**Financing Component:**
```
If payment timing ≠ transfer timing AND >12 months
→ Adjust for time value of money
```

### 7. Step4_Price_Allocation
**Purpose:** Allocate transaction price to performance obligations

**Method:** Relative Standalone Selling Price (SSP)

**Formula:**
```
Allocated Amount = Transaction Price × (SSP of PO / Total SSP)
```

**SSP Determination:**
- Observable price (if sold separately)
- Adjusted market assessment
- Expected cost plus margin
- Residual approach (if highly variable/uncertain)

**Columns:**
- Performance obligation
- Standalone selling price (SSP)
- % of total
- Allocated amount
- Discount allocation

### 8. Step5_Revenue_Recognition
**Purpose:** Recognize revenue when/as obligations satisfied

**Over Time Recognition:**
- Customer receives & consumes benefits
- Customer controls asset as created
- No alternative use + enforceable right to payment

**Progress Measurement:**
- Output methods (units delivered, milestones)
- Input methods (costs incurred, labor hours)

**Point in Time Recognition:**
- Control transfers (indicators):
  - Present right to payment
  - Legal title transferred
  - Physical possession transferred
  - Customer accepted
  - Risks & rewards transferred

**Columns:**
- Period
- Performance obligation
- % complete
- Revenue recognized (cumulative)
- Revenue recognized (period)

### 9. Contract_Assets_Liabilities
**Purpose:** Track contract balances

**Contract Asset:**
- Right to consideration (conditional on something other than time)
- Example: Unbilled revenue, work-in-progress

**Contract Liability:**
- Obligation to transfer goods/services
- Example: Advance from customer, deferred revenue

**Receivable:**
- Unconditional right to consideration
- Only passage of time required

**Movements:**
- Opening balance
- Revenue recognized
- Cash received
- Closing balance

### 10. Variable_Consideration
**Purpose:** Track variable consideration estimates

**Types:**
- Discounts
- Rebates
- Refunds
- Credits
- Price concessions
- Performance bonuses
- Penalties

**Estimation:**
- Expected value method (range of outcomes)
- Most likely amount (binary outcome)

**Constraint:**
- Include only if highly probable no significant reversal
- Reassess each reporting period

### 11. Significant_Judgments
**Purpose:** Document key judgments and estimates

**Common Judgments:**
- Identification of performance obligations
- Determination of SSP
- Variable consideration estimates
- Progress measurement method
- Principal vs agent assessment
- License: Right to use vs right to access

**Documentation:**
- Judgment area
- Facts and circumstances
- Alternatives considered
- Conclusion and rationale
- Impact on revenue

### 12. Principal_vs_Agent
**Purpose:** Assess if entity is principal or agent

**Principal (Gross Revenue):**
- Controls good/service before transfer
- Primary responsibility
- Inventory risk
- Pricing discretion

**Agent (Net Commission):**
- Arranges for another party to provide
- No control before transfer
- Fixed fee or commission

**Indicators:**
- Who has inventory risk?
- Who sets price?
- Who has credit risk?
- Who is primary obligor?

### 13. Licensing_Revenue
**Purpose:** Track revenue from licenses

**License Types:**

**Right to Use (Point in Time):**
- Functional intellectual property
- Standalone functionality
- Revenue at grant date

**Right to Access (Over Time):**
- Symbolic intellectual property
- Entity's ongoing activities affect value
- Revenue over license period

**Examples:**
- Software license (functional) → Right to use
- Brand license (symbolic) → Right to access

### 14. Warranty_Analysis
**Purpose:** Separate warranty types

**Assurance Warranty:**
- Product works as specified
- Not a separate performance obligation
- Provision under Ind AS 37

**Service Warranty:**
- Additional service beyond assurance
- Separate performance obligation
- Allocate transaction price

**Indicators of Service Warranty:**
- Separately priced
- Extended period
- Additional services included

### 15. Disclosure_Schedules
**Purpose:** Prepare required disclosures

**Disclosures Required:**
- Disaggregated revenue
- Contract balances (opening, closing, movements)
- Performance obligations (satisfied/unsatisfied)
- Transaction price allocated to remaining obligations
- Significant judgments
- Practical expedients used
- Assets recognized from costs to obtain/fulfill

### 16. Audit_Notes
**Purpose:** Documentation and references

**Contents:**
- Ind AS 115 key requirements
- 5-step model summary
- Industry-specific guidance
- Significant judgments
- Changes from prior standards
- Audit procedures performed

---

## 🚀 Quick Start Guide

### Step 1: Create Workbook
1. Open new Google Sheet
2. Extensions > Apps Script
3. Copy `indas115.gs` code
4. Run `createIndAS115Workbook()`
5. Authorize when prompted

### Step 2: Configure
1. Go to **Assumptions** sheet
2. Enter entity details
3. Document revenue policies
4. List practical expedients used

### Step 3: Enter Contracts
1. Go to **Contract_Register**
2. Enter each customer contract
3. Specify contract terms and value

### Step 4: Apply 5-Step Model
1. **Step1_Contract_Assessment** - Verify contract exists
2. **Step2_Performance_Obligations** - Identify distinct obligations
3. **Step3_Transaction_Price** - Determine total price
4. **Step4_Price_Allocation** - Allocate to obligations
5. **Step5_Revenue_Recognition** - Recognize revenue

### Step 5: Track Balances
1. **Contract_Assets_Liabilities** - Monitor contract balances
2. **Variable_Consideration** - Update estimates
3. **Disclosure_Schedules** - Prepare disclosures

---

## 📐 Key Formulas & Logic

### Distinct Performance Obligation Test

```
Is promise distinct?
├─ Capable of being distinct?
│  └─ Can customer benefit from it alone or with readily available resources?
│     ├─ YES → Distinct in context?
│     │  └─ Is it separately identifiable?
│     │     ├─ YES → DISTINCT (separate PO)
│     │     └─ NO → NOT DISTINCT (bundle)
│     └─ NO → NOT DISTINCT (bundle)
```

### Transaction Price Allocation

```
Allocated Amount = Transaction Price × (SSP of PO / Σ SSP of all POs)

Example:
Product SSP: ₹80,000
Service SSP: ₹20,000
Total SSP: ₹1,00,000
Contract Price: ₹90,000

Product allocation: ₹90,000 × (80,000/100,000) = ₹72,000
Service allocation: ₹90,000 × (20,000/100,000) = ₹18,000
```

### Over Time Revenue Recognition

```
Revenue = Transaction Price × % Complete

% Complete (Input method) = Costs Incurred / Total Expected Costs
% Complete (Output method) = Units Delivered / Total Units
```

### Contract Asset/Liability

```
Contract Asset = Revenue Recognized - Cash Received (if positive)
Contract Liability = Cash Received - Revenue Recognized (if positive)
Receivable = Unconditional right to payment
```

---

## 🎓 Ind AS 115 Key Concepts

### Control Transfer

**Control = Ability to direct use and obtain benefits**

**Indicators:**
- Present right to payment
- Legal title
- Physical possession
- Risks & rewards
- Acceptance

### Over Time vs Point in Time

**Over Time (if any criterion met):**
1. Customer receives & consumes benefits
2. Customer controls asset as created
3. No alternative use + enforceable right to payment

**Point in Time (default):**
- Control transfers at a point
- Assess indicators

### Variable Consideration Constraint

**Include in transaction price only if:**
- Highly probable
- No significant reversal when uncertainty resolved

**Factors Increasing Reversal Risk:**
- Susceptible to factors outside control
- Long time until resolution
- Limited experience
- Broad range of outcomes
- Large number/range of outcomes

### Principal vs Agent

**Principal:**
- Revenue = Gross amount
- Controls good/service before transfer

**Agent:**
- Revenue = Net commission
- Arranges for another party

---

## 📋 Compliance Checklist

### Contract Assessment
- [ ] All contracts identified
- [ ] Contract criteria assessed
- [ ] Modifications identified and accounted for

### Performance Obligations
- [ ] All promises identified
- [ ] Distinct assessment performed
- [ ] Bundling decisions documented

### Transaction Price
- [ ] Fixed consideration identified
- [ ] Variable consideration estimated
- [ ] Constraint applied
- [ ] Financing component assessed
- [ ] Non-cash consideration measured

### Allocation
- [ ] SSP determined for each PO
- [ ] Allocation calculated
- [ ] Discounts allocated appropriately

### Recognition
- [ ] Timing determined (over time/point in time)
- [ ] Progress measurement method selected
- [ ] Revenue recognized appropriately

### Disclosure
- [ ] Disaggregated revenue disclosed
- [ ] Contract balances disclosed
- [ ] Remaining performance obligations disclosed
- [ ] Significant judgments explained

---

## 🔍 Audit Procedures

### Contract Review
1. Obtain contract register
2. Review sample contracts
3. Assess contract criteria
4. Test modifications

### Performance Obligations
1. Identify all promises
2. Test distinct assessment
3. Review bundling decisions
4. Challenge management judgment

### Transaction Price
1. Verify fixed consideration
2. Test variable consideration estimates
3. Review constraint application
4. Recalculate financing component

### Allocation
1. Test SSP determination
2. Recalculate allocation
3. Verify discount allocation

### Recognition
1. Test timing (over time/point in time)
2. Recalculate progress
3. Verify revenue recognized
4. Test contract balances

---

## 💡 Best Practices

### Contract Management
1. Maintain complete contract register
2. Track modifications systematically
3. Document all judgments
4. Regular management review

### Revenue Recognition
1. Consistent application of policies
2. Regular reassessment of estimates
3. Timely recognition
4. Proper cutoff procedures

### Documentation
1. Document 5-step analysis for each contract type
2. Maintain SSP support
3. Keep variable consideration workings
4. Update judgments regularly

### Controls
1. Contract approval process
2. Revenue recognition review
3. Estimate review and approval
4. Reconciliation to GL
5. Disclosure review

---

## 📚 References

### Standards
- Ind AS 115 - Revenue from Contracts with Customers
- Ind AS 37 - Provisions (for warranties)
- Ind AS 32 & 109 - Financial Instruments

### Guidance
- ICAI Implementation Guide on Ind AS 115
- IASB Educational Material on IFRS 15
- Industry-specific guidance

### Useful Links
- [MCA Ind AS Portal](https://www.mca.gov.in/)
- [ICAI Resources](https://www.icai.org/)
- [IFRS Foundation](https://www.ifrs.org/)

---

## 📄 License

Open source - Free to use, modify, and distribute for audit and compliance purposes.

---

## 📝 Version History

**Version 1.0 (November 2024)**
- Initial release
- Complete 5-step model implementation
- 16 interconnected sheets
- Comprehensive documentation templates
- Industry-agnostic framework

---

**For complex revenue arrangements, always consult with accounting experts and auditors.**

*Simplifying Ind AS 115 compliance, one contract at a time.*
