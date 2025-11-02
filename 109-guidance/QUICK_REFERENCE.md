# IND AS 109 AUDIT BUILDER - QUICK REFERENCE CARD

## 🔄 WORKFLOW DIAGRAM

```
┌─────────────────────────────────────────────────────────────────┐
│                        PHASE 1: SETUP                            │
└─────────────────────────────────────────────────────────────────┘
                                 │
                    ┌────────────▼────────────┐
                    │  Run Apps Script        │
                    │  (One-time setup)       │
                    └────────────┬────────────┘
                                 │
                    ┌────────────▼────────────┐
                    │  11 Sheets Created      │
                    │  with Auto-Formulas     │
                    └─────────────────────────┘

┌─────────────────────────────────────────────────────────────────┐
│                      PHASE 2: DATA ENTRY                         │
└─────────────────────────────────────────────────────────────────┘
                                 │
        ┌────────────────────────┼────────────────────────┐
        │                        │                        │
┌───────▼────────┐    ┌─────────▼─────────┐    ┌────────▼────────┐
│ Input_Variables│    │ Instruments       │    │ Manual Fair     │
│                │    │ Register          │    │ Values (opt.)   │
│ • PD/LGD/EAD   │    │                   │    │                 │
│ • Risk Rates   │    │ • List ALL        │    │ • Override FV   │
│ • Thresholds   │    │   instruments     │    │   if needed     │
└────────────────┘    │ • Classification  │    └─────────────────┘
                      │ • DPD, Rating     │
                      └───────────────────┘

┌─────────────────────────────────────────────────────────────────┐
│                   PHASE 3: AUTO-CALCULATION                      │
└─────────────────────────────────────────────────────────────────┘
                                 │
        ┌────────────────────────┼────────────────────────┐
        │                        │                        │
┌───────▼────────┐    ┌─────────▼─────────┐    ┌────────▼────────┐
│ Classification │    │ Fair Value        │    │ ECL Impairment  │
│ Matrix         │    │ Workings          │    │                 │
│                │    │                   │    │ • Stage 1/2/3   │
│ • Auto-logic   │    │ • FVTPL → P&L     │    │ • PD×LGD×EAD    │
│ • SPPI+BM      │    │ • FVOCI → OCI     │    │ • Provision     │
└────────┬───────┘    └─────────┬─────────┘    └────────┬────────┘
         │                      │                       │
         └──────────────────────┼───────────────────────┘
                                │
                    ┌───────────▼────────────┐
                    │ Amortization Schedule  │
                    │                        │
                    │ • EIR method           │
                    │ • Interest income      │
                    │ • Premium/Discount     │
                    └───────────┬────────────┘

┌─────────────────────────────────────────────────────────────────┐
│                     PHASE 4: JOURNAL ENTRIES                     │
└─────────────────────────────────────────────────────────────────┘
                                 │
                    ┌────────────▼────────────┐
                    │ Period_End_Entries      │
                    │                         │
                    │ JE001: FVTPL Fair Value │
                    │ JE002: FVOCI Fair Value │
                    │ JE003: Interest Income  │
                    │ JE004: ECL Provision    │
                    │ JE005: Amortization     │
                    └────────────┬────────────┘
                                 │
                    ┌────────────▼────────────┐
                    │ Copy to General Ledger  │
                    └─────────────────────────┘

┌─────────────────────────────────────────────────────────────────┐
│                  PHASE 5: VERIFICATION & SIGN-OFF                │
└─────────────────────────────────────────────────────────────────┘
                                 │
        ┌────────────────────────┼────────────────────────┐
        │                        │                        │
┌───────▼────────┐    ┌─────────▼─────────┐    ┌────────▼────────┐
│ Reconciliation │    │ Audit_Notes       │    │ Cover Dashboard │
│                │    │                   │    │                 │
│ • Opening to   │    │ • Control Checks  │    │ • Final Summary │
│   Closing      │    │ • Assertions      │    │ • Export        │
│ • Control =0   │    │ • Sign-off        │    └─────────────────┘
└────────────────┘    └───────────────────┘
```

---

## 🎯 FEATURE MATRIX

| Feature | Functionality | Auto/Manual | Key Formula |
|---------|--------------|-------------|-------------|
| **Classification** | SPPI + Business Model Logic | Auto | `IF(SPPI="Fail","FVTPL",...)` |
| **Fair Value - FVTPL** | Mark-to-Market → P&L | Auto | `FV_End - FV_Open` |
| **Fair Value - FVOCI** | Mark-to-Market → OCI | Auto | Same as FVTPL |
| **ECL Stage 1** | 12-month ECL | Auto | `EAD × PD_Stage1 × LGD` |
| **ECL Stage 2** | Lifetime ECL | Auto | `EAD × PD_Stage2 × LGD` |
| **ECL Stage 3** | Lifetime ECL (NPA) | Auto | `EAD × PD_Stage3 × LGD` |
| **Amortization** | EIR Method | Auto | `Opening × EIR × (Days/365)` |
| **Journal Entries** | 5 Entries Generated | Auto | Links to all sheets |
| **Reconciliation** | Opening to Closing | Auto | `Opening + Movement - Closing` |
| **Control Totals** | 5 Mathematical Checks | Auto | Various |

---

## 📊 INPUT REQUIREMENTS

### CRITICAL Inputs (Must Fill):
1. **Input_Variables Sheet**:
   - ✅ Reporting Date
   - ✅ Previous Reporting Date
   - ✅ Risk-Free Rate
   - ✅ PD for Stage 1, 2, 3
   - ✅ LGD for Secured/Unsecured
   - ✅ DPD Thresholds

2. **Instruments_Register Sheet**:
   - ✅ Instrument ID & Name
   - ✅ Type & Counterparty
   - ✅ Issue & Maturity Dates
   - ✅ Face Value, Coupon, EIR
   - ✅ Opening Balance
   - ✅ Security Type & Rating
   - ✅ DPD (Days Past Due)
   - ✅ SPPI Test Result
   - ✅ Business Model

### OPTIONAL Inputs:
- Manual fair value overrides (Fair_Value_Workings)
- Other adjustments (Amortization_Schedule)
- Audit notes and comments

---

## 🎨 COLOR KEY (At a Glance)

```
┌──────────────────────────┐
│ 🔵 LIGHT BLUE = INPUT    │  Primary user entry cells
├──────────────────────────┤
│ 🟢 LIGHT GREEN = INPUT   │  Adjustment cells
├──────────────────────────┤
│ 🟠 ORANGE = CRITICAL     │  Instruments register data
├──────────────────────────┤
│ 🔷 DARK BLUE = HEADER    │  Sheet titles
├──────────────────────────┤
│ 🟩 GREEN = POSITIVE      │  Gains, Stage 1, AC classification
├──────────────────────────┤
│ 🟨 YELLOW = CAUTION      │  Stage 2, review items
├──────────────────────────┤
│ 🟥 RED = NEGATIVE        │  Losses, Stage 3, errors
└──────────────────────────┘
```

---

## 📋 CONTROL CHECKLIST

Before finalization, verify:

```
┌─────┬────────────────────────────────────────┬──────────┐
│ ☐   │ All input cells filled                 │ Priority │
├─────┼────────────────────────────────────────┼──────────┤
│ ☐   │ Control total #1 = 0 (JE balance)      │   HIGH   │
│ ☐   │ Control total #2 = 0 (Reconciliation)  │   HIGH   │
│ ☐   │ Control total #3 = 0 (Amortization)    │   HIGH   │
│ ☐   │ Control total #4 = 0 (Classification)  │  MEDIUM  │
│ ☐   │ Control total #5 > 0.5 (Stage 3 ECL)   │  MEDIUM  │
├─────┼────────────────────────────────────────┼──────────┤
│ ☐   │ All instruments classified             │   HIGH   │
│ ☐   │ ECL provisions reasonable              │   HIGH   │
│ ☐   │ Fair values supported                  │   HIGH   │
│ ☐   │ SPPI tests documented                  │  MEDIUM  │
│ ☐   │ Business model assessment documented   │  MEDIUM  │
├─────┼────────────────────────────────────────┼──────────┤
│ ☐   │ Journal entries extracted              │   HIGH   │
│ ☐   │ Audit sign-off completed               │   HIGH   │
│ ☐   │ Workbook saved with date               │  MEDIUM  │
└─────┴────────────────────────────────────────┴──────────┘
```

---

## 🔢 FORMULA QUICK REFERENCE

### Classification Logic:
```javascript
=IF(SPPI="Fail", "FVTPL",
  IF(BusinessModel="FVTPL", "FVTPL",
    IF(AND(SPPI="Pass", BusinessModel="Hold to Collect"), "Amortized Cost",
      IF(AND(SPPI="Pass", BusinessModel="Hold to Collect & Sell"), "FVOCI",
        "FVTPL"))))
```

### ECL Calculation:
```javascript
ECL = EAD × PD × LGD

Where:
• EAD = Exposure at Default (Gross Carrying Amount)
• PD = Probability of Default (by stage)
• LGD = Loss Given Default (by security type)
```

### EIR Interest Income:
```javascript
Interest Income = Opening Balance × EIR × (Days / Days_in_Year)
```

### Amortization:
```javascript
Closing Balance = Opening + Interest Income - Cash Received - ECL + Adjustments
```

---

## 📈 TYPICAL VALUES (Industry Benchmarks)

### PD (Probability of Default):
- Stage 1 (Performing): **0.5% - 2%**
- Stage 2 (Underperforming): **10% - 20%**
- Stage 3 (NPA): **80% - 100%**

### LGD (Loss Given Default):
- Secured Assets: **20% - 35%**
- Unsecured Assets: **60% - 75%**
- Sovereign: **5% - 15%**

### DPD Thresholds:
- Stage 1 → Stage 2: **30 days** (rebuttable presumption)
- Stage 2 → Stage 3: **90 days** (RBI NPA norm)

### Risk-Free Rate (India):
- 10Y G-Sec: **6.5% - 7.5%** (as of 2024-25)

### Credit Spreads:
- AAA: **0.25% - 0.50%**
- AA: **0.50% - 1.00%**
- A: **1.00% - 1.50%**
- BBB: **2.00% - 3.00%**

---

## 🚨 COMMON ERRORS & SOLUTIONS

| Error Message | Cause | Solution |
|--------------|-------|----------|
| `#REF!` | Sheet deleted/renamed | Don't rename sheets; re-run script if needed |
| `#VALUE!` | Text in numeric cell | Check input data types |
| `#DIV/0!` | Division by zero | Check for zero balances in denominators |
| `#N/A` | Lookup not found | Verify instrument IDs match across sheets |
| Circular Reference | Formula references itself | Should not occur; check manual edits |

### Control Fails:
- **JE Balance ≠ 0**: Review Period_End_Entries formulas
- **Reconciliation ≠ 0**: Check for missing instruments or broken links
- **Amortization ≠ 0**: Verify all components included

---

## 📞 WHEN TO SEEK PROFESSIONAL HELP

Consult auditors/accountants if:
- ❗ Complex derivatives or embedded derivatives
- ❗ Hedge accounting (cash flow, fair value, net investment)
- ❗ Credit-impaired assets on purchase (POCI)
- ❗ Substantial modification of terms
- ❗ Material ECL models requiring statistical validation
- ❗ Level 3 fair value measurements requiring DCF models
- ❗ Cross-currency instruments
- ❗ Structured products

---

## 🎓 LEARNING PATH

### Beginner (Week 1):
- [ ] Understand Ind AS 109 scope
- [ ] Learn classification principles
- [ ] Review sample data in workbook

### Intermediate (Week 2):
- [ ] Deep dive into ECL model
- [ ] Practice fair value calculations
- [ ] Complete case studies

### Advanced (Week 3):
- [ ] Complex instruments classification
- [ ] ECL model refinement
- [ ] Hedge accounting basics

### Expert (Ongoing):
- [ ] Stay updated on Ind AS amendments
- [ ] Attend ICAI workshops
- [ ] Industry best practices

---

## 🔗 LINKS TO STANDARDS (ICAI)

### Primary Standards:
- **Ind AS 109**: Financial Instruments
  (Classification, Measurement, Impairment, Hedge Accounting)
  
- **Ind AS 107**: Financial Instruments: Disclosures
  (Disclosure requirements in notes to accounts)
  
- **Ind AS 113**: Fair Value Measurement
  (Fair value hierarchy and measurement techniques)
  
- **Ind AS 32**: Financial Instruments: Presentation
  (Equity vs. liability classification)

### Related Guidance:
- **Ind AS 21**: Effects of Changes in Foreign Exchange Rates
- **Ind AS 37**: Provisions, Contingent Liabilities
- **Ind AS 8**: Accounting Policies, Changes in Estimates

---

## 🏆 BEST-IN-CLASS PRACTICES

### Documentation:
✅ Maintain separate file for SPPI test conclusions  
✅ Document business model assessment quarterly  
✅ Keep evidence of fair value sources  
✅ Archive PD/LGD derivation methodology  

### Internal Controls:
✅ Segregation: Data entry ≠ Reviewer  
✅ Monthly ECL provision review  
✅ Quarterly fair value validation  
✅ Annual model back-testing  

### Audit Readiness:
✅ Complete working papers before audit  
✅ All assumptions documented  
✅ Evidence readily available  
✅ Reconciliations prepared  

---

## 💾 FILE MANAGEMENT

### Naming Convention:
```
Ind_AS_109_[Company]_[Period]_[Version]_[Date].xlsx

Examples:
- Ind_AS_109_ABC_Ltd_Q1_FY25_v1.0_20240630.xlsx
- Ind_AS_109_ABC_Ltd_Annual_FY24_Final_20240331.xlsx
```

### Backup Strategy:
- ✅ Daily: Save to Google Drive
- ✅ Weekly: Export PDF copy
- ✅ Monthly: Download Excel backup
- ✅ Quarterly: Archive on secure server

### Version Control:
- v0.1 - v0.9: Draft versions
- v1.0: First complete version
- v1.1, v1.2: Minor updates
- v2.0: Significant changes (e.g., new instruments)
- FINAL: Signed-off version for audit

---

## ⏱️ TIME ESTIMATES

### Initial Setup:
- Script execution: **1-2 minutes**
- Input_Variables: **5-10 minutes**
- Instruments_Register: **30-60 minutes** (depends on count)
- Review & validation: **30 minutes**
- **Total: 1-2 hours** (first time)

### Quarterly Updates:
- Update inputs: **10-15 minutes**
- Update instruments: **20-30 minutes**
- Review calculations: **20 minutes**
- Extract entries: **10 minutes**
- **Total: 1 hour**

### Annual Audit:
- Preparation: **2-3 hours**
- Audit queries response: **3-5 hours**
- Documentation: **2 hours**
- **Total: 7-10 hours**

---

## 📊 OUTPUT DELIVERABLES

From this workbook, you get:

1. **Journal Entries** (Period_End_Entries)
   - Ready to post in general ledger
   - Debit/Credit balanced
   - With narrations and references

2. **Management Reports** (Cover)
   - Executive summary
   - Key metrics dashboard
   - Net financial position

3. **Audit Trail** (Reconciliation)
   - Opening to closing movement
   - Complete trail by classification
   - P&L impact summary

4. **Control Evidence** (Audit_Notes)
   - Mathematical accuracy checks
   - Assertions coverage
   - Sign-off documentation

5. **Compliance Support** (References)
   - Ind AS 109 key provisions
   - Quick reference for queries

---

## 🎯 SUCCESS METRICS

Your implementation is successful when:

✅ **Accuracy**: All control totals pass  
✅ **Completeness**: All instruments classified and measured  
✅ **Compliance**: Ind AS 109 requirements met  
✅ **Efficiency**: Period closure time reduced by 50%+  
✅ **Auditability**: Clear trail, easy to follow  
✅ **Reliability**: Consistent results period-over-period  

---

## 📧 SUPPORT CHANNELS

### For Technical Issues:
- Review troubleshooting section
- Check formula syntax
- Verify data types

### For Accounting Queries:
- Consult References sheet
- Review ICAI guidance
- Engage external auditor

### For Customization:
- Modify Apps Script code
- Adjust formulas
- Add custom validations

---

**END OF QUICK REFERENCE CARD**

Print this document and keep it handy while working with the Ind AS 109 Audit Builder!

**Version**: 1.0  
**Last Updated**: 2024  
**Compatible With**: Google Sheets (Web, Mobile, Desktop)
