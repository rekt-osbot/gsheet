# Fixed Assets Audit Workpaper

> Professional audit working paper template for Property, Plant & Equipment (PPE) verification

## 📋 Overview

This Google Sheets workbook creates a comprehensive audit working paper for fixed assets (Property, Plant & Equipment), covering verification, depreciation testing, additions, disposals, and reconciliation.

**Applicable Standards:**
- **Ind AS 16** - Property, Plant and Equipment
- **AS 10** - Property, Plant and Equipment (IGAAP)

**Purpose:** Audit documentation template (not accounting calculation tool)

---

## 🎯 What Does This Workbook Do?

### Core Functions
1. **Fixed Asset Register** - Complete listing of all PPE
2. **Additions Testing** - Verify new acquisitions
3. **Disposals Testing** - Test retirements and sales
4. **Depreciation Testing** - Recalculate and verify depreciation
5. **Physical Verification** - Document physical inspection
6. **Impairment Assessment** - Test for impairment indicators
7. **Reconciliation** - Opening to closing balance
8. **Audit Conclusions** - Document findings and conclusions

---

## 📊 Sheets Included

### 1. Cover Sheet
- Audit engagement details
- Entity name
- Period under audit
- Audit team members
- Review and approval sign-offs

### 2. Audit_Program
**Purpose:** Audit procedures checklist

**Procedures:**
- [ ] Obtain fixed asset register
- [ ] Agree opening balances to prior year
- [ ] Test additions (sample)
- [ ] Test disposals (sample)
- [ ] Recalculate depreciation
- [ ] Perform physical verification
- [ ] Assess impairment indicators
- [ ] Review capitalization policy
- [ ] Test useful life estimates
- [ ] Verify title/ownership
- [ ] Review insurance coverage
- [ ] Prepare reconciliation
- [ ] Document conclusions

**For Each Procedure:**
- Procedure description
- Performed by
- Date
- Working paper reference
- Conclusion

### 3. Fixed_Asset_Register
**Purpose:** Complete listing from client

**Columns:**
- Asset ID
- Asset description
- Category (Land/Building/Plant/Vehicles/etc.)
- Location
- Date of acquisition
- Original cost
- Accumulated depreciation
- Net book value
- Depreciation rate
- Useful life
- Residual value

**Audit Procedures:**
- Agree totals to trial balance
- Test mathematical accuracy
- Verify classification
- Check for fully depreciated assets still in use

### 4. Additions_Testing
**Purpose:** Test new asset acquisitions

**Sample Selection:**
- All items > materiality threshold
- Random sample of others
- Focus on unusual/complex items

**Testing Columns:**
- Asset description
- Date acquired
- Cost per books
- Invoice reference
- Invoice amount
- Variance
- Capitalization appropriate? (Y/N)
- Comments

**Audit Procedures:**
- Vouch to purchase invoice
- Verify capitalization criteria met
- Check approval
- Verify depreciation commenced
- Test allocation of costs (purchase price + directly attributable costs)

**Capitalization Criteria:**
- Probable future economic benefits
- Cost can be measured reliably
- Meets recognition threshold

### 5. Disposals_Testing
**Purpose:** Test asset retirements and sales

**Testing Columns:**
- Asset description
- Date of disposal
- Original cost
- Accumulated depreciation
- Net book value
- Sale proceeds
- Gain/(Loss) per books
- Gain/(Loss) recalculated
- Variance
- Comments

**Audit Procedures:**
- Vouch to sale invoice/scrap certificate
- Recalculate gain/loss
- Verify approval
- Check removal from register
- Verify accounting entries

**Gain/Loss Calculation:**
```
Sale Proceeds - Net Book Value = Gain/(Loss)
```

### 6. Depreciation_Testing
**Purpose:** Recalculate and verify depreciation

**Testing Columns:**
- Asset description
- Cost
- Useful life
- Depreciation method
- Rate %
- Depreciation per books
- Depreciation recalculated
- Variance
- Comments

**Depreciation Methods:**
- Straight-line method (most common)
- Written down value (WDV) method
- Units of production method

**Formulas:**
```
Straight-line: (Cost - Residual Value) / Useful Life

WDV: Opening NBV × Rate%

Units of Production: (Cost - Residual) × (Units Produced / Total Expected Units)
```

**Audit Procedures:**
- Verify depreciation method appropriate
- Recalculate depreciation
- Check for change in useful life
- Verify residual value estimates
- Test pro-rata depreciation for additions/disposals

### 7. Physical_Verification
**Purpose:** Document physical inspection results

**Verification Columns:**
- Asset ID
- Asset description
- Location per books
- Actual location
- Condition (Good/Fair/Poor)
- In use? (Y/N)
- Verified? (Y/N)
- Discrepancies
- Comments

**Verification Procedures:**
- Select sample for physical inspection
- Verify existence and condition
- Check asset tags/identification
- Assess whether in use
- Identify idle/obsolete assets
- Note any impairment indicators

**Impairment Indicators:**
- Physical damage
- Obsolescence
- Idle/not in use
- Plans to discontinue
- Adverse market changes

### 8. Impairment_Assessment
**Purpose:** Test for impairment

**Indicators of Impairment:**
- Significant decline in market value
- Adverse changes in technology/market/economy
- Physical damage or obsolescence
- Asset idle or plans to dispose
- Worse economic performance than expected

**Testing:**
- Identify assets with indicators
- Obtain management's impairment assessment
- Test recoverable amount calculation
- Verify impairment loss recognized

**Recoverable Amount:**
```
Higher of:
1. Fair Value Less Costs to Sell
2. Value in Use (PV of future cash flows)
```

### 9. Capital_WIP
**Purpose:** Test capital work-in-progress

**Testing Columns:**
- Project description
- Start date
- Expected completion
- Costs incurred to date
- Capitalized? (Y/N)
- Reason if not capitalized
- Comments

**Audit Procedures:**
- Review project status
- Verify costs relate to project
- Check for completed projects not capitalized
- Test for abandoned projects
- Verify no depreciation charged

**Capitalization Trigger:**
- Asset ready for intended use
- All necessary approvals obtained
- Substantially complete

### 10. Title_Verification
**Purpose:** Verify ownership

**Documents to Review:**
- Title deeds (land/buildings)
- Registration certificates (vehicles)
- Purchase agreements
- Lease agreements (if applicable)
- Charge/mortgage documents

**Testing Columns:**
- Asset description
- Document type
- Document reference
- In name of entity? (Y/N)
- Encumbered? (Y/N)
- Comments

### 11. Reconciliation
**Purpose:** Opening to closing balance reconciliation

**Format:**
```
                        Gross Block    Accumulated Dep    Net Block
Opening Balance         ₹ XXX          ₹ XXX              ₹ XXX
Add: Additions              XXX              -                XXX
Less: Disposals            (XXX)           (XXX)            (XXX)
Depreciation for year        -              XXX             (XXX)
Impairment                   -              XXX             (XXX)
                        -------        -------            -------
Closing Balance         ₹ XXX          ₹ XXX              ₹ XXX
                        =======        =======            =======

Agree to:
- Trial Balance
- Financial Statements
- Fixed Asset Register
```

**By Category:**
- Land
- Buildings
- Plant & Machinery
- Furniture & Fixtures
- Vehicles
- Office Equipment
- Computers
- Capital WIP

### 12. Audit_Conclusions
**Purpose:** Document findings and conclusions

**Summary:**
- Total fixed assets tested
- Sample size and selection method
- Exceptions noted
- Adjustments proposed
- Management responses
- Final conclusion

**Conclusion Template:**
```
Based on audit procedures performed:

1. Opening balances agree to prior year: [Yes/No]
2. Additions properly capitalized: [Yes/No/Exceptions noted]
3. Disposals properly accounted: [Yes/No/Exceptions noted]
4. Depreciation calculated correctly: [Yes/No/Exceptions noted]
5. Physical verification satisfactory: [Yes/No/Exceptions noted]
6. No material impairment indicators: [Yes/No/Exceptions noted]
7. Title/ownership verified: [Yes/No/Exceptions noted]

Overall Conclusion: [Satisfactory/Qualified/Adverse]

Adjustments Required: [Yes/No]
If Yes, details: [...]

Signed: _______________
Date: _______________
```

### 13. Exceptions_Log
**Purpose:** Track all exceptions and resolutions

**Columns:**
- Exception #
- Description
- Amount
- Impact (Material/Immaterial)
- Management response
- Auditor assessment
- Resolution
- Status (Open/Closed)

### 14. Audit_Notes
**Purpose:** Reference and guidance

**Contents:**
- Ind AS 16 / AS 10 key requirements
- Capitalization criteria
- Depreciation methods
- Impairment testing
- Common audit issues
- Industry-specific considerations

---

## 🚀 Quick Start Guide

### Step 1: Create Workbook
1. Open new Google Sheet
2. Extensions > Apps Script
3. Copy `far_wp.gs` code
4. Run `createFixedAssetsWorkpaper()`
5. Authorize when prompted

### Step 2: Setup
1. Go to **Cover** sheet
2. Enter audit engagement details
3. Assign team members

### Step 3: Obtain Client Data
1. Request fixed asset register
2. Copy to **Fixed_Asset_Register** sheet
3. Verify totals to trial balance

### Step 4: Perform Audit Procedures
1. Follow **Audit_Program** checklist
2. Test additions in **Additions_Testing**
3. Test disposals in **Disposals_Testing**
4. Recalculate depreciation in **Depreciation_Testing**
5. Document physical verification in **Physical_Verification**

### Step 5: Complete Workpaper
1. Prepare **Reconciliation**
2. Document **Audit_Conclusions**
3. Log any **Exceptions**
4. Obtain review and approval

---

## 📋 Audit Checklist

### Planning
- [ ] Understand entity's fixed asset policies
- [ ] Determine materiality
- [ ] Identify risk areas
- [ ] Plan sample sizes

### Existence & Completeness
- [ ] Physical verification performed
- [ ] Additions tested (existence)
- [ ] Disposals tested (completeness)
- [ ] Capital WIP reviewed

### Valuation
- [ ] Depreciation recalculated
- [ ] Useful lives assessed
- [ ] Impairment tested
- [ ] Residual values reviewed

### Rights & Obligations
- [ ] Title documents verified
- [ ] Leased assets identified
- [ ] Encumbrances disclosed

### Presentation & Disclosure
- [ ] Classification appropriate
- [ ] Reconciliation prepared
- [ ] Disclosures complete
- [ ] Accounting policies disclosed

---

## 💡 Best Practices

### Sampling
1. **Stratify population** - By value, category, location
2. **Risk-based selection** - Focus on high-risk items
3. **Document rationale** - Explain sample selection
4. **Sufficient coverage** - Aim for 60-80% value coverage

### Documentation
1. **Clear cross-references** - Link to supporting documents
2. **Explain variances** - Don't just note, explain
3. **Document judgments** - Especially for estimates
4. **Retain evidence** - Copies of key documents

### Communication
1. **Discuss with management** - Before finalizing conclusions
2. **Timely reporting** - Don't wait until year-end
3. **Clear exceptions** - Specific, not vague
4. **Follow up** - Track resolution of issues

---

## 🔍 Common Audit Issues

### Capitalization Errors
- Revenue expenses capitalized
- Repairs & maintenance capitalized
- Borrowing costs not capitalized (if qualifying asset)

### Depreciation Issues
- Incorrect useful life
- Wrong depreciation method
- Residual value not considered
- Fully depreciated assets not reviewed

### Disposal Issues
- Assets disposed but not removed from register
- Gain/loss not recognized
- Accumulated depreciation not removed

### Impairment
- Indicators present but not assessed
- Recoverable amount not calculated
- Impairment loss not recognized

### Classification
- Capital WIP not transferred when ready
- Leased assets not identified
- Investment property not separated

---

## 📚 References

### Standards
- Ind AS 16 - Property, Plant and Equipment
- AS 10 - Property, Plant and Equipment
- Ind AS 36 - Impairment of Assets
- Ind AS 23 - Borrowing Costs

### Auditing Standards
- SA 500 - Audit Evidence
- SA 501 - Audit Evidence - Specific Considerations
- SA 540 - Auditing Accounting Estimates

---

## 📄 License

Open source - Free to use, modify, and distribute for audit purposes.

---

## 📝 Version History

**Version 1.0 (November 2025)**
- Initial release
- Comprehensive audit program
- Testing templates for additions, disposals, depreciation
- Physical verification documentation
- Reconciliation and conclusions

---

**Professional audit workpaper template for efficient fixed asset audits.**

*Streamlining fixed asset audits, one asset at a time.*
