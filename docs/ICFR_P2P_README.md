# ICFR Procure-to-Pay (P2P) Testing Workpaper

> Internal Control over Financial Reporting (ICFR) testing template for Procure-to-Pay process

## 📋 Overview

This Google Sheets workbook creates a comprehensive ICFR testing workpaper for the Procure-to-Pay (P2P) cycle, covering control identification, design testing, and operating effectiveness testing.

**Process Scope:** Purchase Requisition → Purchase Order → Goods Receipt → Invoice → Payment  
**Control Objectives:** Completeness, Accuracy, Validity, Authorization, Cutoff  
**Testing Approach:** Design + Operating Effectiveness

---

## 🎯 What Does This Workbook Do?

### Core Functions
1. **Process Documentation** - Map P2P process flow
2. **Risk Assessment** - Identify risks and control objectives
3. **Control Matrix** - Document key controls
4. **Design Testing** - Test control design adequacy
5. **Operating Effectiveness** - Test controls in operation
6. **Deficiency Tracking** - Log and track control deficiencies
7. **Management Action Plans** - Track remediation
8. **Testing Conclusions** - Document overall assessment

---

## 📊 Sheets Included

### 1. Cover Sheet
- ICFR testing engagement details
- Entity name
- Period under review
- Testing team
- Review and approval

### 2. Process_Flow
**Purpose:** Document P2P process

**Process Steps:**
1. **Purchase Requisition**
   - Requester identifies need
   - Creates requisition in system
   - Submits for approval

2. **Purchase Order**
   - Approved requisition → PO
   - PO sent to vendor
   - PO recorded in system

3. **Goods Receipt**
   - Goods received at warehouse
   - Inspected for quality/quantity
   - GRN created in system

4. **Invoice Processing**
   - Vendor invoice received
   - 3-way match (PO-GRN-Invoice)
   - Invoice approved for payment

5. **Payment**
   - Payment batch created
   - Approved by authorized signatory
   - Payment executed
   - Accounting entry posted

**Documentation:**
- Process narrative
- Flowchart
- System screenshots
- Key documents

### 3. Risk_Control_Matrix
**Purpose:** Map risks to controls

**Risk Categories:**
- **Completeness** - All transactions recorded
- **Accuracy** - Transactions recorded correctly
- **Validity** - Only valid transactions recorded
- **Authorization** - Proper approval obtained
- **Cutoff** - Transactions in correct period

**Format:**
| Risk | Control Objective | Key Control | Control Type | Frequency |
|------|-------------------|-------------|--------------|-----------|
| Unauthorized purchases | Validity | PO approval | Preventive | Each transaction |
| Incorrect pricing | Accuracy | 3-way match | Detective | Each invoice |
| Duplicate payments | Validity | Duplicate check | Preventive | Each payment |

**Control Types:**
- Preventive (stops error before it occurs)
- Detective (identifies error after it occurs)
- Manual
- Automated
- IT-dependent manual (ITDM)

### 4. Control_Catalog
**Purpose:** Detailed control documentation

**For Each Control:**
- Control ID
- Control description
- Control owner
- Control frequency
- Control type
- Evidence of performance
- Compensating controls (if any)

**Example Controls:**

**C001 - Purchase Requisition Approval**
- Description: All purchase requisitions >₹50,000 require manager approval
- Owner: Department Manager
- Frequency: Each transaction
- Type: Preventive, Manual
- Evidence: Approved requisition in system

**C002 - PO-GRN-Invoice 3-Way Match**
- Description: System performs automatic 3-way match; variances >5% require investigation
- Owner: Accounts Payable team
- Frequency: Each invoice
- Type: Detective, Automated
- Evidence: Match report, variance investigation log

**C003 - Vendor Master Maintenance**
- Description: New vendors require approval; periodic review of vendor master
- Owner: Procurement Manager
- Frequency: Each new vendor; Quarterly review
- Type: Preventive, Manual
- Evidence: Vendor approval form, review report

**C004 - Payment Approval**
- Description: Payments >₹1 lakh require dual approval
- Owner: Finance Manager + CFO
- Frequency: Each payment batch
- Type: Preventive, Manual
- Evidence: Approved payment batch

**C005 - Segregation of Duties**
- Description: Requisitioner ≠ Approver ≠ Receiver ≠ Payment processor
- Owner: IT Administrator
- Frequency: Continuous (system enforced)
- Type: Preventive, Automated
- Evidence: User access matrix

### 5. Design_Testing
**Purpose:** Test if controls are designed adequately

**Testing Approach:**
- Walkthrough
- Inquiry
- Observation
- Inspection of documentation

**Testing Columns:**
- Control ID
- Design test procedure
- Expected result
- Actual result
- Design effective? (Y/N)
- Comments

**Example:**

**Control C001 - PO Approval**
- Procedure: Review approval matrix; inspect sample approved PO
- Expected: All POs >₹50,000 have manager approval
- Actual: Approval matrix in place; sample PO properly approved
- Effective: Yes

**Control C002 - 3-Way Match**
- Procedure: Review system configuration; observe match process
- Expected: System blocks payment if variance >5% without investigation
- Actual: System configured correctly; observed blocking
- Effective: Yes

### 6. OE_Testing_Plan
**Purpose:** Plan operating effectiveness testing

**Sample Selection:**
- Population: All transactions in testing period
- Sample size: Based on control frequency and risk
- Selection method: Random, systematic, or judgmental

**Sample Size Guidelines:**
| Control Frequency | Minimum Sample |
|-------------------|----------------|
| Daily | 25 |
| Weekly | 15 |
| Monthly | 5 |
| Quarterly | 2 |
| Annual | 1 |

**Testing Period:**
- Typically 9-12 months
- Cover all quarters
- Include month-end/year-end

### 7. OE_Testing_Results
**Purpose:** Document operating effectiveness testing

**Testing Columns:**
- Control ID
- Sample item #
- Transaction date
- Transaction reference
- Test procedure
- Expected result
- Actual result
- Exception? (Y/N)
- Comments

**Example:**

**Control C001 - PO Approval (Sample of 25)**
- Item 1: PO#12345, 15-Apr-2024, ₹75,000
  - Procedure: Verify manager approval
  - Expected: Approved by manager
  - Actual: Approved by manager on 15-Apr-2024
  - Exception: No

- Item 2: PO#12567, 22-May-2024, ₹1,20,000
  - Procedure: Verify manager approval
  - Expected: Approved by manager
  - Actual: No approval found
  - Exception: **YES** - Unapproved PO

**Exception Rate:**
```
Exception Rate = Exceptions / Sample Size
Tolerable Rate: Typically 5-10%
```

### 8. Deficiency_Log
**Purpose:** Track control deficiencies

**Deficiency Classification:**

**Design Deficiency:**
- Control doesn't address risk
- Control not operating at right frequency
- Control owner not appropriate

**Operating Deficiency:**
- Control not performed
- Control performed incorrectly
- Control performed late

**Severity:**
- **Significant Deficiency** - Important enough to merit attention
- **Material Weakness** - Reasonable possibility of material misstatement

**Deficiency Columns:**
- Deficiency ID
- Control ID
- Description
- Root cause
- Impact
- Severity (SD/MW)
- Status (Open/Closed)

**Example:**

**DEF-001**
- Control: C001 - PO Approval
- Description: 1 out of 25 POs not approved
- Root cause: System bypass by user
- Impact: Unauthorized purchase of ₹1.2 lakhs
- Severity: Significant Deficiency
- Status: Open

### 9. Management_Action_Plans
**Purpose:** Track remediation of deficiencies

**For Each Deficiency:**
- Deficiency ID
- Management action plan
- Responsible person
- Target completion date
- Status
- Evidence of remediation
- Retest date
- Retest result

**Example:**

**DEF-001 Action Plan**
- Action: Remove system bypass capability; retrain users
- Responsible: IT Manager + Procurement Manager
- Target: 30-Jun-2024
- Status: In Progress
- Evidence: System access log, training records
- Retest: 31-Jul-2024
- Result: [To be updated]

### 10. Automated_Controls_Testing
**Purpose:** Test IT-dependent controls

**IT General Controls (ITGC):**
- Access controls
- Change management
- Backup & recovery
- IT operations

**Application Controls:**
- Input controls (validation, authorization)
- Processing controls (calculations, matching)
- Output controls (reconciliation, review)

**Testing:**
- Obtain IT audit report
- Review ITGC testing
- Test application controls
- Assess dependency

**If ITGC deficiencies:**
- Automated controls may not be reliable
- Increase substantive testing
- Consider compensating controls

### 11. Compensating_Controls
**Purpose:** Document compensating controls

**When Needed:**
- Primary control has deficiency
- Primary control not operating
- Cost-benefit of primary control

**Example:**

**Primary Control:** System-enforced segregation of duties  
**Deficiency:** Small team, cannot fully segregate  
**Compensating Control:** Monthly management review of all transactions >₹1 lakh

**Effectiveness Assessment:**
- Does compensating control address same risk?
- Is it operating effectively?
- Is it sufficient to reduce risk to acceptable level?

### 12. Testing_Conclusions
**Purpose:** Overall ICFR assessment

**Summary:**
- Total controls tested
- Design effective: X out of Y
- Operating effective: X out of Y
- Deficiencies identified: X
  - Significant deficiencies: X
  - Material weaknesses: X
- Overall conclusion

**Conclusion Template:**
```
ICFR Testing - Procure-to-Pay Process
Period: [Date range]

Controls Tested: [Number]
- Design Effective: [Number] ([%])
- Operating Effective: [Number] ([%])

Deficiencies:
- Significant Deficiencies: [Number]
- Material Weaknesses: [Number]

Overall Conclusion:
[Effective / Effective with deficiencies / Ineffective]

Key Findings:
1. [Finding 1]
2. [Finding 2]
3. [Finding 3]

Management Actions:
[Summary of action plans]

Signed: _______________
Date: _______________
```

### 13. Audit_Notes
**Purpose:** Reference and guidance

**Contents:**
- ICFR framework (COSO, SOX)
- P2P process best practices
- Common control deficiencies
- Testing methodologies
- Deficiency classification guidance

---

## 🚀 Quick Start Guide

### Step 1: Create Workbook
1. Open new Google Sheet
2. Extensions > Apps Script
3. Copy `ifc_p2p.gs` code
4. Run `createICFRP2PWorkpaper()`
5. Authorize when prompted

### Step 2: Understand Process
1. Review **Process_Flow** sheet
2. Conduct walkthrough with process owners
3. Update process documentation

### Step 3: Identify Controls
1. Review **Risk_Control_Matrix**
2. Document controls in **Control_Catalog**
3. Obtain control evidence

### Step 4: Test Design
1. Follow **Design_Testing** procedures
2. Document results
3. Identify design deficiencies

### Step 5: Test Operating Effectiveness
1. Plan testing in **OE_Testing_Plan**
2. Execute tests in **OE_Testing_Results**
3. Calculate exception rates

### Step 6: Document Deficiencies
1. Log deficiencies in **Deficiency_Log**
2. Obtain management action plans
3. Track remediation

### Step 7: Conclude
1. Prepare **Testing_Conclusions**
2. Communicate to management
3. Follow up on action plans

---

## 📋 ICFR Testing Checklist

### Planning
- [ ] Understand P2P process
- [ ] Identify key controls
- [ ] Assess risks
- [ ] Determine sample sizes
- [ ] Plan testing timeline

### Design Testing
- [ ] Walkthrough performed
- [ ] Controls documented
- [ ] Design effectiveness assessed
- [ ] Design deficiencies identified

### Operating Effectiveness Testing
- [ ] Samples selected
- [ ] Testing performed
- [ ] Exceptions documented
- [ ] Exception rates calculated
- [ ] Operating deficiencies identified

### Deficiency Management
- [ ] Deficiencies classified
- [ ] Root causes identified
- [ ] Management action plans obtained
- [ ] Remediation tracked
- [ ] Retesting planned

### Conclusion
- [ ] Overall assessment documented
- [ ] Findings communicated
- [ ] Report prepared
- [ ] Follow-up scheduled

---

## 💡 Best Practices

### Control Identification
1. **Focus on key controls** - Don't test every control
2. **Risk-based approach** - Prioritize high-risk areas
3. **Entity-level controls** - Don't forget tone at the top
4. **IT controls** - Test ITGC for automated controls

### Testing
1. **Independent testing** - Tester ≠ Control owner
2. **Sufficient samples** - Follow statistical guidelines
3. **Random selection** - Avoid bias
4. **Document thoroughly** - Evidence is key

### Deficiency Assessment
1. **Consider magnitude** - How big is the impact?
2. **Consider likelihood** - How likely to occur?
3. **Consider compensating controls** - Do they mitigate?
4. **Aggregate deficiencies** - Multiple small = one big?

### Communication
1. **Timely reporting** - Don't wait until year-end
2. **Clear descriptions** - Specific, not vague
3. **Constructive tone** - Focus on improvement
4. **Follow up** - Track action plans

---

## 🔍 Common P2P Control Deficiencies

### Authorization
- Unapproved purchase orders
- Approval limits not enforced
- Approver = Requester (SoD violation)

### Accuracy
- 3-way match not performed
- Price variances not investigated
- Quantity discrepancies not resolved

### Validity
- Duplicate payments
- Payments to invalid vendors
- Fictitious vendors

### Completeness
- Goods received not recorded
- Invoices not recorded timely
- Accruals not complete

### Cutoff
- Goods received in wrong period
- Invoices recorded in wrong period
- Payments in wrong period

---

## 📚 References

### Frameworks
- COSO Internal Control Framework
- SOX Section 404 (if applicable)
- PCAOB AS 2201 (Auditing Standards)

### Guidance
- ICAI Guidance on ICFR
- COSO Guidance on Monitoring
- Industry best practices

---

## 📄 License

Open source - Free to use, modify, and distribute for ICFR testing purposes.

---

## 📝 Version History

**Version 1.0 (November 2025)**
- Initial release
- Complete P2P control framework
- Design and OE testing templates
- Deficiency tracking
- Management action plan tracking

---

**Professional ICFR testing workpaper for efficient P2P control testing.**

*Strengthening internal controls, one process at a time.*
