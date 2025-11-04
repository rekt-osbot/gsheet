# H1 REVIEW - TREASURY & SYSTEMS AUDIT PLAN
## SNVA TravelTech Pvt. Ltd. | FY2025-26

**Audit Period**: April - September 2025  
**Fieldwork Period**: October - November 2025  
**Lead Auditor**: Team Member 1 (TM1) - Treasury | Team Member 3 (TM3) - Systems  

---

## PART A: TREASURY & CASH MANAGEMENT

**Lead**: TM1

### OBJECTIVES

1. Assess effectiveness of treasury operations
2. Ensure bank reconciliations are accurate and timely
3. Verify adherence to bank authorization matrix
4. Validate cash flow monitoring processes
5. Ensure forex transactions comply with FEMA regulations

### SCOPE

**In-Scope Activities:**
1. Bank account management
2. Bank reconciliations (monthly)
3. Fund transfers and payments
4. Cash flow monitoring
5. Foreign inward remittances (FIRC/BRC)
6. Fixed deposit management (if any)
7. Petty cash (if applicable)

### RISK ASSESSMENT

**High Risk:**
- Unauthorized payments (weak authorization controls)
- Unreconciled bank items (fraud/error concealment)
- Forex non-compliance (FEMA violations)

**Medium Risk:**
- Delayed bank reconciliations
- Cash flow mismanagement

**Low Risk:**
- Routine bank transactions
- FD interest accruals

### CONTROL TESTING

#### Test T1: Bank Reconciliation Accuracy and Timeliness
**Objective**: Verify monthly bank reconciliations are accurate and timely

**Procedure:**
1. Select 3 months: April, June, September 2025
2. For each month and each bank account:
   - Obtain bank reconciliation statement
   - Verify prepared by Finance Executive
   - Verify reviewed by Finance Manager
   - Check reconciliation date (should be within 5 days of month-end)
   - Recalculate reconciliation:
     * Balance per bank statement
     * Add: Deposits in transit
     * Less: Outstanding cheques
     * Equals: Balance per books
   - Investigate reconciling items >30 days old
   - Verify subsequent clearance (next month)

**Sample**: All bank accounts for 3 selected months

**Workpaper**: H1-TRY-01 to H1-TRY-03

**Expected Evidence:**
- Bank reconciliation statements
- Finance Manager review sign-off
- Bank statements
- Tally bank ledgers

#### Test T2: Bank Payment Authorization
**Objective**: Ensure payments follow authorization matrix

**Procedure:**
1. Select 15 high-value bank payments (>₹5 lakhs)
2. Stratify across payment types:
   - Vendor payments
   - Tax payments
   - Salary payments
   - Other payments
3. For each payment:
   - Verify Finance Head initiation
   - Check dual approval per authorization matrix
   - Verify approval evidence (email, bank portal log)
   - Trace to bank statement
   - Verify supporting documentation

**Workpaper**: H1-TRY-04

**Expected Evidence:**
- Payment approval emails
- Bank portal approval logs
- Bank statements
- Supporting documents

#### Test T3: Foreign Inward Remittances (FIRC/BRC)
**Objective**: Ensure FEMA compliance for forex receipts

**Procedure:**
1. Obtain list of all foreign currency receipts in H1
2. Test all if <10, else sample 10
3. For each receipt:
   - Verify bank statement showing forex receipt
   - Check FIRC/BRC certificate obtained
   - Verify certificate details match receipt
   - Check certificate filed properly
   - Verify Senior Accounts Executive validation
   - Check Tally recording at correct exchange rate
   - Verify timely realization (within 9 months of export)

**Workpaper**: H1-TRY-05

**Expected Evidence:**
- Bank statements
- FIRC/BRC certificates
- Senior Accounts Executive sign-off
- Tally entries

#### Test T4: Fixed Deposit Management
**Objective**: Verify FD accounting and interest accrual

**Procedure:**
1. Obtain list of FDs as of Sept 30, 2025
2. For each FD:
   - Verify FD certificate/statement
   - Recalculate interest accrued (Apr-Sep)
   - Verify monthly interest accrual entries in Tally
   - Check Finance Manager review
   - Verify FD maturity tracking

**Workpaper**: H1-TRY-06

**Expected Evidence:**
- FD certificates
- Interest calculation worksheet
- Tally accrual entries
- Finance Manager review

#### Test T5: Cash Flow Monitoring
**Objective**: Assess cash flow management processes

**Procedure:**
1. Inquire about cash flow monitoring practices
2. Obtain cash flow reports (if prepared)
3. Check frequency of monitoring
4. Verify management review
5. Assess adequacy of cash flow forecasting

**Workpaper**: H1-TRY-07

**Expected Evidence:**
- Cash flow reports/statements
- Management review notes
- Forecasting models (if any)

#### Test T6: Petty Cash Management
**Objective**: Verify petty cash controls (if applicable)

**Procedure:**
1. Determine if petty cash exists
2. If yes:
   - Verify petty cash policy
   - Check imprest amount
   - Review sample of petty cash vouchers
   - Verify approvals
   - Check periodic counts
   - Verify reconciliation with books

**Workpaper**: H1-TRY-08

**Expected Evidence:**
- Petty cash policy
- Petty cash vouchers
- Count sheets
- Reconciliation

---

## PART B: TALLY ERP & SYSTEM CONTROLS

**Lead**: TM3

### OBJECTIVES

1. Assess Tally ERP controls for financial data integrity
2. Verify user access controls and segregation of duties
3. Test period locking and audit trail functionality
4. Perform data analytics to identify anomalies
5. Evaluate IT general controls

### SCOPE

**In-Scope System Controls:**
1. User access management
2. Segregation of duties in Tally
3. Period locking controls
4. Audit trail and change tracking
5. Data backup and recovery
6. Master data controls (customer, vendor, GL)
7. Voucher approval mechanisms (if used)

**Data Analytics:**
1. Invoice sequence testing
2. Purchase register anomaly detection
3. Duplicate payment detection
4. Weekend/holiday transaction review
5. Round amount analysis
6. Trend and variance analysis

### CONTROL TESTING

#### Test S1: User Access Rights and Segregation of Duties
**Objective**: Ensure appropriate user access and segregation

**Procedure:**
1. Obtain list of all Tally users
2. For each user:
   - Document role and responsibilities
   - Review Tally access rights
   - Verify access appropriate for role
3. Assess segregation of duties:
   - Can same user create and approve?
   - Can same user create vendor and make payment?
   - Can same user post revenue and receipts?
4. Identify any conflicts
5. Verify compensating controls (supervisory review)

**Workpaper**: H1-SYS-01

**Expected Evidence:**
- Tally user list with access rights
- Role-responsibility matrix
- Segregation of duties assessment

#### Test S2: Period Locking Controls
**Objective**: Verify books are locked after month-end

**Procedure:**
1. Check period lock status for Apr-Sep 2025
2. For each month:
   - Verify lock date
   - Check who applied the lock
   - Verify lock prevents backdated entries
3. Test if lock can be overridden (should require authorization)
4. Review any unlock instances and reasons

**Workpaper**: H1-SYS-02

**Expected Evidence:**
- Tally period lock report
- Lock application logs
- Unlock authorization (if any)

#### Test S3: Audit Trail Review
**Objective**: Identify unauthorized changes to transactions

**Procedure:**
1. Enable and extract Tally audit trail for H1
2. Review for:
   - Deleted vouchers
   - Modified vouchers (after initial posting)
   - Master data changes (customer, vendor, GL)
3. For significant changes:
   - Verify authorization
   - Check business justification
   - Assess appropriateness
4. Identify any suspicious patterns

**Workpaper**: H1-SYS-03

**Expected Evidence:**
- Tally audit trail report
- Authorization for changes
- Justification documentation

#### Test S4: Master Data Controls
**Objective**: Ensure master data integrity

**Procedure:**
1. Customer Master:
   - Review new customers added in H1
   - Check authorization for additions
   - Verify data accuracy
2. Vendor Master:
   - Review new vendors (covered in P2P test)
   - Check for duplicate vendors
3. GL Master:
   - Review new GL accounts created
   - Verify authorization
   - Check Chart of Accounts integrity

**Workpaper**: H1-SYS-04

**Expected Evidence:**
- Master data change logs
- Authorization approvals
- Data validation checks

#### Test S5: Data Backup and Recovery
**Objective**: Assess data backup controls

**Procedure:**
1. Inquire about backup policy
2. Verify backup frequency (daily, weekly?)
3. Check backup storage location (onsite, cloud?)
4. Verify backup testing/restoration (when last tested?)
5. Assess disaster recovery plan

**Workpaper**: H1-SYS-05

**Expected Evidence:**
- Backup policy document
- Backup logs
- Restoration test results
- DR plan

#### Test S6: Invoice Sequence Analysis
**Objective**: Detect missing or duplicate invoices

**Procedure:**
1. Extract sales register from Tally (Apr-Sep)
2. Analyze invoice numbering:
   - Check for gaps
   - Check for duplicates
   - Verify sequential numbering
3. For any gaps:
   - Obtain explanation
   - Verify cancellation approval
4. Generate exception report

**Workpaper**: H1-SYS-06 (also supports H1-REV-01)

**Expected Evidence:**
- Sales register
- Sequence analysis report
- Cancellation approvals for gaps

#### Test S7: Purchase Register Anomaly Detection
**Objective**: Identify unusual purchase transactions

**Procedure:**
1. Extract purchase register from Tally (Apr-Sep)
2. Scan for anomalies:
   - Round amount transactions (e.g., exactly ₹50,000)
   - Weekend/holiday transactions
   - Duplicate invoices (same vendor, amount, date)
   - Unusually high amounts
3. Investigate flagged transactions
4. Assess if anomalies indicate control issues

**Workpaper**: H1-SYS-07

**Expected Evidence:**
- Purchase register
- Anomaly report
- Investigation notes

#### Test S8: Duplicate Payment Detection
**Objective**: Identify potential duplicate payments

**Procedure:**
1. Extract payment vouchers from Tally (Apr-Sep)
2. Analyze for duplicates:
   - Same vendor, same amount, close dates
   - Same invoice number paid twice
3. Investigate flagged items
4. Verify if genuine duplicates or legitimate

**Workpaper**: H1-SYS-08

**Expected Evidence:**
- Payment register
- Duplicate analysis report
- Investigation results

#### Test S9: Trend and Variance Analysis
**Objective**: Identify unusual trends requiring investigation

**Procedure:**
1. Revenue Analysis:
   - Monthly revenue trend (Apr-Sep)
   - Revenue by customer
   - Revenue by stream
2. Expense Analysis:
   - Monthly expense trend by category
   - Expense by vendor
3. Variance Analysis:
   - Compare to budget (if available)
   - Compare to prior year H1
4. Investigate significant variances (>20%)

**Workpaper**: H1-SYS-09 to H1-SYS-11

**Expected Evidence:**
- Trend charts
- Variance analysis
- Management explanations

#### Test S10: Journal Entry Review
**Objective**: Identify unusual manual journal entries

**Procedure:**
1. Extract all manual journal entries for H1
2. Focus on:
   - Large amount JEs
   - Revenue/expense JEs
   - Period-end JEs (last 3 days of month)
3. For flagged JEs:
   - Review description and purpose
   - Verify supporting documentation
   - Check authorization
   - Assess appropriateness

**Workpaper**: H1-SYS-12

**Expected Evidence:**
- Journal entry register
- Supporting documents
- Authorization approvals

---

## SAMPLING STRATEGY

### Treasury
- **Bank Reconciliations**: 3 months × all accounts
- **Payment Authorization**: 15 high-value payments
- **FIRC/BRC**: All forex receipts (or sample 10 if >10)
- **FDs**: All FDs

### Systems
- **User Access**: All users
- **Period Locks**: All months (Apr-Sep)
- **Audit Trail**: 100% review for significant changes
- **Data Analytics**: 100% of transactions (automated)

---

## KEY CONTROLS TO TEST

### Treasury Controls

| Control # | Control Description | Control Owner | Test Procedure | Workpaper |
|-----------|-------------------|---------------|----------------|-----------|
| TRY-C1 | Monthly bank reconciliation | Finance Exec + Finance Mgr | Test T1 | H1-TRY-01 |
| TRY-C2 | Payment authorization per matrix | Finance Head + Approver | Test T2 | H1-TRY-04 |
| TRY-C3 | FIRC/BRC verification for forex | Sr. Accounts Executive | Test T3 | H1-TRY-05 |
| TRY-C4 | FD interest accrual | Finance Executive | Test T4 | H1-TRY-06 |
| TRY-C5 | Cash flow monitoring | Finance Manager | Test T5 | H1-TRY-07 |

### System Controls

| Control # | Control Description | Control Owner | Test Procedure | Workpaper |
|-----------|-------------------|---------------|----------------|-----------|
| SYS-C1 | User access rights appropriate | IT/Finance Manager | Test S1 | H1-SYS-01 |
| SYS-C2 | Period locking after month-end | Finance Manager | Test S2 | H1-SYS-02 |
| SYS-C3 | Audit trail enabled and reviewed | Finance Manager | Test S3 | H1-SYS-03 |
| SYS-C4 | Master data changes authorized | Finance/Admin Head | Test S4 | H1-SYS-04 |
| SYS-C5 | Data backup performed regularly | IT Team | Test S5 | H1-SYS-05 |

---

## TIMELINE

| Week | Activity | TM1 (Treasury) | TM3 (Systems) |
|------|----------|----------------|---------------|
| 1 | Process understanding | Treasury walkthrough | Tally system review |
| 2 | Data collection | Bank statements, FIRC | Tally data extraction |
| 3 | Control testing | Tests T1-T3 | Tests S1-S3 |
| 4 | Control testing | Tests T4-T6 | Tests S4-S5 |
| 5 | Analytics | Support TM3 | Tests S6-S10 |
| 6 | Findings compilation | Treasury findings | Systems findings |

---

## SUCCESS CRITERIA

- [ ] All treasury tests completed
- [ ] All system tests completed
- [ ] Data analytics performed
- [ ] Anomalies investigated
- [ ] Findings documented
- [ ] Workpapers peer-reviewed

---

**Document Control:**
- Plan Version: 1.0
- Date: October 2025
- Prepared by: TM1 (Treasury) & TM3 (Systems)
- Reviewed by: IA Manager
- Approved by: IA Manager
