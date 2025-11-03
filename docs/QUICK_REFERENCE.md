# Quick Reference Guide

> Fast lookup for all workbooks in the Ind AS Audit Builder Suite

## 📋 Workbook Selection Guide

### Choose Your Workbook

| If you need to... | Use this workbook | File | Function to run |
|-------------------|-------------------|------|-----------------|
| Account for financial instruments | Ind AS 109 | `indas109.gs` | `createIndAS109WorkingPapers()` |
| Account for leases | Ind AS 116 | `indas116.gs` | `createIndAS116Workbook()` |
| Recognize revenue | Ind AS 115 | `indas115.gs` | `createIndAS115Workbook()` |
| Calculate deferred tax | Deferred Tax | `deferredtax.gs` | `createDeferredTaxWorkbook()` |
| Manage TDS compliance | TDS Tracker | `tds_compliance.gs` | `createTDSComplianceWorkbook()` |
| Audit fixed assets | Fixed Assets WP | `far_wp.gs` | `createFixedAssetsWorkpaper()` |
| Test P2P controls | ICFR P2P | `ifc_p2p.gs` | `createICFRP2PWorkpaper()` |

---

## 🎯 Complexity & Time Estimates

| Workbook | Complexity | Setup Time | Learning Curve | Best For |
|----------|------------|------------|----------------|----------|
| TDS Compliance | ⭐⭐ Medium | 30 min | Easy | Tax teams, CAs |
| Fixed Assets WP | ⭐⭐ Medium | 20 min | Easy | Auditors |
| ICFR P2P | ⭐⭐ Medium | 30 min | Medium | Internal audit |
| Deferred Tax | ⭐⭐⭐ Medium-High | 45 min | Medium | Finance teams |
| Ind AS 115 | ⭐⭐⭐ Medium-High | 60 min | Medium | Revenue accounting |
| Ind AS 116 | ⭐⭐⭐⭐ High | 60 min | Medium-High | Lease accounting |
| Ind AS 109 | ⭐⭐⭐⭐⭐ Very High | 90 min | High | Treasury, finance |

---

## 📊 Feature Comparison

| Feature | 109 | 116 | 115 | DT | TDS | FA | P2P |
|---------|-----|-----|-----|----|----|----|----|
| Auto calculations | ✅ | ✅ | ✅ | ✅ | ✅ | ⚠️ | ⚠️ |
| Sample data | ❌ | ❌ | ❌ | ❌ | ✅ | ❌ | ❌ |
| Journal entries | ✅ | ✅ | ✅ | ✅ | ❌ | ❌ | ❌ |
| Reconciliation | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ❌ |
| Audit program | ❌ | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ |
| Known issues | ⚠️ | ⚠️ | ✅ | ⚠️ | ✅ | ✅ | ✅ |

Legend: ✅ Full support | ⚠️ Partial/Issues | ❌ Not applicable

---

## 🚀 Quick Start Commands

### Installation (All Workbooks)

```
1. Open Google Sheets → New Blank Sheet
2. Extensions → Apps Script
3. Copy relevant .gs file content
4. Paste into editor
5. Save (Ctrl+S / Cmd+S)
6. Select main function from dropdown
7. Click Run ▶
8. Authorize (first time only)
9. Return to sheet
```

### Sample Data (TDS Only)

```
After creating TDS workbook:
1. Extensions → Apps Script
2. Select: populateSampleData
3. Click Run ▶
4. Review Dashboard
```

---

## 📐 Key Formulas by Workbook

### Ind AS 109
```
Fair Value Gain/Loss = Current FV - Previous FV
ECL = EAD × PD × LGD
Interest Income = Opening Carrying Amount × EIR
```

### Ind AS 116
```
ROU Asset = Lease Liability + Initial Costs - Incentives
Interest Expense = Opening Liability × IBR × Time
Depreciation = ROU Asset / Lease Term
```

### Ind AS 115
```
Allocated Amount = Transaction Price × (SSP of PO / Total SSP)
Revenue = Transaction Price × % Complete
Contract Asset = Revenue Recognized - Cash Received
```

### Deferred Tax
```
Temporary Difference = Book Value - Tax Base
DTA = Deductible Difference × Tax Rate
DTL = Taxable Difference × Tax Rate
```

### TDS
```
TDS Amount = IF(Gross > Threshold, Gross × Rate%, 0)
Interest = TDS Amount × 1% × (Delay Days / 30)
```

---

## 🎨 Color Coding (All Workbooks)

| Color | Meaning | Action |
|-------|---------|--------|
| 🟦 Light Blue | Input cell | Fill with your data |
| ⬜ White/Gray | Calculated | Auto-filled, don't edit |
| 🟩 Green | Positive/OK | Review and confirm |
| 🟨 Yellow | Warning/Pending | Needs attention |
| 🟥 Red | Error/Exception | Fix immediately |
| 🟦 Dark Blue | Header | Section title |

---

## 📋 Common Input Fields

### All Workbooks Need
- Entity name
- Reporting period
- Currency
- Preparer name

### Financial Workbooks Need
- Tax rates
- Discount rates
- Accounting policies

### Audit Workbooks Need
- Audit team
- Materiality
- Sample sizes

---

## 🔍 Troubleshooting Quick Fixes

| Problem | Quick Fix |
|---------|-----------|
| Authorization error | Normal first time - follow prompts |
| Function not found | Select from dropdown before Run |
| Nothing happens | Save script, refresh sheet, retry |
| #REF! errors | Don't delete sheets manually |
| Slow performance | Reduce data range, use filters |
| Formula errors | Check input cells are filled |

---

## 📚 Documentation Links

- [Main README](README.md) - Project overview
- [Ind AS 109](INDAS109_README.md) - Financial instruments
- [Ind AS 116](INDAS116_README.md) - Leases
- [Ind AS 115](INDAS115_README.md) - Revenue
- [Deferred Tax](DEFERRED_TAX_README.md) - Income taxes
- [TDS Compliance](TDS_COMPLIANCE_README.md) - TDS management
- [Fixed Assets](FIXED_ASSETS_README.md) - PPE audit
- [ICFR P2P](ICFR_P2P_README.md) - Controls testing
- [Known Issues](todo.md) - Bug tracker

---

## 🎓 Learning Path

### Beginner
1. Start with **TDS Compliance** (easiest, sample data included)
2. Try **Fixed Assets WP** (audit template)
3. Move to **Deferred Tax** (calculations)

### Intermediate
4. **Ind AS 115** (revenue recognition)
5. **ICFR P2P** (controls testing)

### Advanced
6. **Ind AS 116** (lease accounting)
7. **Ind AS 109** (financial instruments)

---

## 💡 Pro Tips

### Efficiency
- Use Ctrl+F (Cmd+F) to find sheets quickly
- Freeze rows/columns for easier navigation
- Use filters on large data sheets
- Create bookmarks for frequently used sheets

### Accuracy
- Always start with Assumptions/Cover sheet
- Fill input cells completely before reviewing calculations
- Use Reconciliation sheets to verify totals
- Check Audit_Notes for guidance

### Collaboration
- Share with "Can comment" for review
- Use comments for questions
- Version history for tracking changes
- Download backup copies regularly

### Customization
- Copy workbook before modifying
- Document changes in Audit_Notes
- Test formulas after changes
- Keep original as template

---

## 📞 Quick Support

| Issue Type | Where to Look |
|------------|---------------|
| How to use | Workbook-specific README |
| Formula error | Audit_Notes sheet in workbook |
| Known bug | [todo.md](todo.md) |
| Feature request | GitHub Issues |
| General question | GitHub Discussions |

---

## 🔄 Update Frequency

| Workbook | Status | Last Updated | Next Update |
|----------|--------|--------------|-------------|
| Ind AS 109 | Stable | Nov 2024 | Q1 2025 |
| Ind AS 116 | Stable | Nov 2024 | Q1 2025 |
| Ind AS 115 | Stable | Nov 2024 | Q2 2025 |
| Deferred Tax | Issues | Nov 2024 | Q1 2025 (fix) |
| TDS Compliance | Stable | Nov 2024 | Q2 2025 |
| Fixed Assets | Stable | Nov 2024 | Q2 2025 |
| ICFR P2P | Stable | Nov 2024 | Q2 2025 |

---

## 📊 Workbook Stats

| Workbook | Sheets | Formulas | Input Cells | Complexity |
|----------|--------|----------|-------------|------------|
| Ind AS 109 | 12 | 200+ | 50+ | Very High |
| Ind AS 116 | 14 | 180+ | 40+ | High |
| Ind AS 115 | 16 | 150+ | 60+ | High |
| Deferred Tax | 12 | 120+ | 30+ | Medium |
| TDS Compliance | 12 | 250+ | 100+ | Medium |
| Fixed Assets | 14 | 80+ | 50+ | Medium |
| ICFR P2P | 13 | 50+ | 80+ | Medium |

---

**Keep this guide handy for quick reference!**

*Last updated: November 2025*