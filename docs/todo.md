# Known Issues and Planned Improvements

> This file tracks known issues, their impact, and planned fixes for the Ind AS Audit Builder Suite

**Last Updated:** November 2024
**Status:** All known issues resolved - no outstanding items

---

## ✅ Resolved Issues

### High Priority (Resolved)

**Deferred Tax - Movement Analysis Flaws**
- **Issue:** Used hardcoded percentages instead of actual data
- **Impact:** Unreliable movement analysis
- **Resolution:** Fixed calculation logic to use dynamic data references
- **Status:** ✅ Resolved in v1.0.1
- **Details:** Movement analysis now correctly pulls from Temp_Differences sheet

**Ind AS 116 - Interest Calculation**
- **Issue:** Used average balance method instead of true EIR
- **Impact:** Minor variance in interest expense
- **Resolution:** Implemented proper EIR calculation methodology
- **Status:** ✅ Resolved in v1.0.1
- **Details:** Now uses effective interest rate method consistently

**Ind AS 109 - ECL Discounting**
- **Issue:** ECL not discounted to present value
- **Impact:** Overstated ECL for long-term exposures
- **Resolution:** Added present value discounting to ECL calculations
- **Status:** ✅ Resolved in v1.0.1
- **Details:** ECL now properly discounted using appropriate discount rates

---

## 📋 Version 1.1 Planned Enhancements (Q1 2025)

### New Features
- [ ] Sample data for all workbooks (currently only TDS has sample data)
- [ ] Video tutorials for installation and usage
- [ ] Enhanced dashboard features
- [ ] Export to Excel functionality

### Performance Improvements
- [ ] Code optimization for large datasets
- [ ] Reduced formula calculation time
- [ ] Memory usage optimization

---

## 🎯 Version 2.0 Roadmap (Q2 2025)

### New Standards
- [ ] Ind AS 19 - Employee Benefits
- [ ] Ind AS 36 - Impairment of Assets
- [ ] Ind AS 21 - Foreign Exchange
- [ ] Ind AS 37 - Provisions

### Advanced Features
- [ ] API integrations with accounting software
- [ ] Multi-entity consolidation support
- [ ] Automated data import from Excel/CSV
- [ ] Custom reporting templates

---

## 🐛 Issue Reporting

If you encounter any issues not listed above:

1. **Check Documentation** - Review the specific workbook README
2. **Verify Installation** - Ensure proper script installation
3. **Check Formulas** - Verify input cells are filled correctly
4. **Report Issues** - Open a GitHub issue with:
   - Workbook name and version
   - Steps to reproduce
   - Expected vs actual behavior
   - Screenshot if applicable

---

## ✅ Quality Assurance

### Testing Status
- **Manual Testing:** ✅ Complete
- **Formula Validation:** ✅ Complete
- **Cross-sheet References:** ✅ Complete
- **Error Handling:** ✅ Complete
- **User Acceptance:** ⏳ In Progress

### Known Limitations
- Requires Google account for Google Sheets
- Internet connection needed for script execution
- Mobile browsers not recommended
- Large datasets may impact performance

---

**All critical issues have been resolved. The suite is production-ready for professional use.**

*Last updated: November 2024*