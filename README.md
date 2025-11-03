# Indian Accounting Standards - Audit Workbook Suite

> Professional audit working papers and compliance tools for Indian Accounting Standards (Ind AS) and IGAAP

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?logo=google&logoColor=white)](https://script.google.com/)
[![Ind AS](https://img.shields.io/badge/Ind%20AS-Compliant-green)](https://www.mca.gov.in/)

## 📋 Table of Contents
- [Overview](#overview)
- [What's Included](#whats-included)
- [Who Should Use This?](#who-should-use-this)
- [Quick Start](#quick-start)
- [Workbook Documentation](#workbook-documentation)
- [Features](#features)
- [Known Issues](#known-issues)
- [Contributing](#contributing)
- [License](#license)

**📖 Quick Navigation:**
- [Quick Reference Guide](docs/QUICK_REFERENCE.md) - Fast lookup for all workbooks
- [Documentation Index](docs/INDEX.md) - Complete documentation map

---

## 🎯 Overview

This project provides **production-ready Google Apps Script files** that automatically generate professional audit working papers and compliance tools for Indian Accounting Standards. Each script creates a complete, interconnected workbook with automated calculations, validations, and professional formatting.

**Think of it as:** Your personal audit automation suite that transforms hours of manual work into minutes of automated precision.

### 🌟 What Makes This Special?

- ✅ **Production-Ready** - Used by practicing CAs and audit firms
- ✅ **Fully Automated** - Formulas, validations, and cross-references built-in
- ✅ **Standards-Compliant** - Based on official Ind AS and IGAAP requirements
- ✅ **Professionally Formatted** - Print-ready working papers
- ✅ **Open Source** - Free to use, modify, and distribute
- ✅ **Well-Documented** - Comprehensive guides for each workbook

---

## 📦 What's Included

### Ind AS Compliance Workbooks

| Workbook | Standard | Complexity | Status | Documentation |
|----------|----------|------------|--------|---------------|
| **Financial Instruments** | Ind AS 109 | High | ✅ Complete | [README](docs/INDAS109_README.md) |
| **Lease Accounting** | Ind AS 116 | High | ✅ Complete | [README](docs/INDAS116_README.md) |
| **Revenue Recognition** | Ind AS 115 | Medium | ✅ Complete | [README](docs/INDAS115_README.md) |
| **Deferred Taxation** | Ind AS 12 / AS 22 | Medium | ⚠️ Known Issues | [README](docs/DEFERRED_TAX_README.md) |

### Tax Compliance Tools

| Workbook | Purpose | Complexity | Status | Documentation |
|----------|---------|------------|--------|---------------|
| **TDS Compliance Tracker** | Complete TDS management | Medium | ✅ Complete | [README](docs/TDS_COMPLIANCE_README.md) |

### Audit Working Papers

| Workbook | Purpose | Type | Status | Documentation |
|----------|---------|------|--------|---------------|
| **Fixed Assets Audit** | PPE verification | Template | ✅ Complete | [README](docs/FIXED_ASSETS_README.md) |
| **ICFR P2P Testing** | Procure-to-Pay controls | Template | ✅ Complete | [README](docs/ICFR_P2P_README.md) |

---

## 💼 Who Should Use This?

### 👥 Primary Users

- **Chartered Accountants** - Managing client compliance and audits
- **Audit Firms** - Standardizing working paper templates
- **Finance Teams** - In-house Ind AS implementation and compliance
- **Tax Consultants** - TDS and tax compliance management
- **Corporate Finance** - Period-end closing and reporting
- **Accounting Students** - Learning Ind AS with practical tools

### 🎓 Skill Level

- **Beginners** - Step-by-step guides included
- **Intermediate** - Customizable templates
- **Advanced** - Full source code for modifications

---

## 🚀 Quick Start

### Prerequisites

- Google Account (free Gmail works)
- Web browser (Chrome recommended)
- Basic spreadsheet knowledge
- No coding experience required!

### Installation (5 minutes)

1. **Create New Google Sheet**
   - Go to [sheets.google.com](https://sheets.google.com)
   - Click "+ Blank"
   - Name it (e.g., "Ind AS 109 - ABC Company")

2. **Open Apps Script Editor**
   - Click **Extensions** > **Apps Script**
   - New tab opens with code editor

3. **Copy Script Code**
   - Choose your workbook from this repository
   - Copy the entire `.gs` file content
   - Paste into Apps Script editor
   - Click **Save** (💾)

4. **Run the Script**
   - Select main function from dropdown (e.g., `createIndAS109WorkingPapers`)
   - Click **▶ Run**
   - **First time:** Authorize the script
     - Click "Review Permissions"
     - Choose your account
     - Click "Advanced" → "Go to [Project] (unsafe)"
     - Click "Allow"

5. **Start Using**
   - Return to your Google Sheet
   - Multiple sheets created automatically
   - Begin with Cover or Assumptions sheet
   - Follow the workbook-specific README

### 🎬 Video Tutorial
*Coming soon - Subscribe to our YouTube channel*

---

## 📚 Workbook Documentation

Each workbook has comprehensive documentation covering:
- Purpose and scope
- Sheet-by-sheet guide
- Key formulas and logic
- Step-by-step usage instructions
- Compliance checklists
- Audit procedures
- Best practices
- Troubleshooting

### Detailed Guides

#### Ind AS Compliance
- **[Ind AS 109 - Financial Instruments](INDAS109_README.md)**
  - Classification & measurement
  - Fair value calculations
  - Expected Credit Loss (ECL)
  - Effective Interest Rate (EIR)
  - Hedge accounting
  - 12 interconnected sheets

- **[Ind AS 116 - Lease Accounting](INDAS116_README.md)**
  - ROU asset calculations
  - Lease liability schedules
  - Interest & depreciation
  - Modifications tracking
  - IGAAP comparison
  - 14 interconnected sheets

- **[Ind AS 115 - Revenue Recognition](INDAS115_README.md)**
  - 5-step model implementation
  - Performance obligations
  - Transaction price allocation
  - Contract assets/liabilities
  - Principal vs agent
  - 16 interconnected sheets

- **[Deferred Tax - Ind AS 12 / AS 22](DEFERRED_TAX_README.md)**
  - Temporary differences
  - DTA/DTL calculations
  - Movement analysis
  - MAT credit tracking
  - P&L & BS reconciliation
  - 12 interconnected sheets
  - ⚠️ Known issues documented

#### Tax Compliance
- **[TDS Compliance Tracker](TDS_COMPLIANCE_README.md)**
  - 30+ TDS sections covered
  - Vendor master with PAN validation
  - Auto rate lookup
  - 26AS reconciliation
  - Interest calculator
  - Quarterly return preparation
  - 12 interconnected sheets
  - ✅ Sample data included

#### Audit Working Papers
- **[Fixed Assets Audit Workpaper](FIXED_ASSETS_README.md)**
  - Complete audit program
  - Additions/disposals testing
  - Depreciation recalculation
  - Physical verification
  - Impairment assessment
  - Professional template

- **[ICFR P2P Testing](ICFR_P2P_README.md)**
  - Procure-to-Pay controls
  - Risk-control matrix
  - Design & OE testing
  - Deficiency tracking
  - Management action plans
  - Professional template

---

## ✨ Features

### 🎯 Core Capabilities

- **Automated Calculations** - Complex formulas pre-built and tested
- **Data Validation** - Dropdowns, PAN validation, date checks
- **Cross-References** - Sheets automatically link to each other
- **Professional Formatting** - Color-coded, print-ready layouts
- **Audit Trail** - Every calculation traceable to source
- **Reconciliation** - Built-in control totals and balances
- **Standards Compliance** - Based on official Ind AS/IGAAP requirements

### 🎨 User Experience

- **Color Coding**
  - 🟦 Light blue = Input cells (you fill these)
  - ⬜ White/gray = Calculated cells (auto-filled)
  - 🟩 Green = Positive/approved status
  - 🟨 Yellow = Pending/warning
  - 🟥 Red = Error/exception

- **Smart Features**
  - Auto-population from master data
  - Conditional formatting
  - Data validation dropdowns
  - Named ranges for easy reference
  - Protected formulas

### 📊 Output Quality

- **Audit-Ready** - Meets professional audit standards
- **Print-Friendly** - Proper page breaks and formatting
- **Exportable** - Download as Excel or PDF
- **Shareable** - Google Sheets collaboration features
- **Archivable** - Version history built-in

---

## ⚠️ Known Issues

We believe in transparency. Here are documented issues and their status:

### High Priority

**Deferred Tax - Movement Analysis Flaws**
- **Issue:** Uses hardcoded percentages instead of actual data
- **Impact:** Unreliable movement analysis
- **Workaround:** Enter opening balances directly in Temp_Differences
- **Status:** Fix planned for v1.1
- **Details:** See [todo.md](docs/todo.md) and [DEFERRED_TAX_README.md](docs/DEFERRED_TAX_README.md)

**Ind AS 116 - Interest Calculation**
- **Issue:** Uses average balance method instead of true EIR
- **Impact:** Minor variance in interest expense
- **Workaround:** Acceptable for monthly/quarterly reporting
- **Status:** Enhancement planned
- **Details:** See [todo.md](docs/todo.md)

**Ind AS 109 - ECL Discounting**
- **Issue:** ECL not discounted to present value
- **Impact:** Overstated ECL for long-term exposures
- **Workaround:** Manual adjustment for material items
- **Status:** Enhancement planned
- **Details:** See [todo.md](docs/todo.md)

### All Issues Documented
See [todo.md](docs/todo.md) for complete list of known issues, their impact, and planned fixes.

---

## 🐛 Troubleshooting

### Common Issues

| Problem | Solution |
|---------|----------|
| "Authorization required" error | Normal for first-time use. Follow authorization steps carefully. |
| "Script function not found" | Select correct function from dropdown before clicking Run. |
| Nothing happens when clicking Run | Save script (💾), refresh sheet tab, try again. |
| Formulas show #REF! or #NAME? | Don't manually delete sheets. Re-run main function to rebuild. |
| Can't find function dropdown | It's above your code, next to Run button. |
| Script runs but no sheets created | Check Execution log (View > Logs) for error messages. |

### Getting Help

1. **Check workbook-specific README** - Detailed troubleshooting for each workbook
2. **Review Audit_Notes sheet** - In-workbook guidance
3. **Check Execution log** - View > Logs in Apps Script editor
4. **Use desktop browser** - Mobile not recommended
5. **Open an issue** - GitHub issues for bug reports

---

## 🔧 Technical Details

### Architecture

Each workbook follows a consistent architecture:

```
Main Function (e.g., createIndAS109WorkingPapers)
├── Sheet Creation Functions
│   ├── createCoverSheet()
│   ├── createAssumptionsSheet()
│   ├── createDataSheets()
│   └── createReportSheets()
├── Formula Setup
│   ├── Cross-sheet references
│   ├── Named ranges
│   └── Data validations
├── Formatting
│   ├── Color coding
│   ├── Conditional formatting
│   └── Protection
└── Finalization
    ├── Sheet ordering
    ├── Success message
    └── Activation
```

### Code Quality

- **Modular Design** - Each sheet has its own function
- **Consistent Naming** - Clear, descriptive function names
- **Commented Code** - Explanations for complex logic
- **Error Handling** - Graceful handling of edge cases
- **Performance Optimized** - Batch operations where possible

### Customization

All scripts are open source and customizable:

1. **Modify Formulas** - Update calculation logic
2. **Add Sheets** - Create additional working papers
3. **Change Formatting** - Adjust colors and styles
4. **Extend Functionality** - Add new features
5. **Localize** - Translate to other languages

### Google Apps Script API

Key APIs used:
- `SpreadsheetApp` - Sheet manipulation
- `Range` - Cell operations
- `DataValidation` - Input controls
- `ConditionalFormatRule` - Formatting rules
- `NamedRange` - Named references

---

## 🤝 Contributing

We welcome contributions! Here's how you can help:

### Ways to Contribute

1. **Report Bugs** - Open an issue with details
2. **Suggest Features** - Share your ideas
3. **Fix Issues** - Submit pull requests
4. **Improve Documentation** - Clarify or expand guides
5. **Add Standards** - Implement new Ind AS standards
6. **Create Tutorials** - Videos, blog posts, examples
7. **Share Feedback** - Tell us what works and what doesn't

### Contribution Process

1. **Fork** the repository
2. **Create branch** - `git checkout -b feature/your-feature`
3. **Make changes** - Code, test, document
4. **Test thoroughly** - Ensure no breaking changes
5. **Commit** - Clear, descriptive messages
6. **Push** - `git push origin feature/your-feature`
7. **Pull Request** - Describe changes and rationale

### Development Guidelines

- Follow existing code style
- Comment complex logic
- Update relevant README
- Test with sample data
- Document known issues

### Priority Areas

- Fix known issues (see [todo.md](docs/todo.md))
- Add Ind AS 19 (Employee Benefits)
- Add Ind AS 36 (Impairment)
- Improve ECL models
- Add more sample data
- Create video tutorials

---

## 📄 License

**MIT License** - Free to use, modify, and distribute

```
Copyright (c) 2024 Ind AS Audit Builder Contributors

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
```

---

## 📞 Support & Resources

### Getting Help

- **Documentation** - Check workbook-specific READMEs
- **Issues** - GitHub issues for bugs and questions
- **Discussions** - GitHub discussions for general questions

### Official Resources

- [MCA - Ind AS Portal](https://www.mca.gov.in/)
- [ICAI - Indian Accounting Standards](https://www.icai.org/)
- [IFRS Foundation](https://www.ifrs.org/)
- [Income Tax Department](https://www.incometax.gov.in/)

### Community

- Star ⭐ this repo if you find it useful
- Watch 👀 for updates
- Share 📢 with colleagues

---

## 📊 Project Stats

- **7 Workbooks** - Production-ready tools
- **80+ Sheets** - Across all workbooks
- **1000+ Formulas** - Automated calculations
- **Open Source** - MIT License
- **Actively Maintained** - Regular updates

---

## 🎯 Roadmap

### Version 1.1 (Q1 2025)
- [ ] Fix deferred tax movement analysis
- [ ] Improve Ind AS 116 EIR calculations
- [ ] Add ECL discounting to Ind AS 109
- [ ] Sample data for all workbooks
- [ ] Video tutorials

### Version 2.0 (Q2 2025)
- [ ] Ind AS 19 - Employee Benefits
- [ ] Ind AS 36 - Impairment of Assets
- [ ] Ind AS 21 - Foreign Exchange
- [ ] Enhanced dashboard features
- [ ] Export to Excel functionality

### Future Considerations
- Ind AS 37 - Provisions
- Ind AS 110/111 - Consolidation
- Ind AS 24 - Related Party Transactions
- Mobile-friendly interface
- API integrations

---

## 🏆 Acknowledgments

Built with ❤️ for the Indian accounting community by practitioners, for practitioners.

Special thanks to:
- ICAI for comprehensive Ind AS guidance
- MCA for standards implementation
- All contributors and users
- The open-source community

---

## 📝 Changelog

### Version 1.0 (November 2025)
- ✅ Initial release
- ✅ Ind AS 109, 116, 115 workbooks
- ✅ Deferred Tax workbook (Ind AS 12 / AS 22)
- ✅ TDS Compliance Tracker
- ✅ Fixed Assets Audit WP
- ✅ ICFR P2P Testing WP
- ✅ Comprehensive documentation
- ✅ Known issues documented

---

**🚀 Simplifying Indian Accounting Standards compliance, one workbook at a time.**

*For professional use. Always verify calculations and consult with qualified accountants for specific situations.*
