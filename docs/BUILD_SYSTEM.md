# Build System Documentation

## Overview

This project uses a simple Node.js build system to maintain modular source code while distributing single-file Google Apps Scripts to end users.

## Architecture

### Development Structure (What You Work With)

```
src/
├── common/                    # Shared code (DRY principle)
│   ├── formatting.gs          # Colors, formatting helpers
│   ├── utilities.gs           # Common utilities
│   └── namedRanges.gs         # Named range functions
└── workbooks/                 # Individual workbook logic
    ├── deferredtax.gs
    ├── far_wp.gs
    ├── ifc_p2p.gs
    ├── indas109.gs
    ├── indas115.gs
    ├── indas116.gs
    └── tds_compliance.gs
```

### Distribution Structure (What Users Get)

```
dist/
├── deferredtax_standalone.gs      # common/* + workbooks/deferredtax.gs
├── far_wp_standalone.gs           # common/* + workbooks/far_wp.gs
├── ifc_p2p_standalone.gs          # common/* + workbooks/ifc_p2p.gs
├── indas109_standalone.gs         # common/* + workbooks/indas109.gs
├── indas115_standalone.gs         # common/* + workbooks/indas115.gs
├── indas116_standalone.gs         # common/* + workbooks/indas116.gs
└── tds_compliance_standalone.gs   # common/* + workbooks/tds_compliance.gs
```

## Build Process

### How It Works

The `build.js` script:

1. Reads all files from `src/common/` (in order):
   - `utilities.gs`
   - `formatting.gs`
   - `namedRanges.gs`

2. For each file in `src/workbooks/`:
   - Concatenates common files
   - Appends the workbook-specific code
   - Writes to `dist/[workbook]_standalone.gs`

### Running the Build

```bash
# One-time build
npm run build

# Or directly with node
node build.js
```

### Output

```
Building standalone Google Apps Script files...

Building deferredtax.gs...
  ✓ Created deferredtax_standalone.gs
Building far_wp.gs...
  ✓ Created far_wp_standalone.gs
...

✓ Build complete! All standalone files are in the dist/ folder.
```

## Development Workflow

### 1. Making Changes to Common Code

**Scenario:** You want to add a new color to the color scheme.

```javascript
// Edit: src/common/formatting.gs
const COLORS = {
  HEADER_BG: "#1a237e",
  // ... existing colors ...
  NEW_COLOR: "#ff5722"  // Add this
};
```

**Build:**
```bash
npm run build
```

**Result:** All 7 standalone files now include the new color.

### 2. Making Changes to a Specific Workbook

**Scenario:** You want to fix a bug in the Ind AS 109 workbook.

```javascript
// Edit: src/workbooks/indas109.gs
function createECLImpairmentSheet(ss) {
  // Your bug fix here
}
```

**Build:**
```bash
npm run build
```

**Result:** Only `dist/indas109_standalone.gs` is updated with your fix.

### 3. Adding a New Workbook

**Steps:**

1. Create `src/workbooks/newworkbook.gs`
2. Write your workbook-specific code
3. Run `npm run build`
4. The build script automatically detects and builds `dist/newworkbook_standalone.gs`

No changes to `build.js` needed!

## Common Code Reference

### utilities.gs

**Functions:**
- `clearExistingSheets(ss)` - Safely clear all sheets
- `onOpen()` - Create custom menu (auto-detects workbook type)
- `showAbout()` - Display about dialog

**Usage in workbooks:**
```javascript
function createMyWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  clearExistingSheets(ss);  // Use common function
  // ... rest of your code
}
```

### formatting.gs

**Constants:**
- `COLORS` - Color scheme object
- `FONT_SIZES` - Font size constants

**Functions:**
- `formatHeader(sheet, row, startCol, endCol, text, bgColor)`
- `formatSubHeader(sheet, row, startCol, values, bgColor)`
- `formatInputCell(range, bgColor)`
- `formatCurrency(range)`
- `formatPercentage(range)`
- `formatDate(range)`
- `setColumnWidths(sheet, widths)`
- `protectSheet(sheet, warningOnly)`

**Usage in workbooks:**
```javascript
function createCoverSheet(ss) {
  const sheet = ss.insertSheet('Cover');
  
  // Use common formatting
  formatHeader(sheet, 1, 1, 5, 'MY WORKBOOK', COLORS.HEADER_BG);
  setColumnWidths(sheet, [50, 200, 150, 100]);
  formatCurrency(sheet.getRange('B5:B10'));
}
```

### namedRanges.gs

**Functions:**
- `setupNamedRanges(ss)` - Default implementation (can be overridden)
- `createNamedRange(ss, name, range)` - Helper function

**Usage in workbooks:**
```javascript
// Override in your workbook if needed
function setupNamedRanges(ss) {
  createNamedRange(ss, 'TaxRate', ss.getSheetByName('Assumptions').getRange('B5'));
  createNamedRange(ss, 'EntityName', ss.getSheetByName('Assumptions').getRange('B3'));
}
```

## Best Practices

### DO ✅

1. **Edit source files in `src/`**, never edit `dist/` directly
2. **Run build after every change** to keep dist/ in sync
3. **Test the generated standalone file** in Google Sheets
4. **Commit both `src/` and `dist/`** to version control
5. **Use common functions** instead of duplicating code
6. **Add comments** to explain complex logic

### DON'T ❌

1. **Don't edit `dist/` files directly** - they'll be overwritten
2. **Don't duplicate code** - extract to `src/common/` instead
3. **Don't forget to build** before testing or committing
4. **Don't break backward compatibility** in common code without testing all workbooks

## Troubleshooting

### Build fails with "Cannot find module"

**Solution:** Make sure you're in the project root directory:
```bash
cd gsheet
npm run build
```

### Changes not appearing in standalone file

**Solution:** Make sure you ran the build:
```bash
npm run build
```

### Function conflicts between common and workbook code

**Solution:** Rename the function in your workbook to be more specific:
```javascript
// Instead of:
function formatHeader() { }

// Use:
function formatIndAS109Header() { }
```

### Want to exclude a file from build

**Solution:** Move it out of `src/workbooks/` or rename it to not end in `.gs`

## Advanced: Customizing the Build

### Changing Common File Order

Edit `build.js`:
```javascript
const commonFiles = [
  path.join(srcDir, 'common', 'utilities.gs'),
  path.join(srcDir, 'common', 'formatting.gs'),
  path.join(srcDir, 'common', 'namedRanges.gs'),
  path.join(srcDir, 'common', 'mynewfile.gs')  // Add here
];
```

### Adding Build Metadata

Edit `build.js` to add a header:
```javascript
let combinedCode = `/**
 * Built: ${new Date().toISOString()}
 * Source: ${workbookFile}
 */\n\n`;

commonFiles.forEach(commonFile => {
  // ... rest of code
});
```

### Watch Mode (Future Enhancement)

To automatically rebuild on file changes, you could add:
```javascript
// In build.js
const chokidar = require('chokidar');

if (process.argv.includes('--watch')) {
  chokidar.watch('src/**/*.gs').on('change', () => {
    console.log('File changed, rebuilding...');
    // Run build logic
  });
}
```

Then use:
```bash
npm run watch
```

## Benefits Recap

| Aspect | Before (Monolithic) | After (Modular + Build) |
|--------|---------------------|-------------------------|
| **Bug Fix** | Edit 7 files | Edit 1 common file, build |
| **New Feature** | Copy-paste to 7 files | Add to common, build |
| **Code Review** | Review 7 large files | Review small, focused files |
| **Testing** | Test 7 files | Test affected files only |
| **User Experience** | Single file (good) | Single file (unchanged) |
| **Developer Experience** | Tedious, error-prone | Fast, maintainable |

## Conclusion

This build system gives you the best of both worlds:
- **Developers** work with clean, modular code
- **Users** get simple, single-file scripts

It's a one-time setup that pays dividends every time you need to make a change.
