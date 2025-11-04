const fs = require('fs');
const path = require('path');

const srcDir = './src';
const distDir = './dist';
const packageJson = JSON.parse(fs.readFileSync('./package.json', 'utf8'));

// Ensure dist directory exists
if (!fs.existsSync(distDir)) {
  fs.mkdirSync(distDir);
}

// Common files to be included in every build (order matters!)
const commonFiles = [
  path.join(srcDir, 'common', 'utilities.gs'),
  path.join(srcDir, 'common', 'formatting.gs'),
  path.join(srcDir, 'common', 'dataValidation.gs'),
  path.join(srcDir, 'common', 'conditionalFormatting.gs'),
  path.join(srcDir, 'common', 'sheetBuilders.gs'),
  path.join(srcDir, 'common', 'namedRanges.gs'),
  path.join(srcDir, 'common', 'errorHandling.gs'),
  path.join(srcDir, 'common', 'testing.gs'),
  path.join(srcDir, 'common', 'configBuilder.gs'),
  path.join(srcDir, 'common', 'sampleData.gs')
];

const workbookDir = path.join(srcDir, 'workbooks');
const workbookFiles = fs.readdirSync(workbookDir);

console.log('Building standalone Google Apps Script files...\n');

// Loop through each workbook file
workbookFiles.forEach(workbookFile => {
  console.log(`Building ${workbookFile}...`);
  
  // Create header with metadata
  const workbookName = workbookFile.replace('.gs', '');
  let combinedCode = `/**
 * @name ${workbookName}
 * @version ${packageJson.version}
 * @built ${new Date().toISOString()}
 * @description Standalone script. Do not edit directly - edit source files in src/ folder.
 * 
 * This file is auto-generated from:
 * - Common utilities (src/common/*.gs)
 * - Workbook-specific code (src/workbooks/${workbookFile})
 * 
 * To make changes:
 * 1. Edit source files in src/ folder
 * 2. Run: npm run build
 * 3. Copy the generated file from dist/ folder to Google Apps Script
 */

`;
  
  // 1. Add common code
  commonFiles.forEach(commonFile => {
    if (fs.existsSync(commonFile)) {
      combinedCode += fs.readFileSync(commonFile, 'utf8') + '\n\n';
    }
  });
  
  // 2. Add workbook-specific code
  combinedCode += fs.readFileSync(path.join(workbookDir, workbookFile), 'utf8');
  
  // 3. Write to a new standalone file in the dist folder
  const outputFileName = workbookFile.replace('.gs', '_standalone.gs');
  fs.writeFileSync(path.join(distDir, outputFileName), combinedCode);
  
  console.log(`  ✓ Created ${outputFileName}`);
});

console.log('\n✓ Build complete! All standalone files are in the dist/ folder.');
