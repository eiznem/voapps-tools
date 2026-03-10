#!/usr/bin/env node
// scripts/update-version.js
// Comprehensive version synchronization across all project files
// Usage: npm run update-version

const fs = require('fs');
const path = require('path');

// ANSI color codes for pretty output
const colors = {
  reset: '\x1b[0m',
  green: '\x1b[32m',
  yellow: '\x1b[33m',
  blue: '\x1b[34m',
  red: '\x1b[31m'
};

function log(message, color = 'reset') {
  console.log(`${colors[color]}${message}${colors.reset}`);
}

// Read version.js as source of truth
const versionPath = path.join(__dirname, '..', 'version.js');
const versionModule = require(versionPath);
const VERSION = versionModule.VERSION;
const VERSION_NAME = versionModule.VERSION_NAME;
const RELEASE_DATE = versionModule.RELEASE_DATE;
const AUTHOR = versionModule.AUTHOR;

log(`\n${'='.repeat(50)}`, 'blue');
log(`VoApps Tools - Version Update Script`, 'blue');
log(`${'='.repeat(50)}`, 'blue');
log(`\nSource: version.js`);
log(`Version: ${VERSION}`, 'yellow');
log(`Name: ${VERSION_NAME}`, 'yellow');
log(`Date: ${RELEASE_DATE}`, 'yellow');
log(`\nUpdating all project files...\n`);

let filesUpdated = 0;
let filesSkipped = 0;
let errors = [];

// ============================================
// 1. UPDATE package.json
// ============================================
try {
  const packagePath = path.join(__dirname, '..', 'package.json');
  const packageJson = JSON.parse(fs.readFileSync(packagePath, 'utf8'));
  
  if (packageJson.version !== VERSION) {
    packageJson.version = VERSION;
    fs.writeFileSync(packagePath, JSON.stringify(packageJson, null, 2) + '\n', 'utf8');
    log(`✅ package.json → ${VERSION}`, 'green');
    filesUpdated++;
  } else {
    log(`⏭️  package.json (already ${VERSION})`, 'yellow');
    filesSkipped++;
  }
} catch (err) {
  errors.push(`package.json: ${err.message}`);
  log(`❌ package.json: ${err.message}`, 'red');
}

// ============================================
// 2. UPDATE server.js
// ============================================
try {
  const serverPath = path.join(__dirname, '..', 'server.js');
  let serverContent = fs.readFileSync(serverPath, 'utf8');
  
  let updated = false;
  
  // Ensure proper import
  if (!serverContent.includes("require('./version')")) {
    const lastRequire = serverContent.lastIndexOf("const ");
    const insertPoint = serverContent.indexOf('\n', lastRequire) + 1;
    serverContent = serverContent.slice(0, insertPoint) + 
                   `const { VERSION, VERSION_NAME } = require('./version');\n` +
                   serverContent.slice(insertPoint);
    updated = true;
  }
  
  if (updated) {
    fs.writeFileSync(serverPath, serverContent, 'utf8');
    log(`✅ server.js → imports version.js`, 'green');
    filesUpdated++;
  } else {
    log(`⏭️  server.js (already imports version.js)`, 'yellow');
    filesSkipped++;
  }
} catch (err) {
  errors.push(`server.js: ${err.message}`);
  log(`❌ server.js: ${err.message}`, 'red');
}

// ============================================
// 3. UPDATE main.js
// ============================================
try {
  const mainPath = path.join(__dirname, '..', 'main.js');
  let mainContent = fs.readFileSync(mainPath, 'utf8');
  
  let updated = false;
  
  if (!mainContent.includes("require('./version')")) {
    const insertPoint = mainContent.indexOf('\n') + 1;
    mainContent = mainContent.slice(0, insertPoint) + 
                 `const { VERSION } = require('./version');\n` +
                 mainContent.slice(insertPoint);
    updated = true;
  }
  
  const titlePattern = /mainWindow\.setTitle\(['"`].*?['"`]\)/g;
  if (titlePattern.test(mainContent)) {
    mainContent = mainContent.replace(
      titlePattern,
      `mainWindow.setTitle(\`VoApps Tools v\${VERSION}\`)`
    );
    updated = true;
  }
  
  if (updated) {
    fs.writeFileSync(mainPath, mainContent, 'utf8');
    log(`✅ main.js → v${VERSION}`, 'green');
    filesUpdated++;
  } else {
    log(`⏭️  main.js (no updates needed)`, 'yellow');
    filesSkipped++;
  }
} catch (err) {
  errors.push(`main.js: ${err.message}`);
  log(`❌ main.js: ${err.message}`, 'red');
}

// ============================================
// 4. UPDATE index.html
// ============================================
try {
  const indexPath = path.join(__dirname, '..', 'index.html');
  let indexContent = fs.readFileSync(indexPath, 'utf8');
  
  let updated = false;
  
  const patterns = [
    { regex: /<title>VoApps Tools v[\d.]+<\/title>/g, replacement: `<title>VoApps Tools v${VERSION}</title>` },
    { regex: /<h1>VoApps Tools v[\d.]+<\/h1>/g, replacement: `<h1>VoApps Tools v${VERSION}</h1>` },
    { regex: /Version: v?[\d.]+/g, replacement: `Version: v${VERSION}` },
    { regex: /<span id="version">v?[\d.]+<\/span>/g, replacement: `<span id="version">v${VERSION}</span>` }
  ];
  
  patterns.forEach(({ regex, replacement }) => {
    if (regex.test(indexContent)) {
      indexContent = indexContent.replace(regex, replacement);
      updated = true;
    }
  });
  
  if (updated) {
    fs.writeFileSync(indexPath, indexContent, 'utf8');
    log(`✅ index.html → v${VERSION}`, 'green');
    filesUpdated++;
  } else {
    log(`⏭️  index.html (no version strings found)`, 'yellow');
    filesSkipped++;
  }
} catch (err) {
  errors.push(`index.html: ${err.message}`);
  log(`❌ index.html: ${err.message}`, 'red');
}

// ============================================
// 5. UPDATE trendAnalyzer.js
// ============================================
try {
  const analyzerPath = path.join(__dirname, '..', 'trendAnalyzer.js');
  let analyzerContent = fs.readFileSync(analyzerPath, 'utf8');
  
  let updated = false;
  
  const patterns = [
    { regex: /const VERSION = ['"][\d.]+['"];/g, replacement: `const VERSION = '${VERSION}';` },
    { regex: /\/\/ trendAnalyzer\.js - v[\d.]+/g, replacement: `// trendAnalyzer.js - v${VERSION}` }
  ];
  
  patterns.forEach(({ regex, replacement }) => {
    if (regex.test(analyzerContent)) {
      analyzerContent = analyzerContent.replace(regex, replacement);
      updated = true;
    }
  });
  
  if (updated) {
    fs.writeFileSync(analyzerPath, analyzerContent, 'utf8');
    log(`✅ trendAnalyzer.js → v${VERSION}`, 'green');
    filesUpdated++;
  } else {
    log(`⏭️  trendAnalyzer.js (no version strings found)`, 'yellow');
    filesSkipped++;
  }
} catch (err) {
  errors.push(`trendAnalyzer.js: ${err.message}`);
  log(`❌ trendAnalyzer.js: ${err.message}`, 'red');
}

// ============================================
// 6. UPDATE README.md
// ============================================
try {
  const readmePath = path.join(__dirname, '..', 'README.md');
  let readmeContent = fs.readFileSync(readmePath, 'utf8');
  
  let updated = false;
  
  const patterns = [
    { regex: /badge\/version-[\d.]+-blue/g, replacement: `badge/version-${VERSION}-blue` },
    { regex: /VoApps Tools-[\d.]+-arm64\.dmg/g, replacement: `VoApps Tools-${VERSION}-arm64.dmg` },
    { regex: /\*\*Latest Version:\*\* \[v[\d.]+\]/g, replacement: `**Latest Version:** [v${VERSION}]` },
    { regex: /\*\*Version:\*\* [\d.]+/g, replacement: `**Version:** ${VERSION}` },
    { regex: /\*\*Last Updated:\*\* .+/g, replacement: `**Last Updated:** ${RELEASE_DATE}` }
  ];
  
  patterns.forEach(({ regex, replacement }) => {
    if (regex.test(readmeContent)) {
      readmeContent = readmeContent.replace(regex, replacement);
      updated = true;
    }
  });
  
  if (updated) {
    fs.writeFileSync(readmePath, readmeContent, 'utf8');
    log(`✅ README.md → v${VERSION}`, 'green');
    filesUpdated++;
  } else {
    log(`⏭️  README.md (no version strings found)`, 'yellow');
    filesSkipped++;
  }
} catch (err) {
  errors.push(`README.md: ${err.message}`);
  log(`❌ README.md: ${err.message}`, 'red');
}

// ============================================
// SUMMARY
// ============================================
log(`\n${'='.repeat(50)}`, 'blue');
log(`Update Summary`, 'blue');
log(`${'='.repeat(50)}`, 'blue');
log(`\n✅ Files Updated: ${filesUpdated}`, 'green');
log(`⏭️  Files Skipped: ${filesSkipped}`, 'yellow');

if (errors.length > 0) {
  log(`❌ Errors: ${errors.length}`, 'red');
  errors.forEach(err => log(`   ${err}`, 'red'));
} else {
  log(`❌ Errors: 0`, 'green');
}

log(`\n${'='.repeat(50)}`, 'blue');
log(`Next Steps:`, 'blue');
log(`${'='.repeat(50)}`, 'blue');
log(`1. Review changes: git diff`);
log(`2. Test app: npm start`);
log(`3. Build app: npm run build:mac`);
log(`4. Commit: git commit -am "Release v${VERSION}"`, 'yellow');
log(`5. Tag: git tag -a v${VERSION} -m "Release v${VERSION}"`, 'yellow');
log(`6. Push: git push origin main v${VERSION}`, 'yellow');
log(``);

if (errors.length > 0) {
  process.exit(1);
}