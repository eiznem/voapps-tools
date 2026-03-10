#!/usr/bin/env node
/**
 * postinstall.js
 *
 * 1. Rebuilds native modules (DuckDB) against the correct Electron ABI.
 * 2. Restores the sharp Proxy stub so @xenova/transformers can initialize on
 *    platforms where the sharp native binary is unavailable (e.g. Windows
 *    when the app is built on macOS, or any platform where sharp was not
 *    pre-built for the Electron runtime).
 *
 * Why we need the sharp stub:
 *   @xenova/transformers optionally imports sharp for image resizing. When
 *   sharp's native binary is missing, the original sharp.js throws an Error.
 *   That throw propagates through @xenova/transformers init and prevents the
 *   AI pipeline from loading entirely.  Our stub replaces the throw with a
 *   fully-chainable Proxy so every property access succeeds silently, letting
 *   @xenova/transformers continue to load (it only uses sharp for image input,
 *   which we never pass — we always send a pre-decoded Float32Array).
 *
 *   IMPORTANT: npm install sharp (or any package that depends on sharp) will
 *   re-download the package and overwrite lib/sharp.js with the original
 *   throwing version.  This postinstall always restores our stub after any
 *   npm install, keeping the build hermetic.
 */

'use strict';
const fs   = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const root = path.join(__dirname, '..');

// ─── 1. Restore the sharp Proxy stub ────────────────────────────────────────
const sharpJs = path.join(root, 'node_modules', 'sharp', 'lib', 'sharp.js');
const SHARP_STUB = `// Copyright 2013 Lovell Fuller and others.
// SPDX-License-Identifier: Apache-2.0

'use strict';

const platformAndArch = require('./platform')();

/* istanbul ignore next */
try {
  module.exports = require(\`../build/Release/sharp-\${platformAndArch}.node\`);
} catch (err) {
  // Instead of hard-crashing, export a fully-chainable Proxy stub so packages
  // that optionally use sharp (e.g. @xenova/transformers) can fully initialize.
  //
  // CRITICAL: get and apply MUST both return \`stub\` (the Proxy), NOT the bare
  // \`noop\` function.  sharp/lib/utility.js does:
  //   const format = sharp.format();
  //   format.heif.output.alias = ['avif', 'heic'];
  // If get/apply return the plain noop, then format.heif is a plain function
  // (not Proxied) so format.heif.output === undefined → TypeError.
  // By returning stub everywhere, every sub-property access stays on the Proxy
  // and property assignments are absorbed by the set trap.
  console.warn('[sharp] Native binary unavailable for', platformAndArch, '— image processing disabled.');
  let stub;
  stub = new Proxy(function () { return stub; }, {
    apply ()      { return stub; },
    get (_, prop) {
      if (prop === '__sharpUnavailable') return true;
      if (prop === 'then') return undefined; // not a thenable/Promise
      return stub;  // always return Proxy so property chains never hit undefined
    },
    set ()        { return true; } // silently absorb all property assignments
  });
  module.exports = stub;
}
`;

try {
  if (fs.existsSync(sharpJs)) {
    const current = fs.readFileSync(sharpJs, 'utf8');
    if (!current.includes('__sharpUnavailable')) {
      fs.writeFileSync(sharpJs, SHARP_STUB, 'utf8');
      console.log('[postinstall] ✅ sharp Proxy stub restored');
    } else {
      console.log('[postinstall] ⏭️  sharp stub already in place');
    }
  } else {
    console.log('[postinstall] ⚠️  sharp not installed — stub not needed yet');
  }
} catch (stubErr) {
  console.warn('[postinstall] ⚠️  Could not restore sharp stub:', stubErr.message);
}

// ─── 2. Rebuild DuckDB for Electron ABI ─────────────────────────────────────
const bin = path.join(root, 'node_modules', '.bin',
  process.platform === 'win32' ? 'electron-rebuild.cmd' : 'electron-rebuild');

console.log('[postinstall] Rebuilding native modules for Electron (target: DuckDB)...');

try {
  execSync(`"${bin}" -f -w duckdb`, { stdio: 'inherit', cwd: root });
  console.log('[postinstall] ✅ DuckDB rebuilt for Electron successfully');
} catch (e) {
  console.warn('');
  console.warn('[postinstall] ⚠️  electron-rebuild failed — DuckDB will be disabled at runtime.');
  console.warn('[postinstall]    This is expected if build tools are not installed.');
  console.warn('[postinstall]    Windows: install Visual Studio Build Tools 2019+');
  console.warn('[postinstall]    Mac:     xcode-select --install');
  console.warn('[postinstall]    App will still start; AI caching and DB features will be unavailable.');
  console.warn('');
  // Exit 0 so npm install does not fail — missing DuckDB is handled gracefully at runtime
  process.exitCode = 0;
}
