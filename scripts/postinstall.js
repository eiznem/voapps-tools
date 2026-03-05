#!/usr/bin/env node
/**
 * postinstall.js
 *
 * Rebuilds native modules (DuckDB) against the correct Electron ABI.
 *
 * Why this is needed:
 *   npm install downloads pre-built binaries targeting the *system* Node.js ABI.
 *   Electron 28 uses a different internal module ABI than Node.js 18 even though
 *   it embeds Node 18. electron-rebuild re-targets the binary to Electron's ABI,
 *   making DuckDB actually loadable inside the packaged Electron app.
 *
 * Failure handling:
 *   On systems without Visual Studio Build Tools (Windows) or Xcode (Mac),
 *   electron-rebuild may fail to compile from source if no pre-built binary is
 *   available. That's OK — the app detects the failure and disables DB-dependent
 *   features gracefully (AI transcript caching is skipped; core features still work).
 */

'use strict';
const path = require('path');
const { execSync } = require('child_process');

const root = path.join(__dirname, '..');
const bin  = path.join(root, 'node_modules', '.bin',
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
