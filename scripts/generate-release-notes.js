'use strict';
// generate-release-notes.js — called by GitHub Actions to write release_notes.md
// Reads the current VERSION and CHANGELOG from version.js and formats Markdown.

const fs   = require('fs');
const path = require('path');

const v   = require(path.join(__dirname, '..', 'version.js'));
const ver = v.VERSION;
const c   = v.CHANGELOG[ver] || {};
const title = c.title || '';

const lines = [];
lines.push(`## What's New in ${ver}${title ? ' \u2014 ' + title : ''}`);
lines.push('');

if (c.features && c.features.length) {
  lines.push('### \u2728 New Features');
  c.features.forEach(f => lines.push('- ' + f));
  lines.push('');
}

if (c.changes && c.changes.length) {
  lines.push('### \uD83D\uDCCA Improvements');
  c.changes.forEach(ch => lines.push('- ' + ch));
  lines.push('');
}

if (c.fixes && c.fixes.length) {
  lines.push('### \uD83D\uDC1B Fixes');
  c.fixes.forEach(f => lines.push('- ' + f));
  lines.push('');
}

lines.push('---');
lines.push('');
lines.push('## \uD83E\uDD16 AI Models (Optional)');
lines.push(
  'To use local AI transcription and intent classification without downloading from Hugging Face, ' +
  'download **VoApps-Tools-Models.zip** (under Assets below) and extract the `Xenova` folder to:'
);
lines.push('- **Windows:** `%APPDATA%\\voapps-tools\\models\\`');
lines.push('- **macOS:** `~/Library/Application Support/VoApps Tools/models/`');
lines.push('');
lines.push('---');
lines.push('**Full Changelog**: https://github.com/eiznem/voapps-tools/releases');

const out = path.join(process.cwd(), 'release_notes.md');
fs.writeFileSync(out, lines.join('\n'), 'utf8');
process.stdout.write(`Generated release notes for v${ver} -> ${out}\n`);
