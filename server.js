/**
 * VoApps Tools — Local Server (Electron)
 * Version: 4.0.5
 *
 * NEW IN v4.0.3:
 * - Bulk INSERT with transaction (ON CONFLICT DO NOTHING) replaces per-row SELECT+INSERT;
 *   ~100-200× faster database saves on Windows
 * - Windows DLL path injection for onnxruntime-node so Whisper/transformers can load
 * - Fixed tryBareImport re-throwing onnxruntime init errors as "Unexpected error"
 * - All hardcoded version strings removed from index.html (now fetched from /api/ping)
 *
 * FROM v4.0.0:
 * - AI Message Analysis: transcribes DDVM recordings via local Whisper or OpenAI Whisper API
 * - Intent & summary: local nli-deberta-v3-small (free) or GPT-4o-mini; results cached in DuckDB
 * - Caller number match detection, URL mentions, and Voice Append detection from voapps_voice_append
 * - New message_transcriptions DuckDB table; voapps_voice_append column in campaign_results
 * - Report Output modal: Analysis Tabs moved above CSV Columns, tightened spacing
 *
 * FROM v3.4.2:
 * - Timezone-aware campaign date filter (covers Hawaii UTC-10 through Eastern UTC-4)
 * - Extended API query buffer to +2 days to cover Hawaii campaigns
 * - Optional detail tabs (TN Health, Variability Analysis, Number Summary) with localStorage preference
 *
 * FROM v3.3.1:
 * - All timestamps normalized to user-selected timezone (default: UTC-7 VoApps Time)
 *
 * FROM v3.3.0:
 * - Windows x64 support with native HTTPS fetch
 * - Configurable output folder (Documents by default)
 * - Fixed success rate calculation (only counts delivery attempts with timestamps)
 * - Streaming database export to prevent OOM on large datasets
 * - Cross-platform settings storage
 *
 * FROM v3.2.0:
 * - Delivery Intelligence Platform with TN Health Classification
 * - Attempt Index tracking, Variability Score, Retry Decay Curve
 * - Day-of-week recommendations for accounts/messages
 * - Global Insights with timezone detection
 *
 * FROM v3.1.0:
 * - Live log streaming with SSE (Server-Sent Events)
 * - Real-time progress bar updates
 * - Bulk export CSV enrichment with caller_number and message_id
 *
 * FROM v3.0.0:
 * - DuckDB database integration for faster analysis
 * - Smart data checking (skip duplicate API calls)
 * - Three output modes: CSV only, Database only, or Both
 * - Database management endpoints
 * - SQL query interface
 */

"use strict";

const http = require("http");
const https = require("https");

// Fix for SSL certificate issues in Electron
// - On macOS: Use mac-ca to access Apple Keychain certificate store (fixes VPN issues)
// - On Windows: Use permissive agent for trusted VoApps domains
const TRUSTED_DOMAINS = ['directdropvoicemail.voapps.com', 'voapps.com'];
let httpsAgent = null;

function isTrustedDomain(hostname) {
  return TRUSTED_DOMAINS.some(d => hostname.includes(d));
}

// macOS: Load system certificates from Keychain
// This fixes SSL issues when behind corporate VPNs or proxy servers
if (process.platform === 'darwin') {
  try {
    const macCa = require('mac-ca');
    // Add macOS Keychain certificates to the trusted CA list
    macCa.addToGlobalAgent();
    console.log('[SSL] macOS detected - added Keychain certificates to trusted store');
  } catch (e) {
    console.warn('[SSL] mac-ca not available, using default certificate store:', e.message);
  }
}

// Windows: Create permissive agent for trusted VoApps domains
if (process.platform === 'win32') {
  httpsAgent = new https.Agent({
    rejectUnauthorized: false  // Skip SSL verification for VoApps API (trusted domain)
  });
  console.log('[SSL] Windows detected - using permissive HTTPS agent for VoApps API');
}
const fs = require("fs");
const fsp = require("fs/promises");
const path = require("path");
const os = require("os");
const { parse: parseUrl } = require("url");
const { createWriteStream } = require('fs');
const { generateTrendAnalysis, inferMessageIntent } = require("./trendAnalyzer");
const { Worker } = require('worker_threads');
const { VERSION, VERSION_NAME } = require('./version');

// ============================================================================
// WINDOWS — pre-patch DLL search path for onnxruntime-node
// ============================================================================
// @xenova/transformers statically imports onnxruntime-node, which in turn
// loads onnxruntime_binding.node via require().  Electron/ASAR redirects the
// .node file load to app.asar.unpacked, but the sibling DLLs
// (onnxruntime.dll, onnxruntime_providers_shared.dll) are also in that
// directory.  Windows' LoadLibraryW does NOT automatically search the
// directory of the .node being loaded — it checks the EXE directory, system
// dirs, and PATH.  Adding the unpacked bin directory to PATH before the first
// import of @xenova/transformers ensures the DLLs are found.
if (process.platform === 'win32') {
  try {
    const ortRelPath = path.join(
      'node_modules', 'onnxruntime-node', 'bin', 'napi-v3', 'win32', 'x64'
    );
    const ortDirPacked   = path.join(__dirname, ortRelPath);
    // When running from a packaged .asar, __dirname ends in "app.asar/…"
    const ortDirUnpacked = ortDirPacked.replace(/(app\.asar)([/\\])/, '$1.unpacked$2');
    const ortDir = fs.existsSync(ortDirUnpacked) ? ortDirUnpacked
                 : fs.existsSync(ortDirPacked)   ? ortDirPacked
                 : null;
    if (ortDir && !(process.env.PATH || '').includes(ortDir)) {
      process.env.PATH = ortDir + path.delimiter + (process.env.PATH || '');
      console.log(`[ONNX] Added DLL search path: ${ortDir}`);
    }
  } catch (e) {
    console.warn('[ONNX] Could not patch DLL search path:', e.message);
  }
}

// ============================================================================
// AI MESSAGE ANALYSIS — model status, download, transcription pipeline
// ============================================================================

// Electron patches the global import() function to redirect ASAR paths.
// That patch can prevent ESM-only packages like @xenova/transformers from loading
// in the main process.  new Function() creates an import() call in a fresh
// function scope that bypasses the patch and goes through V8's native ESM loader.
const esImport = new Function('p', 'return import(p)');

// Cached @xenova/transformers module reference — set on first successful load.
// Keeps us from re-importing the module on every transcription/classification call.
let _xenovaMod = null;

/**
 * Return the path that @xenova/transformers uses as its FileCache root.
 * - Dev:     <project>/node_modules/@xenova/transformers/.cache
 * - Packaged: resources/app.asar.unpacked/node_modules/@xenova/transformers/.cache
 *
 * This function is the single source of truth so both getXenovaMod (which sets
 * env.cacheDir) and getAiModelStatus (which checks for downloaded model dirs)
 * always agree on the same location.
 */
function xenovaCacheDir() {
  const pkgDir = xenovaPkgDir();          // resolves to unpacked dir in production
  return path.join(pkgDir, '.cache') + path.sep;
}

/**
 * Import @xenova/transformers once, configure ONNX Runtime log severity so the
 * thousands of "Removing initializer" C++ optimizer warnings are suppressed, then
 * cache the module reference for all subsequent calls.
 *
 * Also explicitly pins env.cacheDir to the known on-disk path so that the
 * FileCache lookup always points to the correct location regardless of how
 * import.meta.url resolves inside a packaged Electron asar.
 */
async function getXenovaMod(log, fallbackEntryPath = null) {
  if (_xenovaMod) return _xenovaMod;
  try {
    const mod = fallbackEntryPath
      ? await esImport(fallbackEntryPath)
      : await esImport('@xenova/transformers');
    // Suppress ONNX Runtime C++ INFO/WARNING logs (graph optimiser "Removing initializer" spam).
    // logSeverityLevel: 0=verbose 1=info 2=warning 3=error — set to 3 to show errors only.
    if (mod.env?.onnx !== undefined) mod.env.onnx = { logSeverityLevel: 3 };
    else if (mod.env) mod.env.onnx = { logSeverityLevel: 3 };

    // Explicitly set cacheDir so FileCache always looks at the real on-disk path.
    // When esImport() bypasses Electron's asar patching, import.meta.url inside
    // env.js can resolve to an asar-virtual path, making cacheDir point inside
    // the archive rather than to app.asar.unpacked — causing all ONNX cache lookups
    // to fail with a "file not found" miss even when files are physically present.
    if (mod.env) {
      const cacheDir = xenovaCacheDir();
      mod.env.cacheDir = cacheDir;
      if (log) log(`[AI] Cache dir: ${cacheDir}`);
    }

    _xenovaMod = mod;
    return mod;
  } catch (e) {
    if (log) log(`[AI] ⚠️  getXenovaMod failed: ${e.message}`, true);
    throw e;
  }
}

const AI_MODEL_STATUS = { stt: { downloaded: false }, intent: { downloaded: false } };

function getAiModelStatus() {
  // Check the @xenova/transformers FileCache directory for downloaded model folders.
  // The FileCache stores files as: <cacheDir>/Xenova/<model-name>/<files>
  // (NOT the Python HuggingFace Hub format at ~/.cache/huggingface/hub/models--...)
  const cacheRoot = xenovaCacheDir();
  const sttModelDir    = path.join(cacheRoot, 'Xenova', 'whisper-base');
  const intentModelDir = path.join(cacheRoot, 'Xenova', 'nli-deberta-v3-small');
  AI_MODEL_STATUS.stt.downloaded    = fs.existsSync(sttModelDir);
  AI_MODEL_STATUS.intent.downloaded = fs.existsSync(intentModelDir);
  return AI_MODEL_STATUS;
}

/**
 * Return the on-disk path to the @xenova/transformers package directory.
 *
 * In development __dirname is the project root, so the package is at
 * __dirname/node_modules/@xenova/transformers.
 *
 * In a packaged Electron app the source is inside app.asar (a virtual archive
 * that Node's native ESM loader cannot read via file:// URLs).  electron-builder
 * extracts asarUnpack entries to app.asar.unpacked/ next to the archive, so we
 * look there instead.  Both paths are tried; whichever exists wins.
 */
function xenovaPkgDir() {
  const inAsar = __dirname.includes('app.asar');
  if (inAsar) {
    // Primary: extracted via asarUnpack
    const unpackedDir = path.join(process.resourcesPath, 'app.asar.unpacked',
      'node_modules', '@xenova', 'transformers');
    if (fs.existsSync(path.join(unpackedDir, 'package.json'))) return unpackedDir;
    // Fallback: resources root (user-installed after first run)
    const resourcesDir = path.join(process.resourcesPath,
      'node_modules', '@xenova', 'transformers');
    if (fs.existsSync(path.join(resourcesDir, 'package.json'))) return resourcesDir;
  }
  return path.join(__dirname, 'node_modules', '@xenova', 'transformers');
}

/**
 * Return a writable directory to use as the cwd for npm install.
 * Inside a packaged ASAR __dirname is a virtual path — spawning npm with
 * cwd set to it fails with ENOTDIR.  Use process.resourcesPath instead,
 * which is the real on-disk directory containing app.asar.
 */
function xenovaInstallDir() {
  return __dirname.includes('app.asar') ? process.resourcesPath : __dirname;
}

/**
 * Auto-install @xenova/transformers via npm, streaming output to the log function.
 */
function installXenovaTransformers(log) {
  return new Promise((resolve, reject) => {
    const { spawn } = require('child_process');
    const npmBin = process.platform === 'win32' ? 'npm.cmd' : 'npm';
    const appDir = xenovaInstallDir();
    // Quick check: if npm.cmd can't be found at all, give a helpful message
    if (process.platform === 'win32') {
      const { execSync } = require('child_process');
      let npmFound = false;
      try { execSync('where.exe npm.cmd', { timeout: 2000 }); npmFound = true; } catch (e) { /* not in PATH */ }
      if (!npmFound) {
        return reject(new Error(
          'Node.js/npm not found. AI features require Node.js — please install it from https://nodejs.org and restart the app.'
        ));
      }
    }

    log('[AI] Installing @xenova/transformers (this may take a minute)…');
    const child = spawn(npmBin, ['install', '@xenova/transformers', '--no-audit', '--no-fund'], {
      cwd: appDir,
      stdio: ['ignore', 'pipe', 'pipe']
    });

    child.stdout.on('data', (data) => {
      data.toString().split('\n').filter(l => l.trim()).forEach(line => log(`[npm] ${line}`));
    });
    child.stderr.on('data', (data) => {
      data.toString().split('\n').filter(l => l.trim()).forEach(line => log(`[npm] ${line}`));
    });
    child.on('close', (code) => {
      if (code === 0) {
        log('[AI] ✅ @xenova/transformers installed successfully');
        resolve();
      } else {
        reject(new Error(`npm install exited with code ${code}`));
      }
    });
    child.on('error', reject);
  });
}

async function downloadAiModelBackground(type, log = (msg, isError = false) => isError ? console.error(msg) : console.log(msg)) {
  const label = type === 'stt' ? 'Whisper (STT)' : 'Intent (nli-deberta-v3-small)';
  const modelId = type === 'stt' ? 'Xenova/whisper-base' : 'Xenova/nli-deberta-v3-small';
  log(`[AI] Starting ${label} model download: ${modelId}`);

  // Import @xenova/transformers — auto-install if missing.
  // IMPORTANT: After installation we import via explicit file URL rather than the bare specifier,
  // because Node.js/Electron caches the "not found" result for bare specifiers at startup and
  // that cache entry persists even after npm installs the package in the same session.
  // File URL imports have a separate cache key and always resolve fresh from disk.
  let pipeline;

  const tryBareImport = async () => {
    try { return (await getXenovaMod(log)).pipeline; }
    catch (e) {
      if (e.code === 'ERR_MODULE_NOT_FOUND' || e.message?.includes('Cannot find package')) {
        return null; // not installed — fall through to auto-install path
      }
      // Module found but failed to initialize (e.g. onnxruntime-node DLL issue on Windows).
      // Log the full stack for diagnosis, then return null so the file-URL fallback is tried.
      log(`[AI] ⚠️  @xenova/transformers init error: ${e.message}`, true);
      if (e.stack) log(`[AI]    Stack: ${e.stack}`, true);
      if (process.platform === 'win32') {
        log('[AI]    Windows: if this is a DLL error, the app will retry via file-URL import.', true);
      }
      return null;
    }
  };

  pipeline = await tryBareImport();

  if (!pipeline) {
    const pkgDir = xenovaPkgDir();
    const alreadyOnDisk = fs.existsSync(path.join(pkgDir, 'package.json'));

    if (!alreadyOnDisk) {
      // Package not installed at all — auto-install it
      try {
        await installXenovaTransformers(log);
      } catch (installErr) {
        log(`[AI] ❌ Install failed: ${installErr.message}`, true);
        return;
      }
    }
    // Package is on disk (either just installed, or was present but the bare
    // import failed due to an init error like the sharp stub issue).
    // Import via file URL to bypass the ESM specifier cache.
    try {
      const { pathToFileURL } = require('url');
      const pkgJson = JSON.parse(await fsp.readFile(path.join(pkgDir, 'package.json'), 'utf8'));
      const exp = pkgJson.exports?.['.'];
      const mainFile = (typeof exp === 'string' ? exp
        : exp?.import || exp?.default || exp?.require
          || pkgJson.module || pkgJson.main
          || 'src/transformers.js');
      const entryPath = path.join(pkgDir, mainFile.replace(/^\.\//, ''));
      log('[AI] Loading module from disk…');
      ({ pipeline } = await getXenovaMod(log, pathToFileURL(entryPath).href));
    } catch (retryErr) {
      log(`[AI] ❌ Failed to load: ${retryErr.message}`, true);
      if (!alreadyOnDisk) log('[AI] Please restart VoApps Tools and try again.', true);
      return;
    }
  }

  try {
    // Track last reported % per file to avoid flooding the log
    const lastPct = {};
    const progress_callback = (data) => {
      if (data.status === 'initiate') {
        log(`[AI] Fetching: ${data.file}`);
      } else if (data.status === 'progress' && data.file) {
        const pct = Math.round(data.progress || 0);
        const prev = lastPct[data.file] ?? -1;
        if (pct - prev >= 10 || pct === 100) {
          lastPct[data.file] = pct;
          log(`[AI] ${data.file}: ${pct}%`);
        }
      } else if (data.status === 'done' && data.file) {
        log(`[AI] ✅ ${data.file} ready`);
      }
    };

    if (type === 'stt') {
      log('[AI] Loading Whisper base model (~142 MB)…');
      await pipeline('automatic-speech-recognition', modelId, { progress_callback });
    } else {
      log('[AI] Loading nli-deberta-v3-small model (~85 MB)…');
      await pipeline('zero-shot-classification', modelId, { progress_callback });
    }
    log(`[AI] ✅ ${label} model downloaded and cached`);
  } catch (e) {
    // "Unsupported model type: whisper" is a misleading error that surfaces on
    // Windows when the real cause is a network failure.  What actually happens:
    //   1. AutoModelForSpeechSeq2Seq.from_pretrained() fails → network error
    //   2. Falls through to AutoModelForCTC.from_pretrained()
    //   3. AutoModelForCTC has no 'whisper' mapping → throws "Unsupported model type"
    //   4. The last error wins, hiding the real network error
    // Similarly, "fetch failed" on any model = huggingface.co is unreachable.
    const isNetworkMasked = e.message?.includes('Unsupported model type');
    const isFetchFailed   = e.message?.toLowerCase().includes('fetch failed')
                         || e.message?.toLowerCase().includes('network')
                         || e.message?.toLowerCase().includes('enotfound')
                         || e.message?.toLowerCase().includes('econnrefused');
    if (isNetworkMasked || isFetchFailed) {
      log(`[AI] ❌ Failed to download ${label} model: could not reach huggingface.co`, true);
      log(`[AI]    Likely cause: the network is blocking huggingface.co (firewall, proxy, or no internet).`, true);
      log(`[AI]    Workaround: switch Transcription and Intent modes to "OpenAI" in Settings.`, true);
    } else {
      log(`[AI] ❌ Failed to download ${label} model: ${e.message}`, true);
    }
  }
}

/**
 * Transcribe and analyze messages using AI (local Whisper or OpenAI Whisper API).
 * Results are cached in message_transcriptions DuckDB table.
 * Returns transcriptMap: { "accountId:messageId" -> {transcript, intent, intent_summary, mentioned_phone, mentions_url} }
 */
async function transcribeAndAnalyzeMessages(messageInfo, aiSettings, log) {
  const transcriptMap = {};
  if (!dbReady) log('[AI] Note: Database not ready — transcripts will run but will not be cached this session');

  const { transcriptionMode = 'local', intentMode = 'local', openaiApiKey = '', notify = null } = aiSettings;
  let quotaExceeded = false;   // set to true on first 429 — skip remaining messages

  // Collect unique messages that have a file_url
  const uniqueMessages = new Map(); // key -> { messageId, accountId, file_url }
  for (const [key, info] of Object.entries(messageInfo)) {
    if (info.file_url && !uniqueMessages.has(key)) {
      const [accountId, messageId] = key.split(':');
      uniqueMessages.set(key, { messageId, accountId, file_url: info.file_url });
    }
  }

  if (uniqueMessages.size === 0) {
    log('[AI] No message audio URLs found — skipping AI analysis');
    return transcriptMap;
  }

  const sttLabel    = transcriptionMode === 'openai' ? 'OpenAI Whisper' : 'local Whisper';
  const intentLabel = intentMode === 'openai' ? 'GPT-4o-mini' : 'local nli-deberta';
  log(`[AI] Analyzing ${uniqueMessages.size} unique message(s) — STT: ${sttLabel}, intent: ${intentLabel}`);

  for (const [key, { messageId, accountId, file_url }] of uniqueMessages) {
    try {
      // Check cache (skip when DB not available — e.g. Windows)
      let cached = null;
      if (dbReady) {
        cached = await runQuery(
          `SELECT transcript, intent, intent_summary, mentioned_phone, mentions_url FROM message_transcriptions WHERE message_id = ? AND account_id = ?`,
          [messageId, accountId]
        );
      }
      if (cached && cached.length > 0) {
        const c = cached[0];
        const cachedTranscript = c.transcript || '';
        // Skip stale hallucination results — re-transcribe rather than use garbage cache
        if (isWhisperHallucination(cachedTranscript) && !c.intent) {
          const msgName = messageInfo[key]?.name || messageId;
          log(`[AI]   ♻️  Stale hallucination in cache for ${msgName} — re-transcribing`);
          // Fall through to re-transcribe
        } else {
          transcriptMap[key] = {
            transcript: cachedTranscript,
            intent: c.intent || '',
            intent_summary: c.intent_summary || '',
            mentioned_phone: c.mentioned_phone || '',
            mentions_url: !!c.mentions_url
          };
          const msgName = messageInfo[key]?.name || messageId;
          log(`[AI]   [cached] ${msgName} → ${c.intent || 'unknown'}`);
          continue;
        }
      }

      // Download audio to temp file
      const tmpFile = path.join(os.tmpdir(), `voapps_msg_${messageId}.mp3`);
      try {
        await downloadFile(file_url, tmpFile, log);
      } catch (dlErr) {
        log(`[AI]   ⚠️  Could not download audio for ${messageId}: ${dlErr.message}`);
        continue;
      }

      const t0 = Date.now();
      let transcript = '';

      // Skip remaining messages if OpenAI quota was already exceeded
      if (quotaExceeded) {
        try { fs.unlinkSync(tmpFile); } catch (_) {}
        continue;
      }

      // Transcribe
      if (transcriptionMode === 'openai') {
        if (!openaiApiKey) { log('[AI]   ⚠️  OpenAI API key not set — skipping transcription'); fs.unlinkSync(tmpFile); continue; }
        transcript = await transcribeWithOpenAI(tmpFile, openaiApiKey, log, () => {
          quotaExceeded = true;
          if (notify) {
            notify(
              'OpenAI quota exceeded — add credits to your account and run the analysis again.',
              'Add Credits',
              'https://platform.openai.com/docs/guides/error-codes/api-errors'
            );
          }
        });
        if (transcript === '__QUOTA_EXCEEDED__') transcript = '';
      } else {
        transcript = await transcribeWithLocalWhisper(tmpFile, log);
      }

      // Clean up temp file
      try { fs.unlinkSync(tmpFile); } catch (_) {}

      if (!transcript) {
        log(`[AI]   ⚠️  No transcript for ${messageId}`);
        continue;
      }

      // Extract metadata from transcript (no API needed)
      const mentionedPhone = extractMentionedPhone(transcript);
      const mentionsUrl    = detectUrlMention(transcript);

      // Intent
      let intent = '';
      let intentModelLabel = '';
      let sttModelLabel = transcriptionMode === 'openai' ? 'openai-whisper-1' : 'whisper-base-local';

      // Resolve message name early — used both as a classifier hint and for logging.
      const msgName = messageInfo[key]?.name || messageId;

      if (intentMode === 'openai') {
        if (!openaiApiKey) { log('[AI]   ⚠️  OpenAI API key not set — using name-based intent'); }
        else {
          const result = await classifyIntentWithOpenAI(transcript, openaiApiKey, log, msgName);
          intent = result.intent;
          intentModelLabel = 'openai-gpt4o-mini';
        }
      } else {
        const result = await classifyIntentLocal(transcript, log, msgName);
        intent = result.intent;
        intentModelLabel = 'nli-deberta-v3-local';
      }

      // Final fallback: if neither AI path returned an intent, derive it from
      // the message name alone (no model needed, purely pattern-based).
      if (!intent) {
        const nameHint = inferMessageIntent(msgName);
        if (nameHint && nameHint !== 'general notice' && nameHint !== 'unknown') {
          intent = nameHint;
          intentModelLabel = 'name-based';
          log(`[AI]   💡 No AI intent returned — using name-based: ${intent}`);
        }
      }

      const elapsed = ((Date.now() - t0) / 1000).toFixed(1);
      const urlNote = mentionsUrl ? ' | mentions URL' : '';
      log(`[AI]   [transcribed] ${msgName} → ${intent || 'unknown'} | ${elapsed}s${urlNote}`);

      // Cache in DB (skip when DB not available — e.g. Windows; results still returned in-memory)
      if (dbReady) {
        await runQuery(
          `INSERT OR REPLACE INTO message_transcriptions
             (message_id, account_id, audio_url, transcript, intent, intent_summary,
              mentioned_phone, mentions_url, stt_model, intent_model, transcribed_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)`,
          [messageId, accountId, file_url, transcript, intent, '',
           mentionedPhone || null, mentionsUrl ? 1 : 0, sttModelLabel, intentModelLabel]
        );
      }

      transcriptMap[key] = { transcript, intent, intent_summary: '', mentioned_phone: mentionedPhone || '', mentions_url: mentionsUrl };

    } catch (err) {
      log(`[AI]   ⚠️  Error processing message ${messageId}: ${err.message}`);
    }
  }

  return transcriptMap;
}

/**
 * Scan CSV rows for VoApps result codes 408/409/410 and log warnings.
 * Returns a summary object for use in API responses.
 */
function checkInvalidResultCodes(rows, log) {
  const INVALID_CODES = {
    '408': 'Invalid Caller Number',
    '409': 'Invalid Message ID',
    '410': 'Prohibited Self Call'
  };
  const found = {};  // code -> { count, campaigns: Set }

  for (const row of rows) {
    const code = String(row.voapps_code || '').trim();
    if (INVALID_CODES[code]) {
      if (!found[code]) found[code] = { count: 0, campaigns: new Set() };
      found[code].count++;
      const cid = row.campaign_id || row.campaign_name || 'unknown';
      found[code].campaigns.add(cid);
    }
  }

  const alerts = [];
  for (const [code, info] of Object.entries(found)) {
    const label = INVALID_CODES[code];
    const campaignList = [...info.campaigns].slice(0, 10).join(', ');
    const moreNote = info.campaigns.size > 10 ? ` (+${info.campaigns.size - 10} more)` : '';
    const msg = `⚠️  ${label} (${code}): ${info.count.toLocaleString()} result(s) in campaigns: ${campaignList}${moreNote}`;
    log(msg, true);  // log as warning
    alerts.push({ code, label, count: info.count, campaigns: [...info.campaigns] });
  }

  if (alerts.length > 0) {
    log('⚠️  Review your campaign configuration for the codes above. Invalid results do not count toward delivery but may indicate misconfigured caller numbers or message IDs.', true);
  }

  return alerts;
}

async function downloadFile(url, destPath, log, maxRedirects = 5) {
  return new Promise((resolve, reject) => {
    const urlObj = new URL(url);
    const mod = urlObj.protocol === 'https:' ? require('https') : require('http');
    const file = fs.createWriteStream(destPath);
    mod.get(url, (res) => {
      // Follow redirects (301, 302, 303, 307, 308)
      if ([301, 302, 303, 307, 308].includes(res.statusCode)) {
        file.close();
        const location = res.headers?.location;
        if (!location || maxRedirects <= 0) {
          reject(new Error(`HTTP ${res.statusCode} redirect loop or no Location header`));
          return;
        }
        if (log) log(`[AI]   🔍 Audio redirect ${res.statusCode} → ${location.slice(0, 80)}`);
        // Recurse with decremented redirect count and a fresh stream
        downloadFile(location, destPath, log, maxRedirects - 1).then(resolve).catch(reject);
        return;
      }
      if (res.statusCode !== 200) { file.close(); reject(new Error(`HTTP ${res.statusCode}`)); return; }
      const contentType = res.headers?.['content-type'] || 'unknown';
      const contentLength = res.headers?.['content-length'];
      if (log) log(`[AI]   🔍 Audio download: type=${contentType}, size=${contentLength ? Math.round(contentLength/1024)+'KB' : 'unknown'}`);
      res.pipe(file);
      file.on('finish', () => { file.close(); resolve(); });
    }).on('error', (e) => { file.close(); fs.unlink(destPath, () => {}); reject(e); });
  });
}

async function transcribeWithOpenAI(audioPath, apiKey, log, onQuotaExceeded = null) {
  try {
    let FormData;
    try {
      FormData = require('form-data');
    } catch (e) {
      await installNpmPackage('form-data', log);
      FormData = require('form-data');
    }
    const form = new FormData();
    form.append('file', fs.createReadStream(audioPath), path.basename(audioPath));
    form.append('model', 'whisper-1');

    // crossPlatformFetch serializes the body via req.write(), which converts a FormData
    // object to "[object Object]". npm form-data must be piped directly into the request
    // so the multipart stream is properly transmitted with correct boundaries.
    return new Promise((resolve) => {
      const req = https.request({
        hostname: 'api.openai.com',
        port: 443,
        path: '/v1/audio/transcriptions',
        method: 'POST',
        headers: { ...form.getHeaders(), Authorization: `Bearer ${apiKey}` },
        timeout: 60000
      }, (res) => {
        let data = '';
        res.on('data', chunk => data += chunk);
        res.on('end', () => {
          if (res.statusCode === 429) {
            log(`[AI]   ⚠️  OpenAI STT quota exceeded (429) — add credits and retry`);
            if (onQuotaExceeded) onQuotaExceeded();
            resolve('__QUOTA_EXCEEDED__');
          } else if (res.statusCode < 200 || res.statusCode >= 300) {
            log(`[AI]   ⚠️  OpenAI STT failed: OpenAI STT error ${res.statusCode}: ${data}`);
            resolve('');
          } else {
            try {
              resolve(JSON.parse(data).text || '');
            } catch (e) {
              log(`[AI]   ⚠️  OpenAI STT response parse error: ${e.message}`);
              resolve('');
            }
          }
        });
      });
      req.on('error', (err) => {
        log(`[AI]   ⚠️  OpenAI STT failed: ${err.message}`);
        resolve('');
      });
      req.on('timeout', () => {
        req.destroy();
        log('[AI]   ⚠️  OpenAI STT failed: Request timeout (60s)');
        resolve('');
      });
      form.pipe(req);
    });
  } catch (e) {
    log(`[AI]   ⚠️  OpenAI STT failed: ${e.message}`);
    return '';
  }
}

/**
 * Generic npm package installer — streams output through the log callback.
 * Works the same way as installXenovaTransformers but for any package name.
 */
function installNpmPackage(pkgName, log) {
  return new Promise((resolve, reject) => {
    const { spawn } = require('child_process');
    const npmBin = process.platform === 'win32' ? 'npm.cmd' : 'npm';
    log(`[AI] Installing ${pkgName}…`);
    const child = spawn(npmBin, ['install', pkgName, '--no-audit', '--no-fund'], {
      cwd: __dirname, stdio: ['ignore', 'pipe', 'pipe']
    });
    child.stdout.on('data', d => d.toString().split('\n').filter(l => l.trim()).forEach(l => log(`[npm] ${l}`)));
    child.stderr.on('data', d => d.toString().split('\n').filter(l => l.trim()).forEach(l => log(`[npm] ${l}`)));
    child.on('close', code => code === 0 ? resolve() : reject(new Error(`npm install exited with code ${code}`)));
    child.on('error', reject);
  });
}

/**
 * Decode a RIFF/WAVE PCM file directly from a Buffer — no external library needed.
 * Supports 8/16/24/32-bit mono or stereo PCM (audioFormat = 1).
 * Returns a 16 kHz mono Float32Array ready for Whisper.
 */
function decodeWavPcm(buffer, log) {
  const riff = buffer.slice(0, 4).toString('ascii');
  const wave = buffer.slice(8, 12).toString('ascii');
  if (riff !== 'RIFF' || wave !== 'WAVE') throw new Error('Not a valid RIFF/WAVE file');

  let offset = 12;
  let sampleRate = 0, numChannels = 0, bitsPerSample = 0;
  let dataOffset = 0, dataSize = 0;

  while (offset + 8 <= buffer.length) {
    const chunkId   = buffer.slice(offset, offset + 4).toString('ascii');
    const chunkSize = buffer.readUInt32LE(offset + 4);
    offset += 8;

    if (chunkId === 'fmt ') {
      const audioFormat = buffer.readUInt16LE(offset);
      numChannels       = buffer.readUInt16LE(offset + 2);
      sampleRate        = buffer.readUInt32LE(offset + 4);
      bitsPerSample     = buffer.readUInt16LE(offset + 14);
      if (audioFormat !== 1) {
        throw new Error(`WAV audio format ${audioFormat} not supported (only PCM=1). Re-export the message as uncompressed WAV.`);
      }
    } else if (chunkId === 'data') {
      dataOffset = offset;
      dataSize   = chunkSize;
      break;  // done scanning — data chunk found
    }

    offset += chunkSize + (chunkSize & 1); // advance (chunks are word-aligned)
  }

  if (!sampleRate || !dataOffset) throw new Error('WAV file missing fmt or data chunk');

  const bytesPerSample    = bitsPerSample >> 3;
  const frameSize         = numChannels * bytesPerSample;
  const samplesPerChannel = Math.floor(dataSize / frameSize);

  log(`[AI]   🔍 Audio decoded: ${sampleRate} Hz, ${samplesPerChannel} samples, ${(samplesPerChannel / sampleRate).toFixed(1)}s (WAV PCM ${bitsPerSample}-bit, ${numChannels}ch)`);

  // Mix channels to mono Float32 in [-1, 1]
  const scale = bitsPerSample === 8 ? 128 : Math.pow(2, bitsPerSample - 1);
  const out   = new Float32Array(samplesPerChannel);

  for (let i = 0; i < samplesPerChannel; i++) {
    const basePos = dataOffset + i * frameSize;
    let sum = 0;
    for (let ch = 0; ch < numChannels; ch++) {
      const pos = basePos + ch * bytesPerSample;
      let val;
      if (bitsPerSample === 8) {
        val = buffer.readUInt8(pos) - 128;  // 8-bit WAV is unsigned; center at 0
      } else if (bitsPerSample === 16) {
        val = buffer.readInt16LE(pos);
      } else if (bitsPerSample === 24) {
        // 3-byte little-endian signed integer
        let u = buffer.readUInt8(pos) | (buffer.readUInt8(pos + 1) << 8) | (buffer.readUInt8(pos + 2) << 16);
        if (u & 0x800000) u |= 0xFF000000;  // sign-extend to 32-bit
        val = u | 0;
      } else if (bitsPerSample === 32) {
        val = buffer.readInt32LE(pos);
      } else {
        val = 0;
      }
      sum += val;
    }
    out[i] = (sum / numChannels) / scale;
  }

  return resampleTo16kHz(out, sampleRate);
}

/**
 * Linear-interpolation resampler — converts a Float32Array from sourceSampleRate to 16 kHz.
 * Whisper expects 16 kHz mono audio.
 */
function resampleTo16kHz(samples, sourceSampleRate) {
  if (sourceSampleRate === 16000) return samples;
  const ratio     = sourceSampleRate / 16000;
  const outLength = Math.ceil(samples.length / ratio);
  const output    = new Float32Array(outLength);
  for (let i = 0; i < outLength; i++) {
    const srcIdx = i * ratio;
    const lo = Math.floor(srcIdx);
    const hi = Math.min(lo + 1, samples.length - 1);
    const t  = srcIdx - lo;
    output[i] = samples[lo] * (1 - t) + samples[hi] * t;
  }
  return output;
}

/**
 * Decode an audio file (WAV or MP3) to a 16 kHz mono Float32Array for Whisper.
 *
 * WAV/RIFF files: decoded natively via decodeWavPcm() — no external library needed.
 *   VoApps DDVM message files are typically 16 kHz mono PCM WAV, so no resampling needed.
 *
 * MP3 files: decoded via mpg123-decoder (pure-WASM, auto-installed on first use).
 *
 * Why we can't pass file paths to the Whisper pipeline directly: @xenova/transformers
 * falls back to AudioContext when given a path, but AudioContext is a browser API
 * unavailable in Node.js / Electron main process. A pre-decoded Float32Array bypasses this.
 */
async function decodeMp3File(audioPath, log) {
  // Read once — used for format detection and decoding
  const fileBuffer = fs.readFileSync(audioPath);
  const magic = fileBuffer.slice(0, 4).toString('ascii');

  // ── WAV/RIFF: parse PCM directly (no external library) ──────────────────────
  if (magic === 'RIFF') {
    log(`[AI]   🔍 Detected WAV format — using native PCM decoder`);
    return decodeWavPcm(fileBuffer, log);
  }

  // ── MP3: use mpg123-decoder (WASM-based, no native compilation) ─────────────
  // Helper to import mpg123-decoder via file URL (bypasses ESM specifier cache after install)
  const importMpg123 = async () => {
    const { pathToFileURL } = require('url');
    const pkgDir  = path.join(__dirname, 'node_modules', 'mpg123-decoder');
    const pkgJson = JSON.parse(await fsp.readFile(path.join(pkgDir, 'package.json'), 'utf8'));
    const exp = pkgJson.exports?.['.'];
    const mainFile = (typeof exp === 'string' ? exp
      : exp?.import || exp?.default || exp?.require
        || pkgJson.module || pkgJson.main
        || 'src/mpg123-decoder.js');
    const entryPath = path.join(pkgDir, mainFile.replace(/^\.\//, ''));
    return await esImport(pathToFileURL(entryPath).href);
  };

  let MPEGDecoder;
  try {
    ({ MPEGDecoder } = await import('mpg123-decoder'));
  } catch (e) {
    if (e.code === 'ERR_MODULE_NOT_FOUND' || e.message?.includes('Cannot find')) {
      // Auto-install mpg123-decoder on first use (~71 KB WASM, no native compilation)
      await installNpmPackage('mpg123-decoder', log);
      ({ MPEGDecoder } = await importMpg123());
    } else {
      throw e;
    }
  }

  const decoder = new MPEGDecoder();
  await decoder.ready;
  // decode() is synchronous — takes Uint8Array/Buffer, returns { channelData, samplesDecoded, sampleRate }
  const { channelData, samplesDecoded, sampleRate } = decoder.decode(fileBuffer);
  decoder.free();

  // Diagnostic: log decode results so we can verify audio is healthy before transcription
  log(`[AI]   🔍 Audio decoded: ${sampleRate} Hz, ${samplesDecoded} samples, ${(samplesDecoded / (sampleRate || 1)).toFixed(1)}s (MP3)`);

  if (!sampleRate || samplesDecoded === 0) {
    throw new Error('MP3 decode produced no samples — file may be corrupt or unsupported');
  }

  // Collapse stereo → mono (average L + R channels)
  let samples;
  if (Array.isArray(channelData) && channelData.length > 1) {
    const L = channelData[0];
    const R = channelData[1];
    samples = new Float32Array(L.length);
    for (let i = 0; i < L.length; i++) samples[i] = (L[i] + R[i]) * 0.5;
  } else {
    samples = Array.isArray(channelData) ? channelData[0] : channelData;
  }

  // Resample to 16 kHz — Whisper's required input sample rate
  return resampleTo16kHz(samples, sampleRate);
}

/**
 * Detect Whisper hallucination output — transcripts that contain ONLY non-speech
 * event tokens (music, chiming, applause, etc.) with no real words.
 * These occur when the audio is music, telephony tones, or spectrally unusual content.
 * Returns true if the transcript should be discarded.
 */
function isWhisperHallucination(text) {
  if (!text || !text.trim()) return true;
  // Strip all known Whisper event annotations: [Music], (chiming), [APPLAUSE], etc.
  const stripped = text
    .replace(/\[[\w\s]+\]/gi, '')      // [Music], [MUSIC], [Applause], [Laughter]
    .replace(/\([\w\s]+\)/gi, '')      // (chiming), (whistling), (whimsical music), (beeping)
    .replace(/\s+/g, ' ')
    .trim();
  // If nothing real is left, it's a hallucination
  if (!stripped) return true;
  // If real content is very short (< 3 chars after stripping), also discard
  if (stripped.length < 3) return true;
  return false;
}

/**
 * Stitch together transcript segments that were transcribed with overlapping audio windows.
 * Finds the longest word-level sequence that appears at the tail of one segment and the
 * head of the next, then concatenates without duplicating the overlap.
 *
 * If no overlap is found (e.g. segments cover non-overlapping audio), the segments are
 * simply joined with a space.
 */
function stitchTranscriptSegments(texts) {
  if (!texts.length) return '';
  let result = texts[0];
  for (let i = 1; i < texts.length; i++) {
    const next = (texts[i] || '').trim();
    if (!next) continue;
    const rWords = result.trim().split(/\s+/);
    const nWords = next.split(/\s+/);
    // Search for the longest common word-ngram (up to 15 words) at the seam.
    // Normalise to lowercase, strip punctuation for comparison only.
    const norm = (w) => w.toLowerCase().replace(/[^a-z0-9]/g, '');
    const rNorm = rWords.map(norm);
    const nNorm = nWords.map(norm);
    let stitchAt = 0;
    const maxSearch = Math.min(rWords.length, nWords.length, 15);
    outer: for (let len = maxSearch; len >= 2; len--) {
      const rTail = rNorm.slice(-len);
      for (let offset = 0; offset <= Math.min(nNorm.length - len, 5); offset++) {
        const nChunk = nNorm.slice(offset, offset + len);
        if (rTail.every((w, j) => w === nChunk[j])) {
          stitchAt = offset + len; // unique content starts here in next
          break outer;
        }
      }
    }
    if (stitchAt > 0) {
      result = result + ' ' + nWords.slice(stitchAt).join(' ');
    } else {
      result = result + ' ' + next;
    }
  }
  return result.trim();
}

async function transcribeWithLocalWhisper(audioPath, log) {
  try {
    // Import @xenova/transformers — auto-install if missing, file-URL fallback if specifier
    // cache is stale (e.g. package was just installed in this session).
    // getXenovaMod() also configures ONNX log severity to suppress optimizer noise.
    let pipelineFn;
    try {
      ({ pipeline: pipelineFn } = await getXenovaMod(log));
    } catch (e) {
      const isNotFound = e.code === 'ERR_MODULE_NOT_FOUND' || e.message?.includes('Cannot find package');
      const pkgDir = xenovaPkgDir();
      const alreadyOnDisk = fs.existsSync(path.join(pkgDir, 'package.json'));

      if (!isNotFound) {
        log(`[AI]   ⚠️  @xenova/transformers init error: ${e.message}`, true);
        if (e.stack) log(`[AI]      ${e.stack.split('\n').slice(1, 4).join(' | ')}`, true);
      } else if (!alreadyOnDisk) {
        log('[AI]   Installing @xenova/transformers…');
        await installXenovaTransformers(log);
      }
      // File-URL import bypasses the ESM specifier cache — also sets ONNX log severity.
      try {
        const { pathToFileURL } = require('url');
        const pkgJson = JSON.parse(await fsp.readFile(path.join(pkgDir, 'package.json'), 'utf8'));
        const exp = pkgJson.exports?.['.'];
        const mainFile = (typeof exp === 'string' ? exp
          : exp?.import || exp?.default || exp?.require
            || pkgJson.module || pkgJson.main
            || 'src/transformers.js');
        const entryPath = path.join(pkgDir, mainFile.replace(/^\.\//, ''));
        ({ pipeline: pipelineFn } = await getXenovaMod(log, pathToFileURL(entryPath).href));
      } catch (retryErr) {
        log(`[AI]   ❌ File-URL import also failed: ${retryErr.message}`, true);
        return '';
      }
    }

    // Node.js / Electron main process has no AudioContext — passing a file path to the
    // pipeline would fail with "Unable to load audio … AudioContext is not available".
    // Instead: decode the MP3 ourselves to a 16 kHz mono Float32Array and pass that directly.
    const audioData = await decodeMp3File(audioPath, log);

    // --- Amplitude checks on the FULL resampled audio --------------------------------
    // The earlier diagnostic only covers the first 1000 samples (≈62 ms), which is often
    // silence/intro.  We need the full-file peak for two reasons:
    //
    //   1. Some MP3 encoders / decoders produce values outside [-1, 1].  Whisper's
    //      WhisperFeatureExtractor computes a log-mel spectrogram that is calibrated for
    //      audio in that range.  Amplitude 10× too large shifts every log-mel bin by
    //      log₁₀(10) = 1, producing feature vectors completely outside the training
    //      distribution → the model hallucinates "(chiming)" / "(whistling)" etc.
    //
    //   2. Genuinely silent files (blank templates, unrecorded placeholders) should be
    //      skipped rather than wasting Whisper time producing noise-hallucinations.
    // ---------------------------------------------------------------------------------
    let peakAmp = 0;
    for (let i = 0; i < audioData.length; i++) {
      const abs = Math.abs(audioData[i]);
      if (abs > peakAmp) peakAmp = abs;
    }
    log(`[AI]   🔍 Peak amplitude (full resampled audio): ${peakAmp.toFixed(4)}`);

    if (peakAmp < 0.001) {
      log('[AI]   ⚠️  Audio is silent — no transcription');
      return '';
    }
    if (peakAmp > 1.0) {
      // Peak-normalize: bring the loudest sample to exactly ±1.0
      for (let i = 0; i < audioData.length; i++) audioData[i] /= peakAmp;
      log(`[AI]   🔍 Normalized amplitude from ${peakAmp.toFixed(4)} → 1.0`);
    }

    // Log audio duration for diagnostics
    const audioDurationSec = audioData.length / 16000;
    log(`[AI]   🔍 Audio duration: ${audioDurationSec.toFixed(1)}s (${audioData.length} samples @ 16kHz)`);

    const transcriber = await pipelineFn('automatic-speech-recognition', 'Xenova/whisper-base');
    // Pass the Float32Array directly — @xenova/transformers v2.x WhisperFeatureExtractor
    // requires a raw Float32Array, not a { data, sampling_rate } wrapper object.
    //
    // no_repeat_ngram_size: 3 — prevents Whisper's decoder from entering a repetition loop
    // (e.g. "registros" → "rosrosros...") that silently truncates the transcript.
    //
    // For audio > 30s we do NOT use the library's built-in chunk_length_s/stride_length_s.
    // That creates a very short final chunk (audio_len − 30s) padded with lots of silence,
    // which destabilises the decoder and causes the repetition loop even with no_repeat_ngram.
    // Instead we split manually with a generous 8s overlap so every segment is ≥ 18s of real
    // audio (well within Whisper-base's sweet spot), then stitch the segments at the text level.
    const WHISPER_SR      = 16000;
    const MANUAL_CHUNK_S  = 30;  // each audio segment fed to Whisper
    const MANUAL_OVERLAP_S = 8;  // overlap between consecutive segments for reliable stitching
    const passOpts = { no_repeat_ngram_size: 3 };

    let rawText;
    if (audioDurationSec > 30) {
      const segTexts = [];
      let segStart = 0;
      let segIdx   = 0;
      while (segStart < audioData.length) {
        const segEnd   = Math.min(segStart + MANUAL_CHUNK_S * WHISPER_SR, audioData.length);
        const segAudio = audioData.slice(segStart, segEnd);
        const segSec   = segAudio.length / WHISPER_SR;
        const segRes   = await transcriber(segAudio, passOpts);
        const segText  = (segRes.text || '')
          .replace(/(\s*\[S\])+/g, '')
          .replace(/\[BLANK_AUDIO\]/gi, '')
          .trim();
        log(`[AI]   🔍 Segment ${++segIdx} (${(segStart/WHISPER_SR)|0}–${(segEnd/WHISPER_SR)|0}s, ${segSec.toFixed(1)}s): ${JSON.stringify(segText.slice(0, 120))}`);
        segTexts.push(segText);
        if (segEnd >= audioData.length) break;
        segStart += (MANUAL_CHUNK_S - MANUAL_OVERLAP_S) * WHISPER_SR;
      }
      rawText = stitchTranscriptSegments(segTexts);
    } else {
      const result = await transcriber(audioData, passOpts);
      rawText = result.text || '';
    }
    log(`[AI]   🔍 Raw Whisper output: ${JSON.stringify(rawText.slice(0, 200))}`);

    // Strip Whisper silence/padding tokens — [S] appears when a chunk contains silence
    // after the speech ends (common at chunk boundaries or end of short audio padded to 30s).
    // Also strip [BLANK_AUDIO] and similar filler tokens that add no transcript value.
    const cleanText = rawText
      .replace(/(\s*\[S\])+/g, '')
      .replace(/\[BLANK_AUDIO\]/gi, '')
      .trim();

    // Discard hallucination-only output (e.g. "[Music] (chiming) [Music]" with no real words)
    if (isWhisperHallucination(cleanText)) {
      log(`[AI]   ⚠️  Whisper produced only non-speech tokens ("${rawText.slice(0, 80).trim()}") — discarding`);
      return '';
    }
    return cleanText;
  } catch (e) {
    if (e.code === 'ERR_MODULE_NOT_FOUND' || e.message?.includes('Cannot find')) {
      log('[AI]   ⚠️  @xenova/transformers not installed. Use OpenAI Whisper mode or download the model first.');
    } else {
      log(`[AI]   ⚠️  Local Whisper failed: ${e.message}`);
    }
    return '';
  }
}

// ---------------------------------------------------------------------------
// Pattern-based intent detection (runs before AI models)
// ---------------------------------------------------------------------------

// LCM = Limited Content Message: a message that intentionally avoids identifying
// the debt — just asks the recipient to call back. No payment/loan/debt terms.
const LCM_CALLBACK_PATTERN  = /please (return my call|call (?:me|us) back)\b|call (me|us)(?: or \w+)? back at\b|personal business matter/i;
const LCM_DEBT_TERMS        = /payment|past due|loan|account balance|debt|owe\b|credit|mortgage|delinquent|overdrawn|amount due/i;

function detectLCM(transcript) {
  if (!transcript) return false;
  return LCM_CALLBACK_PATTERN.test(transcript) && !LCM_DEBT_TERMS.test(transcript);
}

// Modified Zortman: message contains formal debt-collection disclosure language.
const MODIFIED_ZORTMAN_PATTERN = /this is an attempt to collect a debt|debt collector\b|any information obtained will be used for that purpose|fair debt collection practices act|fdcpa/i;

function detectModifiedZortman(transcript) {
  return MODIFIED_ZORTMAN_PATTERN.test(transcript || '');
}

// ---------------------------------------------------------------------------
// Shared intent label set (used by both local NLI and OpenAI)
// ---------------------------------------------------------------------------
const INTENT_LABELS = [
  // ── Collections ─────────────────────────────────────────────────────────
  'first payment default',
  'delinquent loan payment',
  'pre-charge-off collections',
  'overdrawn account',
  'friendly payment reminder',
  'healthcare or third-party collections',   // EBO, Early Out, 3rd-party, debt buyers
  // ── Consumer / Direct Lending ────────────────────────────────────────────
  'loan application follow-up',
  'pre-approval or refinance offer',
  'CPI or insurance notification',
  'title perfection reminder',
  'holiday skip payment offer',
  'unused rewards notification',
  // ── Servicing ───────────────────────────────────────────────────────────
  'dormant account notification',
  'fraud alert',
  'card services notification',              // card activation, card issues
  'system update or downtime notice',
  // ── Marketing / Outreach ─────────────────────────────────────────────────
  'new member welcome',
  'product or rate promotion',
  'educational event or workshop',
  'marketing or lead response',
  'disaster recovery or closure notice',
  // ── General ──────────────────────────────────────────────────────────────
  'general notice',
];

// Intent category groups used for the category-mismatch guard.
// If the message name clearly maps to one group but the AI returns the other,
// the name-based hint wins — AI models (especially local NLI on Spanish text)
// frequently confuse soft-spoken collections messages with marketing.
const COLLECTIONS_INTENTS = new Set([
  'first payment default', 'delinquent loan payment', 'pre-charge-off collections',
  'overdrawn account', 'friendly payment reminder', 'healthcare or third-party collections',
]);
const MARKETING_INTENTS = new Set([
  'marketing or lead response', 'product or rate promotion',
  'new member welcome', 'educational event or workshop',
  'pre-approval or refinance offer', 'holiday skip payment offer',
  'unused rewards notification',
]);

async function classifyIntentWithOpenAI(transcript, apiKey, log, messageName = '') {
  // Pattern-detect before calling the model
  if (detectLCM(transcript))             return { intent: 'LCM' };
  if (detectModifiedZortman(transcript)) return { intent: 'Modified Zortman' };

  // Use the message name as a classification hint when available.
  // inferMessageIntent maps common naming conventions (e.g. "Negative Shares 45+ DPD"
  // → "overdrawn account") so the AI isn't misled by a generic-sounding transcript.
  const nameHint = messageName ? inferMessageIntent(messageName) : '';
  const nameContext = messageName
    ? `\nMessage title: "${messageName}".` +
      (nameHint && nameHint !== 'general notice' && nameHint !== 'unknown'
        ? ` Based on this title alone, it is likely: "${nameHint}".`
        : '')
    : '';

  try {
    const body = JSON.stringify({
      model: 'gpt-4o-mini',
      messages: [{
        role: 'user',
        content: `You are classifying a short voicemail message left by a financial institution. Return ONLY valid JSON with:
- "intent": a concise label (3–6 words) describing the message purpose. Prefer one of these if it fits well: ${JSON.stringify(INTENT_LABELS)} — but use your own label if none of them accurately describes the message.
${nameContext}
Transcript:
"""
${transcript.slice(0, 1000)}
"""

Respond with only valid JSON, no markdown.`
      }],
      max_tokens: 30,
      temperature: 0
    });
    const response = await crossPlatformFetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${apiKey}` },
      body
    });
    if (!response.ok) throw new Error(`OpenAI chat error ${response.status}`);
    const data = await response.json();
    const raw = data.choices?.[0]?.message?.content || '{}';
    const parsed = JSON.parse(raw.replace(/```json|```/g, '').trim());
    const aiIntent = parsed.intent || '';

    // Category-mismatch guard: if the name says "collections" but the AI says
    // "marketing" (or vice versa), trust the name — the AI can be confused by
    // polite/indirect language, especially in Spanish.
    if (nameHint && COLLECTIONS_INTENTS.has(nameHint) && MARKETING_INTENTS.has(aiIntent)) {
      log(`[AI]   💡 Category mismatch (name: collections, GPT: marketing) — using name: ${nameHint}`);
      return { intent: nameHint };
    }

    return { intent: aiIntent };
  } catch (e) {
    log(`[AI]   ⚠️  OpenAI intent failed: ${e.message}`);
    return { intent: '' };
  }
}

async function classifyIntentLocal(transcript, log, messageName = '') {
  // Pattern-detect before running the model (deterministic, no model needed)
  if (detectLCM(transcript))             return { intent: 'LCM' };
  if (detectModifiedZortman(transcript)) return { intent: 'Modified Zortman' };

  // Pre-compute name-based hint so we can use it as a tiebreaker or fallback.
  const nameHint = messageName ? inferMessageIntent(messageName) : '';

  try {
    const { pipeline } = await getXenovaMod(log);
    const classifier = await pipeline('zero-shot-classification', 'Xenova/nli-deberta-v3-small');
    const result = await classifier(normalizeSttText(transcript).slice(0, 500), INTENT_LABELS);
    const topLabel = result.labels?.[0] || '';
    const topScore = result.scores?.[0] || 0;

    // When the NLI model is uncertain (score below threshold) and the message
    // name gives a specific, non-generic intent, trust the name-based hint.
    // The local NLI model can confuse collections messages with marketing when
    // the transcript uses polite/indirect language (e.g. "urgent message about
    // your account" without hard collections keywords in the top tokens).
    const CONFIDENCE_THRESHOLD = 0.30;
    if (nameHint && nameHint !== 'general notice' && nameHint !== 'unknown' &&
        topScore < CONFIDENCE_THRESHOLD) {
      log(`[AI]   💡 Low NLI confidence (${topScore.toFixed(2)}) — using name-based intent: ${nameHint}`);
      return { intent: nameHint };
    }

    // Category-mismatch guard: regardless of confidence, if the name clearly
    // signals a collections intent but the NLI returns marketing (common with
    // Spanish transcripts where the model is primarily English-trained), trust
    // the name.
    if (nameHint && COLLECTIONS_INTENTS.has(nameHint) && MARKETING_INTENTS.has(topLabel)) {
      log(`[AI]   💡 Category mismatch (name: collections, NLI: marketing) — using name: ${nameHint}`);
      return { intent: nameHint };
    }

    return { intent: topLabel };
  } catch (e) {
    if (e.code === 'ERR_MODULE_NOT_FOUND' || e.message?.includes('Cannot find')) {
      log('[AI]   ⚠️  @xenova/transformers not installed — using extractive intent only.');
    } else {
      log(`[AI]   ⚠️  Local intent classification failed: ${e.message}`);
    }
    // On complete model failure, fall back to name-based intent if available.
    if (nameHint && nameHint !== 'general notice' && nameHint !== 'unknown') {
      log(`[AI]   💡 Model unavailable — using name-based intent: ${nameHint}`);
      return { intent: nameHint };
    }
    return { intent: '' };
  }
}

// ---------------------------------------------------------------------------
// STT normalization dictionary
// Ordered [pattern, replacement] pairs applied to raw Whisper output before
// summary generation and NLI classification. Add entries as new errors appear.
// ---------------------------------------------------------------------------
const STT_CORRECTIONS = [
  // ── English phoneme confusions ──────────────────────────────────────────
  [/\bpassed due\b/gi,       'past due'],       // Whisper hears "past" as "passed"
  [/\bover draft\b/gi,       'overdraft'],

  // ── Spanish financial terms (Whisper merges/mangles these) ───────────────
  [/\bsuprestemo\b/gi,       'su préstamo'],    // "su préstamo" → "Suprestemo"
  [/\bsuprestamo\b/gi,       'su préstamo'],
  [/\bprestamo\b/gi,         'préstamo'],       // accent frequently dropped
];

// DB-backed correction entries loaded into memory at startup and refreshed on change.
// Each entry: { raw_text, corrected } — applied as whole-word case-insensitive replacements.
let sttCorrectionCache = [];

async function loadSttCorrectionCache() {
  if (!dbReady) return;
  try {
    const rows = await runQuery('SELECT raw_text, corrected FROM stt_dictionary ORDER BY created_at');
    sttCorrectionCache = rows || [];
  } catch (e) {
    sttCorrectionCache = [];
  }
}

function normalizeSttText(text) {
  if (!text) return text;
  let out = text;
  // Apply hardcoded corrections first
  for (const [pattern, replacement] of STT_CORRECTIONS) {
    out = out.replace(pattern, replacement);
  }
  // Apply user-defined DB corrections (whole-word, case-insensitive)
  for (const { raw_text, corrected } of sttCorrectionCache) {
    if (!raw_text) continue;
    const escaped = raw_text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    out = out.replace(new RegExp(`\\b${escaped}\\b`, 'gi'), corrected);
  }
  return out;
}

function extractMentionedPhone(transcript) {
  if (!transcript) return null;
  // Match common spoken phone number patterns: "8005551234", "800-555-1234", "(800) 555-1234"
  const digitMatch = transcript.match(/\b(?:\+?1[-.\s]?)?(?:\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4})\b/);
  if (digitMatch) return digitMatch[0].replace(/\D/g, '').replace(/^1(\d{10})$/, '$1');

  // Match spelled-out numbers like "eight hundred five five five one two three four"
  // (basic — converts common word groups to digits)
  const numberWords = { zero:0, one:1, two:2, three:3, four:4, five:5, six:6, seven:7, eight:8, nine:9,
    ten:10, eleven:11, twelve:12, hundred:100, thousand:1000 };
  const wordNumRegex = /\b(zero|one|two|three|four|five|six|seven|eight|nine)\b/gi;
  const spokenDigits = [];
  let m;
  const lc = transcript.toLowerCase();
  // Simple: collect consecutive digit words near "call", "dial", "number", "reach"
  const triggerIdx = Math.max(lc.indexOf('call us'), lc.indexOf('dial'), lc.indexOf('our number'), lc.indexOf('reach us'));
  const searchStr = triggerIdx >= 0 ? lc.slice(triggerIdx, triggerIdx + 200) : lc.slice(0, 200);
  wordNumRegex.lastIndex = 0;
  while ((m = wordNumRegex.exec(searchStr)) !== null) {
    spokenDigits.push(numberWords[m[1].toLowerCase()]);
  }
  if (spokenDigits.length >= 7) {
    const num = spokenDigits.join('');
    return num.length >= 10 ? num.slice(0, 10) : null;
  }
  return null;
}

function detectUrlMention(transcript) {
  if (!transcript) return false;
  const lc = transcript.toLowerCase();
  return /\.(com|org|net|io|gov|edu)\b/.test(lc) ||
    /\bwww\b/.test(lc) ||
    /\bhttps?:\/\//.test(lc) ||
    /\bvisit (us at|our (website|site|page))\b/.test(lc) ||
    /\bgo to (our|the) (website|site|page)\b/.test(lc) ||
    /\bcheck (us )?out (at|online)\b/.test(lc);
}

/**
 * Run generateTrendAnalysis in a worker thread so the main/UI thread stays responsive.
 */
function runAnalysisInWorker(inputData, outputPath, minConsec, minSpan, messageMap, callerMap, accountMap, userTz, userTzLabel, includeDetailTabs = true, transcriptMap = {}) {
  return new Promise((resolve, reject) => {
    const worker = new Worker(path.join(__dirname, 'analysisWorker.js'), {
      workerData: { inputData, outputPath, minConsec, minSpan, messageMap, callerMap, accountMap, userTz, userTzLabel, includeDetailTabs, transcriptMap },
      // Allow up to 6GB heap for large dataset analysis
      resourceLimits: { maxOldGenerationSizeMb: 6144 }
    });
    worker.on('message', (msg) => {
      if (msg.ok) resolve();
      else reject(new Error(msg.error || 'Worker analysis failed'));
    });
    worker.on('error', reject);
    worker.on('exit', (code) => {
      if (code !== 0) reject(new Error(`Analysis worker exited with code ${code}`));
    });
  });
}

// Cross-platform fetch implementation
// On Windows, we always use the https module with a permissive agent due to SSL certificate issues
// On other platforms, we use native fetch
async function crossPlatformFetch(url, options = {}) {
  const urlObj = new URL(url);
  const isWindows = process.platform === 'win32';
  const isTrusted = isTrustedDomain(urlObj.hostname);

  // On Windows for trusted domains, always use https module with permissive agent
  // This avoids SSL certificate store issues that are common on Windows
  if (!isWindows && typeof fetch === 'function') {
    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 30000);

      const response = await fetch(url, {
        ...options,
        signal: controller.signal
      });

      clearTimeout(timeoutId);
      return response;
    } catch (fetchError) {
      console.warn('[Fetch] Native fetch failed:', fetchError.message);
      // Fall through to https module
    }
  }

  // Use https module (required on Windows, fallback on other platforms)
  return new Promise((resolve, reject) => {
    const requestOptions = {
      hostname: urlObj.hostname,
      port: urlObj.port || 443,
      path: urlObj.pathname + urlObj.search,
      method: options.method || 'GET',
      headers: options.headers || {},
      timeout: 30000,
      // On Windows, use permissive agent for trusted VoApps domains
      agent: (isWindows && isTrusted && httpsAgent) ? httpsAgent : undefined
    };

    const req = https.request(requestOptions, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        resolve({
          ok: res.statusCode >= 200 && res.statusCode < 300,
          status: res.statusCode,
          statusText: res.statusMessage,
          headers: {
            get: (name) => res.headers[name.toLowerCase()],
            forEach: (cb) => Object.entries(res.headers).forEach(([k, v]) => cb(v, k))
          },
          text: async () => data,
          json: async () => JSON.parse(data)
        });
      });
    });

    req.on('error', (err) => {
      // If SSL certificate error on Windows, provide a helpful message
      if (err.message?.includes('certificate') || err.code === 'UNABLE_TO_GET_ISSUER_CERT_LOCALLY') {
        console.error('[HTTPS] SSL Certificate Error - this may be a Windows CA store issue');
        console.error('[HTTPS] Try running: set NODE_TLS_REJECT_UNAUTHORIZED=0 (not recommended for production)');
      }
      reject(err);
    });

    req.on('timeout', () => {
      req.destroy();
      reject(new Error('Request timeout (30s)'));
    });

    if (options.signal) {
      options.signal.addEventListener('abort', () => {
        req.destroy();
        reject(new Error('Request aborted'));
      });
    }

    if (options.body) {
      req.write(options.body);
    }

    req.end();
  });
}

// DuckDB integration - safely load with fallback for Windows compatibility
let duckdb = null;
let duckdbLoadError = null;
try {
  duckdb = require('duckdb');
  console.log('[DuckDB] Native module loaded successfully');
} catch (e) {
  duckdbLoadError = e.message;
  console.warn('[DuckDB] Failed to load native module:', e.message);
  console.warn('[DuckDB] Database features will be disabled. CSV-only mode available.');
}
const crypto = require('crypto');

const PORT = process.env.PORT ? Number(process.env.PORT) : 3000;
const HOST = "127.0.0.1";
const VOAPPS_API_BASE = process.env.VOAPPS_API_BASE || "https://directdropvoicemail.voapps.com/api/v1";

const MAX_ROWS_PER_FILE = 500000; // Split at 500K rows for optimal analysis

// Database configuration - cross-platform path
const DB_DIR = process.platform === 'win32'
  ? path.join(os.homedir(), 'AppData', 'Local', 'VoApps Tools')
  : path.join(os.homedir(), 'Library', 'Application Support', 'VoApps Tools');
const DB_PATH = path.join(DB_DIR, 'voapps_data.duckdb');
let db = null;
let dbReady = false;

let serverInstance = null;
let serverUrl = null;

const lastArtifacts = { csvPath: null, logPath: null, errorPath: null, analysisPath: null };
function getLastArtifacts() { return { ...lastArtifacts }; }

const jobs = new Map();

// =============================================================================
// SSE (Server-Sent Events) FUNCTIONS
// =============================================================================

/**
 * Send progress update to connected SSE client
 */
function sendProgress(jobId, update) {
  const job = jobs.get(jobId);
  if (!job) return;

  // Update job state
  Object.assign(job, update);

  // Calculate progress if current/total provided
  if (update.current !== undefined && job.total > 0) {
    job.progress = Math.round((update.current / job.total) * 100);
  }

  // Send SSE update if stream is open
  if (job.stream && !job.stream.destroyed) {
    try {
      const data = {
        type: 'progress',
        jobId,
        progress: job.progress,
        status: job.status,
        current: job.current,
        total: job.total,
        message: job.message
      };
      job.stream.write(`data: ${JSON.stringify(data)}\n\n`);
    } catch (e) {
      job.stream = null;  // Stream broken, clear it
    }
  }
}

/**
 * Send a persistent notification (action toast) to the connected SSE client.
 * The renderer will display this as a non-dismissing toast with an action button.
 */
function sendNotify(jobId, message, actionLabel, url) {
  const job = jobs.get(jobId);
  if (!job) return;
  if (job.stream && !job.stream.destroyed) {
    try {
      job.stream.write(`data: ${JSON.stringify({ type: 'notify', jobId, message, actionLabel, url })}\n\n`);
    } catch (e) {
      job.stream = null;
    }
  }
}

/**
 * Send log message to connected SSE client
 */
function sendLog(jobId, message, isError = false) {
  const job = jobs.get(jobId);
  if (!job) return;

  // Add to logs array (keep last 100)
  job.logs = job.logs || [];
  job.logs.push({ message, isError, time: Date.now() });
  if (job.logs.length > 100) {
    job.logs.shift();
  }

  // Send as SSE if stream open
  if (job.stream && !job.stream.destroyed) {
    try {
      const data = {
        type: 'log',
        jobId,
        message,
        isError
      };
      job.stream.write(`data: ${JSON.stringify(data)}\n\n`);
    } catch (e) {
      job.stream = null;
    }
  }
}

// =============================================================================
// DATABASE FUNCTIONS
// =============================================================================

/**
 * Check if database features are available
 */
function isDatabaseAvailable() {
  return duckdb !== null;
}

/**
 * Initialize DuckDB database
 */
async function initDatabase() {
  if (dbReady) return;

  // Check if DuckDB is available
  if (!duckdb) {
    console.warn('[DuckDB] Cannot initialize - native module not loaded:', duckdbLoadError);
    throw new Error(`Database features unavailable — DuckDB failed to load (${duckdbLoadError}). Please use CSV output mode instead.`);
  }

  try {
    // Ensure directory exists
    const dbDir = path.dirname(DB_PATH);
    await fsp.mkdir(dbDir, { recursive: true });

    // Create database connection
    db = new duckdb.Database(DB_PATH);
    
    // Create schema
    await runQuery(`
      CREATE TABLE IF NOT EXISTS campaign_results (
        row_id VARCHAR PRIMARY KEY,
        number VARCHAR NOT NULL,
        account_id VARCHAR NOT NULL,
        account_name VARCHAR,
        campaign_id VARCHAR,
        campaign_name VARCHAR,
        caller_number VARCHAR,
        caller_number_name VARCHAR,
        message_id VARCHAR,
        message_name VARCHAR,
        message_description VARCHAR,
        voapps_result VARCHAR,
        voapps_code VARCHAR,
        voapps_timestamp VARCHAR,
        campaign_url VARCHAR,
        target_date VARCHAR,
        voapps_voice_append VARCHAR,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
      )
    `);

    // Migration: add account_name column if it doesn't exist (for existing DBs)
    try {
      await runQuery(`ALTER TABLE campaign_results ADD COLUMN IF NOT EXISTS account_name VARCHAR`);
    } catch (e) {
      // Ignore if already exists or not supported
    }
    // Migration: add voapps_voice_append column if it doesn't exist (for existing DBs)
    try {
      await runQuery(`ALTER TABLE campaign_results ADD COLUMN IF NOT EXISTS voapps_voice_append VARCHAR`);
    } catch (e) {
      // Ignore if already exists or not supported
    }

    // Message transcription cache (AI Message Analysis)
    await runQuery(`
      CREATE TABLE IF NOT EXISTS message_transcriptions (
        message_id      VARCHAR NOT NULL,
        account_id      VARCHAR NOT NULL,
        audio_url       VARCHAR,
        transcript      VARCHAR,
        intent          VARCHAR,
        intent_summary  VARCHAR,
        mentioned_phone VARCHAR,
        mentions_url    BOOLEAN DEFAULT false,
        stt_model       VARCHAR,
        intent_model    VARCHAR,
        transcribed_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (message_id, account_id)
      )
    `);

    // STT normalization dictionary (user-managed corrections applied at analysis time)
    await runQuery(`
      CREATE TABLE IF NOT EXISTS stt_dictionary (
        id         VARCHAR PRIMARY KEY,
        raw_text   VARCHAR NOT NULL,
        corrected  VARCHAR NOT NULL,
        note       VARCHAR DEFAULT '',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
      )
    `);

    // Create indexes for faster queries
    await runQuery(`CREATE INDEX IF NOT EXISTS idx_number ON campaign_results(number)`);
    await runQuery(`CREATE INDEX IF NOT EXISTS idx_account ON campaign_results(account_id)`);
    await runQuery(`CREATE INDEX IF NOT EXISTS idx_campaign ON campaign_results(campaign_id)`);
    await runQuery(`CREATE INDEX IF NOT EXISTS idx_target_date ON campaign_results(target_date)`);
    await runQuery(`CREATE INDEX IF NOT EXISTS idx_timestamp ON campaign_results(voapps_timestamp)`);

    dbReady = true;
    console.log(`[DuckDB] Database initialized at ${DB_PATH}`);
  } catch (err) {
    console.error(`[DuckDB] Initialization error:`, err);
    throw err;
  }
}

/**
 * Run a query and return results
 */
function runQuery(sql, params = []) {
  return new Promise((resolve, reject) => {
    if (!db) {
      return reject(new Error('Database not initialized'));
    }
    
    db.all(sql, ...params, (err, rows) => {
      if (err) reject(err);
      else resolve(rows);
    });
  });
}

/**
 * Generate MD5 hash for row identification
 * Based on: number + account_id + campaign_id + voapps_timestamp
 */
function generateRowId(row) {
  const data = `${row.number}|${row.account_id}|${row.campaign_id}|${row.voapps_timestamp}`;
  return crypto.createHash('md5').update(data).digest('hex');
}

/**
 * Timezone configuration - maps user-friendly names to IANA timezones
 * VoApps Time is a constant UTC-7 (no DST), used for consistent day slicing
 */
const TIMEZONE_CONFIG = {
  'America/New_York': { label: 'ET', name: 'Eastern Time' },
  'America/Chicago': { label: 'CT', name: 'Central Time' },
  'America/Denver': { label: 'MT', name: 'Mountain Time' },
  'America/Los_Angeles': { label: 'PT', name: 'Pacific Time' },
  'UTC': { label: 'UTC', name: 'UTC' },
  'VoApps': { label: 'VoApps', name: 'VoApps Time (UTC-7)' }  // Constant UTC-7, no DST
};

/**
 * Get the current offset for a timezone at a specific date (DST-aware)
 * @param {string} timezone - IANA timezone name or 'VoApps'
 * @param {Date} date - The date to check offset for
 * @returns {string} - Offset string like "-05:00" or "-04:00"
 */
function getTimezoneOffsetForDate(timezone, date) {
  // VoApps Time is always UTC-7 (no DST)
  if (timezone === 'VoApps') {
    return '-07:00';
  }

  // UTC is always +00:00
  if (timezone === 'UTC') {
    return '+00:00';
  }

  try {
    // Use Intl.DateTimeFormat to get the actual offset at the given date
    const formatter = new Intl.DateTimeFormat('en-US', {
      timeZone: timezone,
      timeZoneName: 'longOffset'
    });

    const parts = formatter.formatToParts(date);
    const tzPart = parts.find(p => p.type === 'timeZoneName');

    if (tzPart && tzPart.value) {
      // Extract offset from "GMT-05:00" or "GMT-04:00" format
      const match = tzPart.value.match(/GMT([+-]\d{2}:\d{2})/);
      if (match) {
        return match[1];
      }
    }

    // Fallback: calculate offset manually
    const utcDate = new Date(date.toLocaleString('en-US', { timeZone: 'UTC' }));
    const tzDate = new Date(date.toLocaleString('en-US', { timeZone: timezone }));
    const diffMinutes = (tzDate - utcDate) / 60000;
    const hours = Math.floor(Math.abs(diffMinutes) / 60);
    const minutes = Math.abs(diffMinutes) % 60;
    const sign = diffMinutes >= 0 ? '+' : '-';
    return `${sign}${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
  } catch (e) {
    // Default to UTC-7 if timezone is invalid
    return '-07:00';
  }
}

/**
 * Normalize timestamp to user's selected timezone (DST-aware)
 * Supports IANA timezone names (e.g., "America/New_York") and "VoApps" for constant UTC-7
 *
 * @param {string} timestamp - The timestamp to normalize
 * @param {string} targetTimezone - Target timezone (IANA name or 'VoApps'), defaults to user setting
 *
 * Examples:
 *   "2025-03-15 09:47:05 -05:00" with ET → "2025-03-15 09:47:05 -04:00" (DST active)
 *   "2025-01-15 09:47:05 -05:00" with ET → "2025-01-15 09:47:05 -05:00" (no DST)
 *   "2025-11-16 14:47:05 UTC" with VoApps → "2025-11-16 07:47:05 -07:00" (constant)
 */
function normalizeToVoAppsTime(timestamp, targetTimezone = null) {
  if (!timestamp || typeof timestamp !== 'string') return timestamp;

  const ts = timestamp.trim();
  if (!ts) return timestamp;

  // Use user's selected timezone if not specified
  const timezone = targetTimezone || getTimezone();

  try {
    let date;

    // Parse timestamp based on format
    if (ts.includes(' UTC')) {
      // Format: "2025-11-16 14:47:05 UTC"
      const dateStr = ts.replace(' UTC', '');
      date = new Date(dateStr + 'Z');  // Treat as UTC
    } else {
      // Format: "2025-11-16 09:47:05 -05:00" or similar
      const match = ts.match(/^(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2}:\d{2})\s*([+-]\d{2}:\d{2})$/);
      if (match) {
        const [, datePart, timePart, offsetPart] = match;
        // Convert to ISO format for parsing
        date = new Date(`${datePart}T${timePart}${offsetPart}`);
      } else {
        // Try direct parsing as fallback
        date = new Date(ts);
      }
    }

    if (isNaN(date.getTime())) {
      return timestamp;  // Return original if parsing failed
    }

    // Get the DST-aware offset for this specific date
    const offset = getTimezoneOffsetForDate(timezone, date);
    const offsetHours = parseTimezoneOffset(offset);

    // Convert to target timezone
    const utcMs = date.getTime();
    const targetOffsetMs = offsetHours * 60 * 60 * 1000;
    const targetDate = new Date(utcMs + targetOffsetMs);

    // Format as YYYY-MM-DD HH:MM:SS [offset]
    const year = targetDate.getUTCFullYear();
    const month = String(targetDate.getUTCMonth() + 1).padStart(2, '0');
    const day = String(targetDate.getUTCDate()).padStart(2, '0');
    const hours = String(targetDate.getUTCHours()).padStart(2, '0');
    const minutes = String(targetDate.getUTCMinutes()).padStart(2, '0');
    const seconds = String(targetDate.getUTCSeconds()).padStart(2, '0');

    // Normalise +00:00 to the more readable "UTC" label
    const offsetLabel = offset === '+00:00' ? 'UTC' : offset;
    return `${year}-${month}-${day} ${hours}:${minutes}:${seconds} ${offsetLabel}`;
  } catch (e) {
    return timestamp;  // Return original on any error
  }
}

/**
 * Parse a timezone offset string like "-07:00" or "+05:30" into hours
 * @param {string} offset - Timezone offset string (e.g., "-07:00", "+00:00")
 * @returns {number} - Offset in hours (e.g., -7, 0, 5.5)
 */
function parseTimezoneOffset(offset) {
  if (!offset || typeof offset !== 'string') return -7; // Default to UTC-7

  const match = offset.match(/^([+-])(\d{2}):(\d{2})$/);
  if (!match) return -7;

  const sign = match[1] === '+' ? 1 : -1;
  const hours = parseInt(match[2], 10);
  const minutes = parseInt(match[3], 10);

  return sign * (hours + minutes / 60);
}

/**
 * Get timezone label for display (e.g., "ET" for America/New_York)
 * @param {string} timezone - IANA timezone name or 'VoApps'
 * @returns {string} - Human-readable timezone label
 */
function getTimezoneLabel(timezone) {
  if (TIMEZONE_CONFIG[timezone]) {
    return TIMEZONE_CONFIG[timezone].label;
  }
  // Fallback for legacy offset format
  const legacyLabels = {
    '-05:00': 'ET',
    '-06:00': 'CT',
    '-07:00': 'MT',
    '-08:00': 'PT',
    '+00:00': 'UTC'
  };
  return legacyLabels[timezone] || timezone;
}

/**
 * Check which phone numbers already have data in date range
 * Returns: { hasData: Map<number, boolean>, totalRows: number }
 */
async function checkExistingData(numbers, accountIds, startDate, endDate) {
  if (!dbReady) await initDatabase();
  
  const numberList = numbers.map(n => `'${n}'`).join(',');
  const accountList = accountIds.map(a => `'${a}'`).join(',');
  
  const sql = `
    SELECT 
      number,
      COUNT(*) as row_count
    FROM campaign_results
    WHERE number IN (${numberList})
      AND account_id IN (${accountList})
      AND target_date >= '${startDate}'
      AND target_date <= '${endDate}'
    GROUP BY number
  `;
  
  const results = await runQuery(sql);
  
  const hasData = new Map();
  let totalRows = 0;
  
  for (const row of results) {
    hasData.set(row.number, true);
    totalRows += Number(row.row_count); // Convert BigInt to Number
  }
  
  return { hasData, totalRows };
}

/**
 * Insert rows into database with duplicate handling.
 *
 * Performance: uses a single BEGIN/COMMIT transaction with batched multi-row
 * INSERT … ON CONFLICT (row_id) DO NOTHING statements.
 * Previously this was 2 queries per row (SELECT + INSERT/UPDATE); now it is
 * ≈ ceil(n/500) + 2 queries total — a 100–200× speed improvement for large
 * imports on Windows where each DuckDB round-trip has higher latency.
 */
async function insertRows(rows, logger = null) {
  if (!dbReady) await initDatabase();
  if (rows.length === 0) return { inserted: 0, updated: 0, skipped: 0 };

  // Deduplicate by row_id within this batch.  DuckDB's ON CONFLICT DO NOTHING
  // handles conflicts against existing table rows but throws a PRIMARY KEY error
  // when the same row_id appears twice within a single VALUES list.  Source CSVs
  // can contain duplicate rows (same data → same MD5), so we filter them out here.
  const seen = new Set();
  rows = rows.filter(row => {
    const id = generateRowId(row);
    if (seen.has(id)) return false;
    seen.add(id);
    return true;
  });

  // Rows per INSERT statement.  17 columns × 500 rows = 8,500 bound params,
  // comfortably within DuckDB's limit.
  const BATCH_SIZE = 500;

  try {
    // Count existing rows before so we can report accurate inserted/skipped stats
    const countBefore = Number(
      (await runQuery('SELECT COUNT(*) as cnt FROM campaign_results'))[0]?.cnt ?? 0
    );

    await runQuery('BEGIN');

    const totalBatches = Math.ceil(rows.length / BATCH_SIZE);
    for (let i = 0; i < rows.length; i += BATCH_SIZE) {
      const batch = rows.slice(i, i + BATCH_SIZE);

      // Build "(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?), ..." for this batch
      const placeholders = batch.map(() =>
        '(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'
      ).join(',');

      const params = [];
      for (const row of batch) {
        params.push(
          generateRowId(row),
          row.number              || '',
          row.account_id         || '',
          row.account_name       || '',
          row.campaign_id        || '',
          row.campaign_name      || '',
          row.caller_number      || '',
          row.caller_number_name || '',
          row.message_id         || '',
          row.message_name       || '',
          row.message_description|| '',
          row.voapps_result      || '',
          row.voapps_code        || '',
          row.voapps_timestamp   || '',
          row.campaign_url       || '',
          row.target_date        || '',
          row.voapps_voice_append|| ''
        );
      }

      await runQuery(`
        INSERT INTO campaign_results (
          row_id, number, account_id, account_name, campaign_id, campaign_name,
          caller_number, caller_number_name, message_id, message_name, message_description,
          voapps_result, voapps_code, voapps_timestamp, campaign_url, target_date,
          voapps_voice_append
        ) VALUES ${placeholders}
        ON CONFLICT (row_id) DO NOTHING
      `, params);

      if (logger && totalBatches > 1) {
        const batchNum = Math.floor(i / BATCH_SIZE) + 1;
        logger(`[DuckDB] Batch ${batchNum}/${totalBatches}: ${Math.min(i + BATCH_SIZE, rows.length).toLocaleString()}/${rows.length.toLocaleString()} rows processed`);
      }
    }

    await runQuery('COMMIT');

    const countAfter = Number(
      (await runQuery('SELECT COUNT(*) as cnt FROM campaign_results'))[0]?.cnt ?? 0
    );

    const inserted = countAfter - countBefore;
    const skipped  = rows.length - inserted;
    return { inserted, updated: 0, skipped };

  } catch (err) {
    // Roll back the transaction, then fall back to per-row inserts so the
    // caller still gets SOME data written even if the bulk path failed.
    try { await runQuery('ROLLBACK'); } catch (_) {}
    if (logger) logger(`[DuckDB] Bulk insert failed (${err.message}) — falling back to row-by-row mode`);

    let inserted = 0, skipped = 0;
    for (const row of rows) {
      try {
        await runQuery(`
          INSERT INTO campaign_results (
            row_id, number, account_id, account_name, campaign_id, campaign_name,
            caller_number, caller_number_name, message_id, message_name, message_description,
            voapps_result, voapps_code, voapps_timestamp, campaign_url, target_date,
            voapps_voice_append
          ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
          ON CONFLICT (row_id) DO NOTHING
        `, [
          generateRowId(row),
          row.number              || '',
          row.account_id         || '',
          row.account_name       || '',
          row.campaign_id        || '',
          row.campaign_name      || '',
          row.caller_number      || '',
          row.caller_number_name || '',
          row.message_id         || '',
          row.message_name       || '',
          row.message_description|| '',
          row.voapps_result      || '',
          row.voapps_code        || '',
          row.voapps_timestamp   || '',
          row.campaign_url       || '',
          row.target_date        || '',
          row.voapps_voice_append|| ''
        ]);
        inserted++;
      } catch (rowErr) {
        if (logger) logger(`[DuckDB] Row error: ${rowErr.message}`);
        skipped++;
      }
    }
    return { inserted, updated: 0, skipped };
  }
}

/**
 * Get database statistics
 */
async function getDatabaseStats() {
  // Check if DuckDB is available at all
  if (!isDatabaseAvailable()) {
    return {
      ready: false,
      available: false,
      error: duckdbLoadError || 'DuckDB native module not loaded',
      totalRows: 0,
      uniqueNumbers: 0,
      uniqueCampaigns: 0,
      dateRange: null,
      dbSize: 0,
      dbPath: null
    };
  }

  if (!dbReady) {
    try {
      await initDatabase();
    } catch (err) {
      return {
        ready: false,
        available: true,
        error: err.message,
        totalRows: 0,
        uniqueNumbers: 0,
        uniqueCampaigns: 0,
        dateRange: null,
        dbSize: 0
      };
    }
  }
  
  try {
    const [totalResult] = await runQuery(`SELECT COUNT(*) as count FROM campaign_results`);
    const [numbersResult] = await runQuery(`SELECT COUNT(DISTINCT number) as count FROM campaign_results`);
    const [campaignsResult] = await runQuery(`SELECT COUNT(DISTINCT campaign_id) as count FROM campaign_results`);
    const dateRangeResult = await runQuery(`
      SELECT 
        MIN(target_date) as min_date,
        MAX(target_date) as max_date
      FROM campaign_results
    `);
    
    // Get database file size
    let dbSize = 0;
    try {
      const stats = await fsp.stat(DB_PATH);
      dbSize = stats.size;
    } catch (err) {
      // File might not exist yet
    }
    
    // Convert BigInt to Number for JSON serialization
    return {
      ready: true,
      totalRows: Number(totalResult.count),
      uniqueNumbers: Number(numbersResult.count),
      uniqueCampaigns: Number(campaignsResult.count),
      dateRange: dateRangeResult[0] || { min_date: null, max_date: null },
      dbSize,
      dbPath: DB_PATH
    };
  } catch (err) {
    return {
      ready: false,
      error: err.message,
      totalRows: 0,
      uniqueNumbers: 0,
      uniqueCampaigns: 0,
      dateRange: null,
      dbSize: 0
    };
  }
}

/**
 * Clear all data from database with backup and VACUUM
 * @param {boolean} createBackup - Whether to create a backup before clearing (default: true)
 */
async function clearDatabase(createBackup = true) {
  if (!dbReady) await initDatabase();

  try {
    let backupPath = null;

    // Create backup before clearing if requested
    if (createBackup) {
      try {
        backupPath = DB_PATH.replace('.duckdb', `_backup_${Date.now()}.duckdb`);
        await fsp.copyFile(DB_PATH, backupPath);
        console.log(`[Database] Backup created: ${backupPath}`);
      } catch (backupErr) {
        console.warn(`[Database] Failed to create backup: ${backupErr.message}`);
        // Continue with clear even if backup fails
      }
    }

    // Get row count before clearing
    const beforeResult = await runQuery(`SELECT COUNT(*) as count FROM campaign_results`);
    const rowCount = beforeResult[0]?.count || 0;

    // Clear table
    await runQuery(`DELETE FROM campaign_results`);
    console.log(`[Database] Deleted ${rowCount.toLocaleString()} rows`);

    // VACUUM to reclaim disk space
    // Note: DuckDB VACUUM requires closing and reopening the connection
    console.log(`[Database] Running VACUUM to reclaim disk space...`);
    await runQuery(`VACUUM`);
    console.log(`[Database] VACUUM complete`);

    return {
      success: true,
      backupPath,
      rowCount,
      message: `Database cleared (${rowCount.toLocaleString()} rows deleted, space reclaimed)`
    };
  } catch (err) {
    console.error(`[Database] Clear failed: ${err.message}`);
    return {
      success: false,
      error: err.message
    };
  }
}

/**
 * Compact database (VACUUM)
 */
async function compactDatabase() {
  if (!dbReady) await initDatabase();
  
  try {
    await runQuery(`VACUUM`);
    return {
      success: true,
      message: 'Database compacted successfully'
    };
  } catch (err) {
    return {
      success: false,
      error: err.message
    };
  }
}

/**
 * Export database to CSV using streaming to avoid OOM on large datasets
 */
async function exportDatabase(outputPath) {
  if (!isDatabaseAvailable()) {
    return { success: false, error: 'Database not available' };
  }
  if (!dbReady) await initDatabase();

  try {
    // Get count first to check if there's data
    const countResult = await runQuery('SELECT COUNT(*) as cnt FROM campaign_results');
    const totalRows = Number(countResult[0]?.cnt || 0);

    if (totalRows === 0) {
      return { success: false, error: 'No data to export' };
    }

    // Get column names from schema
    const schemaResult = await runQuery(`
      SELECT column_name FROM information_schema.columns
      WHERE table_name = 'campaign_results'
      ORDER BY ordinal_position
    `);
    const headers = schemaResult.map(r => r.column_name);

    // Create write stream for memory-efficient export
    const writeStream = fs.createWriteStream(outputPath, { encoding: 'utf-8' });

    // Write headers
    writeStream.write(headers.join(',') + '\n');

    // Stream rows in batches to avoid OOM
    const BATCH_SIZE = 50000;
    let offset = 0;
    let rowCount = 0;

    while (offset < totalRows) {
      const batchRows = await runQuery(`
        SELECT * FROM campaign_results
        ORDER BY target_date DESC, voapps_timestamp DESC
        LIMIT ${BATCH_SIZE} OFFSET ${offset}
      `);

      if (batchRows.length === 0) break;

      // Write batch to stream
      for (const row of batchRows) {
        const line = headers.map(h => {
          let val = row[h];
          if (val === null || val === undefined) return '';
          // Normalize timestamps to VoApps Time (UTC-7)
          if (h === 'voapps_timestamp' && val) {
            val = normalizeToVoAppsTime(val);
          }
          const str = String(val);
          return str.includes(',') || str.includes('"') ? `"${str.replace(/"/g, '""')}"` : str;
        }).join(',');
        writeStream.write(line + '\n');
      }

      rowCount += batchRows.length;
      offset += BATCH_SIZE;

      // Log progress for large exports
      if (totalRows > BATCH_SIZE) {
        console.log(`[Export] ${rowCount.toLocaleString()} / ${totalRows.toLocaleString()} rows written`);
      }
    }

    // Close the stream
    await new Promise((resolve, reject) => {
      writeStream.end((err) => err ? reject(err) : resolve());
    });

    return {
      success: true,
      path: outputPath,
      rowCount
    };
  } catch (err) {
    return {
      success: false,
      error: err.message
    };
  }
}

// =============================================================================
// HELPER FUNCTIONS (from v2.4.1)
// =============================================================================

function sendJson(res, status, obj) {
  // Handle BigInt serialization by converting to string
  const body = JSON.stringify(obj, (key, value) =>
    typeof value === 'bigint' ? value.toString() : value
  );
  res.writeHead(status, { 
    "Content-Type": "application/json; charset=utf-8", 
    "Cache-Control": "no-store" 
  });
  res.end(body);
}

async function readJson(req) {
  return new Promise((resolve) => {
    let data = "";
    req.on("data", (chunk) => (data += chunk));
    req.on("end", () => {
      try { resolve(data ? JSON.parse(data) : {}); }
      catch { resolve({}); }
    });
  });
}

function safeJoinPublic(filePath) {
  const pub = path.join(__dirname, "public");
  const full = path.normalize(path.join(pub, filePath));
  return full.startsWith(pub) ? full : null;
}

function guessContentType(file) {
  const ext = path.extname(file).toLowerCase();
  const map = {
    ".html": "text/html; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".svg": "image/svg+xml",
  };
  return map[ext] || "application/octet-stream";
}

function dateToYMD(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function getLogTimestamp() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  const h = String(now.getHours()).padStart(2, '0');
  const min = String(now.getMinutes()).padStart(2, '0');
  const s = String(now.getSeconds()).padStart(2, '0');
  return `[${y}-${m}-${d} ${h}:${min}:${s}]`;
}

function getFilenameSuffix(directory, prefix) {
  const now = new Date();
  const dateStr = dateToYMD(now);
  
  let maxCounter = 0;
  try {
    const files = fs.readdirSync(directory);
    const pattern = new RegExp(`^${prefix}_${dateStr}_(\\d{3})`);
    
    for (const file of files) {
      const match = file.match(pattern);
      if (match) {
        const counter = parseInt(match[1]);
        if (counter > maxCounter) maxCounter = counter;
      }
    }
  } catch (err) {
    // Directory doesn't exist yet or can't read
  }
  
  const nextCounter = maxCounter + 1;
  const counterStr = String(nextCounter).padStart(3, '0');
  return `${dateStr}_${counterStr}`;
}

// Settings management
const SETTINGS_PATH = process.platform === 'win32'
  ? path.join(os.homedir(), 'AppData', 'Local', 'VoApps Tools', 'settings.json')
  : path.join(os.homedir(), 'Library', 'Application Support', 'VoApps Tools', 'settings.json');

function getDefaultOutputFolder() {
  // Use Documents folder as default instead of Downloads
  return path.join(os.homedir(), 'Documents', 'VoApps Tools');
}

function loadSettings() {
  try {
    if (fs.existsSync(SETTINGS_PATH)) {
      const data = fs.readFileSync(SETTINGS_PATH, 'utf-8');
      return JSON.parse(data);
    }
  } catch (e) {
    console.warn('[Settings] Failed to load settings:', e.message);
  }
  return {
    outputFolder: getDefaultOutputFolder(),
    timezone: '-07:00'  // Default to VoApps Time (UTC-7)
  };
}

function saveSettings(settings) {
  try {
    const dir = path.dirname(SETTINGS_PATH);
    fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(SETTINGS_PATH, JSON.stringify(settings, null, 2), 'utf-8');
    console.log('[Settings] Saved to', SETTINGS_PATH);
    return true;
  } catch (e) {
    console.error('[Settings] Failed to save settings:', e.message);
    return false;
  }
}

function getOutputFolder() {
  const settings = loadSettings();
  return settings.outputFolder || getDefaultOutputFolder();
}

function setOutputFolder(folderPath) {
  const settings = loadSettings();
  settings.outputFolder = folderPath;
  return saveSettings(settings);
}

function getTimezone() {
  const settings = loadSettings();
  // Default to VoApps Time (constant UTC-7) for consistency
  // Migrate legacy offset format to IANA timezone names
  const tz = settings.timezone || 'VoApps';
  const legacyMapping = {
    '-05:00': 'America/New_York',
    '-06:00': 'America/Chicago',
    '-07:00': 'VoApps',  // Keep as VoApps for constant UTC-7
    '-08:00': 'America/Los_Angeles',
    '+00:00': 'UTC'
  };
  return legacyMapping[tz] || tz;
}

function setTimezone(timezone) {
  const settings = loadSettings();
  settings.timezone = timezone;
  return saveSettings(settings);
}

function getOpenaiKey() {
  const settings = loadSettings();
  return settings.openaiApiKey || '';
}

function setOpenaiKey(key) {
  const settings = loadSettings();
  settings.openaiApiKey = key;
  return saveSettings(settings);
}

function getAiSettings() {
  const settings = loadSettings();
  return {
    enabled: settings.enableAiAnalysis || false,
    transcriptionMode: settings.aiTranscriptionMode || 'local',
    intentMode: settings.aiIntentMode || 'local',
    openaiApiKey: settings.openaiApiKey || ''
  };
}

function setAiSettings(updates) {
  const settings = loadSettings();
  if (updates.enabled !== undefined) settings.enableAiAnalysis = updates.enabled;
  if (updates.transcriptionMode !== undefined) settings.aiTranscriptionMode = updates.transcriptionMode;
  if (updates.intentMode !== undefined) settings.aiIntentMode = updates.intentMode;
  if (updates.openaiApiKey !== undefined) settings.openaiApiKey = updates.openaiApiKey;
  return saveSettings(settings);
}

function createOutputFolders() {
  const base = getOutputFolder();
  const logs = path.join(base, "Logs");
  const output = path.join(base, "Output");
  const phoneHistory = path.join(output, "Phone Number History");
  const combineCampaigns = path.join(output, "Combine Campaigns");
  const bulkExport = path.join(output, "Bulk Campaign Export");
  const executiveSummary = path.join(output, "Executive Summary");

  for (const dir of [base, logs, output, phoneHistory, combineCampaigns, bulkExport, executiveSummary]) {
    fs.mkdirSync(dir, { recursive: true });
  }

  return { base, logs, phoneHistory, combineCampaigns, bulkExport, executiveSummary };
}

function createLogger(logPath, errorPath, verbosity = "normal", jobId = null) {
  const logStream = fs.createWriteStream(logPath, { flags: "a" });
  // Lazily create the error stream — only open the file on the first actual error write.
  // This prevents empty error files from being created when no errors occur.
  let errorStream = null;
  function getErrorStream() {
    if (!errorStream && errorPath) {
      errorStream = fs.createWriteStream(errorPath, { flags: "a" });
    }
    return errorStream;
  }

  function log(message, isError = false) {
    if (verbosity === "none") return;

    const timestamp = getLogTimestamp();
    const line = `${timestamp} ${message}\n`;

    // Safety check: only write if streams are still writable
    if (isError) {
      const es = getErrorStream();
      if (es && !es.writableEnded) {
        try {
          es.write(line);
        } catch (e) {
          // Stream closed, ignore
        }
      }
    }
    if (logStream && !logStream.writableEnded) {
      try {
        logStream.write(line);
      } catch (e) {
        // Stream closed, ignore
      }
    }

    // Send to SSE stream if jobId provided
    if (jobId) {
      sendLog(jobId, message, isError);
    }
  }

  function close() {
    logStream.end();
    if (errorStream) errorStream.end();
  }

  return { log, close };
}

// =============================================================================
// CSV PARSING AND WRITING (from v2.4.1)
// =============================================================================

function parseCsv(text) {
  const lines = text.split(/\r?\n/);
  if (lines.length === 0) return { headers: [], rows: [] };

  // Parse a single line handling quoted fields with commas
  function parseLine(line) {
    const values = [];
    let current = '';
    let inQuotes = false;

    for (let i = 0; i < line.length; i++) {
      const char = line[i];
      const nextChar = line[i + 1];

      if (char === '"') {
        if (inQuotes && nextChar === '"') {
          // Escaped quote
          current += '"';
          i++; // Skip next quote
        } else {
          // Toggle quote mode
          inQuotes = !inQuotes;
        }
      } else if (char === ',' && !inQuotes) {
        // End of field
        values.push(current.trim());
        current = '';
      } else {
        current += char;
      }
    }

    // Push last field
    values.push(current.trim());

    return values;
  }

  const headers = parseLine(lines[0]);
  const rows = [];

  for (let i = 1; i < lines.length; i++) {
    if (!lines[i].trim()) continue;
    const values = parseLine(lines[i]);
    const row = {};
    headers.forEach((h, idx) => {
      row[h] = values[idx] || "";
    });
    rows.push(row);
  }

  return { headers, rows };
}

async function writeCsv(filePath, rows, headers, logger = null, maxRowsPerFile = null) {
  if (!Array.isArray(rows) || rows.length === 0) {
    throw new Error("No rows to write");
  }

  const shouldSplit = maxRowsPerFile && rows.length > maxRowsPerFile;
  
  if (!shouldSplit) {
    // Single file
    const csvContent = [
      headers.join(','),
      ...rows.map(row => headers.map(h => {
        const val = row[h];
        // Handle null/undefined
        if (val === null || val === undefined) return '';
        // Convert to string
        const str = String(val);
        // Escape if needed
        return str.includes(',') || str.includes('"') ? `"${str.replace(/"/g, '""')}"` : str;
      }).join(','))
    ].join('\n');

    await fsp.writeFile(filePath, csvContent, 'utf-8');
    
    if (logger) {
      logger(`✅ CSV written: ${filePath}`);
      logger(`   ${rows.length.toLocaleString()} rows`);
    }

    return {
      success: true,
      files: [filePath],
      totalRows: rows.length,
      fileCount: 1,
      wasSplit: false
    };
  }

  // Multi-file split
  const dir = path.dirname(filePath);
  const ext = path.extname(filePath);
  const base = path.basename(filePath, ext);

  const files = [];
  let fileIndex = 1;
  let startIdx = 0;

  while (startIdx < rows.length) {
    const endIdx = Math.min(startIdx + maxRowsPerFile, rows.length);
    const chunk = rows.slice(startIdx, endIdx);

    const partFilename = `${base}_part${fileIndex}${ext}`;
    const partPath = path.join(dir, partFilename);

    const csvContent = [
      headers.join(','),
      ...chunk.map(row => headers.map(h => {
        const val = row[h];
        // Handle null/undefined
        if (val === null || val === undefined) return '';
        // Convert to string
        const str = String(val);
        // Escape if needed
        return str.includes(',') || str.includes('"') ? `"${str.replace(/"/g, '""')}"` : str;
      }).join(','))
    ].join('\n');

    await fsp.writeFile(partPath, csvContent, 'utf-8');
    files.push(partPath);

    if (logger) {
      logger(`✅ Part ${fileIndex} written: ${partFilename}`);
      logger(`   ${chunk.length.toLocaleString()} rows`);
    }

    startIdx = endIdx;
    fileIndex++;
  }

  if (logger) {
    logger(`📊 Split into ${files.length} files (${rows.length.toLocaleString()} total rows)`);
  }

  return {
    success: true,
    files,
    totalRows: rows.length,
    fileCount: files.length,
    wasSplit: true
  };
}

// =============================================================================
// API CALLING (from v2.4.1)
// =============================================================================

function maskApiKey(key) {
  if (!key || key.length < 8) return "***";
  return key.slice(0, 4) + "..." + key.slice(-4);
}

async function callVoAppsApi(endpoint, apiKey, logger = null, verbosity = "normal") {
  const url = `${VOAPPS_API_BASE}${endpoint}`;
  const maskedKey = maskApiKey(apiKey);

  if (verbosity === "verbose" && logger) {
    logger(`[API] ${url}`);
    logger(`      curl -H "Authorization: Bearer ${maskedKey}" -H "Content-Type: application/json" -H "Accept: application/json" "${url}"`);
  }

  console.log(`[API Request] ${url}`);

  try {
    // Use cross-platform fetch for Windows compatibility
    const response = await crossPlatformFetch(url, {
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json",
        Accept: "application/json"
      }
    });

    if (!response.ok) {
      let errorDetails = '';
      let responseHeaders = {};

      try {
        // Capture headers
        response.headers.forEach((value, key) => {
          responseHeaders[key] = value;
        });

        const contentType = response.headers.get('content-type') || 'unknown';
        const text = await response.text();
        errorDetails = text || response.statusText || 'No error details';

        // Enhanced logging
        console.error(`[VoApps API Error] ${response.status} ${url}`);
        console.error(`[VoApps API Error] Content-Type: ${contentType}`);
        console.error(`[VoApps API Error] Response Headers:`, responseHeaders);
        console.error(`[VoApps API Error] Response Body (length ${text.length}): ${text.substring(0, 500)}`);

      } catch (e) {
        errorDetails = response.statusText || 'Could not read error response';
        console.error(`[VoApps API Error] Failed to read response:`, e.message);
      }

      throw new Error(`HTTP ${response.status}: ${errorDetails}`);
    }

    return await response.json();
  } catch (err) {
    console.error(`[API Error] ${url}:`, err.message);
    throw err;
  }
}

async function retryableApiCall(endpoint, apiKey, logger = null, verbosity = "normal") {
  const delays = [3000, 10000, 60000];

  for (let attempt = 0; attempt < delays.length; attempt++) {
    try {
      return await callVoAppsApi(endpoint, apiKey, logger, verbosity);
    } catch (err) {
      if (logger) {
        logger(`❌ API call failed (attempt ${attempt + 1}/${delays.length}): ${err.message}`, true);
      }

      if (attempt < delays.length - 1) {
        const delay = delays[attempt];
        if (logger) logger(`⏳ Retrying in ${delay / 1000}s...`);
        await new Promise(resolve => setTimeout(resolve, delay));
      } else {
        throw err;
      }
    }
  }
}

// =============================================================================
// CAMPAIGN AND ACCOUNT FETCHING (from v2.4.1)
// =============================================================================

async function fetchAllAccounts(apiKey, logger = null, verbosity = "normal", filter = null) {
  try {
    let endpoint = "/accounts";
    if (filter) {
      endpoint += `?filter=${filter}`;
    }
    
    const data = await retryableApiCall(endpoint, apiKey, logger, verbosity);
    
    if (!Array.isArray(data.accounts)) {
      throw new Error("Invalid accounts response");
    }

    if (logger) {
      logger(`✅ Fetched ${data.accounts.length} accounts${filter ? ` (${filter})` : ''}`);
    }

    return data.accounts;
  } catch (err) {
    if (logger) {
      logger(`❌ Failed to fetch accounts: ${err.message}`, true);
    }
    throw err;
  }
}

/**
 * Fetch account names from API
 * Returns a map of accountId -> accountName
 */
async function fetchAccountNames(apiKey, accountIds, logger = null, verbosity = "normal") {
  const accountNames = {};

  try {
    const data = await retryableApiCall('/accounts', apiKey, logger, verbosity);

    if (Array.isArray(data.accounts)) {
      for (const account of data.accounts) {
        const accountId = String(account.id);
        if (accountIds.includes(accountId)) {
          accountNames[accountId] = account.name || '';
        }
      }
    }

    if (logger && verbosity !== 'minimal') {
      logger(`✅ Fetched account names for ${Object.keys(accountNames).length} account(s)`);
    }
  } catch (err) {
    if (logger) {
      logger(`⚠️  Could not fetch account names: ${err.message}`);
    }
  }

  return accountNames;
}

/**
 * Fetch timezone settings for accounts
 * VoApps API returns timezone as IANA format (e.g., "America/Denver")
 */
async function fetchAccountTimezones(apiKey, accountIds, logger = null, verbosity = "normal") {
  const accountTimezones = {};

  try {
    // Fetch all accounts to get timezone info
    const data = await retryableApiCall('/accounts', apiKey, logger, verbosity);

    if (Array.isArray(data.accounts)) {
      for (const account of data.accounts) {
        const accountId = String(account.id);
        if (accountIds.includes(accountId) && account.timezone) {
          accountTimezones[accountId] = account.timezone;
        }
      }
    }

    if (logger && verbosity !== 'minimal') {
      const tzCount = Object.keys(accountTimezones).length;
      logger(`✅ Fetched timezone settings for ${tzCount} account(s)`);
    }
  } catch (err) {
    if (logger) {
      logger(`⚠️  Could not fetch account timezones: ${err.message}`);
    }
  }

  return accountTimezones;
}

async function fetchCallerNumbers(apiKey, accountIds, logger = null, verbosity = "normal") {
  const callerNumberNames = {};

  for (const accountId of accountIds) {
    try {
      const data = await retryableApiCall(`/accounts/${accountId}/caller_numbers?filter=all`, apiKey, logger, verbosity);
      
      if (Array.isArray(data.caller_numbers)) {
        for (const caller of data.caller_numbers) {
          const key = `${accountId}:${caller.number}`;
          callerNumberNames[key] = caller.name || '';
        }
      }

      if (logger) {
        logger(`✅ Fetched ${data.caller_numbers?.length || 0} caller numbers for account ${accountId}`);
      }
    } catch (err) {
      if (logger) {
        logger(`⚠️  Could not fetch caller numbers for account ${accountId}: ${err.message}`);
      }
    }
  }

  return callerNumberNames;
}

async function fetchAllCampaigns(apiKey, accountIds, startDate, endDate, logger = null, verbosity = "normal", jobId = null) {
  const bufferDays = 7;
  const start = new Date(startDate);
  start.setDate(start.getDate() - bufferDays);
  const bufferedStart = dateToYMD(start);

  const end = new Date(endDate);
  end.setDate(end.getDate() + 2);
  const bufferedEnd = dateToYMD(end);

  if (logger) {
    logger(`📅 Requested range: ${startDate} to ${endDate}`);
    logger(`📅 Buffer applied: ${bufferedStart} to ${bufferedEnd} (widened for lead time)`);
    logger(`📅 Will filter results by target_date: ${startDate} to ${endDate}`);
  }

  const allCampaigns = [];

  for (const accountId of accountIds) {
    if (jobId && jobs.get(jobId)?.cancelled) {
      throw new Error("Cancelled");
    }

    if (logger) {
      logger(`\n🔍 Fetching campaigns for account ${accountId}...`);
    }

    let page = 1;  // VoApps API uses 1-based pagination
    let hasMore = true;
    let accountCampaigns = [];

    while (hasMore) {
      if (jobId && jobs.get(jobId)?.cancelled) {
        throw new Error("Cancelled");
      }

      try {
        const endpoint = `/accounts/${accountId}/campaigns?created_date_start=${bufferedStart}&created_date_end=${bufferedEnd}&page=${page}&filter=all`;
        const fullUrl = `${VOAPPS_API_BASE}${endpoint}`;
        
        // Detailed logging (like v2.4.1)
        if (logger && verbosity !== "minimal") {
          const maskedKey = apiKey ? `${apiKey.substring(0, 4)}...${apiKey.substring(apiKey.length - 4)}` : '[none]';
          logger(`   → Page ${page}: GET ${fullUrl}`);
          logger(`   → Auth: Bearer ${maskedKey}`);
          
          // Generate curl command for debugging
          const curlCmd = `curl -X GET '${fullUrl}' -H 'Authorization: Bearer ${maskedKey}' -H 'Content-Type: application/json' -H 'Accept: application/json'`;
          logger(`   → curl: ${curlCmd}`);
        }
        
        const data = await retryableApiCall(endpoint, apiKey, logger, verbosity);

        if (!Array.isArray(data.campaigns)) {
          throw new Error("Invalid campaigns response");
        }

        accountCampaigns.push(...data.campaigns);

        if (logger && verbosity !== "minimal") {
          logger(`   Page ${page}: ${data.campaigns.length} campaigns`);
        }

        // If we got fewer than 25 campaigns, this is the last page
        if (data.campaigns.length < 25) {
          hasMore = false;
          if (logger && verbosity !== "minimal") {
            logger(`   ✓ Pagination complete: ${accountCampaigns.length} total campaigns fetched`);
          }
        }
        
        page++;
      } catch (err) {
        if (logger) {
          logger(`❌ Error fetching campaigns for account ${accountId}: ${err.message}`, true);
        }
        throw err;
      }
    }

    const filtered = accountCampaigns.filter(c => {
      const targetDate = c?.target_date;
      if (!targetDate) return false;

      try {
        // For plain YYYY-MM-DD strings, compare directly (avoids UTC midnight timezone shift).
        // For full datetime strings (e.g. '2026-03-02T08:30:00-07:00'), check whether
        // the UTC time falls on the selected date in ANY US timezone (UTC-4 EDT through UTC-10 HST).
        // This ensures campaigns targeted on the last date of the range are not excluded
        // simply because their local delivery time maps to a later UTC date.
        if (/^\d{4}-\d{2}-\d{2}$/.test(targetDate)) {
          return targetDate >= startDate && targetDate <= endDate;
        }

        const targetUTC = new Date(targetDate);
        if (isNaN(targetUTC.getTime())) return false;

        // US timezone offsets (hours behind UTC): EDT=4, EST=5, CDT=5, CST=6,
        // MDT=6, MST=7, PDT=7, PST=8, AKDT=8, AKST=9, HST=10
        const usOffsets = [4, 5, 6, 7, 8, 9, 10];
        for (const offsetHours of usOffsets) {
          const localDateStr = new Date(targetUTC.getTime() - offsetHours * 3600 * 1000)
            .toISOString().slice(0, 10);
          if (localDateStr >= startDate && localDateStr <= endDate) return true;
        }
        return false;
      } catch (e) {
        return false;
      }
    });

    // Add account_id to each campaign
    const filteredWithAccountId = filtered.map(c => ({
      ...c,
      account_id: accountId
    }));

    if (logger) {
      logger(`✅ Account ${accountId}: ${accountCampaigns.length} total, ${filtered.length} in range (${startDate} to ${endDate})`);
    }

    allCampaigns.push(...filteredWithAccountId);
  }

  return allCampaigns;
}

async function fetchCampaignDetail(apiKey, accountId, campaignId, signal) {
  const url = `${VOAPPS_API_BASE}/accounts/${accountId}/campaigns/${campaignId}`;
  return await retryableApiCall(`/accounts/${accountId}/campaigns/${campaignId}`, apiKey, null, "minimal");
}

// =============================================================================
// MAIN SEARCH FUNCTIONS (Enhanced for v3.0.0)
// =============================================================================

async function runNumberSearch(config) {
  const {
    api_key,
    numbers,
    account_ids,
    start_date,
    end_date,
    include_caller = true,
    include_message_meta = true,
    output_mode = "csv", // NEW: "csv", "database", or "both"
    job_id = null,
    client_prefix = "" // Optional prefix for output files
  } = config;

  // Build filename prefix
  const filePrefix = client_prefix ? `${client_prefix}_` : "";

  const folders = createOutputFolders();
  const suffix = getFilenameSuffix(folders.logs, 'voapps_log');
  const logPath = path.join(folders.logs, `voapps_log_${suffix}.txt`);
  const errorPath = path.join(folders.logs, `voapps_errors_${suffix}.txt`);

  const { log, close } = createLogger(logPath, errorPath, "normal", job_id);

  // Initialize job tracking with SSE support
  if (job_id) {
    jobs.set(job_id, {
      id: job_id,
      cancelled: false,
      stream: null,
      progress: 0,
      status: 'starting',
      current: 0,
      total: 0,
      message: 'Initializing search...',
      logs: [],
      startTime: Date.now()
    });
    sendProgress(job_id, { status: 'starting', message: 'Initializing phone number search...' });
  }

  lastArtifacts.logPath = logPath;
  lastArtifacts.errorPath = errorPath;

  try {
    log(`=== VoApps Tools v${VERSION} - Phone Number Search ===`);
    log(`Output Mode: ${output_mode}`);
    log(`Numbers: ${numbers.length}`);
    log(`Accounts: ${account_ids.join(", ")}`);
    log(`Date Range: ${start_date} to ${end_date}`);

    // Check database for existing data if database mode
    if (output_mode === "database" || output_mode === "both") {
      if (!dbReady) await initDatabase();
      
      const { hasData, totalRows } = await checkExistingData(numbers, account_ids, start_date, end_date);
      
      const numbersWithData = numbers.filter(n => hasData.get(n));
      const numbersNeedingFetch = numbers.filter(n => !hasData.get(n));
      
      if (numbersWithData.length > 0) {
        log(`\n📊 Database Check:`);
        log(`   ${numbersWithData.length} numbers already in database (${totalRows.toLocaleString()} rows)`);
        log(`   ${numbersNeedingFetch.length} numbers need fresh data`);
        
        // If all numbers already have data, skip API calls
        if (numbersNeedingFetch.length === 0) {
          log(`\n✅ All requested data already in database - skipping API calls`);
          
          if (output_mode === "both" || output_mode === "csv") {
            // Export database to CSV
            const csvPath = path.join(folders.phoneHistory, `${filePrefix}phone_search_${suffix}.csv`);
            const exportResult = await exportDatabaseSubset(numbers, account_ids, start_date, end_date, csvPath);
            
            if (exportResult.success) {
              lastArtifacts.csvPath = csvPath;
              log(`\n✅ Exported ${exportResult.rowCount.toLocaleString()} rows to CSV`);
            }
          }
          
          close();
          return {
            csvPath: lastArtifacts.csvPath,
            logPath,
            matches: totalRows,
            wasSplit: false,
            fileCount: 1,
            fromDatabase: true
          };
        }
      }
    }

    // Fetch campaigns
    const campaigns = await fetchAllCampaigns(api_key, account_ids, start_date, end_date, log, "normal", job_id);

    if (campaigns.length === 0) {
      log("\n⚠️  No campaigns found in date range");
      close();
      throw new Error("No campaigns found in specified date range");
    }

    log(`\n📊 Found ${campaigns.length} campaigns to search`);

    // Fetch caller numbers, messages, account timezones, and account names
    const callerNumberNames = include_caller
      ? await fetchCallerNumbers(api_key, account_ids, log, "normal")
      : {};

    // Fetch account timezones for discrepancy detection
    const accountTimezones = await fetchAccountTimezones(api_key, account_ids, log, "normal");

    // Fetch account names
    const accountNames = await fetchAccountNames(api_key, account_ids, log, "normal");

    const messageInfo = {};
    if (include_message_meta) {
      // Fetch message metadata for all accounts
      let _loggedMsgFields = false;
      for (const accountId of account_ids) {
        try {
          const data = await retryableApiCall(`/accounts/${accountId}/messages?filter=all`, api_key, log, "normal");
          if (Array.isArray(data.messages)) {
            if (!_loggedMsgFields && data.messages.length > 0) {
              log(`[AI] Messages API fields: ${Object.keys(data.messages[0]).join(', ')}`);
              _loggedMsgFields = true;
            }
            for (const msg of data.messages) {
              const key = `${accountId}:${msg.id}`;
              messageInfo[key] = {
                name: msg.name || '',
                description: msg.description || '',
                file_url: msg.file_url || msg.audio_url || msg.recording_url || msg.url || ''
              };
            }
          }
        } catch (err) {
          log(`⚠️  Could not fetch messages for account ${accountId}: ${err.message}`);
        }
      }
    }

    // Search campaigns
    const allMatches = [];
    // Normalize phone numbers - strip non-digits, handle +1 prefix
    const normalizedNumbers = numbers.map(n => {
      let num = String(n).replace(/\D/g, '');
      // Remove leading 1 if 11 digits starting with 1
      if (num.length === 11 && num.startsWith('1')) {
        num = num.slice(1);
      }
      return num;
    }).filter(n => n.length === 10);
    const numberSet = new Set(normalizedNumbers);

    log(`\n🔍 Searching ${campaigns.length} campaigns for ${numberSet.size} numbers...`);

    // Initialize progress for search phase
    if (job_id) {
      sendProgress(job_id, {
        status: 'running',
        total: campaigns.length,
        current: 0,
        message: 'Searching campaigns...'
      });
    }

    for (let i = 0; i < campaigns.length; i++) {
      if (job_id && jobs.get(job_id)?.cancelled) {
        throw new Error("Cancelled");
      }

      const campaign = campaigns[i];
      const accountId = campaign.account_id;
      const campaignId = campaign.id;
      
      log(`\n[${i + 1}/${campaigns.length}] Campaign ${campaignId} - ${campaign.name || 'Unnamed'}`);

      // Update progress
      if (job_id) {
        sendProgress(job_id, {
          current: i + 1,
          message: `Searching: ${campaign.name || `Campaign ${campaignId}`}`
        });
      }
      try {
        // Get campaign detail to extract export URL
        const detail = await fetchCampaignDetail(api_key, accountId, campaignId);
        const exportUrl = detail.export || detail.campaign?.export || null;
        
        if (!exportUrl) {
          log(`   ⚠️  No export URL available`);
          continue;
        }

        // Fetch CSV from S3 (NO authentication - it's a pre-signed URL)
        const expResp = await fetch(exportUrl);
        if (!expResp.ok) {
          throw new Error(`Failed to download CSV: HTTP ${expResp.status}`);
        }
        
        const csvText = await expResp.text();
        const { rows } = parseCsv(csvText);

        // Get campaign-level caller_number and message_id for fallback
        const campaignCallerNumber = campaign.caller_number || '';
        const campaignMessageId = campaign.message_id ? String(campaign.message_id) : '';

        const matches = rows.filter(row => {
          const rawNum = String(row.number || row.phone_number || "").replace(/\D/g, '');
          const num = rawNum.length === 11 && rawNum.startsWith('1') ? rawNum.slice(1) : rawNum;
          return num.length === 10 && numberSet.has(num);
        });

        if (matches.length > 0) {
          log(`   ✅ ${matches.length} matches found`);

          for (const row of matches) {
            const rawNum = String(row.number || row.phone_number || "").replace(/\D/g, '');
            const num = rawNum.length === 11 && rawNum.startsWith('1') ? rawNum.slice(1) : rawNum;

            // Use CSV value first, then campaign-level fallback
            const callerNum = row.voapps_caller_number || row.caller_number || campaignCallerNumber;
            const messageId = row.voapps_message_id || row.message_id || campaignMessageId;
            const callerKey = `${accountId}:${callerNum}`;
            const messageKey = `${accountId}:${messageId}`;

            allMatches.push({
              number: num,
              account_id: accountId,
              account_name: accountNames[accountId] || '',
              campaign_id: campaignId,
              campaign_name: campaign.name || '',
              caller_number: callerNum,
              caller_number_name: callerNumberNames[callerKey] || '',
              message_id: messageId,
              message_name: messageInfo[messageKey]?.name || '',
              message_description: messageInfo[messageKey]?.description || '',
              voapps_result: row.voapps_result || '',
              voapps_code: row.voapps_code || '',
              voapps_timestamp: normalizeToVoAppsTime(row.voapps_timestamp || ''),
              campaign_url: `https://directdropvoicemail.voapps.com/accounts/${accountId}/campaigns/${campaignId}`,
              target_date: campaign.target_date || ''
            });
          }
        }
      } catch (err) {
        log(`   ❌ Error: ${err.message}`, true);
      }
    }

    log(`\n📊 Search complete: ${allMatches.length.toLocaleString()} total matches`);

    // Save to database if requested
    if (output_mode === "database" || output_mode === "both") {
      log(`\n💾 Saving to database...`);
      const dbResult = await insertRows(allMatches, log);
      log(`   ✅ Database: ${dbResult.inserted} inserted, ${dbResult.updated} updated`);
    }

    // Save to CSV if requested
    let csvPath = null;
    let allCsvFiles = [];
    let wasSplit = false;
    let fileCount = 1;

    if (output_mode === "csv" || output_mode === "both") {
      csvPath = path.join(folders.phoneHistory, `${filePrefix}phone_search_${suffix}.csv`);

      const headers = [
        'number', 'account_id', 'account_name', 'campaign_id', 'campaign_name',
        'caller_number', 'caller_number_name', 'message_id', 'message_name', 'message_description',
        'voapps_result', 'voapps_code', 'voapps_timestamp', 'campaign_url'
      ];

      const csvResult = await writeCsv(csvPath, allMatches, headers, log, MAX_ROWS_PER_FILE);
      allCsvFiles = csvResult.files;
      wasSplit = csvResult.wasSplit;
      fileCount = csvResult.fileCount;

      lastArtifacts.csvPath = csvPath;
    }

    // Check for invalid result codes (408/409/410)
    const invalidCodeAlerts = checkInvalidResultCodes(allMatches, log);

    log(`\n✅ Search complete!`);

    // Send completion signal
    if (job_id) {
      sendProgress(job_id, {
        status: 'complete',
        progress: 100,
        message: `Search complete! Found ${allMatches.length.toLocaleString()} matches`
      });
    }

    close();

    return {
      csvPath,
      allCsvFiles,
      logPath,
      matches: allMatches.length,
      invalidCodeAlerts,
      wasSplit,
      fileCount
    };
  } catch (err) {
    log(`\n❌ Fatal error: ${err.message}`, true);
    close();
    throw err;
  }
}

/**
 * Generate Executive Summary - aggregate campaign statistics into deliverability report
 * Produces a CSV with per-campaign statistics:
 * campaign_id, campaign_name, account_id, target_date, records, deliverable,
 * successful_deliveries, expired, canceled, duplicate, unsuccessful_attempts, restricted, delivery_pct
 */
async function generateExecutiveSummary(config) {
  const {
    api_key,
    account_ids,
    start_date,
    end_date,
    job_id = null
  } = config;

  const folders = createOutputFolders();
  const suffix = getFilenameSuffix(folders.logs, 'voapps_log');
  const logPath = path.join(folders.logs, `voapps_log_${suffix}.txt`);
  const errorPath = path.join(folders.logs, `voapps_errors_${suffix}.txt`);

  const { log, close } = createLogger(logPath, errorPath, "normal", job_id);

  // Initialize job tracking with SSE support
  if (job_id) {
    jobs.set(job_id, {
      id: job_id,
      cancelled: false,
      stream: null,
      progress: 0,
      status: 'starting',
      current: 0,
      total: 0,
      message: 'Generating Executive Summary...',
      logs: [],
      logPath,
      errorPath
    });
  }

  // Get user's timezone preference
  const userTimezone = getTimezone();
  const userTimezoneLabel = getTimezoneLabel(userTimezone);

  log(`📊 Executive Summary Report`);
  log(`Accounts: ${account_ids.join(", ")}`);
  log(`Date Range: ${start_date} to ${end_date}`);
  log(`Timezone: ${userTimezoneLabel} (${userTimezone})`);

  try {
    // Fetch account names for the report
    const accountNames = await fetchAccountNames(api_key, account_ids, log, "normal");

    // Fetch campaigns
    const campaigns = await fetchAllCampaigns(api_key, account_ids, start_date, end_date, log, "normal", job_id);

    if (campaigns.length === 0) {
      log("\n⚠️  No campaigns found in date range");
      close();
      throw new Error("No campaigns found in specified date range");
    }

    log(`\n📊 Found ${campaigns.length} campaigns to analyze`);

    // Initialize progress
    if (job_id) {
      sendProgress(job_id, {
        status: 'running',
        total: campaigns.length,
        current: 0,
        message: 'Analyzing campaigns...'
      });
    }

    const summaryRows = [];
    let totalRecords = 0;

    for (let i = 0; i < campaigns.length; i++) {
      if (job_id && jobs.get(job_id)?.cancelled) {
        throw new Error("Cancelled");
      }

      const campaign = campaigns[i];
      const accountId = campaign.account_id;
      const campaignId = campaign.id;

      log(`\n[${i + 1}/${campaigns.length}] ${campaign.name || 'Unnamed'}`);

      // Update progress
      if (job_id) {
        sendProgress(job_id, {
          current: i + 1,
          message: `Analyzing: ${campaign.name || `Campaign ${campaignId}`}`
        });
      }

      try {
        // Get campaign detail to extract export URL
        const detail = await fetchCampaignDetail(api_key, accountId, campaignId);
        const exportUrl = detail.export || detail.campaign?.export || null;

        if (!exportUrl) {
          log(`   ⚠️  No export URL available`);
          continue;
        }

        // Fetch CSV from S3 (NO authentication)
        const expResp = await fetch(exportUrl);
        if (!expResp.ok) {
          throw new Error(`Failed to download CSV: HTTP ${expResp.status}`);
        }

        const csvText = await expResp.text();
        const { rows } = parseCsv(csvText);

        // Count results by voapps_code
        // Based on VoApps result codes:
        // 200 = Successfully delivered -> successful_deliveries (deliverable)
        // 300 = Expired -> expired (deliverable)
        // 301 = Canceled -> canceled (deliverable)
        // 400 = Unsuccessful delivery attempt -> unsuccessful_attempts (deliverable)
        // 401 = Not a wireless number -> (not deliverable, just a record)
        // 402 = Duplicate number -> duplicate (not deliverable)
        // 403 = Not a valid US number -> (not deliverable, just a record)
        // 404 = Undeliverable -> (not deliverable, just a record)
        // 405 = Not in service -> unsuccessful_attempts (deliverable)
        // 406 = Voicemail not setup -> unsuccessful_attempts (deliverable)
        // 407 = Voicemail full -> unsuccessful_attempts (deliverable)
        // 408 = Invalid caller number -> unfinished (deliverable)
        // 409 = Invalid message id -> unfinished (deliverable)
        // 410 = Prohibited self call -> unfinished (deliverable)
        // 500-504 = Restricted variants -> restricted (not deliverable)
        const resultCounts = {
          successful_deliveries: 0,
          expired: 0,
          canceled: 0,
          unsuccessful_attempts: 0,
          duplicate: 0,
          restricted: 0,
          unfinished: 0
        };

        for (const row of rows) {
          const code = String(row.voapps_code || '').trim();

          switch (code) {
            case '200':
              resultCounts.successful_deliveries++;
              break;
            case '300':
              resultCounts.expired++;
              break;
            case '301':
              resultCounts.canceled++;
              break;
            case '400':  // Unsuccessful delivery attempt
            case '405':  // Not in service
            case '406':  // Voicemail not setup
            case '407':  // Voicemail full
              resultCounts.unsuccessful_attempts++;
              break;
            case '402':
              resultCounts.duplicate++;
              break;
            case '408':  // Invalid caller number
            case '409':  // Invalid message id
            case '410':  // Prohibited self call
              resultCounts.unfinished++;
              break;
            case '500':  // Restricted
            case '501':  // Restricted for frequency
            case '502':  // Restricted geographical region
            case '503':  // Restricted individual number
            case '504':  // Restricted WebRecon
              resultCounts.restricted++;
              break;
            // 401, 403, 404 are just counted as records (not categorized)
            default:
              break;
          }
        }

        const records = rows.length;
        totalRecords += records;

        // Deliverable = successful + expired + canceled + unsuccessful + unfinished
        // (NOT duplicate, restricted, or the non-deliverable codes 401/403/404)
        const deliverable = resultCounts.successful_deliveries +
                           resultCounts.expired +
                           resultCounts.canceled +
                           resultCounts.unsuccessful_attempts +
                           resultCounts.unfinished;

        // Delivery % = successful / deliverable (avoid divide by zero)
        const deliveryPct = deliverable > 0 ? resultCounts.successful_deliveries / deliverable : 0;

        // Normalize target_date to user's selected timezone
        const normalizedTargetDate = campaign.target_date
          ? normalizeToVoAppsTime(campaign.target_date, userTimezone)
          : '';

        summaryRows.push({
          campaign_id: campaignId,
          campaign_name: campaign.name || '',
          account_id: accountId,
          account_name: accountNames[accountId] || '',
          target_date: normalizedTargetDate,
          records: records,
          deliverable: deliverable,
          successful_deliveries: resultCounts.successful_deliveries,
          expired: resultCounts.expired,
          canceled: resultCounts.canceled,
          duplicate: resultCounts.duplicate,
          unsuccessful_attempts: resultCounts.unsuccessful_attempts,
          unfinished: resultCounts.unfinished,
          restricted: resultCounts.restricted,
          delivery_pct: deliveryPct,
          campaign_url: `https://directdropvoicemail.voapps.com/accounts/${accountId}/campaigns/${campaignId}`
        });

        log(`   ✅ ${records.toLocaleString()} records, ${(deliveryPct * 100).toFixed(1)}% delivery rate`);
      } catch (err) {
        log(`   ❌ Error: ${err.message}`, true);
      }
    }

    log(`\n📊 Total: ${summaryRows.length} campaigns, ${totalRecords.toLocaleString()} records`);

    // Write CSV to Executive Summary folder
    const suffix = getFilenameSuffix(folders.executiveSummary, 'executive_summary');
    const csvPath = path.join(folders.executiveSummary, `executive_summary_${suffix}.csv`);

    const headers = [
      'campaign_id', 'campaign_name', 'account_id', 'account_name', 'target_date', 'records',
      'deliverable', 'successful_deliveries', 'expired', 'canceled', 'duplicate',
      'unsuccessful_attempts', 'unfinished', 'restricted', 'delivery_pct', 'campaign_url'
    ];

    let csvContent = headers.join(',') + '\n';
    for (const row of summaryRows) {
      const values = headers.map(h => {
        const val = row[h];
        if (h === 'delivery_pct') return (val * 100).toFixed(2) + '%';
        if (typeof val === 'string' && (val.includes(',') || val.includes('"'))) {
          return `"${val.replace(/"/g, '""')}"`;
        }
        return val;
      });
      csvContent += values.join(',') + '\n';
    }

    fs.writeFileSync(csvPath, csvContent, 'utf8');
    log(`\n✅ Saved: ${path.basename(csvPath)}`);

    lastArtifacts.csvPath = csvPath;
    lastArtifacts.logPath = logPath;

    close();
    return {
      csvPath,
      logPath,
      campaignCount: summaryRows.length,
      totalRecords
    };
  } catch (err) {
    log(`\n❌ Fatal error: ${err.message}`, true);
    close();
    throw err;
  }
}

async function runCombineCampaigns(config) {
  const {
    api_key,
    account_ids,
    start_date,
    end_date,
    include_caller = true,
    include_message_meta = true,
    generate_trend_analysis = false,
    min_consec_unsuccessful = 4,
    min_run_span_days = 30,
    include_detail_tabs = false, // TN Health, Variability Analysis, Number Summary tabs
    output_mode = "csv", // "csv", "database", or "both"
    job_id = null,
    client_prefix = "", // Optional prefix for output files
    // AI settings come from the frontend request payload (UI/localStorage).
    // getAiSettings() reads from disk and never sees the UI toggle state.
    ai_enabled = false,
    ai_transcription_mode = 'local',
    ai_intent_mode = 'local'
  } = config;

  // Build filename prefix
  const filePrefix = client_prefix ? `${client_prefix}_` : "";

  const folders = createOutputFolders();
  const suffix = getFilenameSuffix(folders.logs, 'voapps_log');
  const logPath = path.join(folders.logs, `voapps_log_${suffix}.txt`);
  const errorPath = path.join(folders.logs, `voapps_errors_${suffix}.txt`);

  const { log, close } = createLogger(logPath, errorPath, "normal", job_id);

  // Initialize job tracking with SSE support
  if (job_id) {
    jobs.set(job_id, {
      id: job_id,
      cancelled: false,
      stream: null,
      progress: 0,
      status: 'starting',
      current: 0,
      total: 0,
      message: 'Initializing combine...',
      logs: [],
      startTime: Date.now()
    });
    sendProgress(job_id, { status: 'starting', message: 'Initializing combine campaigns...' });
  }

  lastArtifacts.logPath = logPath;
  lastArtifacts.errorPath = errorPath;

  try {
    log(`=== VoApps Tools v${VERSION} - Combine Campaigns ===`);
    log(`Output Mode: ${output_mode}`);
    log(`Accounts: ${account_ids.join(", ")}`);
    log(`Date Range: ${start_date} to ${end_date}`);
    if (generate_trend_analysis) {
      log(`Trend Analysis: Enabled (min_consec=${min_consec_unsuccessful}, min_span=${min_run_span_days} days)`);
    }

    // Fetch campaigns
    const campaigns = await fetchAllCampaigns(api_key, account_ids, start_date, end_date, log, "normal", job_id);

    if (campaigns.length === 0) {
      log("\n⚠️  No campaigns found in date range");
      close();
      throw new Error("No campaigns found in specified date range");
    }

    log(`\n📊 Found ${campaigns.length} campaigns to combine`);

    // Fetch caller numbers, messages, account timezones, and account names
    const callerNumberNames = include_caller
      ? await fetchCallerNumbers(api_key, account_ids, log, "normal")
      : {};

    // Fetch account timezones for discrepancy detection
    const accountTimezones = await fetchAccountTimezones(api_key, account_ids, log, "normal");

    // Fetch account names
    const accountNames = await fetchAccountNames(api_key, account_ids, log, "normal");

    const messageInfo = {};
    if (include_message_meta) {
      let _loggedMsgFields = false;
      for (const accountId of account_ids) {
        try {
          const data = await retryableApiCall(`/accounts/${accountId}/messages?filter=all`, api_key, log, "normal");
          if (Array.isArray(data.messages)) {
            if (!_loggedMsgFields && data.messages.length > 0) {
              log(`[AI] Messages API fields: ${Object.keys(data.messages[0]).join(', ')}`);
              _loggedMsgFields = true;
            }
            for (const msg of data.messages) {
              const key = `${accountId}:${msg.id}`;
              messageInfo[key] = {
                name: msg.name || '',
                description: msg.description || '',
                file_url: msg.file_url || msg.audio_url || msg.recording_url || msg.url || ''
              };
            }
          }
        } catch (err) {
          log(`⚠️  Could not fetch messages for account ${accountId}: ${err.message}`);
        }
      }
    }

    // Fetch all campaign reports
    const allRows = [];

    log(`\n📥 Downloading campaign reports...`);

    // Initialize progress for download phase
    if (job_id) {
      sendProgress(job_id, {
        status: 'running',
        total: campaigns.length,
        current: 0,
        message: 'Downloading campaigns...'
      });
    }

    for (let i = 0; i < campaigns.length; i++) {
      if (job_id && jobs.get(job_id)?.cancelled) {
        throw new Error("Cancelled");
      }

      const campaign = campaigns[i];
      const accountId = campaign.account_id;
      const campaignId = campaign.id;
      
      log(`\n[${i + 1}/${campaigns.length}] ${campaign.name || 'Unnamed'}`);

      // Update progress
      if (job_id) {
        sendProgress(job_id, {
          current: i + 1,
          message: `Downloading: ${campaign.name || `Campaign ${campaignId}`}`
        });
      }
      try {
        // Get campaign detail to extract export URL
        const detail = await fetchCampaignDetail(api_key, accountId, campaignId);
        const exportUrl = detail.export || detail.campaign?.export || null;
        
        if (!exportUrl) {
          log(`   ⚠️  No export URL available`);
          continue;
        }

        // Fetch CSV from S3 (NO authentication)
        const expResp = await fetch(exportUrl);
        if (!expResp.ok) {
          throw new Error(`Failed to download CSV: HTTP ${expResp.status}`);
        }
        
        const csvText = await expResp.text();
        const { rows } = parseCsv(csvText);

        // Get campaign-level caller_number and message_id for fallback
        const campaignCallerNumber = campaign.caller_number || '';
        const campaignMessageId = campaign.message_id ? String(campaign.message_id) : '';

        for (const row of rows) {
          const rawNum = String(row.number || row.phone_number || "").replace(/\D/g, '');
          const num = rawNum.length === 11 && rawNum.startsWith('1') ? rawNum.slice(1) : rawNum;

          if (num.length !== 10) continue; // Skip invalid numbers

          // Use CSV value first, then campaign-level fallback
          const callerNum = row.voapps_caller_number || row.caller_number || campaignCallerNumber;
          const messageId = row.voapps_message_id || row.message_id || campaignMessageId;
          const callerKey = `${accountId}:${callerNum}`;
          const messageKey = `${accountId}:${messageId}`;

          allRows.push({
            number: num,
            account_id: accountId,
            account_name: accountNames[accountId] || '',
            campaign_id: campaignId,
            campaign_name: campaign.name || '',
            caller_number: callerNum,
            caller_number_name: callerNumberNames[callerKey] || '',
            message_id: messageId,
            message_name: messageInfo[messageKey]?.name || '',
            message_description: messageInfo[messageKey]?.description || '',
            voapps_result: row.voapps_result || '',
            voapps_code: row.voapps_code || '',
            voapps_timestamp: normalizeToVoAppsTime(row.voapps_timestamp || ''),
            campaign_url: `https://directdropvoicemail.voapps.com/accounts/${accountId}/campaigns/${campaignId}`,
            target_date: campaign.target_date || ''
          });
        }

        log(`   ✅ ${rows.length.toLocaleString()} rows`);
      } catch (err) {
        log(`   ❌ Error: ${err.message}`, true);
      }
    }

    log(`\n📊 Total rows: ${allRows.length.toLocaleString()}`);

    // Fill missing voapps_timestamp from nearest same-campaign record (by row proximity),
    // falling back to target_date if no campaign record has any timestamp.
    {
      const ctsMap = new Map();
      for (let i = 0; i < allRows.length; i++) {
        const ts = allRows[i].voapps_timestamp;
        if (ts) {
          const cid = allRows[i].campaign_id || '__none__';
          if (!ctsMap.has(cid)) ctsMap.set(cid, []);
          ctsMap.get(cid).push({ idx: i, ts });
        }
      }
      // Binary search helper — find nearest ts by row index within a campaign
      const nearestTs = (list, targetIdx) => {
        if (!list || list.length === 0) return null;
        let lo = 0, hi = list.length - 1;
        while (lo < hi) { const mid = (lo + hi) >> 1; if (list[mid].idx < targetIdx) lo = mid + 1; else hi = mid; }
        const prev = lo > 0 ? list[lo - 1] : null;
        const next = lo < list.length ? list[lo] : null;
        if (!prev) return next.ts;
        if (!next) return prev.ts;
        return Math.abs(prev.idx - targetIdx) <= Math.abs(next.idx - targetIdx) ? prev.ts : next.ts;
      };
      let filled = 0;
      for (let i = 0; i < allRows.length; i++) {
        if (!allRows[i].voapps_timestamp) {
          const cid = allRows[i].campaign_id || '__none__';
          const ts = nearestTs(ctsMap.get(cid), i);
          if (ts) { allRows[i].voapps_timestamp = ts; filled++; }
          else if (allRows[i].target_date) {
            allRows[i].voapps_timestamp = `${allRows[i].target_date} 00:00:00 UTC`;
            filled++;
          }
        }
      }
      ctsMap.clear();
      if (filled > 0) log(`   ✅ Filled ${filled.toLocaleString()} missing timestamp(s) from campaign proximity`);
    }

    // Save to database if requested
    if (output_mode === "database" || output_mode === "both") {
      log(`\n💾 Saving to database...`);
      const dbResult = await insertRows(allRows, log);
      log(`   ✅ Database: ${dbResult.inserted} inserted, ${dbResult.updated} updated`);
    }

    // Save to CSV if requested
    let csvPath = null;
    let allCsvFiles = [];
    let wasSplit = false;
    let fileCount = 1;
    let tempAnalysisCsvFiles = []; // temp files created only for analysis in database-only mode

    const CSV_HEADERS = [
      'number', 'account_id', 'account_name', 'campaign_id', 'campaign_name',
      'caller_number', 'caller_number_name', 'message_id', 'message_name', 'message_description',
      'voapps_result', 'voapps_code', 'voapps_timestamp', 'campaign_url'
    ];

    if (output_mode === "csv" || output_mode === "both") {
      csvPath = path.join(folders.combineCampaigns, `${filePrefix}combined_${suffix}.csv`);
      const csvResult = await writeCsv(csvPath, allRows, CSV_HEADERS, log, MAX_ROWS_PER_FILE);
      allCsvFiles = csvResult.files;
      wasSplit = csvResult.wasSplit;
      fileCount = csvResult.fileCount;
      lastArtifacts.csvPath = csvPath;
    } else if (generate_trend_analysis) {
      // Database-only output but analysis was requested: write a temporary CSV so the
      // worker can read from disk (same pattern as CSV mode). Deleted after analysis.
      log(`\n📊 Writing temporary CSV for analysis (database-only output)...`);
      const tmpCsvPath = path.join(folders.combineCampaigns, `${filePrefix}combined_${suffix}_analysis_tmp.csv`);
      const csvResult = await writeCsv(tmpCsvPath, allRows, CSV_HEADERS, log, MAX_ROWS_PER_FILE);
      allCsvFiles = csvResult.files;
      tempAnalysisCsvFiles = csvResult.files;
    }

    // Check for invalid result codes (408/409/410) before clearing allRows
    const invalidCodeAlerts = checkInvalidResultCodes(allRows, log);

    // Capture row count then free the in-memory rows array before spawning the worker.
    // This is critical for large datasets: the worker reads from the CSV files on disk,
    // so holding allRows in memory at the same time would double memory usage and OOM.
    const totalRows = allRows.length;
    // Capture AI scoping data BEFORE clearing allRows — the AI block needs to know which
    // messages/accounts were actually represented in this dataset, but runs after the clear.
    const _aiUsedMessageKeys = new Set(
      allRows.filter(r => r.account_id && r.message_id).map(r => `${r.account_id}:${r.message_id}`)
    );
    const _aiActiveAccountIds = new Set(allRows.map(r => r.account_id).filter(Boolean));
    allRows.length = 0;

    // Generate trend analysis if requested
    let analysisPath = null;
    if (generate_trend_analysis && allCsvFiles.length > 0) {
      log(`\n📊 Generating trend analysis...`);

      const analysisFilename = `${filePrefix}number_analysis_${suffix}.xlsx`;
      analysisPath = path.join(folders.combineCampaigns, analysisFilename);

      // Always pass file paths (never raw rows) so the worker reads from disk
      // while the main thread keeps no large array in memory.
      const userTimezone = getTimezone();
      const userTimezoneLabel = getTimezoneLabel(userTimezone);

      // AI Message Analysis (optional, opt-in)
      // Re-fetch message list right before transcription regardless of include_message_meta:
      //   1. The initial messageInfo fetch is gated by include_message_meta (CSV column selection)
      //      but AI needs file_urls even when those columns aren't selected.
      //   2. Pre-signed S3 URLs expire after 1 hour — fetching fresh here avoids stale URLs
      //      on long-running combine jobs.
      // Build aiSettings from the request payload values (enable/mode come from the UI
      // via localStorage, so they're sent in the request body).  The OpenAI API key is
      // stored server-side via /api/settings/openai-key, so we still read that from disk.
      const aiSettings = {
        enabled: ai_enabled,
        transcriptionMode: ai_transcription_mode,
        intentMode: ai_intent_mode,
        openaiApiKey: getAiSettings().openaiApiKey,
        notify: (message, actionLabel, url) => sendNotify(jobId, message, actionLabel, url)
      };
      if (ai_enabled) {
        log(`[AI] Message Analysis enabled (stt: ${ai_transcription_mode}, intent: ${ai_intent_mode})`);
      }
      let transcriptMap = {};
      if (aiSettings.enabled) {
        try {
          log('[AI] Fetching fresh message audio URLs...');
          const aiMessageInfo = { ...messageInfo }; // keep any name/description already collected
          for (const accountId of account_ids) {
            try {
              const freshData = await retryableApiCall(`/accounts/${accountId}/messages?filter=all`, api_key, log, "normal");
              if (Array.isArray(freshData.messages)) {
                for (const msg of freshData.messages) {
                  const key = `${accountId}:${msg.id}`;
                  const existing = aiMessageInfo[key] || {};
                  aiMessageInfo[key] = {
                    name: existing.name || msg.name || '',
                    description: existing.description || msg.description || '',
                    file_url: msg.file_url || msg.audio_url || msg.recording_url || msg.url || ''
                  };
                }
              }
            } catch (urlErr) {
              log(`[AI] ⚠️  Could not fetch messages for account ${accountId}: ${urlErr.message}`);
            }
          }
          // Only transcribe messages that were actually used in this dataset, not every
          // message available for these accounts. This avoids wasting time and API/compute
          // resources on messages that didn't appear in any of the campaign rows being analyzed.
          //
          // Two-tier scoping strategy:
          //   Tier 1 (precise) — if campaign rows have reliable message_id values, filter by
          //     matching "accountId:messageId" keys. This is the ideal path.
          //   Tier 2 (fallback) — campaign.message_id is sometimes null/0 (API returns no
          //     message_id for some campaign types), so all rows end up with message_id=''.
          //     In that case, fall back to account-level scoping: include all messages that
          //     belong to accounts which have at least one row in this dataset. Messages from
          //     accounts with zero dataset rows are still excluded.
          //
          // NOTE: allRows was cleared (allRows.length = 0) earlier to free memory before the
          // worker spawns. The sets _aiUsedMessageKeys and _aiActiveAccountIds were captured
          // from allRows just before that clear, so they still contain the right data.
          const filteredAiMessageInfo = {};
          const totalCount = Object.keys(aiMessageInfo).length;
          if (_aiUsedMessageKeys.size > 0) {
            // Tier 1: precise message-level filter
            for (const [key, info] of Object.entries(aiMessageInfo)) {
              if (_aiUsedMessageKeys.has(key)) filteredAiMessageInfo[key] = info;
            }
          } else {
            // Tier 2: fallback — message IDs not available in row data; scope by active accounts
            for (const [key, info] of Object.entries(aiMessageInfo)) {
              const accountId = key.split(':')[0];
              if (_aiActiveAccountIds.has(accountId)) filteredAiMessageInfo[key] = info;
            }
            if (_aiActiveAccountIds.size > 0) {
              log(`[AI] Note: campaign rows have no message IDs — scoping by account (${_aiActiveAccountIds.size} active account(s))`);
            }
          }
          const usedCount = Object.keys(filteredAiMessageInfo).length;
          if (totalCount > usedCount) {
            log(`[AI] Scoped to ${usedCount} message(s) used in this dataset (${totalCount - usedCount} available but unused — skipped)`);
          }
          if (usedCount > 0) {
            transcriptMap = await transcribeAndAnalyzeMessages(filteredAiMessageInfo, aiSettings, log);
          } else {
            log('[AI] No messages from this dataset have audio URLs — skipping transcription');
          }
        } catch (aiErr) {
          log(`[AI] ⚠️  AI analysis failed, continuing without transcripts: ${aiErr.message}`);
        }
      }

      await runAnalysisInWorker(
        allCsvFiles,
        analysisPath,
        min_consec_unsuccessful,
        min_run_span_days,
        messageInfo,
        callerNumberNames,
        accountTimezones,
        userTimezone,
        userTimezoneLabel,
        include_detail_tabs,
        transcriptMap
      );

      lastArtifacts.analysisPath = analysisPath;
      log(`✅ Analysis generated: ${analysisFilename}`);

      // Remove temp CSVs that were created only to feed the analysis worker
      if (tempAnalysisCsvFiles.length > 0) {
        for (const f of tempAnalysisCsvFiles) {
          try { fs.unlinkSync(f); } catch {}
        }
        log(`🗑️  Removed ${tempAnalysisCsvFiles.length} temporary analysis CSV file(s)`);
      }
    }

    log(`\n✅ Combine complete!`);

    // Send completion signal
    if (job_id) {
      sendProgress(job_id, {
        status: 'complete',
        progress: 100,
        message: `Combine complete! Total: ${totalRows.toLocaleString()} rows`
      });
    }

    close();

    return {
      csvPath,
      allCsvFiles,
      logPath,
      analysisPath,
      totalRows,
      wasSplit,
      fileCount
    };
  } catch (err) {
    log(`\n❌ Fatal error: ${err.message}`, true);
    close();
    err.logPath = logPath;
    throw err;
  }
}

async function runBulkCampaignExport(config) {
  const {
    api_key,
    account_ids,
    start_date,
    end_date,
    job_id = null
  } = config;

  const folders = createOutputFolders();
  const suffix = getFilenameSuffix(folders.logs, 'voapps_log');
  const logPath = path.join(folders.logs, `voapps_log_${suffix}.txt`);
  const errorPath = path.join(folders.logs, `voapps_errors_${suffix}.txt`);

  const { log, close } = createLogger(logPath, errorPath, "normal", job_id);

  // Initialize job tracking with SSE support
  if (job_id) {
    jobs.set(job_id, {
      id: job_id,
      cancelled: false,
      stream: null,
      progress: 0,
      status: 'starting',
      current: 0,
      total: 0,
      message: 'Initializing export...',
      logs: [],
      startTime: Date.now()
    });
    sendProgress(job_id, { status: 'starting', message: 'Initializing bulk export...' });
  }

  lastArtifacts.logPath = logPath;
  lastArtifacts.errorPath = errorPath;

  try {
    log(`=== VoApps Tools v${VERSION} - Bulk Campaign Export ===`);
    log(`Accounts: ${account_ids.join(", ")}`);
    log(`Date Range: ${start_date} to ${end_date}`);

    // Fetch campaigns
    const campaigns = await fetchAllCampaigns(api_key, account_ids, start_date, end_date, log, "normal", job_id);

    if (campaigns.length === 0) {
      log("\n⚠️  No campaigns found in date range");
      close();
      throw new Error("No campaigns found in specified date range");
    }

    log(`\n📊 Found ${campaigns.length} campaigns to export`);

    const stats = {
      total: campaigns.length,
      exported: 0,
      failed: 0,
      totalRows: 0
    };

    log(`\n📥 Exporting campaigns...`);

    // Initialize progress for export phase
    if (job_id) {
      sendProgress(job_id, {
        status: 'running',
        total: campaigns.length,
        current: 0,
        message: 'Exporting campaigns...'
      });
    }

    for (let i = 0; i < campaigns.length; i++) {
      if (job_id && jobs.get(job_id)?.cancelled) {
        throw new Error("Cancelled");
      }

      const campaign = campaigns[i];
      const accountId = campaign.account_id;
      const campaignId = campaign.id;
      const targetDate = campaign.target_date || 'unknown';
      const year = targetDate.split('-')[0];
      const month = targetDate.split('-')[1];

      const campaignDir = path.join(folders.bulkExport, year, month);
      fs.mkdirSync(campaignDir, { recursive: true });

      const safeName = (campaign.name || 'Unnamed').replace(/[^a-zA-Z0-9-_ ]/g, '_').substring(0, 100);
      const filename = `${safeName}_${targetDate}.csv`;
      const filePath = path.join(campaignDir, filename);

      log(`\n[${i + 1}/${campaigns.length}] ${campaign.name || 'Unnamed'}`);

      // Update progress
      if (job_id) {
        sendProgress(job_id, {
          current: i + 1,
          message: `Exporting: ${campaign.name || `Campaign ${campaignId}`}`
        });
      }
      try {
        // Get campaign detail to extract export URL
        const detail = await fetchCampaignDetail(api_key, accountId, campaignId);
        const exportUrl = detail.export || detail.campaign?.export || null;

        if (!exportUrl) {
          stats.failed++;
          log(`   ⚠️  No export URL available (Campaign ID: ${campaignId}, Account: ${accountId})`, true);
          log(`   Campaign may still be processing or incomplete in VoApps`, true);
          continue;
        }

        // Get campaign-level caller_number and message_id for fallback
        const campaignCallerNumber = campaign.caller_number || '';
        const campaignMessageId = campaign.message_id ? String(campaign.message_id) : '';

        // Fetch CSV from S3 (NO authentication)
        const expResp = await fetch(exportUrl);
        if (!expResp.ok) {
          throw new Error(`Failed to download CSV: HTTP ${expResp.status}`);
        }

        const csvText = await expResp.text();
        const { headers, rows } = parseCsv(csvText);

        // Check if CSV has voapps_caller_number and voapps_message_id columns
        const hasCallerNumberCol = headers.includes('voapps_caller_number');
        const hasMessageIdCol = headers.includes('voapps_message_id');

        // Determine output headers - add columns if missing
        const outputHeaders = [...headers];
        if (!hasCallerNumberCol) {
          outputHeaders.push('voapps_caller_number');
        }
        if (!hasMessageIdCol) {
          outputHeaders.push('voapps_message_id');
        }

        // Enrich rows with campaign-level data if CSV columns are missing
        const enrichedRows = rows.map(row => {
          const enriched = { ...row };

          // If CSV doesn't have voapps_caller_number, use campaign-level value
          if (!hasCallerNumberCol) {
            enriched.voapps_caller_number = campaignCallerNumber;
          }

          // If CSV doesn't have voapps_message_id, use campaign-level value
          if (!hasMessageIdCol) {
            enriched.voapps_message_id = campaignMessageId;
          }

          return enriched;
        });

        // Write enriched CSV
        const csvContent = [
          outputHeaders.join(','),
          ...enrichedRows.map(row => outputHeaders.map(h => {
            const val = row[h];
            if (val === null || val === undefined) return '';
            const str = String(val);
            return str.includes(',') || str.includes('"') ? `"${str.replace(/"/g, '""')}"` : str;
          }).join(','))
        ].join('\n');

        await fsp.writeFile(filePath, csvContent, 'utf-8');

        stats.exported++;
        stats.totalRows += rows.length;

        // Log whether we enriched the data
        const enrichmentNote = (!hasCallerNumberCol || !hasMessageIdCol)
          ? ` (enriched: caller=${!hasCallerNumberCol ? campaignCallerNumber : 'from CSV'}, msg=${!hasMessageIdCol ? campaignMessageId : 'from CSV'})`
          : '';
        log(`   ✅ ${rows.length.toLocaleString()} rows → ${year}/${month}/${filename}${enrichmentNote}`);

        // Check for invalid result codes (408/409/410) in this campaign's rows
        checkInvalidResultCodes(enrichedRows, log);
      } catch (err) {
        stats.failed++;
        log(`   ❌ Error: ${err.message}`, true);
      }
    }

    log(`\n📊 Export Summary:`);
    log(`   Total: ${stats.total}`);
    log(`   Exported: ${stats.exported}`);
    log(`   Failed: ${stats.failed}`);
    log(`   Total Rows: ${stats.totalRows.toLocaleString()}`);

    log(`\n✅ Bulk export complete!`);
    
    // Send completion signal
    if (job_id) {
      sendProgress(job_id, {
        status: 'complete',
        progress: 100,
        current: stats.total,
        message: 'Export complete!'
      });
    }
    
    close();

    return {
      bulkExportPath: folders.bulkExport,
      logPath,
      errorPath,
      stats
    };
  } catch (err) {
    log(`\n❌ Fatal error: ${err.message}`, true);
    close();
    throw err;
  }
}

/**
 * Helper: Export subset of database to CSV using streaming to avoid OOM
 */
async function exportDatabaseSubset(numbers, accountIds, startDate, endDate, outputPath) {
  if (!isDatabaseAvailable()) {
    return { success: false, error: 'Database not available' };
  }
  if (!dbReady) await initDatabase();

  const numberList = numbers.map(n => `'${n}'`).join(',');
  const accountList = accountIds.map(a => `'${a}'`).join(',');

  // Build base WHERE clause
  const whereClause = `
    WHERE number IN (${numberList})
      AND account_id IN (${accountList})
      AND target_date >= '${startDate}'
      AND target_date <= '${endDate}'
  `;

  // Get count first
  const countResult = await runQuery(`SELECT COUNT(*) as cnt FROM campaign_results ${whereClause}`);
  const totalRows = Number(countResult[0]?.cnt || 0);

  if (totalRows === 0) {
    return { success: false, error: 'No data found' };
  }

  // Get column names, excluding internal columns
  const schemaResult = await runQuery(`
    SELECT column_name FROM information_schema.columns
    WHERE table_name = 'campaign_results'
      AND column_name NOT IN ('row_id', 'created_at', 'updated_at')
    ORDER BY ordinal_position
  `);
  const headers = schemaResult.map(r => r.column_name);

  // Create write stream for memory-efficient export
  const writeStream = fs.createWriteStream(outputPath, { encoding: 'utf-8' });
  writeStream.write(headers.join(',') + '\n');

  // Stream rows in batches to avoid OOM
  const BATCH_SIZE = 50000;
  let offset = 0;
  let rowCount = 0;

  while (offset < totalRows) {
    const batchRows = await runQuery(`
      SELECT ${headers.join(', ')} FROM campaign_results
      ${whereClause}
      ORDER BY target_date DESC, voapps_timestamp DESC
      LIMIT ${BATCH_SIZE} OFFSET ${offset}
    `);

    if (batchRows.length === 0) break;

    for (const row of batchRows) {
      const line = headers.map(h => {
        let val = row[h];
        if (val === null || val === undefined) return '';
        // Normalize timestamps to VoApps Time (UTC-7)
        if (h === 'voapps_timestamp' && val) {
          val = normalizeToVoAppsTime(val);
        }
        const str = String(val);
        return str.includes(',') || str.includes('"') ? `"${str.replace(/"/g, '""')}"` : str;
      }).join(',');
      writeStream.write(line + '\n');
    }

    rowCount += batchRows.length;
    offset += BATCH_SIZE;
  }

  await new Promise((resolve, reject) => {
    writeStream.end((err) => err ? reject(err) : resolve());
  });

  return {
    success: true,
    path: outputPath,
    rowCount
  };
}

// =============================================================================
// HTTP SERVER (Enhanced for v3.0.0)
// =============================================================================

function createHttpServer() {
  return http.createServer(async (req, res) => {
    try {
      const { pathname } = parseUrl(req.url, true);

      // CORS headers
      res.setHeader('Access-Control-Allow-Origin', '*');
      res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
      res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

      if (req.method === "OPTIONS") {
        res.writeHead(200);
        return res.end();
      }

    // Ping endpoint
    if (req.method === "GET" && pathname === "/api/ping") {
      return sendJson(res, 200, { ok: true, message: "VoApps Tools Server", version: VERSION, versionName: VERSION_NAME });
    }

    // Settings - Get output folder
    if (req.method === "GET" && pathname === "/api/settings/output-folder") {
      return sendJson(res, 200, { ok: true, folder: getOutputFolder() });
    }

    // Settings - Set output folder
    if (req.method === "POST" && pathname === "/api/settings/output-folder") {
      try {
        const body = await readJson(req);
        const { folder } = body;
        if (!folder) {
          return sendJson(res, 400, { ok: false, error: "Folder path required" });
        }
        // Validate the folder exists or can be created
        try {
          fs.mkdirSync(folder, { recursive: true });
        } catch (e) {
          return sendJson(res, 400, { ok: false, error: `Cannot create folder: ${e.message}` });
        }
        const success = setOutputFolder(folder);
        if (success) {
          return sendJson(res, 200, { ok: true, folder });
        } else {
          return sendJson(res, 500, { ok: false, error: "Failed to save settings" });
        }
      } catch (e) {
        return sendJson(res, 500, { ok: false, error: e.message });
      }
    }

    // Settings - Ensure output folder exists (creates it if needed)
    if (req.method === "POST" && pathname === "/api/settings/ensure-folder") {
      try {
        const folder = getOutputFolder();
        fs.mkdirSync(folder, { recursive: true });
        return sendJson(res, 200, { ok: true, folder, exists: true });
      } catch (e) {
        return sendJson(res, 500, { ok: false, error: e.message });
      }
    }

    // Settings - Get timezone
    if (req.method === "GET" && pathname === "/api/settings/timezone") {
      const timezone = getTimezone();
      const label = getTimezoneLabel(timezone);
      return sendJson(res, 200, { ok: true, timezone, label });
    }

    // Settings - Set timezone
    if (req.method === "POST" && pathname === "/api/settings/timezone") {
      try {
        const body = await readJson(req);
        const { timezone } = body;
        if (!timezone) {
          return sendJson(res, 400, { ok: false, error: "Timezone required" });
        }
        // Validate format - accept IANA names, 'VoApps', 'UTC', or legacy offset format
        const validTimezones = ['VoApps', 'UTC', 'America/New_York', 'America/Chicago', 'America/Denver', 'America/Los_Angeles'];
        const isValidIANA = validTimezones.includes(timezone);
        const isValidOffset = /^[+-]\d{2}:\d{2}$/.test(timezone);
        if (!isValidIANA && !isValidOffset) {
          return sendJson(res, 400, { ok: false, error: "Invalid timezone. Use IANA name (e.g., 'America/New_York') or 'VoApps'" });
        }
        const success = setTimezone(timezone);
        if (success) {
          const label = getTimezoneLabel(timezone);
          return sendJson(res, 200, { ok: true, timezone, label });
        } else {
          return sendJson(res, 500, { ok: false, error: "Failed to save settings" });
        }
      } catch (e) {
        return sendJson(res, 500, { ok: false, error: e.message });
      }
    }

    // Settings - Get OpenAI API key
    if (req.method === "GET" && pathname === "/api/settings/openai-key") {
      const key = getOpenaiKey();
      return sendJson(res, 200, { ok: true, key });
    }

    // Settings - Set OpenAI API key
    if (req.method === "POST" && pathname === "/api/settings/openai-key") {
      try {
        const body = await readJson(req);
        const { key } = body;
        if (!key || typeof key !== 'string') return sendJson(res, 400, { ok: false, error: 'key required' });
        const success = setOpenaiKey(key.trim());
        return sendJson(res, success ? 200 : 500, { ok: success });
      } catch (e) {
        return sendJson(res, 500, { ok: false, error: e.message });
      }
    }

    // AI - Download local model
    if (req.method === "POST" && pathname === "/api/ai/download-model") {
      try {
        const body = await readJson(req);
        const { type, job_id } = body; // 'stt' or 'intent'
        if (type !== 'stt' && type !== 'intent') {
          return sendJson(res, 400, { ok: false, error: 'type must be stt or intent' });
        }

        // Register a job so the SSE stream endpoint can attach and replay logs
        const jobId = job_id || `dl_${Date.now()}`;
        jobs.set(jobId, { status: 'running', progress: 0, logs: [], stream: null, startTime: Date.now() });

        // Log function pipes through the SSE stream
        const log = (msg, isError = false) => {
          isError ? console.error(msg) : console.log(msg);
          sendLog(jobId, msg, isError);
        };

        // Fire-and-forget — logs stream live via SSE; on finish signal complete/error
        downloadAiModelBackground(type, log).then(() => {
          sendProgress(jobId, { status: 'complete', progress: 100 });
        }).catch(e => {
          sendLog(jobId, `[AI] Unexpected error: ${e.message}`, true);
          sendProgress(jobId, { status: 'error', progress: 0 });
        });

        return sendJson(res, 200, { ok: true, job_id: jobId, message: `${type === 'stt' ? 'Whisper' : 'Intent'} model download started.` });
      } catch (e) {
        return sendJson(res, 500, { ok: false, error: e.message });
      }
    }

    // AI - Model download status
    if (req.method === "GET" && pathname === "/api/ai/model-status") {
      const status = getAiModelStatus();
      return sendJson(res, 200, { ok: true, ...status });
    }

    // AI - Transcript cache stats
    if (req.method === "GET" && pathname === "/api/ai/cache-stats") {
      try {
        const rows = await runQuery('SELECT COUNT(*) as count FROM message_transcriptions');
        return sendJson(res, 200, { ok: true, count: Number(rows[0]?.count) || 0 });
      } catch (e) {
        return sendJson(res, 200, { ok: true, count: 0 }); // DB not ready — not an error
      }
    }

    // AI - Clear transcript cache
    if (req.method === "POST" && pathname === "/api/ai/clear-cache") {
      try {
        await runQuery('DELETE FROM message_transcriptions');
        console.log('[AI] Transcript cache cleared');
        return sendJson(res, 200, { ok: true, message: 'Transcript cache cleared' });
      } catch (e) {
        // If DB isn't ready there's nothing to clear — treat as success
        if (!db || e.message === 'Database not initialized') {
          return sendJson(res, 200, { ok: true, message: 'Cache is empty (database not initialized)' });
        }
        return sendJson(res, 500, { ok: false, error: e.message });
      }
    }

    // AI - List all cached transcriptions
    if (req.method === "GET" && pathname === "/api/ai/cache-list") {
      try {
        const rows = await runQuery(
          `SELECT message_id, account_id, audio_url, transcript, intent, intent_summary,
                  mentioned_phone, mentions_url, stt_model, intent_model, transcribed_at
           FROM message_transcriptions ORDER BY transcribed_at DESC`
        );
        return sendJson(res, 200, { ok: true, entries: rows || [] });
      } catch (e) {
        return sendJson(res, 200, { ok: true, entries: [] });
      }
    }

    // AI - Update a cached transcription entry
    if (req.method === "POST" && pathname === "/api/ai/cache-update") {
      try {
        const body = await readJson(req);
        const { message_id, account_id, transcript, intent, intent_summary } = body;
        if (!message_id || !account_id) return sendJson(res, 400, { ok: false, error: 'message_id and account_id required' });
        await runQuery(
          `UPDATE message_transcriptions SET transcript = ?, intent = ?, intent_summary = ? WHERE message_id = ? AND account_id = ?`,
          [transcript ?? '', intent ?? '', intent_summary ?? '', message_id, account_id]
        );
        console.log(`[AI] Cache entry updated: ${message_id} / ${account_id}`);
        return sendJson(res, 200, { ok: true });
      } catch (e) {
        return sendJson(res, 500, { ok: false, error: e.message });
      }
    }

    // AI - Delete a single cached transcription entry
    if (req.method === "POST" && pathname === "/api/ai/cache-delete-entry") {
      try {
        const body = await readJson(req);
        const { message_id, account_id } = body;
        if (!message_id || !account_id) return sendJson(res, 400, { ok: false, error: 'message_id and account_id required' });
        await runQuery(
          `DELETE FROM message_transcriptions WHERE message_id = ? AND account_id = ?`,
          [message_id, account_id]
        );
        console.log(`[AI] Cache entry deleted: ${message_id} / ${account_id}`);
        return sendJson(res, 200, { ok: true });
      } catch (e) {
        return sendJson(res, 500, { ok: false, error: e.message });
      }
    }

    // ── STT Dictionary CRUD ──────────────────────────────────────────────────

    if (req.method === "GET" && pathname === "/api/ai/dict-list") {
      try {
        const rows = await runQuery('SELECT id, raw_text, corrected, note, created_at FROM stt_dictionary ORDER BY created_at');
        return sendJson(res, 200, { ok: true, entries: rows || [] });
      } catch (e) {
        return sendJson(res, 200, { ok: true, entries: [] });
      }
    }

    if (req.method === "POST" && pathname === "/api/ai/dict-add") {
      try {
        const body = await readJson(req);
        const { raw_text, corrected, note } = body;
        if (!raw_text || !corrected) return sendJson(res, 400, { ok: false, error: 'raw_text and corrected are required' });
        const id = Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
        await runQuery(
          `INSERT INTO stt_dictionary (id, raw_text, corrected, note) VALUES (?, ?, ?, ?)`,
          [id, raw_text.trim(), corrected.trim(), (note || '').trim()]
        );
        await loadSttCorrectionCache();
        return sendJson(res, 200, { ok: true, id });
      } catch (e) {
        return sendJson(res, 500, { ok: false, error: e.message });
      }
    }

    if (req.method === "POST" && pathname === "/api/ai/dict-update") {
      try {
        const body = await readJson(req);
        const { id, raw_text, corrected, note } = body;
        if (!id) return sendJson(res, 400, { ok: false, error: 'id is required' });
        await runQuery(
          `UPDATE stt_dictionary SET raw_text = ?, corrected = ?, note = ? WHERE id = ?`,
          [raw_text.trim(), corrected.trim(), (note || '').trim(), id]
        );
        await loadSttCorrectionCache();
        return sendJson(res, 200, { ok: true });
      } catch (e) {
        return sendJson(res, 500, { ok: false, error: e.message });
      }
    }

    if (req.method === "POST" && pathname === "/api/ai/dict-delete") {
      try {
        const body = await readJson(req);
        const { id } = body;
        if (!id) return sendJson(res, 400, { ok: false, error: 'id is required' });
        await runQuery('DELETE FROM stt_dictionary WHERE id = ?', [id]);
        await loadSttCorrectionCache();
        return sendJson(res, 200, { ok: true });
      } catch (e) {
        return sendJson(res, 500, { ok: false, error: e.message });
      }
    }

    if (req.method === "GET" && pathname === "/api/ai/dict-stats") {
      try {
        const rows = await runQuery('SELECT COUNT(*) as count FROM stt_dictionary');
        return sendJson(res, 200, { ok: true, count: Number(rows[0]?.count) || 0 });
      } catch (e) {
        return sendJson(res, 200, { ok: true, count: 0 });
      }
    }

    // Accounts endpoint
    if (req.method === "POST" && pathname === "/api/accounts") {
      try {
        const body = await readJson(req);
        console.log('[API] /api/accounts - Starting request, platform:', process.platform);
        const accounts = await fetchAllAccounts(body.api_key || "", null, "minimal", body.filter);
        console.log('[API] /api/accounts - Success, found', accounts?.length || 0, 'accounts');
        return sendJson(res, 200, { ok: true, accounts });
      } catch (e) {
        const errorDetail = `${e.message} | Code: ${e.code || 'none'}`;
        console.error('[API Error - /api/accounts]', errorDetail, e.stack);
        return sendJson(res, 500, { ok: false, error: e.message, code: e.code, platform: process.platform });
      }
    }

    // Cancel job endpoint
    if (req.method === "POST" && pathname === "/api/cancel") {
      try {
        const body = await readJson(req);
        const { job_id } = body;
        if (job_id && jobs.has(job_id)) {
          jobs.get(job_id).cancelled = true;
          return sendJson(res, 200, { ok: true, message: "Job cancelled" });
        }
        return sendJson(res, 404, { ok: false, error: "Job not found" });
      } catch (e) {
        return sendJson(res, 500, { ok: false, error: e.message });
      }
    }

    // SSE stream endpoint for real-time progress
    if (req.method === "GET" && pathname.startsWith("/api/stream/")) {
      const jobId = pathname.substring("/api/stream/".length);
      const job = jobs.get(jobId);
      
      if (!job) {
        return sendJson(res, 404, { ok: false, error: "Job not found" });
      }

      // Set SSE headers
      res.writeHead(200, {
        'Content-Type': 'text/event-stream',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Access-Control-Allow-Origin': '*'
      });

      // Send initial connection event
      res.write(`data: ${JSON.stringify({ type: 'connected', jobId })}\n\n`);

      // Send accumulated logs FIRST — client must see all log lines before receiving
      // a 'complete' or 'error' status that would cause it to close the stream
      if (job.logs && job.logs.length > 0) {
        for (const logEntry of job.logs) {
          const data = {
            type: 'log',
            jobId,
            message: logEntry.message,
            isError: logEntry.isError
          };
          res.write(`data: ${JSON.stringify(data)}\n\n`);
        }
      }

      // Send current job state AFTER logs so the client closes only after it has seen everything
      if (job.progress > 0 || job.status !== 'starting') {
        const data = {
          type: 'progress',
          jobId,
          progress: job.progress,
          status: job.status,
          current: job.current,
          total: job.total,
          message: job.message
        };
        res.write(`data: ${JSON.stringify(data)}\n\n`);
      }

      // Store response object in job for live updates
      job.stream = res;

      // Clean up on disconnect
      req.on('close', () => {
        if (job.stream === res) {
          job.stream = null;
        }
      });

      return; // Keep connection open
    }

    // Database stats endpoint
    if (req.method === "GET" && pathname === "/api/database/stats") {
      try {
        const stats = await getDatabaseStats();
        return sendJson(res, 200, { ok: true, stats });
      } catch (e) {
        console.error('[API Error - /api/database/stats]', e.message, e.stack);
        return sendJson(res, 500, { ok: false, error: e.message });
      }
    }

    // Database clear endpoint
    if (req.method === "POST" && pathname === "/api/database/clear") {
      try {
        const body = await readJson(req);
        const createBackup = body.createBackup !== false;  // Default to true
        const result = await clearDatabase(createBackup);
        if (result.success) {
          return sendJson(res, 200, {
            ok: true,
            message: result.message,
            backupPath: result.backupPath,
            rowCount: result.rowCount
          });
        } else {
          return sendJson(res, 500, { ok: false, error: result.error });
        }
      } catch (e) {
        console.error('[API Error - /api/database/clear]', e.message, e.stack);
        return sendJson(res, 500, { ok: false, error: e.message });
      }
    }

    // Database compact endpoint
    if (req.method === "POST" && pathname === "/api/database/compact") {
      try {
        const result = await compactDatabase();
        if (result.success) {
          return sendJson(res, 200, { ok: true, message: result.message });
        } else {
          return sendJson(res, 500, { ok: false, error: result.error });
        }
      } catch (e) {
        console.error('[API Error - /api/database/compact]', e.message, e.stack);
        return sendJson(res, 500, { ok: false, error: e.message });
      }
    }

    // Database export endpoint
    if (req.method === "POST" && pathname === "/api/database/export") {
      try {
        const folders = createOutputFolders();
        const suffix = getFilenameSuffix(folders.combineCampaigns, 'DatabaseExport');
        const outputPath = path.join(folders.combineCampaigns, `DatabaseExport_${suffix}.csv`);
        
        const result = await exportDatabase(outputPath);
        
        if (result.success) {
          lastArtifacts.csvPath = result.path;
          return sendJson(res, 200, { 
            ok: true, 
            message: `Exported ${result.rowCount.toLocaleString()} rows`,
            path: result.path,
            rowCount: result.rowCount
          });
        } else {
          return sendJson(res, 500, { ok: false, error: result.error });
        }
      } catch (e) {
        console.error('[API Error - /api/database/export]', e.message, e.stack);
        return sendJson(res, 500, { ok: false, error: e.message });
      }
    }

    // Database query endpoint
    if (req.method === "POST" && pathname === "/api/database/query") {
      try {
        const body = await readJson(req);
        const { sql } = body;
        
        if (!sql) {
          return sendJson(res, 400, { ok: false, error: "SQL query required" });
        }
        
        // Only allow SELECT queries for safety
        if (!sql.trim().toUpperCase().startsWith('SELECT')) {
          return sendJson(res, 400, { ok: false, error: "Only SELECT queries allowed" });
        }
        
        const rows = await runQuery(sql);
        
        return sendJson(res, 200, { 
          ok: true, 
          rows,
          rowCount: rows.length
        });
      } catch (e) {
        console.error('[API Error - /api/database/query]', e.message, e.stack);
        return sendJson(res, 500, { ok: false, error: e.message });
      }
    }

    // Open database backups folder
    if (req.method === "GET" && pathname === "/api/database/backups-folder") {
      return sendJson(res, 200, { ok: true, path: DB_DIR });
    }

    // Search endpoint
    if (req.method === "POST" && pathname === "/api/search") {
      try {
        const body = await readJson(req);
        const out = await runNumberSearch({
          api_key: body.api_key || "",
          numbers: body.numbers || [],
          account_ids: body.account_ids || [],
          start_date: body.start_date || "",
          end_date: body.end_date || "",
          include_caller: !!(body.include_caller ?? true),
          include_message_meta: !!(body.include_message_meta ?? true),
          output_mode: body.output_mode || "csv",
          job_id: body.job_id || null,
          client_prefix: body.client_prefix || ""
        });

        return sendJson(res, 200, {
          ok: true,
          message: "Search complete",
          artifacts: { 
            csvPath: out.csvPath,
            allCsvFiles: out.allCsvFiles,
            logPath: out.logPath 
          },
          matches: out.matches,
          wasSplit: out.wasSplit,
          fileCount: out.fileCount,
          fromDatabase: out.fromDatabase || false
        });
      } catch (e) {
        console.error('[API Error - /api/search]', e.message, e.stack);
        const cancelled = e.message === "Cancelled";
        return sendJson(res, cancelled ? 499 : 500, { ok: false, error: e.message });
      }
    }

    // Combine endpoint
    if (req.method === "POST" && pathname === "/api/combine") {
      try {
        const body = await readJson(req);
        const out = await runCombineCampaigns({
          api_key: body.api_key || "",
          account_ids: body.account_ids || [],
          start_date: body.start_date || "",
          end_date: body.end_date || "",
          include_caller: !!(body.include_caller ?? true),
          include_message_meta: !!(body.include_message_meta ?? true),
          generate_trend_analysis: !!body.generate_trend_analysis,
          min_consec_unsuccessful: body.min_consec_unsuccessful,
          min_run_span_days: body.min_run_span_days,
          include_detail_tabs: !!body.include_detail_tabs,
          output_mode: body.output_mode || "csv",
          job_id: body.job_id || null,
          client_prefix: body.client_prefix || "",
          // AI settings come from the frontend payload (UI / localStorage).
          // The enable toggle and mode radios are only stored in localStorage,
          // so getAiSettings() (disk-based) would always return enabled:false.
          ai_enabled: body.ai_enabled === true,
          ai_transcription_mode: body.ai_transcription_mode || 'local',
          ai_intent_mode: body.ai_intent_mode || 'local'
        });

        const artifacts = { 
          csvPath: out.csvPath,
          allCsvFiles: out.allCsvFiles,
          logPath: out.logPath 
        };
        if (out.analysisPath) artifacts.analysisPath = out.analysisPath;

        return sendJson(res, 200, {
          ok: true,
          message: "Combine complete",
          artifacts,
          totalRows: out.totalRows,
          wasSplit: out.wasSplit,
          fileCount: out.fileCount
        });
      } catch (e) {
        console.error('[API Error - /api/combine]', e.message, e.stack);
        const cancelled = e.message === "Cancelled";
        const errLogPath = e.logPath || lastArtifacts.logPath || null;
        return sendJson(res, cancelled ? 499 : 500, { ok: false, error: e.message, artifacts: { logPath: errLogPath } });
      }
    }

    // Executive Summary endpoint
    if (req.method === "POST" && pathname === "/api/executive-summary") {
      try {
        const body = await readJson(req);
        const out = await generateExecutiveSummary({
          api_key: body.api_key || "",
          account_ids: body.account_ids || [],
          start_date: body.start_date || "",
          end_date: body.end_date || "",
          job_id: body.job_id || null
        });

        return sendJson(res, 200, {
          ok: true,
          message: "Executive Summary complete",
          artifacts: {
            csvPath: out.csvPath,
            logPath: out.logPath
          },
          stats: {
            campaignCount: out.campaignCount,
            totalRecords: out.totalRecords
          }
        });
      } catch (e) {
        console.error('[API Error - /api/executive-summary]', e.message, e.stack);
        const cancelled = e.message === "Cancelled";
        return sendJson(res, cancelled ? 499 : 500, { ok: false, error: e.message });
      }
    }

    // Analyze CSV endpoint
    if (req.method === "POST" && pathname === "/api/analyze-csv") {
      try {
        const boundary = req.headers['content-type']?.split('boundary=')[1];
        if (!boundary) throw new Error("No boundary found");

        const chunks = [];
        for await (const chunk of req) chunks.push(chunk);
        const bodyBuf = Buffer.concat(chunks);

        // Parse multipart on Buffers to avoid huge string allocations
        const boundaryBuf = Buffer.from(`--${boundary}`);
        const crlf = Buffer.from('\r\n');
        const crlfcrlf = Buffer.from('\r\n\r\n');

        const csvTexts = [];
        let minConsec = 4;
        let minSpan = 30;
        let csvApiKey = '';
        let csvAiEnabled = false;
        let csvAiTranscriptionMode = 'local';
        let csvAiIntentMode = 'local';

        let searchStart = 0;
        while (searchStart < bodyBuf.length) {
          const boundaryPos = bodyBuf.indexOf(boundaryBuf, searchStart);
          if (boundaryPos === -1) break;
          const headerStart = boundaryPos + boundaryBuf.length + crlf.length; // skip \r\n after boundary
          const headerEnd = bodyBuf.indexOf(crlfcrlf, headerStart);
          if (headerEnd === -1) break;
          const header = bodyBuf.slice(headerStart, headerEnd).toString();
          const contentStart = headerEnd + crlfcrlf.length;
          // Find next boundary to know where content ends
          const nextBoundary = bodyBuf.indexOf(boundaryBuf, contentStart);
          const contentEnd = nextBoundary === -1 ? bodyBuf.length : nextBoundary - crlf.length;
          searchStart = nextBoundary === -1 ? bodyBuf.length : nextBoundary;

          if (header.includes('name="csv"')) {
            const text = bodyBuf.slice(contentStart, contentEnd).toString('utf8');
            if (text.trim()) csvTexts.push(text);
          } else if (header.includes('name="min_consec_unsuccessful"')) {
            minConsec = parseInt(bodyBuf.slice(contentStart, contentEnd).toString()) || 4;
          } else if (header.includes('name="min_run_span_days"')) {
            minSpan = parseInt(bodyBuf.slice(contentStart, contentEnd).toString()) || 30;
          } else if (header.includes('name="api_key"')) {
            csvApiKey = bodyBuf.slice(contentStart, contentEnd).toString().trim();
          } else if (header.includes('name="ai_enabled"')) {
            csvAiEnabled = bodyBuf.slice(contentStart, contentEnd).toString().trim() === 'true';
          } else if (header.includes('name="ai_transcription_mode"')) {
            csvAiTranscriptionMode = bodyBuf.slice(contentStart, contentEnd).toString().trim() || 'local';
          } else if (header.includes('name="ai_intent_mode"')) {
            csvAiIntentMode = bodyBuf.slice(contentStart, contentEnd).toString().trim() || 'local';
          }
        }

        if (csvTexts.length === 0) throw new Error("No CSV data found");

        // Parse all CSVs and combine rows (use headers from first file)
        const parsedFiles = csvTexts.map(t => parseCsv(t));
        const headers = parsedFiles[0].headers;
        const allRows = parsedFiles.flatMap(pf => pf.rows);

        const folders = createOutputFolders();
        const outDir = folders.combineCampaigns;

        // Get user's timezone for report
        const userTz = getTimezone();
        const userTzLabel = getTimezoneLabel(userTz);

        const suffix = getFilenameSuffix(outDir, 'UploadedCSV');
        const analysisFilename = `NumberAnalysis_${suffix}.xlsx`;
        const analysisPath = path.join(outDir, analysisFilename);

        // ── AI Message Analysis (optional) ───────────────────────────────────
        // When AI is enabled and an API key is provided, fetch audio for any
        // messages in the dataset that aren't already in the transcription cache,
        // transcribe them, and pass the results to the analysis worker so the
        // report shows real intents instead of name-inferred ones.
        let csvTranscriptMap = {};
        if (csvAiEnabled && csvApiKey) {
          const csvLog = (msg) => console.log('[AI CSV]', msg);
          try {
            // Collect unique account_id:message_id pairs from all CSV rows.
            const uniqueMessageKeys = new Set();
            const uniqueAccountIds  = new Set();
            for (const row of allRows) {
              const aId = (row.account_id || '').trim();
              const mId = (row.message_id  || '').trim();
              if (aId && mId && mId !== 'Unknown' && mId !== '0') {
                uniqueMessageKeys.add(`${aId}:${mId}`);
                uniqueAccountIds.add(aId);
              }
            }
            csvLog(`Found ${uniqueMessageKeys.size} unique message(s) across ${uniqueAccountIds.size} account(s)`);

            if (uniqueMessageKeys.size > 0) {
              // Load already-cached transcriptions from DuckDB.
              const cachedKeys = new Set();
              if (dbReady) {
                const cached = await runQuery(
                  'SELECT message_id, account_id, transcript, intent, intent_summary, mentioned_phone, mentions_url FROM message_transcriptions'
                );
                for (const r of (cached || [])) {
                  const k = `${r.account_id}:${r.message_id}`;
                  cachedKeys.add(k);
                  if (uniqueMessageKeys.has(k)) {
                    csvTranscriptMap[k] = {
                      transcript: r.transcript || '',
                      intent: r.intent || '',
                      intent_summary: r.intent_summary || '',
                      mentioned_phone: r.mentioned_phone || '',
                      mentions_url: !!r.mentions_url
                    };
                  }
                }
              }

              const uncachedKeys = new Set([...uniqueMessageKeys].filter(k => !cachedKeys.has(k)));
              csvLog(`${cachedKeys.size} already cached, ${uncachedKeys.size} need transcription`);

              if (uncachedKeys.size > 0) {
                // Fetch fresh audio URLs from VoApps API for each account.
                const aiMessageInfo = {};
                for (const accountId of uniqueAccountIds) {
                  try {
                    const freshData = await retryableApiCall(
                      `/accounts/${accountId}/messages?filter=all`, csvApiKey, csvLog, 'normal'
                    );
                    if (Array.isArray(freshData?.messages)) {
                      for (const msg of freshData.messages) {
                        const k = `${accountId}:${msg.id}`;
                        if (uncachedKeys.has(k)) {
                          aiMessageInfo[k] = {
                            name: msg.name || '',
                            description: msg.description || '',
                            file_url: msg.file_url || msg.audio_url || msg.recording_url || msg.url || ''
                          };
                        }
                      }
                    }
                  } catch (urlErr) {
                    csvLog(`⚠️  Could not fetch messages for account ${accountId}: ${urlErr.message}`);
                  }
                }

                const toTranscribe = Object.keys(aiMessageInfo).length;
                csvLog(`Fetched audio URLs for ${toTranscribe} message(s) — transcribing...`);

                if (toTranscribe > 0) {
                  const aiSettings = {
                    enabled: true,
                    transcriptionMode: csvAiTranscriptionMode,
                    intentMode: csvAiIntentMode,
                    openaiApiKey: getAiSettings().openaiApiKey,
                  };
                  const newTranscripts = await transcribeAndAnalyzeMessages(aiMessageInfo, aiSettings, csvLog);
                  Object.assign(csvTranscriptMap, newTranscripts);
                }
              }
            }
          } catch (aiErr) {
            console.warn('[AI CSV] AI analysis failed (non-fatal):', aiErr.message);
          }
        }

        // Dynamic row limit: target ~50MB of string data per split file.
        // Estimate avg bytes per row based on column count (each col ~20 chars avg).
        const colCount = headers.length || 14;
        const avgBytesPerRow = colCount * 20;
        const targetBytesPerFile = 50 * 1024 * 1024; // 50MB
        const dynamicRowLimit = Math.max(50000, Math.min(MAX_ROWS_PER_FILE, Math.floor(targetBytesPerFile / avgBytesPerRow)));

        if (allRows.length > dynamicRowLimit) {
          const tempCsvPath = path.join(outDir, `UploadedCSV_${suffix}.csv`);
          const csvResult = await writeCsv(tempCsvPath, allRows, headers, null, dynamicRowLimit);

          await runAnalysisInWorker(csvResult.files, analysisPath, minConsec, minSpan, {}, {}, {}, userTz, userTzLabel, false, csvTranscriptMap);

          lastArtifacts.analysisPath = analysisPath;

          const fileWord = csvTexts.length > 1 ? `${csvTexts.length} files` : '1 file';
          return sendJson(res, 200, {
            ok: true,
            message: `Analysis complete (${allRows.length.toLocaleString()} rows from ${fileWord})`,
            artifacts: { analysisPath }
          });
        }

        await runAnalysisInWorker(allRows, analysisPath, minConsec, minSpan, {}, {}, {}, userTz, userTzLabel, false, csvTranscriptMap);

        lastArtifacts.analysisPath = analysisPath;

        const fileWord = csvTexts.length > 1 ? `${csvTexts.length} files` : '1 file';
        return sendJson(res, 200, {
          ok: true,
          message: `Analysis complete (${allRows.length.toLocaleString()} rows from ${fileWord})`,
          artifacts: { analysisPath }
        });
      } catch (e) {
        console.error('[API Error - /api/analyze-csv]', e.message, e.stack);
        return sendJson(res, 500, { ok: false, error: e.message });
      }
    }

    // Database analysis endpoint
    if (req.method === "POST" && pathname === "/api/analyze-database") {
      try {
        const body = await readJson(req);
        const {
          start_date,
          end_date,
          min_consec_unsuccessful = 4,
          min_run_span_days = 30,
          client_prefix = "",
          include_detail_tabs = false,
          api_key: dbApiKey = '',
          ai_enabled: dbAiEnabled = false,
          ai_transcription_mode: dbAiTranscriptionMode = 'local',
          ai_intent_mode: dbAiIntentMode = 'local'
        } = body;

        if (!dbReady) {
          throw new Error("Database not ready");
        }

        const folders = createOutputFolders();
        const suffix = getFilenameSuffix(folders.logs, 'db_analysis');
        const logPath = path.join(folders.logs, `db_analysis_log_${suffix}.txt`);
        const errorPath = path.join(folders.logs, `db_analysis_errors_${suffix}.txt`);
        const { log, close } = createLogger(logPath, errorPath, "normal", null);

        log(`📊 Delivery Intelligence Report — Database`);
        log(`Date Range: ${start_date} to ${end_date}`);
        log(`Thresholds: min_consec=${min_consec_unsuccessful}, min_span=${min_run_span_days} days`);

        // Stream rows from DB into split CSV temp files to avoid loading 1M+ rows into RAM
        const filePrefix = client_prefix ? `${client_prefix}_` : "";
        const analysisFilename = `${filePrefix}db_analysis_${suffix}.xlsx`;
        const analysisPath = path.join(folders.combineCampaigns, analysisFilename);

        const csvHeaders = [
          'number', 'account_id', 'account_name', 'campaign_id', 'campaign_name',
          'caller_number', 'caller_number_name', 'message_id', 'message_name', 'message_description',
          'voapps_result', 'voapps_code', 'voapps_timestamp', 'campaign_url'
        ];

        const escapeCsvVal = (val) => {
          if (val === null || val === undefined) return '';
          const str = String(val);
          return str.includes(',') || str.includes('"') || str.includes('\n')
            ? `"${str.replace(/"/g, '""')}"` : str;
        };

        const tempCsvFiles = [];
        let totalRows = 0;

        // Count rows first using existing runQuery helper
        const countResult = await runQuery(
          `SELECT COUNT(*) as cnt FROM campaign_results WHERE target_date >= '${start_date}' AND target_date <= '${end_date}'`
        );
        const expectedRows = Number(countResult[0]?.cnt || 0);

        if (expectedRows === 0) {
          close();
          throw new Error(`No data found in database for ${start_date} to ${end_date}`);
        }

        log(`\n💾 Streaming ${expectedRows.toLocaleString()} rows from database...`);

        // Stream in batches using LIMIT/OFFSET — avoids holding everything in RAM
        // and keeps the event loop alive between batches.
        const BATCH_SIZE = 50000;
        let offset = 0;
        let currentStream = null;
        let currentPath = null;
        let currentFileIndex = 1;
        let currentRowCount = 0;

        const openNewCsvFile = () => {
          if (currentStream) currentStream.end();
          currentPath = path.join(folders.combineCampaigns,
            `${filePrefix}db_analysis_temp_${suffix}_part${currentFileIndex}.csv`);
          currentStream = fs.createWriteStream(currentPath, { encoding: 'utf8' });
          currentStream.write(csvHeaders.join(',') + '\n');
          tempCsvFiles.push(currentPath);
          currentFileIndex++;
          currentRowCount = 0;
        };

        openNewCsvFile();

        while (offset < expectedRows) {
          const batchRows = await runQuery(`
            SELECT number, account_id, account_name, campaign_id, campaign_name,
              caller_number, caller_number_name, message_id, message_name, message_description,
              voapps_result, voapps_code, voapps_timestamp, campaign_url
            FROM campaign_results
            WHERE target_date >= '${start_date}' AND target_date <= '${end_date}'
            LIMIT ${BATCH_SIZE} OFFSET ${offset}
          `);

          if (batchRows.length === 0) break;

          for (const row of batchRows) {
            const line = csvHeaders.map(h => escapeCsvVal(row[h])).join(',') + '\n';
            currentStream.write(line);
            totalRows++;
            currentRowCount++;
            if (currentRowCount >= MAX_ROWS_PER_FILE) openNewCsvFile();
          }

          offset += batchRows.length;
          const pct = Math.round((offset / expectedRows) * 100);
          log(`  Streamed ${offset.toLocaleString()} / ${expectedRows.toLocaleString()} rows (${pct}%)`);

          // Yield to event loop between batches
          await new Promise(r => setTimeout(r, 0));
        }

        if (currentStream) currentStream.end();
        log(`✅ Streamed ${totalRows.toLocaleString()} rows into ${tempCsvFiles.length} temp file(s)`);

        // Get user's timezone for report
        const userTz = getTimezone();
        const userTzLabel = getTimezoneLabel(userTz);

        log(`\n📊 Generating Delivery Intelligence Report...`);
        log(`Output: ${analysisFilename}`);

        // ── AI Message Analysis (optional) ───────────────────────────────────
        // Load cached transcripts from DuckDB for any messages in this date
        // range.  If AI is enabled and an API key was provided, also fetch and
        // transcribe any messages that aren't yet cached.
        let dbTranscriptMap = {};
        try {
          const dbLog = (msg) => { log(`[AI] ${msg}`); console.log('[AI DB]', msg); };

          // Find every distinct message used in the date range
          const usedMsgs = await runQuery(`
            SELECT DISTINCT account_id, message_id, message_name
            FROM campaign_results
            WHERE target_date >= '${start_date}' AND target_date <= '${end_date}'
              AND message_id IS NOT NULL AND message_id != '' AND message_id != '0' AND message_id != 'Unknown'
          `);

          const uniqueMessageKeys = new Set((usedMsgs || []).map(r => `${r.account_id}:${r.message_id}`));
          const uniqueAccountIds  = new Set((usedMsgs || []).map(r => String(r.account_id)));

          if (uniqueMessageKeys.size > 0) {
            // Load already-cached transcriptions
            const cachedRows = await runQuery(
              'SELECT message_id, account_id, transcript, intent, intent_summary, mentioned_phone, mentions_url FROM message_transcriptions'
            );
            const cachedKeys = new Set();
            for (const r of (cachedRows || [])) {
              const k = `${r.account_id}:${r.message_id}`;
              cachedKeys.add(k);
              if (uniqueMessageKeys.has(k)) {
                dbTranscriptMap[k] = {
                  transcript: r.transcript || '',
                  intent: r.intent || '',
                  intent_summary: r.intent_summary || '',
                  mentioned_phone: r.mentioned_phone || '',
                  mentions_url: !!r.mentions_url
                };
              }
            }

            const uncachedKeys = new Set([...uniqueMessageKeys].filter(k => !cachedKeys.has(k)));
            dbLog(`${Object.keys(dbTranscriptMap).length} cached, ${uncachedKeys.size} uncached of ${uniqueMessageKeys.size} message(s)`);

            // Fetch + transcribe uncached messages if AI is on and key is set
            if (dbAiEnabled && dbApiKey && uncachedKeys.size > 0) {
              const aiMessageInfo = {};
              // Build name map from the DB query so we can pass names to the classifier
              const nameMap = {};
              for (const r of (usedMsgs || [])) nameMap[`${r.account_id}:${r.message_id}`] = r.message_name || '';

              for (const accountId of uniqueAccountIds) {
                try {
                  const freshData = await retryableApiCall(
                    `/accounts/${accountId}/messages?filter=all`, dbApiKey, dbLog, 'normal'
                  );
                  if (Array.isArray(freshData?.messages)) {
                    for (const msg of freshData.messages) {
                      const k = `${accountId}:${msg.id}`;
                      if (uncachedKeys.has(k)) {
                        aiMessageInfo[k] = {
                          name: msg.name || nameMap[k] || '',
                          description: msg.description || '',
                          file_url: msg.file_url || msg.audio_url || msg.recording_url || msg.url || ''
                        };
                      }
                    }
                  }
                } catch (urlErr) {
                  dbLog(`⚠️  Could not fetch messages for account ${accountId}: ${urlErr.message}`);
                }
              }

              const toTranscribe = Object.keys(aiMessageInfo).length;
              if (toTranscribe > 0) {
                dbLog(`Fetched audio URLs for ${toTranscribe} message(s) — transcribing...`);
                const aiSettings = {
                  enabled: true,
                  transcriptionMode: dbAiTranscriptionMode,
                  intentMode: dbAiIntentMode,
                  openaiApiKey: getAiSettings().openaiApiKey,
                };
                const newTranscripts = await transcribeAndAnalyzeMessages(aiMessageInfo, aiSettings, dbLog);
                Object.assign(dbTranscriptMap, newTranscripts);
              } else {
                dbLog('No audio URLs found for uncached messages — skipping transcription');
              }
            } else if (uncachedKeys.size > 0 && !dbAiEnabled) {
              dbLog(`${uncachedKeys.size} message(s) not yet transcribed — enable AI to transcribe them`);
            }
          }
        } catch (aiErr) {
          console.warn('[AI DB] AI analysis failed (non-fatal):', aiErr.message);
        }

        await runAnalysisInWorker(
          tempCsvFiles,
          analysisPath,
          min_consec_unsuccessful,
          min_run_span_days,
          {},  // messageMap
          {},  // callerMap
          {},  // accountTimezones
          userTz,
          userTzLabel,
          include_detail_tabs,
          dbTranscriptMap
        );

        // Clean up temp CSV files
        for (const f of tempCsvFiles) {
          try { fs.unlinkSync(f); } catch (_) {}
        }

        lastArtifacts.analysisPath = analysisPath;
        lastArtifacts.logPath = logPath;

        log(`\n✅ Complete! ${totalRows.toLocaleString()} rows analyzed.`);
        close();

        return sendJson(res, 200, {
          ok: true,
          message: `Database analysis complete (${totalRows.toLocaleString()} rows)`,
          rowCount: totalRows,
          artifacts: { analysisPath, logPath }
        });
      } catch (e) {
        console.error('[API Error - /api/analyze-database]', e.message, e.stack);
        // Attempt to close the log and return its path so the UI can still open it
        try { if (typeof close === 'function') { log(`\n❌ Error: ${e.message}`); close(); } } catch (_) {}
        const errLogPath = typeof logPath !== 'undefined' ? logPath : null;
        if (errLogPath) lastArtifacts.logPath = errLogPath;
        return sendJson(res, 500, { ok: false, error: e.message, artifacts: { logPath: errLogPath } });
      }
    }

    // Bulk export endpoint
    if (req.method === "POST" && pathname === "/api/bulk-export") {
      try {
        const body = await readJson(req);
        const out = await runBulkCampaignExport({
          api_key: body.api_key || "",
          account_ids: body.account_ids || [],
          start_date: body.start_date || "",
          end_date: body.end_date || "",
          job_id: body.job_id || null
        });

        return sendJson(res, 200, {
          ok: true,
          message: "Bulk export complete",
          artifacts: { 
            archivePath: out.archivePath, 
            logPath: out.logPath,
            errorPath: out.errorPath
          },
          stats: out.stats
        });
      } catch (e) {
        console.error('[API Error - /api/bulk-export]', e.message, e.stack);
        const cancelled = e.message === "Cancelled";
        return sendJson(res, cancelled ? 499 : 500, { ok: false, error: e.message });
      }
    }

    // Static file serving
    if (req.method === "GET") {
      const rel = pathname === "/" ? "/index.html" : pathname;
      const file = safeJoinPublic(rel);
      if (!file) {
        res.writeHead(404);
        return res.end("Not found");
      }
      try {
        const buf = await fsp.readFile(file);
        res.writeHead(200, { "Content-Type": guessContentType(file), "Cache-Control": "no-store" });
        return res.end(buf);
      } catch {
        res.writeHead(404);
        return res.end("Not found");
      }
    }

    res.writeHead(405);
    res.end("Method Not Allowed");
    } catch (err) {
      console.error('[Server Error - Uncaught]', err.message, err.stack);
      if (!res.headersSent) {
        res.writeHead(500);
        res.end(JSON.stringify({ ok: false, error: 'Internal server error' }));
      }
    }
  });
}

/**
 * Clean up empty log files and old log files on startup
 */
async function cleanupLogFiles() {
  const logDir = path.join(os.homedir(), 'Desktop', 'VoApps Reports');

  try {
    // Check if directory exists
    try {
      await fsp.access(logDir);
    } catch {
      return; // Directory doesn't exist, nothing to clean
    }

    const files = await fsp.readdir(logDir);
    const now = Date.now();
    const threeDaysAgo = now - (3 * 24 * 60 * 60 * 1000);

    for (const file of files) {
      // Only process log files
      if (!file.endsWith('.log') && !file.endsWith('_errors.log')) continue;

      const filePath = path.join(logDir, file);

      try {
        const stats = await fsp.stat(filePath);

        // Delete empty files (0 bytes)
        if (stats.size === 0) {
          await fsp.unlink(filePath);
          console.log(`[Cleanup] Deleted empty log file: ${file}`);
          continue;
        }

        // Delete files older than 3 days
        if (stats.mtime.getTime() < threeDaysAgo) {
          await fsp.unlink(filePath);
          console.log(`[Cleanup] Deleted old log file: ${file}`);
        }
      } catch (err) {
        // Skip files that can't be accessed
      }
    }
  } catch (err) {
    console.log('[Cleanup] Log cleanup skipped:', err.message);
  }
}

async function startServer() {
  if (serverInstance && serverUrl) return { url: serverUrl, port: PORT };

  // Clean up old/empty log files on startup
  try {
    await cleanupLogFiles();
  } catch (err) {
    console.log('[Cleanup] Log cleanup failed:', err.message);
  }

  // Initialize database on startup
  try {
    await initDatabase();
    await loadSttCorrectionCache();
  } catch (err) {
    console.error('[DuckDB] Failed to initialize database:', err);
  }

  serverInstance = createHttpServer();
  // Allow long-running requests (DB analysis can take many minutes on large datasets)
  serverInstance.timeout = 0;          // No timeout on the server socket
  serverInstance.keepAliveTimeout = 0; // Don't drop idle connections

  await new Promise((resolve, reject) => {
    serverInstance.once("error", reject);
    serverInstance.listen(PORT, HOST, () => resolve());
  });

  serverUrl = `http://${HOST}:${PORT}`;
  console.log(`[VoApps Tools v${VERSION}] Server listening on ${serverUrl}`);

  return { url: serverUrl, port: PORT };
}

async function stopServer() {
  if (!serverInstance) return;
  
  // Close database connection
  if (db) {
    db.close();
    db = null;
    dbReady = false;
  }
  
  await new Promise((resolve) => serverInstance.close(() => resolve()));
  serverInstance = null;
  serverUrl = null;
}

module.exports = { startServer, stopServer, getLastArtifacts, getDatabaseStats };

if (require.main === module) {
  startServer().catch((e) => {
    console.error("[VoApps Tools] Failed to start:", e);
    process.exit(1);
  });
}