/**
 * VoApps Tools — Local Server (Electron)
 * Version: 2.4.0 - Trend Analyzer with configurable thresholds, analysis-only mode
 */

"use strict";

const http = require("http");
const fs = require("fs");
const fsp = require("fs/promises");
const path = require("path");
const os = require("os");
const { parse: parseUrl } = require("url");
const { generateTrendAnalysis } = require("./trendAnalyzer");

const PORT = process.env.PORT ? Number(process.env.PORT) : 3000;
const HOST = "127.0.0.1";
const VOAPPS_API_BASE = process.env.VOAPPS_API_BASE || "https://directdropvoicemail.voapps.com/api/v1";

let serverInstance = null;
let serverUrl = null;

const lastArtifacts = { csvPath: null, logPath: null, errorPath: null, analysisPath: null };
function getLastArtifacts() { return { ...lastArtifacts }; }

const jobs = new Map();

// Helpers
function sendJson(res, status, obj) {
  const body = JSON.stringify(obj);
  res.writeHead(status, { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" });
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

function normalizePhone(raw) {
  let s = String(raw || "").replace(/\D+/g, "");
  if (s.length === 11 && s.startsWith("1")) s = s.slice(1);
  return s.length === 10 ? s : "";
}

function normalizeNumbers(input) {
  const raw = Array.isArray(input) ? input.join("\n") : String(input || "");
  return raw.split(/[^0-9]+/).map(normalizePhone).filter(Boolean);
}

function ensureDir(p) {
  if (!fs.existsSync(p)) fs.mkdirSync(p, { recursive: true });
}

function csvEscape(v) {
  const s = v === null || v === undefined ? "" : String(v);
  return /[,"\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

async function writeCsv(filePath, rows, headers) {
  const lines = [headers.join(",")];
  for (const r of rows) {
    lines.push(headers.map((h) => csvEscape(r[h])).join(","));
  }
  await fsp.writeFile(filePath, lines.join("\n"), "utf8");
}

async function writeLog(filePath, lines) {
  await fsp.writeFile(filePath, lines.join("\n") + "\n", "utf8");
}

function authHeaders(apiKey) {
  return {
    Authorization: `Bearer ${apiKey}`,
    "Content-Type": "application/json",
    Accept: "application/json",
  };
}

/**
 * Generate curl command for debugging API requests
 */
function generateCurlCommand(url, options = {}, apiKey = null) {
  const method = options.method || 'GET';
  let curl = `curl -X ${method} '${url}'`;
  
  // Add headers
  if (apiKey) {
    const maskedKey = `${apiKey.substring(0, 4)}...${apiKey.substring(apiKey.length - 4)}`;
    curl += ` \\\n  -H 'Authorization: Bearer ${maskedKey}'`;
  }
  
  if (options.headers) {
    Object.entries(options.headers).forEach(([key, value]) => {
      if (key.toLowerCase() !== 'authorization') {
        curl += ` \\\n  -H '${key}: ${value}'`;
      }
    });
  }
  
  // Add body if present
  if (options.body) {
    curl += ` \\\n  -d '${options.body}'`;
  }
  
  return curl;
}

async function fetchJson(url, { method = "GET", headers = {}, body, signal } = {}) {
  const timeoutMs = 120000;
  const controller = new AbortController();
  const t = setTimeout(() => controller.abort(), timeoutMs);

  const signals = [controller.signal];
  if (signal) signals.push(signal);

  const anySignal = new AbortController();
  const onAbort = () => anySignal.abort();
  signals.forEach((s) => s && s.addEventListener("abort", onAbort, { once: true }));

  try {
    const resp = await fetch(url, { method, headers, body, signal: anySignal.signal });
    const text = await resp.text();
    let json = {};
    try { json = text ? JSON.parse(text) : {}; }
    catch { json = { raw: text }; }
    return { ok: resp.ok, status: resp.status, json, text };
  } finally {
    clearTimeout(t);
    signals.forEach((s) => s && s.removeEventListener("abort", onAbort));
  }
}

/**
 * v2.2.0: Fetch JSON with retry logic (3s, 10s, 60s delays)
 */
async function fetchJsonWithRetry(url, options = {}) {
  const delays = [3000, 10000, 60000];
  let lastError = null;

  for (let attempt = 0; attempt < delays.length + 1; attempt++) {
    try {
      const result = await fetchJson(url, options);
      
      if (result.ok) return result;
      
      if (attempt < delays.length) {
        lastError = new Error(`HTTP ${result.status}`);
        await new Promise(resolve => setTimeout(resolve, delays[attempt]));
        continue;
      }
      
      return result;
      
    } catch (error) {
      lastError = error;
      
      if (attempt >= delays.length) {
        return { ok: false, status: 0, json: {}, error: error.message };
      }
      
      await new Promise(resolve => setTimeout(resolve, delays[attempt]));
    }
  }
  
  return { ok: false, status: 0, json: {}, error: lastError?.message || 'Unknown error' };
}

function parseCsvLine(line) {
  const result = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    const next = line[i + 1];

    if (char === '"') {
      if (inQuotes && next === '"') {
        current += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === "," && !inQuotes) {
      result.push(current.trim());
      current = "";
    } else {
      current += char;
    }
  }
  result.push(current.trim());
  return result;
}

function parseCsv(csvText) {
  const lines = csvText.split(/\r?\n/).filter((l) => l.trim());
  if (!lines.length) return { headers: [], rows: [] };

  const headers = parseCsvLine(lines[0]);
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    const cols = parseCsvLine(lines[i]);
    const obj = {};
    headers.forEach((h, idx) => { obj[h] = cols[idx] || ""; });
    rows.push(obj);
  }
  return { headers, rows };
}

// ========== v2.2.0 PAGINATION LOGIC ==========

/**
 * Send log message via SSE if available
 */
function sendSSELog(job_id, message) {
  if (!job_id) return;
  const job = jobs.get(job_id);
  if (job && job.sseResponse) {
    try {
      const data = JSON.stringify({ type: 'log', message });
      job.sseResponse.write(`data: ${data}\n\n`);
    } catch (e) {
      // SSE connection may have closed
    }
  }
}

/**
 * Send progress update via SSE
 */
function sendSSEProgress(job_id, current, total, text) {
  if (!job_id) return;
  const job = jobs.get(job_id);
  if (job && job.sseResponse) {
    try {
      const data = JSON.stringify({ type: 'progress', current, total, text });
      job.sseResponse.write(`data: ${data}\n\n`);
    } catch (e) {
      // SSE connection may have closed
    }
  }
}

/**
 * Check if job is paused
 */
async function checkPaused(job_id) {
  if (!job_id) return false;
  const job = jobs.get(job_id);
  
  // Wait while paused
  while (job && job.paused) {
    await new Promise(resolve => setTimeout(resolve, 500));
  }
  
  return false;
}

/**
 * v2.2.0: Fetch campaigns with pagination support
 * Loops through pages until we get < 25 results (indicating the last page)
 */
async function voappsGetCampaignsForDateRange(apiKey, accountId, start, end, signal, log) {
  // Widen the search window: -30 days before start, +1 day after end (UTC buffer)
  const searchStart = new Date(start);
  searchStart.setDate(searchStart.getDate() - 30);
  
  const searchEnd = new Date(end);
  searchEnd.setDate(searchEnd.getDate() + 1);
  
  const searchStartYMD = dateToYMD(searchStart);
  const searchEndYMD = dateToYMD(searchEnd);
  const userStartYMD = dateToYMD(start);
  const userEndYMD = dateToYMD(end);
  
  log(`[v2.4.0] Fetching campaigns with pagination: ${searchStartYMD} to ${searchEndYMD} (widened for lead time)`);
  log(`[v2.4.0] Will filter results by target_date: ${userStartYMD} to ${userEndYMD}`);
  
  const allCampaigns = [];
  let page = 1;
  let hasMorePages = true;
  
  while (hasMorePages) {
    if (signal.aborted) break;
    
    const url = `${VOAPPS_API_BASE}/accounts/${accountId}/campaigns?created_date_start=${searchStartYMD}&created_date_end=${searchEndYMD}&page=${page}&filter=all`;
    
    const maskedKey = apiKey ? `${apiKey.substring(0, 4)}...${apiKey.substring(apiKey.length - 4)}` : '[none]';
    const curlCmd = generateCurlCommand(url, { headers: authHeaders(apiKey) }, apiKey);
    
    log(`  → Page ${page}: GET ${url}`);
    log(`  → Auth: Bearer ${maskedKey}`);
    log(`  → curl: ${curlCmd}`);
    
    const result = await fetchJsonWithRetry(url, { headers: authHeaders(apiKey), signal });
    
    if (!result.ok) {
      log(`  ← Page ${page}: FAILED (${result.status}) - ${result.json?.error || result.error || 'Unknown error'}`);
      break;
    }
    
    const campaigns = Array.isArray(result.json?.campaigns) ? result.json.campaigns : [];
    log(`  ← Page ${page}: OK (${result.status}) - ${campaigns.length} campaigns returned`);
    
    allCampaigns.push(...campaigns);
    
    // If we got fewer than 25 campaigns, this is the last page
    if (campaigns.length < 25) {
      hasMorePages = false;
      log(`[v2.4.0] ✓ Pagination complete: ${allCampaigns.length} total campaigns fetched`);
    } else {
      page++;
    }
  }
  
  // Filter by target_date to match user's requested range
  const filtered = allCampaigns.filter(c => {
    const targetDate = c?.target_date;
    if (!targetDate) return false;
    
    try {
      const target = new Date(targetDate);
      return target >= start && target <= end;
    } catch (e) {
      return false;
    }
  });
  
  log(`[v2.4.0] ✓ Filtered by target_date: ${filtered.length} campaigns match range ${userStartYMD} to ${userEndYMD}`);
  
  return filtered;
}

function processRecordWithCampaignFallbacks(row, campaign) {
  return {
    ...row,
    voapps_message_id: row.voapps_message_id || row.message_id || campaign?.message_id || "",
    voapps_caller_number: row.voapps_caller_number || campaign?.caller_number || ""
  };
}

/**
 * v2.2.0: Helper functions
 */
function getMonthName(date) {
  const months = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];
  return months[date.getMonth()];
}

function sanitizeFilename(name) {
  return String(name || "")
    .replace(/[^a-z0-9_\-]/gi, "_")
    .replace(/_+/g, "_")
    .substring(0, 100);
}

/**
 * v2.4.0: Get month name with leading zero
 */
function getMonthFolder(date) {
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const monthName = getMonthName(date);
  return `${month} - ${monthName}`;
}

/**
 * v2.4.0: Create standardized output folder structure
 */
function createOutputFolders() {
  const downloads = path.join(os.homedir(), "Downloads");
  const base = path.join(downloads, "VoApps Tools");
  const logs = path.join(base, "Logs");
  const output = path.join(base, "Output");
  const phoneHistory = path.join(output, "Phone Number History");
  const combineCampaigns = path.join(output, "Combine Campaigns");
  const bulkExport = path.join(output, "Bulk Campaign Export");
  
  ensureDir(base);
  ensureDir(logs);
  ensureDir(output);
  ensureDir(phoneHistory);
  ensureDir(combineCampaigns);
  ensureDir(bulkExport);
  
  return {
    base,
    logs,
    output,
    phoneHistory,
    combineCampaigns,
    bulkExport
  };
}

function formatStatistics(stats) {
  const lines = [];
  lines.push("═══════════════════════════════════════");
  lines.push("           STATISTICS REPORT           ");
  lines.push("═══════════════════════════════════════");
  lines.push("");
  lines.push(`Total Campaigns Found: ${stats.totalCampaigns}`);
  lines.push(`Successfully Downloaded: ${stats.successfulDownloads}`);
  lines.push(`Failed Downloads: ${stats.failedDownloads}`);
  lines.push(`Total Records: ${stats.totalRecords}`);
  lines.push(`Time Elapsed: ${stats.timeElapsed}`);
  lines.push("");
  
  if (stats.byAccount && stats.byAccount.length > 0) {
    lines.push("Breakdown by Account:");
    lines.push("─────────────────────────────────────");
    for (const acct of stats.byAccount) {
      lines.push(`  Account ${acct.id}:`);
      lines.push(`    Campaigns: ${acct.campaigns}`);
      lines.push(`    Downloaded: ${acct.downloaded}`);
      lines.push(`    Failed: ${acct.failed}`);
    }
  }
  
  lines.push("═══════════════════════════════════════");
  return lines;
}

// VoApps API (with retry logic)
async function voappsGetAccounts(apiKey, filter, signal) {
  let url = `${VOAPPS_API_BASE}/accounts`;
  if (filter && filter !== "all") url += `?filter=${encodeURIComponent(filter)}`;
  return fetchJsonWithRetry(url, { headers: authHeaders(apiKey), signal });
}

async function voappsGetMessages(apiKey, accountId, signal) {
  const url = `${VOAPPS_API_BASE}/accounts/${accountId}/messages?filter=all`;
  return fetchJsonWithRetry(url, { headers: authHeaders(apiKey), signal });
}

async function voappsGetCampaignDetail(apiKey, accountId, campaignId, signal) {
  const url = `${VOAPPS_API_BASE}/accounts/${accountId}/campaigns/${campaignId}`;
  return fetchJsonWithRetry(url, { headers: authHeaders(apiKey), signal });
}

// Number Search
async function runNumberSearch({ api_key, numbers, account_ids, start_date, end_date, include_caller, include_message_meta, job_id }) {
  const logLines = [];
  const log = (s) => {
    logLines.push(s);
    sendSSELog(job_id, s);
  };

  const selectedNumbers = normalizeNumbers(numbers);
  if (!api_key) throw new Error("Missing api_key");
  if (!selectedNumbers.length) throw new Error("No valid 10-digit numbers");

  const start = new Date(`${start_date}T00:00:00`);
  const end = new Date(`${end_date}T00:00:00`);
  if (isNaN(start.getTime()) || isNaN(end.getTime())) throw new Error("Invalid dates");
  if (start > end) throw new Error("Start date must be <= end date");

  const controller = new AbortController();
  const signal = controller.signal;
  if (job_id) {
    const existing = jobs.get(job_id);
    if (existing) {
      existing.controller = controller;
      existing.paused = false;
    } else {
      jobs.set(job_id, { controller, sseResponse: null, paused: false });
    }
  }

  log(`VoApps Tools — Number Search`);
  log(`Timestamp: ${new Date().toISOString()}`);
  log(`Numbers: ${selectedNumbers.join(", ")}`);
  log(`Range: ${start_date} → ${end_date}`);
  log(`Accounts: ${account_ids.join(", ")}`);
  log('');
  const acctResp = await voappsGetAccounts(api_key, "all", signal);
  if (!acctResp.ok) throw new Error(`Accounts failed (HTTP ${acctResp.status})`);

  const allAccounts = Array.isArray(acctResp.json?.accounts) ? acctResp.json.accounts : [];
  const acctIdSet = new Set(account_ids.map(String));
  const accounts = allAccounts.filter((a) => acctIdSet.has(String(a?.id)));

  const messagesByAccount = new Map();
  if (include_message_meta) {
    for (const a of accounts) {
      const aid = a?.id;
      if (!aid) continue;
      const mr = await voappsGetMessages(api_key, aid, signal);
      const msgs = Array.isArray(mr.json?.messages) ? mr.json.messages : [];
      const map = new Map();
      for (const m of msgs) {
        if (m && m.id != null) map.set(String(m.id), m);
      }
      messagesByAccount.set(String(aid), map);
    }
  }

  const results = [];
  const headers = [
    "number", "account_id", "campaign_id", "campaign_name",
    ...(include_caller ? ["caller_number"] : []),
    "message_id",
    ...(include_message_meta ? ["message_name", "message_description"] : []),
    "voapps_result", "voapps_code", "voapps_timestamp", "campaign_url"
  ];

  const numberSet = new Set(selectedNumbers);

  for (const a of accounts) {
    const aid = String(a?.id);
    log(`--- Account ${aid} ---`);

    const campaigns = await voappsGetCampaignsForDateRange(api_key, aid, start, end, signal, log);

    let campaignIndex = 0;
    for (const c of campaigns) {
      if (signal.aborted) throw new Error("Cancelled");
      await checkPaused(job_id);

      campaignIndex++;
      sendSSEProgress(job_id, campaignIndex, campaigns.length, `Account ${aid}: Campaign ${campaignIndex}/${campaigns.length}`);

      const cid = String(c?.id);
      const cname = c?.name || "";
      if (!cid) continue;

      const detail = await voappsGetCampaignDetail(api_key, aid, cid, signal);
      const exportUrl = detail.json?.export || detail.json?.campaign?.export || null;

      if (!exportUrl) continue;

      try {
        const expResp = await fetch(exportUrl, { signal });
        const csvText = await expResp.text();
        const { rows } = parseCsv(csvText);

        for (const row of rows) {
          const num = normalizePhone(row.number || row.phone_number || "");
          if (!num || !numberSet.has(num)) continue;

          const processedRow = processRecordWithCampaignFallbacks(row, c);
          
          const messageId = String(processedRow.voapps_message_id || "");
          const msgMap = messagesByAccount.get(aid);
          const msg = include_message_meta && messageId && msgMap ? msgMap.get(messageId) : null;

          results.push({
            number: num,
            account_id: aid,
            campaign_id: cid,
            campaign_name: cname,
            ...(include_caller ? { caller_number: processedRow.voapps_caller_number } : {}),
            message_id: messageId,
            ...(include_message_meta ? {
              message_name: msg?.name || "",
              message_description: msg?.description || ""
            } : {}),
            voapps_result: row.voapps_result || "",
            voapps_code: row.voapps_code || "",
            voapps_timestamp: row.voapps_timestamp || "",
            campaign_url: `https://directdropvoicemail.voapps.com/accounts/${aid}/campaigns/${cid}`
          });
        }
      } catch (e) {
        log(`Campaign ${cid}: ${e.message}`);
      }
    }
  }

  const downloads = path.join(os.homedir(), "Downloads");
  const folders = createOutputFolders();
  const outDir = folders.phoneHistory;
  const logDir = folders.logs;

  const stamp = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
  const csvPath = path.join(outDir, `NumberSearch_${stamp}.csv`);
  const logPath = path.join(logDir, `NumberSearch_${stamp}.log`);

  await writeCsv(csvPath, results, headers);
  await writeLog(logPath, logLines);

  lastArtifacts.csvPath = csvPath;
  lastArtifacts.logPath = logPath;
  lastArtifacts.errorPath = null;

  // Send completion message via SSE
  if (job_id) {
    const job = jobs.get(job_id);
    if (job && job.sseResponse) {
      try {
        job.sseResponse.write(`data: ${JSON.stringify({ type: 'complete' })}\n\n`);
        job.sseResponse.end();
      } catch (e) {}
    }
    jobs.delete(job_id);
  }

  return { csvPath, logPath, matches: results.length };
}

// Combine Campaigns
async function runCombineCampaigns({ api_key, account_ids, start_date, end_date, include_caller, include_message_meta, generate_trend_analysis, min_consec_unsuccessful, min_run_span_days, job_id }) {
  const logLines = [];
  const log = (s) => {
    logLines.push(s);
    sendSSELog(job_id, s);
  };

  if (!api_key) throw new Error("Missing api_key");

  const start = new Date(`${start_date}T00:00:00`);
  const end = new Date(`${end_date}T00:00:00`);
  if (isNaN(start.getTime()) || isNaN(end.getTime())) throw new Error("Invalid dates");
  if (start > end) throw new Error("Start must be <= end");

  const controller = new AbortController();
  const signal = controller.signal;
  if (job_id) {
    const existing = jobs.get(job_id);
    if (existing) {
      existing.controller = controller;
      existing.paused = false;
    } else {
      jobs.set(job_id, { controller, sseResponse: null, paused: false });
    }
  }

  log(`VoApps Tools — Combine Campaigns`);
  log(`Timestamp: ${new Date().toISOString()}`);
  log(`Range: ${start_date} → ${end_date}`);
  log(`Accounts: ${account_ids.join(", ")}`);
  log('');
  const acctResp = await voappsGetAccounts(api_key, "all", signal);
  if (!acctResp.ok) throw new Error(`Accounts failed (HTTP ${acctResp.status})`);

  const allAccounts = Array.isArray(acctResp.json?.accounts) ? acctResp.json.accounts : [];
  const acctIdSet = new Set(account_ids.map(String));
  const accounts = allAccounts.filter((a) => acctIdSet.has(String(a?.id)));

  const messagesByAccount = new Map();
  if (include_message_meta) {
    for (const a of accounts) {
      const aid = a?.id;
      if (!aid) continue;
      const mr = await voappsGetMessages(api_key, aid, signal);
      const msgs = Array.isArray(mr.json?.messages) ? mr.json.messages : [];
      const map = new Map();
      for (const m of msgs) {
        if (m && m.id != null) map.set(String(m.id), m);
      }
      messagesByAccount.set(String(aid), map);
    }
  }

  const allRows = [];
  const headers = [
    "number", "account_id", "campaign_id", "campaign_name",
    ...(include_caller ? ["caller_number"] : []),
    "message_id",
    ...(include_message_meta ? ["message_name", "message_description"] : []),
    "voapps_result", "voapps_code", "voapps_timestamp", "campaign_url"
  ];

  for (const a of accounts) {
    const aid = String(a?.id);
    log(`--- Account ${aid} ---`);

    const campaigns = await voappsGetCampaignsForDateRange(api_key, aid, start, end, signal, log);

    let campaignIndex = 0;
    for (const c of campaigns) {
      if (signal.aborted) throw new Error("Cancelled");
      await checkPaused(job_id);

      campaignIndex++;
      sendSSEProgress(job_id, campaignIndex, campaigns.length, `Account ${aid}: Campaign ${campaignIndex}/${campaigns.length}`);

      const cid = String(c?.id);
      const cname = c?.name || "";
      if (!cid) continue;

      const detail = await voappsGetCampaignDetail(api_key, aid, cid, signal);
      const exportUrl = detail.json?.export || detail.json?.campaign?.export || null;

      if (!exportUrl) continue;

      try {
        const expResp = await fetch(exportUrl, { signal });
        const csvText = await expResp.text();
        const { rows } = parseCsv(csvText);

        for (const row of rows) {
          const num = normalizePhone(row.number || row.phone_number || "");
          if (!num) continue;

          const processedRow = processRecordWithCampaignFallbacks(row, c);
          
          const messageId = String(processedRow.voapps_message_id || "");
          const msgMap = messagesByAccount.get(aid);
          const msg = include_message_meta && messageId && msgMap ? msgMap.get(messageId) : null;

          allRows.push({
            number: num,
            account_id: aid,
            campaign_id: cid,
            campaign_name: cname,
            ...(include_caller ? { caller_number: processedRow.voapps_caller_number } : {}),
            message_id: messageId,
            ...(include_message_meta ? {
              message_name: msg?.name || "",
              message_description: msg?.description || ""
            } : {}),
            voapps_result: row.voapps_result || "",
            voapps_code: row.voapps_code || "",
            voapps_timestamp: row.voapps_timestamp || "",
            campaign_url: `https://directdropvoicemail.voapps.com/accounts/${aid}/campaigns/${cid}`
          });
        }

        log(`Campaign ${cid}: ${rows.length} rows`);
      } catch (e) {
        log(`Campaign ${cid}: ${e.message}`);
      }
    }
  }

  const downloads = path.join(os.homedir(), "Downloads");
  const folders = createOutputFolders();
  const outDir = folders.combineCampaigns;
  const logDir = folders.logs;

  const stamp = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
  const csvPath = path.join(outDir, `CombinedCampaigns_${stamp}.csv`);
  const logPath = path.join(logDir, `CombinedCampaigns_${stamp}.log`);

  await writeCsv(csvPath, allRows, headers);
  await writeLog(logPath, logLines);

  let analysisPath = null;
  if (generate_trend_analysis) {
    try {
      log('');
      log('Generating Number Trend Analysis...');
      const analysisFilename = `NumberAnalysis_${stamp}.xlsx`;
      analysisPath = path.join(outDir, analysisFilename);
      
      const minConsec = min_consec_unsuccessful || 4;
      const minSpan = min_run_span_days || 30;
      
      await generateTrendAnalysis(allRows, analysisPath, minConsec, minSpan);
      log(`Analysis saved: ${analysisPath}`);
      
      lastArtifacts.analysisPath = analysisPath;
    } catch (analysisError) {
      log(`Analysis generation failed: ${analysisError.message}`);
      // Don't fail the whole job if analysis fails
    }
  }

  lastArtifacts.csvPath = csvPath;
  lastArtifacts.logPath = logPath;
  lastArtifacts.errorPath = null;

  // Send completion message via SSE
  if (job_id) {
    const job = jobs.get(job_id);
    if (job && job.sseResponse) {
      try {
        job.sseResponse.write(`data: ${JSON.stringify({ type: 'complete' })}\n\n`);
        job.sseResponse.end();
      } catch (e) {}
    }
    jobs.delete(job_id);
  }

  const result = { csvPath, logPath, totalRows: allRows.length };
  if (analysisPath) result.analysisPath = analysisPath;
  
  return result;
}

// v2.2.0: Bulk Campaign Export
async function runBulkCampaignExport({ api_key, account_ids, start_date, end_date, job_id }) {
  const startTime = Date.now();
  const logLines = [];
  const errorLines = [];
  const log = (s) => {
    logLines.push(s);
    sendSSELog(job_id, s);
  };
  const logError = (s) => errorLines.push(s);

  if (!api_key) throw new Error("Missing api_key");

  const start = new Date(`${start_date}T00:00:00`);
  const end = new Date(`${end_date}T00:00:00`);
  if (isNaN(start.getTime()) || isNaN(end.getTime())) throw new Error("Invalid dates");
  if (start > end) throw new Error("Start must be <= end");

  const controller = new AbortController();
  const signal = controller.signal;
  if (job_id) {
    const existing = jobs.get(job_id);
    if (existing) {
      existing.controller = controller;
      existing.paused = false;
    } else {
      jobs.set(job_id, { controller, sseResponse: null, paused: false });
    }
  }

  log(`VoApps Tools — Bulk Campaign Export`);
  log(`Timestamp: ${new Date().toISOString()}`);
  log(`Range: ${start_date} → ${end_date}`);
  log(`Accounts: ${account_ids.join(", ")}`);
  log("");
  const stats = {
    totalCampaigns: 0,
    successfulDownloads: 0,
    failedDownloads: 0,
    totalRecords: 0,
    byAccount: []
  };

  const acctResp = await voappsGetAccounts(api_key, "all", signal);
  if (!acctResp.ok) throw new Error(`Accounts failed (HTTP ${acctResp.status})`);

  const allAccounts = Array.isArray(acctResp.json?.accounts) ? acctResp.json.accounts : [];
  const acctIdSet = new Set(account_ids.map(String));
  const accounts = allAccounts.filter((a) => acctIdSet.has(String(a?.id)));

  const downloads = path.join(os.homedir(), "Downloads");
  const folders = createOutputFolders();
  const baseDir = folders.bulkExport;
  const logDir = folders.logs;

  for (const a of accounts) {
    const aid = String(a?.id);
    log(`=== Account ${aid} ===`);

    const acctStats = { id: aid, campaigns: 0, downloaded: 0, failed: 0 };

    const campaigns = await voappsGetCampaignsForDateRange(api_key, aid, start, end, signal, log);
    
    stats.totalCampaigns += campaigns.length;
    acctStats.campaigns = campaigns.length;

    let campaignIndex = 0;
    for (const c of campaigns) {
      if (signal.aborted) throw new Error("Cancelled");
      await checkPaused(job_id);

      campaignIndex++;
      sendSSEProgress(job_id, stats.successfulDownloads + stats.failedDownloads, stats.totalCampaigns, `Account ${aid}: ${campaignIndex}/${campaigns.length}`);

      const cid = String(c?.id);
      const cname = c?.name || "Unnamed";

      // Use target_date from the campaign
      const createdDate = c?.target_date || c?.created_date || c?.created_at || "";

      if (!cid) continue;

      let year = "Unknown";
      let monthFolder = "Unknown";

      // Try to parse the date from various sources
      if (createdDate) {
        try {
          const date = new Date(createdDate);
          if (!isNaN(date.getTime())) {
            year = String(date.getFullYear());
            monthFolder = getMonthFolder(date);
          }
        } catch (e) {
          log(`  Warning: Could not parse date for campaign ${cid}: ${createdDate}`);
        }
      }

      // Fallback: If still Unknown, log what the campaign object looks like
      if (year === "Unknown" || monthFolder === "Unknown") {
        log(`  Campaign ${cid} date fields: ${JSON.stringify({
          created_date: c?.created_date,
          created_at: c?.created_at,
          date_created: c?.date_created,
          timestamp: c?.timestamp
        })}`);
      }

      const yearDir = path.join(baseDir, year);
      const monthDir = path.join(yearDir, monthFolder);
      ensureDir(yearDir);
      ensureDir(monthDir);

      try {
        const detail = await voappsGetCampaignDetail(api_key, aid, cid, signal);
        const exportUrl = detail.json?.export || detail.json?.campaign?.export || null;

        if (!exportUrl) {
          logError(`Campaign ${cid} (${cname}): No export URL available`);
          acctStats.failed++;
          stats.failedDownloads++;
          continue;
        }

        let csvText = null;
        let downloadError = null;
        
        for (let attempt = 0; attempt < 4; attempt++) {
          try {
            const expResp = await fetch(exportUrl, { signal });
            if (!expResp.ok) throw new Error(`HTTP ${expResp.status}`);
            csvText = await expResp.text();
            break;
          } catch (e) {
            downloadError = e;
            if (attempt < 3) {
              const delays = [3000, 10000, 60000];
              await new Promise(resolve => setTimeout(resolve, delays[attempt]));
            }
          }
        }

        if (!csvText) {
          logError(`Campaign ${cid} (${cname}): Failed to download after retries - ${downloadError?.message}`);
          acctStats.failed++;
          stats.failedDownloads++;
          continue;
        }

        const { rows } = parseCsv(csvText);
        stats.totalRecords += rows.length;

        const safeName = sanitizeFilename(cname);
        const filename = `${safeName}_${cid}.csv`;
        const filepath = path.join(monthDir, filename);

        await fsp.writeFile(filepath, csvText, "utf8");
        
        log(`  ✓ ${filename} (${rows.length} records)`);
        acctStats.downloaded++;
        stats.successfulDownloads++;

      } catch (e) {
        logError(`Campaign ${cid} (${cname}): ${e.message}`);
        acctStats.failed++;
        stats.failedDownloads++;
      }
    }

    stats.byAccount.push(acctStats);
    log("");
  }

  const elapsed = Date.now() - startTime;
  const minutes = Math.floor(elapsed / 60000);
  const seconds = Math.floor((elapsed % 60000) / 1000);
  stats.timeElapsed = `${minutes}m ${seconds}s`;

  const statsLines = formatStatistics(stats);
  logLines.push("");
  logLines.push(...statsLines);

  const stamp = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
  const logPath = path.join(logDir, `BulkExport_${stamp}.log`);
  await writeLog(logPath, logLines);

  let errorPath = null;
  if (errorLines.length > 0) {
    errorPath = path.join(logDir, `BulkExport_${stamp}_ERRORS.log`);
    const errorHeader = [
      "VoApps Tools — Bulk Campaign Export ERRORS",
      `Timestamp: ${new Date().toISOString()}`,
      `Total Errors: ${errorLines.length}`,
      "═══════════════════════════════════════",
      ""
    ];
    await writeLog(errorPath, [...errorHeader, ...errorLines]);
  }

  lastArtifacts.csvPath = baseDir;
  lastArtifacts.logPath = logPath;
  lastArtifacts.errorPath = errorPath;

  // Send statistics to SSE before completion
  statsLines.forEach(line => sendSSELog(job_id, line));

  // Send completion message via SSE
  if (job_id) {
    const job = jobs.get(job_id);
    if (job && job.sseResponse) {
      try {
        job.sseResponse.write(`data: ${JSON.stringify({ type: 'complete' })}\n\n`);
        job.sseResponse.end();
      } catch (e) {}
    }
    jobs.delete(job_id);
  }

  return { 
    archivePath: baseDir, 
    logPath, 
    errorPath,
    stats 
  };
}

// HTTP Server
function createHttpServer() {
  return http.createServer(async (req, res) => {
    const urlObj = parseUrl(req.url || "", true);
    const { pathname } = urlObj;

    if (req.method === "GET" && pathname === "/api/ping") {
      return sendJson(res, 200, { ok: true });
    }

    // SSE endpoint for real-time logging
    if (req.method === "GET" && pathname.startsWith("/api/stream/")) {
      const job_id = pathname.replace("/api/stream/", "");
      
      res.writeHead(200, {
        "Content-Type": "text/event-stream",
        "Cache-Control": "no-cache",
        "Connection": "keep-alive"
      });
      
      // Store SSE response in jobs map
      if (!jobs.has(job_id)) {
        jobs.set(job_id, { sseResponse: res, controller: null });
      } else {
        const job = jobs.get(job_id);
        job.sseResponse = res;
      }
      
      // Send initial connected message
      res.write(`data: ${JSON.stringify({ type: 'connected' })}\n\n`);
      
      // Handle client disconnect
      req.on('close', () => {
        const job = jobs.get(job_id);
        if (job) {
          job.sseResponse = null;
        }
      });
      
      return;
    }

    if (req.method === "POST" && pathname === "/api/accounts") {
      const body = await readJson(req);
      const api_key = body.api_key || "";
      const filter = body.filter || "all";
      if (!api_key) return sendJson(res, 400, { ok: false, error: "Missing api_key" });

      const r = await voappsGetAccounts(api_key, filter);
      if (!r.ok) return sendJson(res, r.status || 500, { ok: false, error: `HTTP ${r.status}` });

      const accounts = Array.isArray(r.json?.accounts) ? r.json.accounts : [];
      return sendJson(res, 200, { ok: true, accounts });
    }

    if (req.method === "POST" && pathname === "/api/cancel") {
      const body = await readJson(req);
      const job_id = body.job_id || "";
      if (!job_id) return sendJson(res, 400, { ok: false, error: "Missing job_id" });

      const job = jobs.get(job_id);
      if (job?.controller) {
        try { job.controller.abort(); } catch (_) {}
        jobs.delete(job_id);
        return sendJson(res, 200, { ok: true, cancelled: true });
      }
      return sendJson(res, 200, { ok: true, cancelled: false });
    }

    if (req.method === "POST" && pathname === "/api/pause") {
      const body = await readJson(req);
      const job_id = body.job_id || "";
      if (!job_id) return sendJson(res, 400, { ok: false, error: "Missing job_id" });

      const job = jobs.get(job_id);
      if (job) {
        job.paused = true;
        return sendJson(res, 200, { ok: true, paused: true });
      }
      return sendJson(res, 200, { ok: true, paused: false });
    }

    if (req.method === "POST" && pathname === "/api/resume") {
      const body = await readJson(req);
      const job_id = body.job_id || "";
      if (!job_id) return sendJson(res, 400, { ok: false, error: "Missing job_id" });

      const job = jobs.get(job_id);
      if (job) {
        job.paused = false;
        return sendJson(res, 200, { ok: true, resumed: true });
      }
      return sendJson(res, 200, { ok: true, resumed: false });
    }

    if (req.method === "POST" && pathname === "/api/shutdown") {
      sendJson(res, 200, { ok: true });
      setTimeout(() => {
        try { stopServer(); } catch (_) {}
        process.exit(0);
      }, 250);
      return;
    }

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
          job_id: body.job_id || null
        });

        return sendJson(res, 200, {
          ok: true,
          message: "Search complete",
          artifacts: { csvPath: out.csvPath, logPath: out.logPath },
          matches: out.matches
        });
      } catch (e) {
        const cancelled = e.message === "Cancelled";
        return sendJson(res, cancelled ? 499 : 500, { ok: false, error: e.message });
      }
    }

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
          job_id: body.job_id || null
        });

        const artifacts = { csvPath: out.csvPath, logPath: out.logPath };
        if (out.analysisPath) artifacts.analysisPath = out.analysisPath;

        return sendJson(res, 200, {
          ok: true,
          message: "Combine complete",
          artifacts,
          totalRows: out.totalRows
        });
      } catch (e) {
        const cancelled = e.message === "Cancelled";
        return sendJson(res, cancelled ? 499 : 500, { ok: false, error: e.message });
      }
    }

    if (req.method === "POST" && pathname === "/api/analyze-csv") {
      try {
        // Parse multipart form data manually
        const boundary = req.headers['content-type']?.split('boundary=')[1];
        if (!boundary) throw new Error("No boundary found");

        const chunks = [];
        for await (const chunk of req) chunks.push(chunk);
        const body = Buffer.concat(chunks).toString();

        // Extract CSV content and form fields
        const parts = body.split(`--${boundary}`);
        let csvText = '';
        let minConsec = 4;
        let minSpan = 30;

        for (const part of parts) {
          if (part.includes('name="csv"')) {
            const contentStart = part.indexOf('\r\n\r\n') + 4;
            const contentEnd = part.lastIndexOf('\r\n');
            csvText = part.substring(contentStart, contentEnd);
          } else if (part.includes('name="min_consec_unsuccessful"')) {
            const val = part.split('\r\n\r\n')[1]?.split('\r\n')[0];
            minConsec = parseInt(val) || 4;
          } else if (part.includes('name="min_run_span_days"')) {
            const val = part.split('\r\n\r\n')[1]?.split('\r\n')[0];
            minSpan = parseInt(val) || 30;
          }
        }

        if (!csvText) throw new Error("No CSV data found");

        // Parse CSV to rows
        const { rows } = parseCsv(csvText);

        // Generate analysis
        const folders = createOutputFolders();
        const outDir = folders.combineCampaigns;
        const stamp = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
        const analysisFilename = `NumberAnalysis_${stamp}.xlsx`;
        const analysisPath = path.join(outDir, analysisFilename);

        await generateTrendAnalysis(rows, analysisPath, minConsec, minSpan);

        lastArtifacts.analysisPath = analysisPath;

        return sendJson(res, 200, {
          ok: true,
          message: "Analysis complete",
          artifacts: { analysisPath }
        });
      } catch (e) {
        return sendJson(res, 500, { ok: false, error: e.message });
      }
    }

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
        const cancelled = e.message === "Cancelled";
        return sendJson(res, cancelled ? 499 : 500, { ok: false, error: e.message });
      }
    }

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
  });
}

async function startServer() {
  if (serverInstance && serverUrl) return { url: serverUrl, port: PORT };

  serverInstance = createHttpServer();

  await new Promise((resolve, reject) => {
    serverInstance.once("error", reject);
    serverInstance.listen(PORT, HOST, () => resolve());
  });

  serverUrl = `http://${HOST}:${PORT}`;
  console.log(`[VoApps Tools] Server listening on ${serverUrl}`);

  return { url: serverUrl, port: PORT };
}

async function stopServer() {
  if (!serverInstance) return;
  await new Promise((resolve) => serverInstance.close(() => resolve()));
  serverInstance = null;
  serverUrl = null;
}

module.exports = { startServer, stopServer, getLastArtifacts };

if (require.main === module) {
  startServer().catch((e) => {
    console.error("[VoApps Tools] Failed to start:", e);
    process.exit(1);
  });
}