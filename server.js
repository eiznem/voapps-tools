/**
 * VoApps Tools — Local Server (Electron)
 * Version: 3.2.0 - Delivery Intelligence Platform
 *
 * NEW IN v3.2.0:
 * - Delivery Intelligence Platform with TN Health Classification
 * - Attempt Index tracking, Variability Score, Retry Decay Curve
 * - Day-of-week recommendations for accounts/messages
 * - Global Insights with timezone detection
 * - Official Excel tables, empty log cleanup
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
 * - Automatic duplicate record handling via MD5 row IDs
 */

"use strict";

const http = require("http");
const fs = require("fs");
const fsp = require("fs/promises");
const path = require("path");
const os = require("os");
const { parse: parseUrl } = require("url");
const { createWriteStream } = require('fs');
const { generateTrendAnalysis } = require("./trendAnalyzer");
const { VERSION, VERSION_NAME } = require('./version');

// DuckDB integration
const duckdb = require('duckdb');
const crypto = require('crypto');

const PORT = process.env.PORT ? Number(process.env.PORT) : 3000;
const HOST = "127.0.0.1";
const VOAPPS_API_BASE = process.env.VOAPPS_API_BASE || "https://directdropvoicemail.voapps.com/api/v1";

const MAX_ROWS_PER_FILE = 500000; // Split at 500K rows for optimal analysis

// Database configuration
const DB_PATH = path.join(os.homedir(), 'Library', 'Application Support', 'VoApps Tools', 'voapps_data.duckdb');
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
 * Initialize DuckDB database
 */
async function initDatabase() {
  if (dbReady) return;
  
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
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
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
 * Insert rows into database with duplicate handling
 */
async function insertRows(rows, logger = null) {
  if (!dbReady) await initDatabase();
  if (rows.length === 0) return { inserted: 0, updated: 0, skipped: 0 };
  
  let inserted = 0;
  let updated = 0;
  let skipped = 0;
  
  // Process in batches to avoid memory issues
  const BATCH_SIZE = 5000;
  
  for (let i = 0; i < rows.length; i += BATCH_SIZE) {
    const batch = rows.slice(i, i + BATCH_SIZE);
    
    for (const row of batch) {
      const rowId = generateRowId(row);
      
      try {
        // Check if row exists
        const existing = await runQuery(
          `SELECT row_id FROM campaign_results WHERE row_id = ?`,
          [rowId]
        );
        
        if (existing.length > 0) {
          // Update existing row
          await runQuery(`
            UPDATE campaign_results SET
              caller_number_name = ?,
              message_name = ?,
              message_description = ?,
              updated_at = CURRENT_TIMESTAMP
            WHERE row_id = ?
          `, [
            row.caller_number_name || '',
            row.message_name || '',
            row.message_description || '',
            rowId
          ]);
          updated++;
        } else {
          // Insert new row
          await runQuery(`
            INSERT INTO campaign_results (
              row_id, number, account_id, campaign_id, campaign_name,
              caller_number, caller_number_name, message_id, message_name, message_description,
              voapps_result, voapps_code, voapps_timestamp, campaign_url, target_date
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
          `, [
            rowId,
            row.number || '',
            row.account_id || '',
            row.campaign_id || '',
            row.campaign_name || '',
            row.caller_number || '',
            row.caller_number_name || '',
            row.message_id || '',
            row.message_name || '',
            row.message_description || '',
            row.voapps_result || '',
            row.voapps_code || '',
            row.voapps_timestamp || '',
            row.campaign_url || '',
            row.target_date || ''
          ]);
          inserted++;
        }
      } catch (err) {
        if (logger) logger(`[DuckDB] Error processing row: ${err.message}`);
        skipped++;
      }
    }
    
    if (logger && batch.length > 0) {
      logger(`[DuckDB] Batch ${Math.floor(i / BATCH_SIZE) + 1}: ${inserted} inserted, ${updated} updated, ${skipped} skipped`);
    }
  }
  
  return { inserted, updated, skipped };
}

/**
 * Get database statistics
 */
async function getDatabaseStats() {
  if (!dbReady) {
    try {
      await initDatabase();
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
 * Export database to CSV
 */
async function exportDatabase(outputPath) {
  if (!dbReady) await initDatabase();
  
  try {
    const rows = await runQuery(`
      SELECT * FROM campaign_results
      ORDER BY target_date DESC, voapps_timestamp DESC
    `);
    
    if (rows.length === 0) {
      return {
        success: false,
        error: 'No data to export'
      };
    }
    
    // Write to CSV
    const headers = Object.keys(rows[0]);
    const csvContent = [
      headers.join(','),
      ...rows.map(row => headers.map(h => {
        const val = row[h];
        // Handle null/undefined
        if (val === null || val === undefined) return '';
        // Convert to string (handles numbers, BigInt, dates, etc.)
        const str = String(val);
        // Escape if contains comma or quote
        return str.includes(',') || str.includes('"') ? `"${str.replace(/"/g, '""')}"` : str;
      }).join(','))
    ].join('\n');
    
    await fsp.writeFile(outputPath, csvContent, 'utf-8');
    
    return {
      success: true,
      path: outputPath,
      rowCount: rows.length
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

function createOutputFolders() {
  const base = path.join(os.homedir(), "Downloads", "VoApps Tools");
  const logs = path.join(base, "Logs");
  const output = path.join(base, "Output");
  const phoneHistory = path.join(output, "Phone Number History");
  const combineCampaigns = path.join(output, "Combine Campaigns");
  const bulkExport = path.join(output, "Bulk Campaign Export");

  for (const dir of [base, logs, output, phoneHistory, combineCampaigns, bulkExport]) {
    fs.mkdirSync(dir, { recursive: true });
  }

  return { base, logs, phoneHistory, combineCampaigns, bulkExport };
}

function createLogger(logPath, errorPath, verbosity = "normal", jobId = null) {
  const logStream = fs.createWriteStream(logPath, { flags: "a" });
  const errorStream = errorPath ? fs.createWriteStream(errorPath, { flags: "a" }) : null;

  function log(message, isError = false) {
    if (verbosity === "none") return;
    
    const timestamp = getLogTimestamp();
    const line = `${timestamp} ${message}\n`;

    // Safety check: only write if streams are still writable
    if (isError && errorStream && !errorStream.writableEnded) {
      try {
        errorStream.write(line);
      } catch (e) {
        // Stream closed, ignore
      }
    }
    if (logStream && !logStream.writableEnded) {
      try {
        logStream.write(line);
      } catch (e) {
        // Stream closed, ignore
      }
    }

    // NEW: Send to SSE stream if jobId provided
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
  
  // Debug logging for troubleshooting
  console.log(`[API Request] ${url}`);
  console.log(`[API Request] Auth header length: ${apiKey.length} chars`);

  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), 30000);

  try {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json",
        Accept: "application/json"
      },
      signal: controller.signal
    });

    clearTimeout(timeoutId);

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
    clearTimeout(timeoutId);
    if (err.name === "AbortError") throw new Error("Request timeout (30s)");
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
  const bufferDays = 30;
  const start = new Date(startDate);
  start.setDate(start.getDate() - bufferDays);
  const bufferedStart = dateToYMD(start);

  const end = new Date(endDate);
  end.setDate(end.getDate() + 1);
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
        const target = new Date(targetDate);
        const startObj = new Date(startDate);
        const endObj = new Date(endDate);
        return target >= startObj && target <= endObj;
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

    // Fetch caller numbers, messages, and account timezones
    const callerNumberNames = include_caller
      ? await fetchCallerNumbers(api_key, account_ids, log, "normal")
      : {};

    // Fetch account timezones for discrepancy detection
    const accountTimezones = await fetchAccountTimezones(api_key, account_ids, log, "normal");

    const messageInfo = {};
    if (include_message_meta) {
      // Fetch message metadata for all accounts
      for (const accountId of account_ids) {
        try {
          const data = await retryableApiCall(`/accounts/${accountId}/messages?filter=all`, api_key, log, "normal");
          if (Array.isArray(data.messages)) {
            for (const msg of data.messages) {
              const key = `${accountId}:${msg.id}`;
              messageInfo[key] = {
                name: msg.name || '',
                description: msg.description || ''
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
    const numberSet = new Set(numbers.map(n => String(n).trim()));

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
              campaign_id: campaignId,
              campaign_name: campaign.name || '',
              caller_number: callerNum,
              caller_number_name: callerNumberNames[callerKey] || '',
              message_id: messageId,
              message_name: messageInfo[messageKey]?.name || '',
              message_description: messageInfo[messageKey]?.description || '',
              voapps_result: row.voapps_result || '',
              voapps_code: row.voapps_code || '',
              voapps_timestamp: row.voapps_timestamp || '',
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
        'number', 'account_id', 'campaign_id', 'campaign_name',
        'caller_number', 'caller_number_name', 'message_id', 'message_name', 'message_description',
        'voapps_result', 'voapps_code', 'voapps_timestamp', 'campaign_url'
      ];

      const csvResult = await writeCsv(csvPath, allMatches, headers, log, MAX_ROWS_PER_FILE);
      allCsvFiles = csvResult.files;
      wasSplit = csvResult.wasSplit;
      fileCount = csvResult.fileCount;

      lastArtifacts.csvPath = csvPath;
    }

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
      wasSplit,
      fileCount
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

    // Fetch caller numbers, messages, and account timezones
    const callerNumberNames = include_caller
      ? await fetchCallerNumbers(api_key, account_ids, log, "normal")
      : {};

    // Fetch account timezones for discrepancy detection
    const accountTimezones = await fetchAccountTimezones(api_key, account_ids, log, "normal");

    const messageInfo = {};
    if (include_message_meta) {
      for (const accountId of account_ids) {
        try {
          const data = await retryableApiCall(`/accounts/${accountId}/messages?filter=all`, api_key, log, "normal");
          if (Array.isArray(data.messages)) {
            for (const msg of data.messages) {
              const key = `${accountId}:${msg.id}`;
              messageInfo[key] = {
                name: msg.name || '',
                description: msg.description || ''
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
            campaign_id: campaignId,
            campaign_name: campaign.name || '',
            caller_number: callerNum,
            caller_number_name: callerNumberNames[callerKey] || '',
            message_id: messageId,
            message_name: messageInfo[messageKey]?.name || '',
            message_description: messageInfo[messageKey]?.description || '',
            voapps_result: row.voapps_result || '',
            voapps_code: row.voapps_code || '',
            voapps_timestamp: row.voapps_timestamp || '',
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

    if (output_mode === "csv" || output_mode === "both") {
      csvPath = path.join(folders.combineCampaigns, `${filePrefix}combined_${suffix}.csv`);

      const headers = [
        'number', 'account_id', 'campaign_id', 'campaign_name',
        'caller_number', 'caller_number_name', 'message_id', 'message_name', 'message_description',
        'voapps_result', 'voapps_code', 'voapps_timestamp', 'campaign_url'
      ];

      const csvResult = await writeCsv(csvPath, allRows, headers, log, MAX_ROWS_PER_FILE);
      allCsvFiles = csvResult.files;
      wasSplit = csvResult.wasSplit;
      fileCount = csvResult.fileCount;

      lastArtifacts.csvPath = csvPath;
    }

    // Generate trend analysis if requested
    let analysisPath = null;
    if (generate_trend_analysis && (output_mode === "csv" || output_mode === "both")) {
      log(`\n📊 Generating trend analysis...`);
      
      const analysisFilename = `${filePrefix}number_analysis_${suffix}.xlsx`;
      analysisPath = path.join(folders.combineCampaigns, analysisFilename);

      // Use allCsvFiles if split, otherwise pass rows directly
      const analysisInput = wasSplit ? allCsvFiles : allRows;
      
      await generateTrendAnalysis(
        analysisInput,
        analysisPath,
        min_consec_unsuccessful,
        min_run_span_days,
        messageInfo,
        callerNumberNames,
        accountTimezones
      );

      lastArtifacts.analysisPath = analysisPath;
      log(`✅ Analysis generated: ${analysisFilename}`);
    }

    log(`\n✅ Combine complete!`);
    
    // Send completion signal
    if (job_id) {
      sendProgress(job_id, {
        status: 'complete',
        progress: 100,
        message: `Combine complete! Total: ${allRows.length.toLocaleString()} rows`
      });
    }
    
    close();

    return {
      csvPath,
      allCsvFiles,
      logPath,
      analysisPath,
      totalRows: allRows.length,
      wasSplit,
      fileCount
    };
  } catch (err) {
    log(`\n❌ Fatal error: ${err.message}`, true);
    close();
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
 * Helper: Export subset of database to CSV
 */
async function exportDatabaseSubset(numbers, accountIds, startDate, endDate, outputPath) {
  if (!dbReady) await initDatabase();
  
  const numberList = numbers.map(n => `'${n}'`).join(',');
  const accountList = accountIds.map(a => `'${a}'`).join(',');
  
  const sql = `
    SELECT * FROM campaign_results
    WHERE number IN (${numberList})
      AND account_id IN (${accountList})
      AND target_date >= '${startDate}'
      AND target_date <= '${endDate}'
    ORDER BY target_date DESC, voapps_timestamp DESC
  `;
  
  const rows = await runQuery(sql);
  
  if (rows.length === 0) {
    return { success: false, error: 'No data found' };
  }
  
  const headers = Object.keys(rows[0]).filter(h => h !== 'row_id' && h !== 'created_at' && h !== 'updated_at');
  const csvContent = [
    headers.join(','),
    ...rows.map(row => headers.map(h => {
      const val = row[h];
      // Handle null/undefined
      if (val === null || val === undefined) return '';
      // Convert to string (handles numbers, BigInt, dates, etc.)
      const str = String(val);
      // Escape if contains comma or quote
      return str.includes(',') || str.includes('"') ? `"${str.replace(/"/g, '""')}"` : str;
    }).join(','))
  ].join('\n');
  
  await fsp.writeFile(outputPath, csvContent, 'utf-8');
  
  return {
    success: true,
    path: outputPath,
    rowCount: rows.length
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

    // Accounts endpoint
    if (req.method === "POST" && pathname === "/api/accounts") {
      try {
        const body = await readJson(req);
        const accounts = await fetchAllAccounts(body.api_key || "", null, "minimal", body.filter);
        return sendJson(res, 200, { ok: true, accounts });
      } catch (e) {
        console.error('[API Error - /api/accounts]', e.message, e.stack);
        return sendJson(res, 500, { ok: false, error: e.message });
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

      // Send current job state
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

      // Send any accumulated logs
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
          output_mode: body.output_mode || "csv",
          job_id: body.job_id || null,
          client_prefix: body.client_prefix || ""
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
        const body = Buffer.concat(chunks).toString();

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

        const { headers, rows } = parseCsv(csvText);
        const folders = createOutputFolders();
        const outDir = folders.combineCampaigns;
        
        if (rows.length > MAX_ROWS_PER_FILE) {
          const suffix = getFilenameSuffix(outDir, 'UploadedCSV');
          const tempCsvPath = path.join(outDir, `UploadedCSV_${suffix}.csv`);
          const csvResult = await writeCsv(tempCsvPath, rows, headers, null, MAX_ROWS_PER_FILE);
          
          const analysisFilename = `NumberAnalysis_${suffix}.xlsx`;
          const analysisPath = path.join(outDir, analysisFilename);
          
          await generateTrendAnalysis(csvResult.files, analysisPath, minConsec, minSpan, {}, {});
          
          lastArtifacts.analysisPath = analysisPath;
          
          return sendJson(res, 200, {
            ok: true,
            message: `Analysis complete (${csvResult.totalRows.toLocaleString()} rows across ${csvResult.fileCount} files)`,
            artifacts: { analysisPath }
          });
        }
        
        const suffix = getFilenameSuffix(outDir, 'NumberAnalysis');
        const analysisFilename = `NumberAnalysis_${suffix}.xlsx`;
        const analysisPath = path.join(outDir, analysisFilename);

        await generateTrendAnalysis(rows, analysisPath, minConsec, minSpan, {}, {});

        lastArtifacts.analysisPath = analysisPath;

        return sendJson(res, 200, {
          ok: true,
          message: "Analysis complete",
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
          client_prefix = ""
        } = body;

        if (!dbReady) {
          throw new Error("Database not ready");
        }

        // Query database for rows in date range
        console.log(`[DB Analysis] Querying database for ${start_date} to ${end_date}...`);

        const query = `
          SELECT
            number, account_id, campaign_id, campaign_name,
            caller_number, caller_number_name, message_id, message_name, message_description,
            voapps_result, voapps_code, voapps_timestamp, campaign_url
          FROM campaign_results
          WHERE voapps_timestamp >= '${start_date}' AND voapps_timestamp <= '${end_date}T23:59:59'
          ORDER BY voapps_timestamp
        `;

        const rows = [];
        await new Promise((resolve, reject) => {
          db.all(query, (err, result) => {
            if (err) return reject(err);
            for (const row of result) {
              rows.push(row);
            }
            resolve();
          });
        });

        console.log(`[DB Analysis] Found ${rows.length.toLocaleString()} rows`);

        if (rows.length === 0) {
          throw new Error("No data found in database for the specified date range");
        }

        // Generate analysis
        const folders = createOutputFolders();
        const suffix = getFilenameSuffix(folders.combineCampaigns, 'db_analysis');
        const filePrefix = client_prefix ? `${client_prefix}_` : "";
        const analysisFilename = `${filePrefix}db_analysis_${suffix}.xlsx`;
        const analysisPath = path.join(folders.combineCampaigns, analysisFilename);

        console.log(`[DB Analysis] Generating analysis: ${analysisFilename}`);

        await generateTrendAnalysis(
          rows,
          analysisPath,
          min_consec_unsuccessful,
          min_run_span_days,
          {},  // messageMap
          {}   // callerMap
        );

        lastArtifacts.analysisPath = analysisPath;

        return sendJson(res, 200, {
          ok: true,
          message: "Database analysis complete",
          rowCount: rows.length,
          artifacts: { analysisPath }
        });
      } catch (e) {
        console.error('[API Error - /api/analyze-database]', e.message, e.stack);
        return sendJson(res, 500, { ok: false, error: e.message });
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
  } catch (err) {
    console.error('[DuckDB] Failed to initialize database:', err);
  }

  serverInstance = createHttpServer();

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