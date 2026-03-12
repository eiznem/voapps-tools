// trendAnalyzer.js - Delivery Intelligence Report
// Analyzes phone numbers, caller numbers, and messages for delivery insights
// Generates comprehensive Excel analysis workbooks from campaign data
//
// Features:
// - Attempt Index tracking per TN (resets after success)
// - Success Probability by attempt number (decay curve)
// - TN Health Classification (Healthy/Delivery Unlikely/Never Delivered)
// - Never Delivered detection
// - Variability Score (message, day, hour, caller diversity)
// - Back-to-back identical message detection
// - Day of week entropy analysis with recommendations
// - Timezone detection from voapps_timestamp
// - Message intent inference from names
// - Day-of-week usage recommendations for accounts/messages
// - List Quality Grade (A-D)
// - Executive Summary with recommendations
// - Official Excel tables with VoApps colors
// - DDVM terminology throughout
// - Correct VoApps result codes in glossary

const ExcelJS = require('exceljs');
const Papa = require('papaparse');
const fs = require('fs');
const path = require('path');
const { generateBusinessReviewSlides } = require('./businessReview');

// Import VERSION from central source of truth
const { VERSION } = require('./version');

// VoApps brand colors (official palette)
const VOAPPS_DARK_NAVY = '0D053F';     // Header background (darkest)
const VOAPPS_PURPLE = '3F2FB8';        // Primary purple
const VOAPPS_PURPLE_MID = '6558C6';    // Mid purple
const VOAPPS_PURPLE_LIGHT = 'B2ACE2';  // Light purple
const VOAPPS_PURPLE_PALE = 'D9D6F1';   // Pale purple for alternating rows
const VOAPPS_PINK = 'FF4B7D';          // Accent pink (primary)
const VOAPPS_PINK_LIGHT = 'FF93B1';    // Light pink
const VOAPPS_PINK_PALE = 'FFB7CB';     // Pale pink
const VOAPPS_BLUSH = 'FAD6D7';         // Blush/rose
const VOAPPS_CREAM = 'FBF7F3';         // Cream background
const VOAPPS_CHARCOAL = '2E2C3E';      // Dark text color

function log(message) {
  const now = new Date();
  const timeStr = now.toLocaleTimeString('en-US', { hour12: true });
  console.log(`[${timeStr}] ${message}`);
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Calculate Shannon entropy for distribution analysis
 */
function calculateEntropy(distribution) {
  const total = distribution.reduce((sum, val) => sum + val, 0);
  if (total === 0) return 0;

  let entropy = 0;
  for (const val of distribution) {
    if (val > 0) {
      const p = val / total;
      entropy -= p * Math.log2(p);
    }
  }
  return entropy;
}

/**
 * Infer message intent from message name
 */
function inferMessageIntent(messageName) {
  if (!messageName) return 'unknown';
  const n = messageName.toLowerCase();

  // ── Special message types (pattern-identified) ────────────────────────────
  if (n.includes('lcm') || n.includes('limited content'))
    return 'LCM';
  if (n.includes('zortman') || n.includes('modified zortman'))
    return 'Modified Zortman';

  // ── Collections ───────────────────────────────────────────────────────────
  if (n.includes('1st payment') || n.includes('first payment') || n.includes('1st pmt'))
    return 'first payment default';
  if (n.includes('charge off') || n.includes('charge-off') || n.includes('pre-charge') ||
      n.includes('60+') || n.includes('60 dpd') || n.includes('90') || n.includes('pre charge'))
    return 'pre-charge-off collections';
  if (n.includes('overdrawn') || n.includes('negative share') || n.includes('nsf') || n.includes('overdraft'))
    return 'overdrawn account';
  if (n.includes('friendly') && (n.includes('remind') || n.includes('payment')))
    return 'friendly payment reminder';
  if (n.includes('hardship') || n.includes('temporary hardship') ||
      n.includes('financial difficulty') || n.includes('assistance program'))
    return 'friendly payment reminder';   // hardship/assistance = soft collections
  if (n.includes('healthcare') || n.includes('medical') || n.includes('hospital') ||
      n.includes('ebo') || n.includes('early out') || /\beo\b/.test(n) ||
      n.includes('3rd party') || n.includes('third party') ||
      n.includes('debt buyer') || n.includes('extended business'))
    return 'healthcare or third-party collections';
  if (n.includes('collect') || n.includes('past due') || n.includes('dpd') ||
      n.includes('delinquent') || n.includes('balance due'))
    return 'delinquent loan payment';

  // ── Consumer / Direct Lending ─────────────────────────────────────────────
  if (n.includes('pre-approv') || n.includes('pre approv') || n.includes('refinanc') ||
      n.includes('blaster') || n.includes('rate promo'))
    return 'pre-approval or refinance offer';
  if (n.includes('application') || n.includes('app follow') || n.includes('pending app') ||
      n.includes('verification') || n.includes('pending verif'))
    return 'loan application follow-up';
  if (n.includes('cpi') || n.includes('force place') || n.includes('force-place') || n.includes('insurance'))
    return 'CPI or insurance notification';
  if (n.includes('title'))
    return 'title perfection reminder';
  if (n.includes('skip pay') || n.includes('holiday skip') || n.includes('skip-pay'))
    return 'holiday skip payment offer';
  if (n.includes('unused reward') || n.includes('rewards'))
    return 'unused rewards notification';
  if (n.includes('lead') || n.includes('remarketing') || n.includes('automatic lead') ||
      n.includes('auto lead'))
    return 'marketing or lead response';

  // ── Servicing ─────────────────────────────────────────────────────────────
  if (n.includes('dormant'))
    return 'dormant account notification';
  if (n.includes('fraud') || n.includes('phish') || n.includes('scam'))
    return 'fraud alert';
  if (n.includes('card activ') || n.includes('card service') || n.includes('card issue'))
    return 'card services notification';
  if (n.includes('downtime') || n.includes('olb') || n.includes('system update') ||
      n.includes('online banking'))
    return 'system update or downtime notice';
  if (n.includes('investment') || n.includes('wealth'))
    return 'product or rate promotion';
  if (n.includes('disaster') || n.includes('closure') || n.includes('holiday closure') ||
      n.includes('emergency'))
    return 'disaster recovery or closure notice';

  // ── Marketing ─────────────────────────────────────────────────────────────
  if (n.includes('welcome') || n.includes('intro') || n.includes('onboard') ||
      n.includes('indirect lending welcome') || n.includes('new member'))
    return 'new member welcome';
  if (n.includes('offer') || n.includes('promo') || n.includes('promotion') ||
      n.includes('rate') || n.includes('cd special') || n.includes('discount'))
    return 'product or rate promotion';
  if (n.includes('event') || n.includes('workshop') || n.includes('seminar') || n.includes('webinar'))
    return 'educational event or workshop';
  if (n.includes('cross sell') || n.includes('ancillary') || n.includes('gap ins') ||
      n.includes('partnership'))
    return 'product or rate promotion';

  // ── Payment reminders (generic, when not more specific) ───────────────────
  if (n.includes('remind') || n.includes('promise') || n.includes('payment reminder'))
    return 'friendly payment reminder';

  // ── Fallback ──────────────────────────────────────────────────────────────
  if (n.includes('follow') || n.includes('follow-up') || n.includes('followup'))
    return 'loan application follow-up';
  if (n.includes('loan') || n.includes('servic') || n.includes('account'))
    return 'delinquent loan payment';

  return 'general notice';
}

/**
 * Normalize a phone number to its last 10 digits for comparison.
 */
function normalizePhone(phone) {
  if (!phone) return '';
  const digits = String(phone).replace(/\D/g, '');
  return digits.slice(-10);
}

/**
 * Compare the phone number mentioned in the transcript against the most-used
 * caller number for that message.
 * Returns '—' when no phone was mentioned, '✅ Match', or '⚠️ Mismatch …'.
 */
function getCallerMatchStatus(mentionedPhone, callerNumbers) {
  if (!mentionedPhone) return '—';
  const mentioned = normalizePhone(mentionedPhone);
  if (!mentioned) return '—';

  // Find the dominant (most-used) caller number for this message
  let dominantCaller = null;
  let maxCount = 0;
  for (const [num, count] of Object.entries(callerNumbers || {})) {
    if (num !== 'Unknown' && count > maxCount) {
      maxCount = count;
      dominantCaller = num;
    }
  }

  if (!dominantCaller) return '—';
  const callerNorm = normalizePhone(dominantCaller);
  if (!callerNorm) return '—';

  return mentioned === callerNorm
    ? '✅ Match'
    : `⚠️ Mismatch (caller: ${dominantCaller}, msg: ${mentionedPhone})`;
}

/**
 * Calculate TN Health classification
 */
function classifyTNHealth(successRate, consecutiveFailures, totalAttempts, recentSuccess14Days) {
  // Delivery Unlikely: Very low success + high consecutive failures
  if (successRate < 0.1 && consecutiveFailures >= 4) return 'Delivery Unlikely';
  if (consecutiveFailures >= 6) return 'Delivery Unlikely';
  if (totalAttempts >= 5 && successRate === 0) return 'Delivery Unlikely';

  // Healthy: Good performance
  return 'Healthy';
}

/**
 * Calculate List Quality Grade
 */
function calculateListGrade(healthyPct, toxicPct, neverDeliveredPct) {
  // A: >80% healthy, <5% delivery unlikely
  if (healthyPct >= 80 && toxicPct < 5 && neverDeliveredPct < 10) return 'A';
  // B: >60% healthy, <10% delivery unlikely
  if (healthyPct >= 60 && toxicPct < 10 && neverDeliveredPct < 20) return 'B';
  // C: >40% healthy, <20% delivery unlikely
  if (healthyPct >= 40 && toxicPct < 20) return 'C';
  // D: Everything else
  return 'D';
}

/**
 * Format date safely
 */
function formatDate(d) {
  if (!d || isNaN(d.getTime())) return 'Unknown';
  return d.toLocaleDateString();
}

/**
 * Format day distribution string
 */
function formatDayDistribution(counts, total) {
  const dayNames = ['SUN', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT'];
  return dayNames.map((day, i) => {
    const pct = total > 0 ? Math.round((counts[i] / total) * 100) : 0;
    return `${day}: ${pct}%`;
  }).join(' | ');
}

/**
 * Extract timezone from voapps_timestamp
 * Examples: "2026-01-16 15:50:21 UTC", "2025-12-10 14:56:39 -07:00"
 */
function extractTimezone(timestamp) {
  if (!timestamp) return null;

  // Check for UTC
  if (timestamp.includes(' UTC')) {
    return { offset: '+00:00', name: 'UTC' };
  }

  // Check for offset like -07:00 or +05:30
  const match = timestamp.match(/([+-]\d{2}:\d{2})$/);
  if (match) {
    const offset = match[1];
    // Map common offsets to timezone names
    const tzNames = {
      '-05:00': 'America/New_York',
      '-06:00': 'America/Chicago',
      '-07:00': 'America/Denver',
      '-08:00': 'America/Los_Angeles',
      '-04:00': 'America/New_York (DST)',
      '-09:00': 'America/Anchorage',
      '-10:00': 'Pacific/Honolulu',
      '+00:00': 'UTC'
    };
    return { offset, name: tzNames[offset] || offset };
  }

  return null;
}

/**
 * Parse a voapps_timestamp into { utcDate, localHour, localDayOfWeek } without
 * relying on JavaScript's system timezone. The timestamp already encodes the local
 * time (either as " UTC" or as " -07:00" etc.), so we read hour/day directly from
 * the string to avoid system-timezone corruption.
 *
 * Formats handled:
 *   "2025-11-16 14:47:05 UTC"        → UTC local time
 *   "2025-11-16 09:47:05 -07:00"     → local time at -07:00
 */
function parseTimestampLocal(timestamp) {
  if (!timestamp || typeof timestamp !== 'string') return null;
  const ts = timestamp.trim();

  // Match: YYYY-MM-DD HH:MM:SS (UTC | ±HH:MM)
  const match = ts.match(/^(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2}):(\d{2})\s+(?:UTC|([+-]\d{2}:\d{2}))$/);
  if (!match) return null;

  const [, year, month, day, hour, minute, second, offsetStr] = match;
  const localHour = parseInt(hour, 10);

  // Compute UTC ms so we can derive day-of-week in the LOCAL timezone
  const offsetMinutes = offsetStr
    ? (parseInt(offsetStr.slice(0, 3), 10) * 60 + parseInt(offsetStr.slice(4), 10) * Math.sign(parseInt(offsetStr, 10) || 1))
    : 0; // UTC = 0 offset

  // Build UTC Date by subtracting the local offset
  const localMs = Date.UTC(
    parseInt(year), parseInt(month) - 1, parseInt(day),
    parseInt(hour), parseInt(minute), parseInt(second)
  );
  const utcMs = localMs - offsetMinutes * 60 * 1000;
  const utcDate = new Date(utcMs);

  // Day of week in local time: shift UTC date back to local then read getUTCDay()
  const localDate = new Date(utcMs + offsetMinutes * 60 * 1000);
  const localDayOfWeek = localDate.getUTCDay();  // 0=Sun … 6=Sat in local time

  // Local date string YYYY-MM-DD (used for daily/weekly trend tracking)
  const lY = localDate.getUTCFullYear();
  const lM = String(localDate.getUTCMonth() + 1).padStart(2, '0');
  const lD = String(localDate.getUTCDate()).padStart(2, '0');
  const localDateStr = `${lY}-${lM}-${lD}`;

  return { utcDate, localHour, localDayOfWeek, localDateStr };
}

/**
 * Normalize a raw voapps_timestamp string: replace "+00:00" suffix with "UTC".
 * All other offsets (e.g. "-05:00") are left unchanged.
 */
function formatTimestamp(ts) {
  if (!ts || typeof ts !== 'string') return ts;
  return ts.replace(/\s\+00:00$/, ' UTC');
}

/**
 * Binary-search a campaign timestamp list (sorted by row index) to find the
 * entry whose index is closest to targetIdx. Returns the timestamp string or null.
 * tsList: Array<{ idx: number, ts: string }>, sorted ascending by idx.
 */
function findNearestTs(tsList, targetIdx) {
  if (!tsList || tsList.length === 0) return null;
  let lo = 0, hi = tsList.length - 1;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (tsList[mid].idx < targetIdx) lo = mid + 1;
    else hi = mid;
  }
  // lo = first entry with idx >= targetIdx
  const prev = lo > 0 ? tsList[lo - 1] : null;
  const next = lo < tsList.length ? tsList[lo] : null;
  if (!prev && !next) return null;
  if (!prev) return next.ts;
  if (!next) return prev.ts;
  const dp = Math.abs(prev.idx - targetIdx);
  const dn = Math.abs(next.idx - targetIdx);
  return dp <= dn ? prev.ts : next.ts;
}

/**
 * Convert a ms-epoch number to a "YYYY-MM-DD HH:MM:SS UTC" timestamp string.
 * Used to reconstruct a filled timestamp from campaignTsMap epoch entries.
 */
function epochToTimestamp(ms) {
  const d = new Date(ms);
  const pad = n => String(n).padStart(2, '0');
  return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())} ` +
         `${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())} UTC`;
}

/**
 * Like findNearestTs but for lists whose entries carry { idx, tsMs } (epoch ms number)
 * instead of { idx, ts } (string).  Returns epoch ms or null.
 * Storing epoch numbers in campaignTsMap instead of strings saves ~80 bytes per entry
 * (~100 MB for a 1.4 M-row dataset) and reduces GC pressure during Phase 2 startup.
 */
function findNearestTsMs(tsList, targetIdx) {
  if (!tsList || tsList.length === 0) return null;
  let lo = 0, hi = tsList.length - 1;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (tsList[mid].idx < targetIdx) lo = mid + 1; else hi = mid;
  }
  const prev = lo > 0 ? tsList[lo - 1] : null;
  const next = lo < tsList.length ? tsList[lo] : null;
  if (!prev && !next) return null;
  if (!prev) return next.tsMs;
  if (!next) return prev.tsMs;
  return Math.abs(prev.idx - targetIdx) <= Math.abs(next.idx - targetIdx) ? prev.tsMs : next.tsMs;
}

/**
 * Get friendly timezone name from IANA or offset
 */
function getTimezoneDisplayName(tzInfo) {
  if (!tzInfo) return 'Unknown';

  const ianaToFriendly = {
    'America/New_York': 'Eastern Time (ET)',
    'America/Chicago': 'Central Time (CT)',
    'America/Denver': 'Mountain Time (MT)',
    'America/Los_Angeles': 'Pacific Time (PT)',
    'America/Anchorage': 'Alaska Time (AKT)',
    'Pacific/Honolulu': 'Hawaii Time (HT)',
    'UTC': 'UTC'
  };

  if (tzInfo.name && ianaToFriendly[tzInfo.name]) {
    return ianaToFriendly[tzInfo.name];
  }

  // If we have an offset, show it with inferred name
  if (tzInfo.offset) {
    const offsetNames = {
      '-05:00': 'Eastern Time (ET)',
      '-06:00': 'Central Time (CT)',
      '-07:00': 'Mountain Time (MT)',
      '-08:00': 'Pacific Time (PT)',
      '-04:00': 'Eastern Daylight Time (EDT)',
      '+00:00': 'UTC'
    };
    return offsetNames[tzInfo.offset] || `UTC${tzInfo.offset}`;
  }

  return tzInfo.name || 'Unknown';
}

/**
 * Check if account/message is only used on specific days
 */
function getDayUsagePattern(dayOfWeekCounts) {
  const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const usedDays = [];
  const total = dayOfWeekCounts.reduce((sum, c) => sum + c, 0);

  if (total === 0) return { limited: false, days: [] };

  for (let i = 0; i < 7; i++) {
    if (dayOfWeekCounts[i] > total * 0.1) { // At least 10% of volume on this day
      usedDays.push(dayNames[i]);
    }
  }

  // Flag as limited only if 1-2 days qualify (at the 10% threshold), excluding Sunday-only patterns
  const nonSundayUsed = usedDays.filter(d => d !== 'Sunday');

  if (nonSundayUsed.length === 0 && usedDays.length === 0) {
    // No day meets 10% threshold — very sparse data
    return { limited: false, days: [] };
  }

  if (nonSundayUsed.length <= 2 && usedDays.length <= 3) {
    // Truly limited to 1-3 days total (not just proportionally light on some days)
    const dayList = usedDays.length > 0 ? usedDays : nonSundayUsed;
    if (dayList.length === 0) return { limited: false, days: usedDays };
    return {
      limited: true,
      days: usedDays,
      recommendation: `Only used on ${dayList.join(' and ')}. Consumers who receive your message on a predictable day pattern begin to recognize and mentally categorize it as routine — reducing the likelihood they listen or call back. Rotating across additional days of the week makes your outreach feel less automated and more timely.`
    };
  }

  // Check for severely disproportionate day usage (one day < 5% when all days used)
  const dominantDays = [];
  const lightDays = [];
  for (let i = 0; i < 7; i++) {
    const pct = dayOfWeekCounts[i] / total;
    if (pct > 0 && pct < 0.05 && dayNames[i] !== 'Sunday') {
      lightDays.push(`${dayNames[i]} (${(pct * 100).toFixed(1)}%)`);
    } else if (pct >= 0.2) {
      dominantDays.push(dayNames[i]);
    }
  }

  if (lightDays.length > 0 && dominantDays.length > 0) {
    return {
      limited: true,
      days: usedDays,
      recommendation: `Volume is heavily concentrated on ${dominantDays.join(' and ')} — ${lightDays.join(', ')} received very little. Spreading attempts more evenly prevents consumers from developing a predictable "this is my weekly voicemail" expectation and increases the chance of catching them in a different mindset.`
    };
  }

  return { limited: false, days: usedDays };
}

/**
 * Map offset to IANA timezone name
 */
function offsetToIANA(offset) {
  const offsetToIanaMap = {
    '-05:00': 'America/New_York',
    '-04:00': 'America/New_York',  // EDT
    '-06:00': 'America/Chicago',
    '-05:00': 'America/Chicago',   // CDT (Note: overlaps with EST - context dependent)
    '-07:00': 'America/Denver',
    '-06:00': 'America/Denver',    // MDT
    '-08:00': 'America/Los_Angeles',
    '-07:00': 'America/Los_Angeles', // PDT
    '-09:00': 'America/Anchorage',
    '-10:00': 'Pacific/Honolulu',
    '+00:00': 'UTC'
  };
  return offsetToIanaMap[offset] || null;
}

/**
 * Check for timezone discrepancies between account settings and results
 * Returns array of discrepancies with recommendations
 */
function detectTimezoneDiscrepancies(accountTimezones, accountResultTimezones) {
  const discrepancies = [];

  // Compare each account's configured timezone with what appears in results
  for (const [accountId, resultsOffset] of Object.entries(accountResultTimezones)) {
    const configuredTz = accountTimezones[accountId];

    if (!configuredTz || !resultsOffset) continue;

    // Get expected offset for the configured timezone
    const ianaToExpectedOffset = {
      'America/New_York': ['-05:00', '-04:00'],      // EST/EDT
      'America/Chicago': ['-06:00', '-05:00'],       // CST/CDT
      'America/Denver': ['-07:00', '-06:00'],        // MST/MDT
      'America/Los_Angeles': ['-08:00', '-07:00'],   // PST/PDT
      'America/Anchorage': ['-09:00', '-08:00'],     // AKST/AKDT
      'Pacific/Honolulu': ['-10:00'],                // HST (no DST)
      'America/Phoenix': ['-07:00'],                 // MST (no DST)
      'UTC': ['+00:00']
    };

    const expectedOffsets = ianaToExpectedOffset[configuredTz] || [];

    if (expectedOffsets.length > 0 && !expectedOffsets.includes(resultsOffset)) {
      discrepancies.push({
        accountId,
        configuredTimezone: configuredTz,
        resultsOffset,
        message: `Account ${accountId}: Configured as ${configuredTz}, but results show ${resultsOffset}`
      });
    }
  }

  return discrepancies;
}

// ============================================================================
// MAIN ANALYSIS FUNCTION
// ============================================================================

/**
 * Auto-fits column widths based on the longest cell content in each column.
 * Numbers, dates, strings, and richText values are all measured.
 * @param {object} sheet   - ExcelJS worksheet
 * @param {number} minWidth - Minimum column width (default 8)
 * @param {number} maxWidth - Maximum column width (default 60)
 */
function autoFitColumns(sheet, minWidth = 8, maxWidth = 60, skipCols = []) {
  const skipSet = new Set(skipCols);
  const colWidths = {};
  sheet.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      if (skipSet.has(colNumber)) return;
      let len = 0;
      const v = cell.value;
      if (v == null) {
        len = 0;
      } else if (typeof v === 'string') {
        // Multi-line strings: measure the longest line only
        len = v.split('\n').reduce((m, l) => Math.max(m, l.length), 0);
      } else if (typeof v === 'number') {
        len = String(v).length;
      } else if (v instanceof Date) {
        len = 19; // 'yyyy-mm-dd hh:mm:ss'
      } else if (typeof v === 'object' && v.richText) {
        len = v.richText.map(r => r.text || '').join('').length;
      } else {
        len = String(v).length;
      }
      if (len > (colWidths[colNumber] || 0)) colWidths[colNumber] = len;
    });
  });
  for (const [col, len] of Object.entries(colWidths)) {
    sheet.getColumn(Number(col)).width = Math.min(maxWidth, Math.max(minWidth, len + 2));
  }
}

/**
 * Generate Delivery Intelligence Analysis Excel Workbook
 * @param {string|Array} csvInput - CSV file path, array of file paths, or array of row objects
 * @param {string} outputPath - Output Excel file path
 * @param {number} minConsecUnsuccessful - Minimum consecutive failures threshold
 * @param {number} minRunSpanDays - Minimum span days for consecutive runs
 * @param {Object} messageMap - Map of message_id to message metadata
 * @param {Object} callerMap - Map of caller_number to caller name
 * @param {Object} accountTimezones - Map of account_id to IANA timezone (e.g., { "12345": "America/Denver" })
 * @param {string} userTimezone - User's selected timezone (IANA name or 'VoApps')
 * @param {string} userTimezoneLabel - User's timezone label (e.g., "VoApps", "ET", "MT")
 */

// Results that represent an actual delivery attempt reaching the carrier.
// Only codes 200/400/405/406/407 — the five deliverable results.
// Excluded: 300 expired, 301 canceled, 401 not wireless, 402 duplicate,
// 403 invalid US number, 404 undeliverable, 408–410 config errors, 500–504 restricted.
const DELIVERY_ATTEMPT_RESULTS = new Set([
  'successfully delivered',        // 200
  'unsuccessful delivery attempt', // 400
  'not in service',                // 405
  'voicemail not setup',           // 406
  'voicemail full',                // 407
]);

async function generateTrendAnalysis(
  csvInput,
  outputPath,
  minConsecUnsuccessful = 4,
  minRunSpanDays = 30,
  messageMap = {},
  callerMap = {},
  accountTimezones = {},
  userTimezone = 'VoApps',
  userTimezoneLabel = 'VoApps',
  includeDetailTabs = false,
  transcriptMap = {}
) {
  log(`Starting Delivery Intelligence Analysis (v${VERSION})`);

  // ── Shared containers populated by whichever input path runs below ──────────
  const numberData = {};
  const timezoneCounts = {};
  const accountResultTimezones = {};
  const CONFIG_ERROR_RESULTS = new Set([
    'invalid message id',
    'invalid caller number',
    'prohibited self call'
  ]);
  const configErrors = {
    'invalid message id':    { byCallerNumber: {}, byMessageId: {}, total: 0 },
    'invalid caller number': { byCallerNumber: {}, byMessageId: {}, total: 0 },
    'prohibited self call':  { byCallerNumber: {}, byMessageId: {}, total: 0 }
  };
  let minDate = null, maxDate = null, fourteenDaysAgo = null;
  let detectedTimezone = 'Unknown (timestamps may be UTC)';
  let timezoneDiscrepancies = [];
  let accountMostCommonOffset = {};
  let totalValidRows = 0;

  // Non-deliverable row counts — for Executive Summary breakdown
  const nonDeliverableCounts = {
    'not a wireless number': 0,   // 401
    'not a valid us number': 0,   // 403
    'duplicate number':      0,   // 402
    'undeliverable':         0,   // 404
    'restricted':            0,   // 500-504 combined
  };
  let notUSPlaceholderRows = 0;   // 403 rows where the number is all-zero digits (e.g. 0000000000)

  // Account / message / caller / time stats — populated inline during row processing
  const accountStats = {};
  const messageStats = {};
  const callerStats = {};
  const globalHourlyStats = {};
  for (let h = 0; h < 24; h++) globalHourlyStats[h] = { successful: 0, unsuccessful: 0, total: 0 };
  const globalDayStats = {};
  for (let d = 0; d < 7; d++) globalDayStats[d] = { successful: 0, unsuccessful: 0, total: 0 };
  const dailyStats = {}; // 'YYYY-MM-DD' → { attempts, successes }

  // Normalise single-file string to one-element array so we always use the file-path path
  if (typeof csvInput === 'string') csvInput = [csvInput];

  if (Array.isArray(csvInput) && csvInput.length > 0 && typeof csvInput[0] === 'string') {
    // ═══════════════════════════════════════════════════════════════════════════
    // FILE-PATH INPUT — Two-phase streaming avoids loading 1M+ rows into RAM.
    // Phase 1 builds a campaign→timestamp index (compact).
    // Phase 2 streams rows directly into numberData (no csvRows array).
    // ═══════════════════════════════════════════════════════════════════════════
    const files = csvInput;
    log(`Processing ${files.length} CSV file(s) with streaming...`);

    // ── Phase 1: light scan — record first & last timestamped row per campaign ─
    log('  Phase 1/2: Scanning campaign timestamps...');
    // campaign_id → { firstIdx, firstMs, lastIdx, lastMs }
    // Storing only the first and last timestamped row per campaign (rather than every
    // row) reduces this map from ~1.26 M objects to one object per campaign, saving
    // ~80-100 MB of heap for a 1.4 M-row dataset.
    const campaignTsMap = new Map();
    let scanIdx = 0;
    for (const csvFile of files) {
      await new Promise((resolve, reject) => {
        Papa.parse(fs.createReadStream(csvFile, { encoding: 'utf8' }), {
          header: true, skipEmptyLines: true,
          transformHeader: h => h.trim().toLowerCase(),
          chunk: results => {
            for (const row of results.data) {
              const ts = (row.voapps_timestamp || '').trim();
              if (ts) {
                const cid = (row.campaign_id || '').trim() || '__none__';
                const parsed = parseTimestampLocal(formatTimestamp(ts));
                const ms = parsed ? parsed.utcDate.getTime() : new Date(ts).getTime();
                if (!isNaN(ms)) {
                  if (!campaignTsMap.has(cid)) {
                    campaignTsMap.set(cid, { firstIdx: scanIdx, firstMs: ms, lastIdx: scanIdx, lastMs: ms });
                  } else {
                    const e = campaignTsMap.get(cid);
                    // scanIdx always increases, so firstIdx never needs updating
                    if (scanIdx > e.lastIdx) { e.lastIdx = scanIdx; e.lastMs = ms; }
                  }
                }
              }
              scanIdx++;
            }
          },
          complete: resolve, error: reject
        });
      });
    }

    // ── Phase 2: stream directly into numberData — no csvRows accumulation ────
    log('  Phase 2/2: Processing rows...');
    let rowIdx = 0;
    for (const csvFile of files) {
      log(`Reading: ${path.basename(csvFile)}`);
      await new Promise((resolve, reject) => {
        Papa.parse(fs.createReadStream(csvFile, { encoding: 'utf8' }), {
          header: true, skipEmptyLines: true,
          transformHeader: h => h.trim().toLowerCase(),
          chunk: results => {
            for (const row of results.data) {
              if (!row.number || !row.voapps_result) { rowIdx++; continue; }

              // ── Resolve and validate phone number ─────────────────────────────
              const num = String(row.number).trim();
              if (num.replace(/\D/g, '').length < 7) { rowIdx++; continue; } // skip malformed (e.g. "0")

              // ── Fill and normalise timestamp ──────────────────────────────────
              let ts = (row.voapps_timestamp || '').trim();
              const _tsOriginal = !!ts; // true only if the row had a real timestamp in the CSV
              if (!ts) {
                // Fill missing timestamp from the campaign's first/last known timestamp.
                // Using the closer of the two (by row index) gives a reasonable estimate
                // without holding all 1.26 M per-row objects in memory.
                const cid = (row.campaign_id || '').trim() || '__none__';
                const e = campaignTsMap.get(cid);
                if (e) {
                  const useFirst = Math.abs(e.firstIdx - rowIdx) <= Math.abs(e.lastIdx - rowIdx);
                  ts = epochToTimestamp(useFirst ? e.firstMs : e.lastMs);
                } else if (row.target_date) {
                  ts = `${row.target_date.trim()} 00:00:00 UTC`; // target_date fallback
                }
              } else {
                ts = formatTimestamp(ts); // normalise +00:00 → UTC
              }

              // ── Parse timestamp ───────────────────────────────────────────────
              const parsed       = ts ? parseTimestampLocal(ts) : null;
              const parsedDate   = parsed ? parsed.utcDate : (ts ? new Date(ts) : new Date(0));
              const pdOk         = parsedDate && !isNaN(parsedDate.getTime());
              const localHour    = parsed ? parsed.localHour    : (pdOk ? parsedDate.getHours() : 0);
              const localDow     = parsed ? parsed.localDayOfWeek : (pdOk ? parsedDate.getDay()  : 0);
              const resultNorm   = String(row.voapps_result || '').trim().toLowerCase();
              const isSuccess    = resultNorm === 'successfully delivered';
              // Only count as a delivery attempt if the row had a real original timestamp.
              // Proximity-inferred timestamps are used for sorting/display only.
              const isDelivery   = _tsOriginal && DELIVERY_ATTEMPT_RESULTS.has(resultNorm);

              // ── Non-deliverable tracking ──────────────────────────────────────
              if (!isDelivery) {
                if      (resultNorm === 'not a wireless number') nonDeliverableCounts['not a wireless number']++;
                else if (resultNorm === 'not a valid us number') {
                  nonDeliverableCounts['not a valid us number']++;
                  if (/^0+$/.test(num.replace(/\D/g, ''))) notUSPlaceholderRows++;
                }
                else if (resultNorm === 'duplicate number')     nonDeliverableCounts['duplicate number']++;
                else if (resultNorm === 'undeliverable')        nonDeliverableCounts['undeliverable']++;
                else if (resultNorm.startsWith('restricted'))   nonDeliverableCounts['restricted']++;
              }

              // ── Timezone tracking ─────────────────────────────────────────────
              if (ts) {
                const tz = extractTimezone(ts);
                if (tz) {
                  timezoneCounts[tz.offset] = (timezoneCounts[tz.offset] || 0) + 1;
                  const aid = row.account_id || 'Unknown';
                  if (!accountResultTimezones[aid]) accountResultTimezones[aid] = {};
                  accountResultTimezones[aid][tz.offset] = (accountResultTimezones[aid][tz.offset] || 0) + 1;
                }
              }

              // ── Config error tracking ─────────────────────────────────────────
              if (CONFIG_ERROR_RESULTS.has(resultNorm)) {
                const entry = configErrors[resultNorm];
                entry.total++;
                const caller = row.caller_number || 'Unknown';
                const msgId  = row.message_id   || 'Unknown';
                entry.byCallerNumber[caller] = (entry.byCallerNumber[caller] || 0) + 1;
                entry.byMessageId[msgId]     = (entry.byMessageId[msgId]     || 0) + 1;
              }

              // ── Global date range ─────────────────────────────────────────────
              if (pdOk) {
                if (!minDate || parsedDate < minDate) minDate = parsedDate;
                if (!maxDate || parsedDate > maxDate) maxDate = parsedDate;
              }

              // ── Build / update numberData entry ───────────────────────────────
              if (!numberData[num]) {
                numberData[num] = {
                  number: num, attempts: [], attemptIndex: 0,
                  consecutiveFailures: 0, totalAttempts: 0,
                  successCount: 0, unsuccessfulCount: 0, lastSuccessTimestamp: null,
                  messageIds: {}, callerNumbers: {}, accountIds: {},
                  hourCounts: new Uint16Array(24), dayOfWeekCounts: new Uint16Array(7),
                  backToBackIdentical: 0, lastMessageId: null,
                  currentSameStreak: 1, maxSameStreak: 0,
                  _fpMs: null, _lpMs: null  // first/last attempt ms-epoch (replaces string storage)
                };
              }
              const nd = numberData[num];
              if (isDelivery) { nd.totalAttempts++; nd.attemptIndex++; }

              // Track first/last attempt epoch ms — formatted to string at write time
              if (ts && pdOk) {
                const pdMs = parsedDate.getTime();
                if (nd._fpMs === null || pdMs < nd._fpMs) nd._fpMs = pdMs;
                if (nd._lpMs === null || pdMs > nd._lpMs) nd._lpMs = pdMs;
              }

              // Only push delivery attempts — non-deliverable rows (401, 403, 404, etc.) are
              // never used in sort / consecRuns / attemptStats and excluding them saves
              // hundreds of thousands of object allocations for large datasets.
              if (isDelivery) {
                nd.attempts.push({
                  ts:          pdOk ? parsedDate.getTime() : 0,
                  isSuccess,
                  attemptIndex: nd.attemptIndex   // already incremented above
                });
              }
              if (isDelivery) {
                const msgId = row.message_id || 'Unknown';
                nd.messageIds[msgId] = (nd.messageIds[msgId] || 0) + 1;
                if (nd.lastMessageId === msgId) {
                  nd.backToBackIdentical++;
                  nd.currentSameStreak++;
                } else {
                  nd.currentSameStreak = 1;
                }
                if (nd.currentSameStreak > nd.maxSameStreak) nd.maxSameStreak = nd.currentSameStreak;
                nd.lastMessageId = msgId;
                nd.callerNumbers[row.caller_number || 'Unknown'] = (nd.callerNumbers[row.caller_number || 'Unknown'] || 0) + 1;
                nd.accountIds[row.account_id || 'Unknown']       = (nd.accountIds[row.account_id || 'Unknown']       || 0) + 1;
                nd.hourCounts[localHour]++; nd.dayOfWeekCounts[localDow]++;
              }
              if (isSuccess) {
                nd.successCount++; nd.consecutiveFailures = 0;
                nd.attemptIndex = 0; nd.lastSuccessTimestamp = pdOk ? parsedDate.getTime() : null;
              } else if (isDelivery) {
                nd.unsuccessfulCount++; nd.consecutiveFailures++;
              }

              // ── Inline account / message / caller / time stats ─────────────────
              if (isDelivery) {
                const aId = row.account_id || 'Unknown';
                if (!accountStats[aId]) accountStats[aId] = { account_id: aId, successful: 0, unsuccessful: 0, total: 0, uniqueNumbers: 0, dayOfWeekCounts: new Uint16Array(7) };
                accountStats[aId].total++; accountStats[aId].dayOfWeekCounts[localDow]++;
                if (isSuccess) accountStats[aId].successful++; else accountStats[aId].unsuccessful++;

                const mId = row.message_id || 'Unknown';
                const mName = row.message_name || messageMap[mId]?.name || '';
                if (!messageStats[mId]) {
                  const txKey = `${aId}:${mId}`;
                  const txData = transcriptMap[txKey] || null;
                  messageStats[mId] = { message_id: mId, account_id: aId, message_name: mName, intent: txData?.intent || inferMessageIntent(mName), intent_summary: txData?.intent_summary || '', transcript: txData?.transcript || '', mentioned_phone: txData?.mentioned_phone || '', mentions_url: txData?.mentions_url || false, voice_append: false, successful: 0, unsuccessful: 0, total: 0, uniqueNumbers: 0, callerNumbers: {}, dayOfWeekCounts: new Uint16Array(7) };
                }
                if (row.voapps_voice_append) messageStats[mId].voice_append = true;
                messageStats[mId].total++; messageStats[mId].dayOfWeekCounts[localDow]++;
                if (isSuccess) messageStats[mId].successful++; else messageStats[mId].unsuccessful++;

                const cNum = row.caller_number || 'Unknown';
                const cName = row.caller_number_name || callerMap[cNum] || '';
                if (!callerStats[cNum]) callerStats[cNum] = { caller_number: cNum, caller_name: cName, successful: 0, unsuccessful: 0, total: 0, uniqueNumbers: 0, dayOfWeekCounts: new Uint16Array(7) };
                callerStats[cNum].total++; callerStats[cNum].dayOfWeekCounts[localDow]++;
                if (isSuccess) callerStats[cNum].successful++; else callerStats[cNum].unsuccessful++;
                messageStats[mId].callerNumbers[cNum] = (messageStats[mId].callerNumbers[cNum] || 0) + 1;

                globalHourlyStats[localHour].total++;
                if (isSuccess) globalHourlyStats[localHour].successful++; else globalHourlyStats[localHour].unsuccessful++;
                globalDayStats[localDow].total++;
                if (isSuccess) globalDayStats[localDow].successful++; else globalDayStats[localDow].unsuccessful++;

                // Daily trend tracking
                const dateKey = parsed ? parsed.localDateStr : (pdOk ? parsedDate.toISOString().slice(0, 10) : null);
                if (dateKey) {
                  const ds = dailyStats[dateKey] || (dailyStats[dateKey] = { attempts: 0, successes: 0 });
                  ds.attempts++;
                  if (isSuccess) ds.successes++;
                }
              }

              totalValidRows++;
              rowIdx++;
            }
          },
          complete: () => { log(`  -> Processed ${totalValidRows.toLocaleString()} rows so far`); resolve(); },
          error: reject
        });
      });
    }

    campaignTsMap.clear();
    log(`Combined ${totalValidRows.toLocaleString()} valid records from ${files.length} file(s)`);

    // ── Sort each number's attempts by timestamp, recompute order-dependent fields ──
    // Streaming may encounter rows out of chronological order within a number,
    // so we sort here before any downstream code relies on ordering.
    log('Sorting per-number attempts and recomputing attempt indices...');
    for (const num in numberData) {
      const nd = numberData[num];
      nd._fpMs = null; nd._lpMs = null; // null not delete — avoids V8 dictionary-mode transition
      if (nd.attempts.length > 1) {
        nd.attempts.sort((a, b) => a.ts - b.ts);
        let aidx = 0, consecFails = 0, lastSuccTs = null;
        for (const att of nd.attempts) {
          att.attemptIndex = ++aidx;
          if (att.isSuccess) { aidx = 0; consecFails = 0; lastSuccTs = att.ts; }
          else consecFails++;
        }
        nd.consecutiveFailures = consecFails;
        nd.attemptIndex = aidx;
        if (lastSuccTs !== null) nd.lastSuccessTimestamp = lastSuccTs; // stored as ms-epoch
      }
    }

    // ── Derive timezone variables from data collected during streaming ────────
    for (const [aid, offsets] of Object.entries(accountResultTimezones)) {
      let maxAcc = 0, best = null;
      for (const [off, cnt] of Object.entries(offsets)) { if (cnt > maxAcc) { maxAcc = cnt; best = off; } }
      accountMostCommonOffset[aid] = best;
    }
    let maxTzCnt = 0, bestOffset = null;
    for (const [off, cnt] of Object.entries(timezoneCounts)) { if (cnt > maxTzCnt) { maxTzCnt = cnt; bestOffset = off; } }
    detectedTimezone = bestOffset ? getTimezoneDisplayName({ offset: bestOffset }) : 'Unknown (timestamps may be UTC)';
    log(`Detected timezone: ${detectedTimezone} (from ${maxTzCnt.toLocaleString()} timestamps)`);
    timezoneDiscrepancies = detectTimezoneDiscrepancies(accountTimezones, accountMostCommonOffset);
    if (timezoneDiscrepancies.length > 0) log(`⚠️  Found ${timezoneDiscrepancies.length} timezone discrepancy(ies) between account settings and results`);
    fourteenDaysAgo = maxDate ? new Date(maxDate.getTime() - 14 * 24 * 60 * 60 * 1000) : null;

  } else {
    // ═══════════════════════════════════════════════════════════════════════════
    // ROW-ARRAY INPUT — existing flow with missing-timestamp pre-fill added.
    // ═══════════════════════════════════════════════════════════════════════════
    const csvRows = Array.isArray(csvInput) ? csvInput : [];
    log(`Processing ${csvRows.length.toLocaleString()} row objects`);
    if (csvRows.length === 0) throw new Error('No data found in CSV input');

    // ── Pre-fill missing timestamps using campaign proximity lookup ───────────
    {
      const ctsMap = new Map();
      for (let i = 0; i < csvRows.length; i++) {
        const ts = (csvRows[i].voapps_timestamp || '').trim();
        if (ts) {
          const cid = (csvRows[i].campaign_id || '').trim() || '__none__';
          if (!ctsMap.has(cid)) ctsMap.set(cid, []);
          ctsMap.get(cid).push({ idx: i, ts: formatTimestamp(ts) });
        }
      }
      for (let i = 0; i < csvRows.length; i++) {
        const ts = (csvRows[i].voapps_timestamp || '').trim();
        if (!ts) {
          // Mark as inferred — proximity-filled timestamps are used for sorting/display only,
          // not for delivery attempt counts or figures.
          csvRows[i]._tsOriginal = false;
          const cid = (csvRows[i].campaign_id || '').trim() || '__none__';
          const filled = findNearestTs(ctsMap.get(cid), i);
          if (filled) {
            csvRows[i].voapps_timestamp = filled;
          } else if (csvRows[i].target_date) {
            csvRows[i].voapps_timestamp = `${csvRows[i].target_date.trim()} 00:00:00 UTC`;
          }
        } else {
          csvRows[i]._tsOriginal = true;
          csvRows[i].voapps_timestamp = formatTimestamp(ts);
        }
      }
      ctsMap.clear();
    }

    // ── Filter and enrich ─────────────────────────────────────────────────────
    const validRows = csvRows.filter(row => row.number && row.voapps_result && String(row.number).replace(/\D/g, '').length >= 7);
    log(`Valid rows: ${validRows.length.toLocaleString()}`);

    for (const row of validRows) {
      const tz = extractTimezone(row.voapps_timestamp);
      if (tz) {
        timezoneCounts[tz.offset] = (timezoneCounts[tz.offset] || 0) + 1;
        const aid = row.account_id || 'Unknown';
        if (!accountResultTimezones[aid]) accountResultTimezones[aid] = {};
        accountResultTimezones[aid][tz.offset] = (accountResultTimezones[aid][tz.offset] || 0) + 1;
      }
    }
    for (const [aid, offsets] of Object.entries(accountResultTimezones)) {
      let maxAcc = 0, best = null;
      for (const [off, cnt] of Object.entries(offsets)) { if (cnt > maxAcc) { maxAcc = cnt; best = off; } }
      accountMostCommonOffset[aid] = best;
    }
    let maxTzCnt2 = 0, bestOffset2 = null;
    for (const [off, cnt] of Object.entries(timezoneCounts)) { if (cnt > maxTzCnt2) { maxTzCnt2 = cnt; bestOffset2 = off; } }
    detectedTimezone = bestOffset2 ? getTimezoneDisplayName({ offset: bestOffset2 }) : 'Unknown (timestamps may be UTC)';
    log(`Detected timezone: ${detectedTimezone} (from ${maxTzCnt2.toLocaleString()} timestamps)`);
    timezoneDiscrepancies = detectTimezoneDiscrepancies(accountTimezones, accountMostCommonOffset);
    if (timezoneDiscrepancies.length > 0) log(`⚠️  Found ${timezoneDiscrepancies.length} timezone discrepancy(ies) between account settings and results`);

    // Enrich, sort, and process in as few passes as possible to keep peak memory low.
    // Key savings vs. original:
    //   • row.parsedMs (epoch number, 8 bytes) replaces row.parsedDate (Date object, ~128 bytes)
    //     → saves ~168 MB for 1.4 M rows
    //   • Three separate passes merged into two (enrich → sort → build numberData)
    //   • validRows[i] nulled out as each row is consumed so GC can collect incrementally
    //   • Only delivery attempts pushed to nd.attempts (same as streaming path)

    log('Parsing timestamps and enriching data...');
    for (const row of validRows) {
      const parsed = parseTimestampLocal(row.voapps_timestamp);
      const _pd    = parsed ? parsed.utcDate : new Date(row.voapps_timestamp);
      const _ok    = _pd && !isNaN(_pd.getTime());
      row.parsedMs      = _ok ? _pd.getTime() : null; // epoch number — no long-lived Date on row
      row.localHour     = parsed ? parsed.localHour      : (_ok ? _pd.getHours() : 0);
      row.localDayOfWeek = parsed ? parsed.localDayOfWeek : (_ok ? _pd.getDay()   : 0);
      row.localDateStr  = parsed ? parsed.localDateStr   : (_ok ? _pd.toISOString().slice(0, 10) : null);
      row.voapps_result_normalized = String(row.voapps_result || '').trim().toLowerCase();
      row.isSuccess        = row.voapps_result_normalized === 'successfully delivered';
      // Only count as a delivery attempt if the row had a real original timestamp.
      // Proximity-inferred timestamps are used for sorting/display only.
      row.isDeliveryAttempt = !!row._tsOriginal &&
        DELIVERY_ATTEMPT_RESULTS.has(row.voapps_result_normalized);

      // Non-deliverable tracking
      if (!row.isDeliveryAttempt) {
        const rn   = row.voapps_result_normalized;
        const num2 = String(row.number).trim();
        if      (rn === 'not a wireless number') nonDeliverableCounts['not a wireless number']++;
        else if (rn === 'not a valid us number') {
          nonDeliverableCounts['not a valid us number']++;
          if (/^0+$/.test(num2.replace(/\D/g, ''))) notUSPlaceholderRows++;
        }
        else if (rn === 'duplicate number')   nonDeliverableCounts['duplicate number']++;
        else if (rn === 'undeliverable')      nonDeliverableCounts['undeliverable']++;
        else if (rn.startsWith('restricted')) nonDeliverableCounts['restricted']++;
      }
    }

    log('Sorting rows by date...');
    validRows.sort((a, b) => (a.parsedMs || 0) - (b.parsedMs || 0));

    // Single pass: build numberData + config errors + date range.
    // validRows[i] is nulled after each row is consumed so GC can reclaim row objects
    // incrementally rather than holding all 1.4 M in memory until the loop ends.
    log('Building number-level data with attempt indexing...');
    let minMs = null, maxMs = null;
    for (let _i = 0; _i < validRows.length; _i++) {
      const row = validRows[_i];
      validRows[_i] = null; // release reference — allows GC to collect this row object

      // Config error tracking (merged from separate loop)
      if (CONFIG_ERROR_RESULTS.has(row.voapps_result_normalized)) {
        const entry = configErrors[row.voapps_result_normalized];
        entry.total++;
        entry.byCallerNumber[row.caller_number || 'Unknown'] = (entry.byCallerNumber[row.caller_number || 'Unknown'] || 0) + 1;
        entry.byMessageId[row.message_id || 'Unknown']       = (entry.byMessageId[row.message_id || 'Unknown']       || 0) + 1;
      }

      // Date range (merged from separate loop)
      if (row.parsedMs) {
        if (minMs === null || row.parsedMs < minMs) minMs = row.parsedMs;
        if (maxMs === null || row.parsedMs > maxMs) maxMs = row.parsedMs;
      }

      const num = row.number;
      if (!numberData[num]) {
        numberData[num] = {
          number: num, attempts: [], attemptIndex: 0,
          consecutiveFailures: 0, totalAttempts: 0,
          successCount: 0, unsuccessfulCount: 0, lastSuccessTimestamp: null,
          messageIds: {}, callerNumbers: {}, accountIds: {},
          hourCounts: new Uint16Array(24), dayOfWeekCounts: new Uint16Array(7),
          backToBackIdentical: 0, lastMessageId: null,
          _fpMs: null, _lpMs: null  // first/last attempt ms-epoch
        };
      }
      const nd = numberData[num];
      if (row.isDeliveryAttempt) { nd.totalAttempts++; nd.attemptIndex++; }

      // Only push delivery attempts — non-deliverable rows never appear in consecRuns
      // or attemptStats, and excluding them saves substantial memory for large datasets.
      if (row.isDeliveryAttempt) {
        nd.attempts.push({
          ts:          row.parsedMs || 0,
          isSuccess:   row.isSuccess,
          attemptIndex: nd.attemptIndex   // already incremented above
        });
      }

      // Track first/last attempt epoch ms — formatted to string at write time
      if (row.parsedMs) {
        if (nd._fpMs === null || row.parsedMs < nd._fpMs) nd._fpMs = row.parsedMs;
        if (nd._lpMs === null || row.parsedMs > nd._lpMs) nd._lpMs = row.parsedMs;
      }
      if (row.isDeliveryAttempt) {
        const msgId = row.message_id || 'Unknown';
        nd.messageIds[msgId] = (nd.messageIds[msgId] || 0) + 1;
        if (nd.lastMessageId && nd.lastMessageId === msgId) nd.backToBackIdentical++;
        nd.lastMessageId = msgId;
        nd.callerNumbers[row.caller_number || 'Unknown'] = (nd.callerNumbers[row.caller_number || 'Unknown'] || 0) + 1;
        nd.accountIds[row.account_id || 'Unknown']       = (nd.accountIds[row.account_id || 'Unknown']       || 0) + 1;
        nd.hourCounts[row.localHour]++; nd.dayOfWeekCounts[row.localDayOfWeek]++;
      }
      if (row.isSuccess) {
        nd.successCount++; nd.consecutiveFailures = 0;
        nd.attemptIndex = 0; nd.lastSuccessTimestamp = row.parsedMs || null;
      } else if (row.isDeliveryAttempt) {
        nd.unsuccessfulCount++; nd.consecutiveFailures++;
      }

      // ── Inline account / message / caller / time stats (same as file-path path) ──
      if (row.isDeliveryAttempt) {
        const num2 = row.number; // already available as `num` above
        const aId = row.account_id || 'Unknown';
        if (!accountStats[aId]) accountStats[aId] = { account_id: aId, successful: 0, unsuccessful: 0, total: 0, uniqueNumbers: 0, dayOfWeekCounts: new Uint16Array(7) };
        accountStats[aId].total++; accountStats[aId].dayOfWeekCounts[row.localDayOfWeek]++;
        if (row.isSuccess) accountStats[aId].successful++; else accountStats[aId].unsuccessful++;

        const mId = row.message_id || 'Unknown';
        const mName = row.message_name || messageMap[mId]?.name || '';
        if (!messageStats[mId]) {
          const txKey = `${aId}:${mId}`;
          const txData = transcriptMap[txKey] || null;
          messageStats[mId] = { message_id: mId, account_id: aId, message_name: mName, intent: txData?.intent || inferMessageIntent(mName), intent_summary: txData?.intent_summary || '', transcript: txData?.transcript || '', mentioned_phone: txData?.mentioned_phone || '', mentions_url: txData?.mentions_url || false, voice_append: false, successful: 0, unsuccessful: 0, total: 0, uniqueNumbers: 0, callerNumbers: {}, dayOfWeekCounts: new Uint16Array(7) };
        }
        if (row.voapps_voice_append) messageStats[mId].voice_append = true;
        messageStats[mId].total++; messageStats[mId].dayOfWeekCounts[row.localDayOfWeek]++;
        if (row.isSuccess) messageStats[mId].successful++; else messageStats[mId].unsuccessful++;

        const cNum = row.caller_number || 'Unknown';
        const cName = row.caller_number_name || callerMap[cNum] || '';
        if (!callerStats[cNum]) callerStats[cNum] = { caller_number: cNum, caller_name: cName, successful: 0, unsuccessful: 0, total: 0, uniqueNumbers: 0, dayOfWeekCounts: new Uint16Array(7) };
        callerStats[cNum].total++; callerStats[cNum].dayOfWeekCounts[row.localDayOfWeek]++;
        if (row.isSuccess) callerStats[cNum].successful++; else callerStats[cNum].unsuccessful++;
        messageStats[mId].callerNumbers[cNum] = (messageStats[mId].callerNumbers[cNum] || 0) + 1;

        globalHourlyStats[row.localHour].total++;
        if (row.isSuccess) globalHourlyStats[row.localHour].successful++; else globalHourlyStats[row.localHour].unsuccessful++;
        globalDayStats[row.localDayOfWeek].total++;
        if (row.isSuccess) globalDayStats[row.localDayOfWeek].successful++; else globalDayStats[row.localDayOfWeek].unsuccessful++;

        // Daily trend tracking
        if (row.localDateStr) {
          const ds = dailyStats[row.localDateStr] || (dailyStats[row.localDateStr] = { attempts: 0, successes: 0 });
          ds.attempts++;
          if (row.isSuccess) ds.successes++;
        }
      }

      totalValidRows++;
    }
    // Convert epoch extremes → Date objects used by shared code below
    minDate         = minMs ? new Date(minMs) : null;
    maxDate         = maxMs ? new Date(maxMs) : null;
    fourteenDaysAgo = maxDate ? new Date(maxDate.getTime() - 14 * 24 * 60 * 60 * 1000) : null;

    // Free raw row arrays — numberData holds all per-number state
    csvRows.length = 0;
    validRows.length = 0;
  }

  const hasConfigErrors = Object.values(configErrors).some(e => e.total > 0);
  const uniqueNumbers = Object.keys(numberData).length;
  log(`Analyzed ${uniqueNumbers.toLocaleString()} unique numbers`);

  // ============================================================================
  // CALCULATE SUCCESS PROBABILITY BY ATTEMPT INDEX
  // ============================================================================

  log('Calculating success probability by attempt index...');
  const attemptStats = {};  // attemptIndex -> { successful, total }

  for (const num in numberData) {
    for (const attempt of numberData[num].attempts) {
      const idx = Math.min(attempt.attemptIndex, 10);  // Cap at 10 for grouping
      if (!attemptStats[idx]) {
        attemptStats[idx] = { successful: 0, total: 0 };
      }
      attemptStats[idx].total++;
      if (attempt.isSuccess) {
        attemptStats[idx].successful++;
      }
    }
  }

  // Calculate probabilities
  const decayCurve = [];
  for (let i = 1; i <= 10; i++) {
    const stats = attemptStats[i] || { successful: 0, total: 0 };
    const prob = stats.total > 0 ? stats.successful / stats.total : 0;
    decayCurve.push({
      attemptIndex: i === 10 ? '10+' : i,
      total: stats.total,
      successful: stats.successful,
      probability: prob
    });
  }

  // ============================================================================
  // COMPUTE DELIVERY CADENCE (interval between consecutive attempts per number)
  // ============================================================================
  log('Computing delivery cadence stats...');
  let cadenceSingleTouch = 0;
  let cadenceBucket_sameDay = 0; // < 1 day
  let cadenceBucket_1to2    = 0; // 1–2 days
  let cadenceBucket_3to5    = 0; // 3–5 days
  let cadenceBucket_6to10   = 0; // 6–10 days
  let cadenceBucket_11to15  = 0; // 11–15 days
  let cadenceBucket_16to30  = 0; // 16–30 days
  let cadenceBucket_over30  = 0; // > 30 days
  const _cadenceMedians = [];
  for (const num in numberData) {
    const atts = numberData[num].attempts;
    if (atts.length === 0) continue;
    if (atts.length === 1) { cadenceSingleTouch++; continue; }
    // Ensure time-ordered — streaming path already sorted in place; row-array path may not be
    if (atts[0].ts > atts[atts.length - 1].ts) atts.sort((a, b) => a.ts - b.ts);
    const intervals = [];
    for (let i = 1; i < atts.length; i++) {
      const diff = (atts[i].ts - atts[i - 1].ts) / 86400000; // ms → days
      if (diff >= 0) intervals.push(diff);
    }
    if (intervals.length === 0) { cadenceSingleTouch++; continue; }
    intervals.sort((a, b) => a - b);
    const cMid = Math.floor(intervals.length / 2);
    const med = intervals.length % 2 === 0
      ? (intervals[cMid - 1] + intervals[cMid]) / 2
      : intervals[cMid];
    _cadenceMedians.push(med);
    if      (med < 1)   cadenceBucket_sameDay++;
    else if (med <= 2)  cadenceBucket_1to2++;
    else if (med <= 5)  cadenceBucket_3to5++;
    else if (med <= 10) cadenceBucket_6to10++;
    else if (med <= 15) cadenceBucket_11to15++;
    else if (med <= 30) cadenceBucket_16to30++;
    else                cadenceBucket_over30++;
  }
  _cadenceMedians.sort((a, b) => a - b);
  let cadenceOverallMedian = null;
  if (_cadenceMedians.length > 0) {
    const cMid = Math.floor(_cadenceMedians.length / 2);
    cadenceOverallMedian = _cadenceMedians.length % 2 === 0
      ? (_cadenceMedians[cMid - 1] + _cadenceMedians[cMid]) / 2
      : _cadenceMedians[cMid];
  }
  _cadenceMedians.length = 0; // free
  const cadenceMultiTouchCount = cadenceBucket_sameDay + cadenceBucket_1to2 +
    cadenceBucket_3to5 + cadenceBucket_6to10 + cadenceBucket_11to15 +
    cadenceBucket_16to30 + cadenceBucket_over30;

  // ============================================================================
  // BUILD ACCOUNT AND MESSAGE LEVEL STATS WITH DAY-OF-WEEK ANALYSIS
  // ============================================================================
  // Stats are built inline during row processing above — no separate pass needed.
  log('Building account and message level stats...');


  // Check for day-of-week patterns in accounts and messages
  const accountDayRecommendations = [];
  const messageDayRecommendations = [];

  for (const accountId in accountStats) {
    const pattern = getDayUsagePattern(accountStats[accountId].dayOfWeekCounts);
    if (pattern.limited) {
      accountDayRecommendations.push({
        type: 'Account',
        id: accountId,
        days: pattern.days.join(', '),
        recommendation: pattern.recommendation
      });
    }
  }

  for (const msgId in messageStats) {
    const pattern = getDayUsagePattern(messageStats[msgId].dayOfWeekCounts);
    if (pattern.limited) {
      messageDayRecommendations.push({
        type: 'Message',
        id: msgId,
        name: messageStats[msgId].message_name,
        days: pattern.days.join(', '),
        recommendation: pattern.recommendation
      });
    }
  }

  // ============================================================================
  // CLASSIFY TN HEALTH AND CALCULATE VARIABILITY SCORES
  // ============================================================================

  log('Classifying TN health and calculating variability scores...');

  let healthyCount = 0, toxicCount = 0, neverDeliveredCount = 0;
  const numberSummaryArray = [];

  for (const num in numberData) {
    const nd = numberData[num];
    // Skip numbers that only ever appeared in non-deliverable rows (e.g. 0000000000 / 403-only).
    // They have no delivery attempt data to analyze and would pollute the summary.
    if (nd.totalAttempts === 0) continue;
    const successRate = nd.totalAttempts > 0 ? nd.successCount / nd.totalAttempts : 0;

    // Check for recent success (within 14 days of max date)
    const recentSuccess = nd.lastSuccessTimestamp && fourteenDaysAgo &&
                          nd.lastSuccessTimestamp >= fourteenDaysAgo;

    // Never Delivered flag
    const neverDelivered = nd.successCount === 0;
    if (neverDelivered) neverDeliveredCount++;

    // TN Health Classification
    const tnHealth = classifyTNHealth(successRate, nd.consecutiveFailures, nd.totalAttempts, recentSuccess);
    if (tnHealth === 'Healthy') healthyCount++;
    else if (tnHealth === 'Delivery Unlikely') toxicCount++;

    // Calculate variability metrics
    const uniqueMessages = Object.keys(nd.messageIds).length;
    const uniqueCallers = Object.keys(nd.callerNumbers).length;

    // Find top message
    let topMsgId = 'Unknown', topMsgCount = 0;
    for (const msgId in nd.messageIds) {
      if (nd.messageIds[msgId] > topMsgCount) {
        topMsgCount = nd.messageIds[msgId];
        topMsgId = msgId;
      }
    }
    const topMsgPct = nd.totalAttempts > 0 ? topMsgCount / nd.totalAttempts : 0;
    const topMsgName = nd.attempts.find(a => a.message_id === topMsgId)?.message_name ||
                       messageMap[topMsgId]?.name || '';

    // Find top caller
    let topCallerNum = 'Unknown', topCallerCount = 0;
    for (const callerNum in nd.callerNumbers) {
      if (nd.callerNumbers[callerNum] > topCallerCount) {
        topCallerCount = nd.callerNumbers[callerNum];
        topCallerNum = callerNum;
      }
    }
    const topCallerPct = nd.totalAttempts > 0 ? topCallerCount / nd.totalAttempts : 0;
    const topCallerName = nd.attempts.find(a => a.caller_number === topCallerNum)?.caller_number_name ||
                          callerMap[topCallerNum] || '';

    // Day of week entropy (max entropy = log2(7) ~ 2.807)
    const dayEntropy = calculateEntropy(nd.dayOfWeekCounts);
    const maxDayEntropy = Math.log2(7);
    const dayEntropyNormalized = maxDayEntropy > 0 ? dayEntropy / maxDayEntropy : 0;

    // Hour variance
    const hourEntropy = calculateEntropy(nd.hourCounts);
    const maxHourEntropy = Math.log2(24);
    const hourEntropyNormalized = maxHourEntropy > 0 ? hourEntropy / maxHourEntropy : 0;

    // Message diversity score (lower top message % = better)
    const msgDiversityScore = uniqueMessages > 1 ? (1 - topMsgPct) * 100 : 0;

    // Caller diversity score
    const callerDiversityScore = uniqueCallers > 1 ? (1 - topCallerPct) * 100 : 0;

    // Back-to-back penalty (0-1, lower is better)
    const backToBackRatio = nd.totalAttempts > 1 ? nd.backToBackIdentical / (nd.totalAttempts - 1) : 0;

    // Calculate variability score only if there are multiple attempts
    let variabilityScore = 0;
    if (nd.totalAttempts > 1) {
      // Base score components
      const msgComponent = msgDiversityScore * 0.30;
      const callerComponent = callerDiversityScore * 0.20;
      const dayComponent = dayEntropyNormalized * 100 * 0.25;
      const hourComponent = hourEntropyNormalized * 100 * 0.15;
      const noBackToBackBonus = (1 - backToBackRatio) * 100 * 0.10;

      variabilityScore = Math.max(0, Math.min(100, Math.round(
        msgComponent + callerComponent + dayComponent + hourComponent + noBackToBackBonus
      )));
    } else {
      // Single attempt - score based on diversity potential
      variabilityScore = 50; // Neutral score for single attempts
    }

    // Day distribution string
    const dayDistribution = formatDayDistribution(nd.dayOfWeekCounts, nd.totalAttempts);

    // First and last attempt — format epoch ms to UTC timestamp string at write time
    const validFirstAttempt = nd._fpMs ? epochToTimestamp(nd._fpMs) : null;
    const validLastAttempt  = nd._lpMs ? epochToTimestamp(nd._lpMs) : null;

    // Infer intent from top message (use AI transcript if available, fall back to name-based)
    const messageIntent = messageStats[topMsgId]?.intent || inferMessageIntent(topMsgName);

    numberSummaryArray.push({
      number: num,
      totalAttempts: nd.totalAttempts,
      successful: nd.successCount,
      unsuccessful: nd.unsuccessfulCount,
      successRate: successRate,
      attemptIndex: nd.attemptIndex,
      consecutiveFailures: nd.consecutiveFailures,
      tnHealth: tnHealth,
      neverDelivered: neverDelivered,
      variabilityScore: variabilityScore,
      topMsgId: topMsgId,
      topMsgName: topMsgName,
      topMsgPct: topMsgPct,
      uniqueMsgCount: uniqueMessages,
      topCallerNum: topCallerNum,
      topCallerName: topCallerName,
      topCallerPct: topCallerPct,
      uniqueCallerCount: uniqueCallers,
      dayDistribution: dayDistribution,
      dayEntropy: dayEntropyNormalized,
      hourEntropy: hourEntropyNormalized,
      backToBackIdentical: nd.backToBackIdentical,
      maxSameStreak: nd.maxSameStreak,
      messageIntent: messageIntent,
      firstAttempt: validFirstAttempt,
      lastAttempt: validLastAttempt,
      lastSuccessTimestamp: nd.lastSuccessTimestamp
    });
  }

  // Sort by total attempts (descending)
  numberSummaryArray.sort((a, b) => b.totalAttempts - a.totalAttempts);

  // Calculate list health percentages
  const healthyPct = (healthyCount / uniqueNumbers) * 100;
  const toxicPct = (toxicCount / uniqueNumbers) * 100;
  const neverDeliveredPct = (neverDeliveredCount / uniqueNumbers) * 100;
  const listGrade = calculateListGrade(healthyPct, toxicPct, neverDeliveredPct);

  log(`TN Health: Healthy=${healthyCount.toLocaleString()}, Delivery Unlikely=${toxicCount.toLocaleString()}, Never Delivered=${neverDeliveredCount.toLocaleString()}`);
  log(`List Grade: ${listGrade}`);

  // ============================================================================
  // BUILD CONSECUTIVE UNSUCCESSFUL RUNS
  // ============================================================================

  log('Building consecutive unsuccessful runs...');
  const consecRuns = [];
  // O(1) lookups instead of O(n) .find() calls inside loops over 271K+ numbers
  const numSummaryMap = new Map(numberSummaryArray.map(ns => [ns.number, ns]));
  const inConsecRuns  = new Set(); // tracks numbers already added to consecRuns

  for (const num in numberData) {
    const nd = numberData[num];
    const attempts = nd.attempts;
    let currentRun = [];

    for (const attempt of attempts) {
      if (!attempt.isSuccess) {
        currentRun.push(attempt);
      } else {
        // Check if the run meets the minimum consecutive failure count
        if (currentRun.length >= minConsecUnsuccessful) {
          const runStart = currentRun[0].ts ? new Date(currentRun[0].ts) : null;
          const runEnd = currentRun[currentRun.length - 1].ts ? new Date(currentRun[currentRun.length - 1].ts) : null;
          const spanDays = runStart && runEnd ? (runEnd - runStart) / (1000 * 60 * 60 * 24) : 0;
          const ns = numSummaryMap.get(num);
          const runHealth = ns?.tnHealth || 'Healthy';
          consecRuns.push({
            number: num,
            count: currentRun.length,
            runStart, runEnd, spanDays,
            tnHealth: runHealth
          });
          inConsecRuns.add(num);
        }
        currentRun = [];
      }
    }

    // Check final run (if the number ends with consecutive failures)
    if (currentRun.length >= minConsecUnsuccessful) {
      const runStart = currentRun[0].ts ? new Date(currentRun[0].ts) : null;
      const runEnd = currentRun[currentRun.length - 1].ts ? new Date(currentRun[currentRun.length - 1].ts) : null;
      const spanDays = runStart && runEnd ? (runEnd - runStart) / (1000 * 60 * 60 * 24) : 0;
      const ns = numSummaryMap.get(num);
      const runHealth2 = ns?.tnHealth || 'Healthy';
      consecRuns.push({
        number: num,
        count: currentRun.length,
        runStart, runEnd, spanDays,
        tnHealth: runHealth2
      });
      inConsecRuns.add(num);
    }
  }

  // Also add all numbers with current consecutive failures >= threshold (that aren't already included)
  for (const ns of numberSummaryArray) {
    if (ns.consecutiveFailures >= minConsecUnsuccessful) {
      if (!inConsecRuns.has(ns.number)) {
        const nd = numberData[ns.number];
        const recentFailures = nd.attempts.slice(-ns.consecutiveFailures);
        const runStart = recentFailures[0]?.ts ? new Date(recentFailures[0].ts) : null;
        const runEnd = recentFailures[recentFailures.length - 1]?.ts ? new Date(recentFailures[recentFailures.length - 1].ts) : null;
        const spanDays = runStart && runEnd ? (runEnd - runStart) / (1000 * 60 * 60 * 24) : 0;
        consecRuns.push({
          number: ns.number,
          count: ns.consecutiveFailures,
          runStart, runEnd, spanDays,
          tnHealth: ns.tnHealth
        });
        inConsecRuns.add(ns.number);
      }
    }
  }

  consecRuns.sort((a, b) => b.count - a.count);
  // Only "Delivery Unlikely" numbers belong on the Suppression Candidates tab.
  const suppressionRuns = consecRuns.filter(r => r.tnHealth === 'Delivery Unlikely' && r.spanDays >= minRunSpanDays);
  log(`  Found ${consecRuns.length.toLocaleString()} consecutive unsuccessful patterns (${suppressionRuns.length.toLocaleString()} Delivery Unlikely → Suppression Candidates tab)`);

  // Free attempt arrays — all stats now extracted, no longer needed
  for (const num in numberData) {
    numberData[num].attempts = null;
  }

  // ── Count unique numbers per account / message / caller ──────────────────────
  // uniqueNumbers was tracked as a plain counter (0) during streaming to avoid
  // large Sets.  nd.messageIds / callerNumbers / accountIds are already keyed by
  // phone number (one entry per unique number per entity), so iterating them
  // gives accurate unique-number counts with O(n) time and O(1) extra memory.
  log('Counting unique numbers per account / message / caller...');
  for (const num in numberData) {
    const nd = numberData[num];
    for (const msgId    in nd.messageIds)     { if (messageStats[msgId])     messageStats[msgId].uniqueNumbers++;     }
    for (const callerNum in nd.callerNumbers) { if (callerStats[callerNum])  callerStats[callerNum].uniqueNumbers++;  }
    for (const acctId   in nd.accountIds)     { if (accountStats[acctId])    accountStats[acctId].uniqueNumbers++;    }
    // Free per-number dictionaries — no longer needed after this pass
    nd.messageIds = null; nd.callerNumbers = null; nd.accountIds = null;
  }
  // numberData itself is no longer needed — all per-number stats are in numberSummaryArray
  // Freeing it here recovers ~50–100 MB before Excel generation begins.
  // (overallSuccessRate uses numberSummaryArray which is already fully built)
  for (const num in numberData) delete numberData[num];

  // ============================================================================
  // PRE-COMPUTE FILTERED SETS (used in Executive Summary + detail tabs)
  // Analysis always runs on ALL numbers; these subsets control what rows appear
  // in detail tabs so that large datasets don't generate unmanageable sheets.
  // ============================================================================

  // Detail tab data — only computed when includeDetailTabs is enabled.
  // When disabled, empty arrays are used so sheet-creation blocks (guarded by
  // the same flag) never reference uninitialized variables.
  const MAX_DETAIL_ROWS = 100_000;
  const filteredHealth = includeDetailTabs
    ? numberSummaryArray
        .filter(ns => ns.tnHealth === 'Delivery Unlikely')
        .sort((a, b) => (a.tnHealth === 'Delivery Unlikely' ? 0 : 1) - (b.tnHealth === 'Delivery Unlikely' ? 0 : 1))
    : [];

  const filteredVariability = includeDetailTabs
    ? numberSummaryArray
        .filter(ns => ns.totalAttempts > 1 && ns.variabilityScore < 60)
        .sort((a, b) => a.variabilityScore - b.variabilityScore)
    : [];

  // Number Summary tab: fail at least one criterion.
  // Variability flag only applies to multi-attempt numbers (same reasoning as above).
  const filteredSummary = includeDetailTabs
    ? numberSummaryArray.filter(ns =>
        ns.tnHealth === 'Delivery Unlikely' ||
        (ns.totalAttempts > 1 && ns.variabilityScore < 60)
      )
    : [];

  if (includeDetailTabs) {
    // Capture true pre-cap counts — used in key metrics and log messages below.
    const healthTotalCount  = filteredHealth.length;
    const varTotalCount     = filteredVariability.length;
    const summaryTotalCount = filteredSummary.length;

    // Cap each detail tab at MAX_DETAIL_ROWS to bound ExcelJS Cell accumulation.
    // ExcelJS holds all rows in memory as Cell objects (~3,200 bytes/row × columns).
    // 3 tabs × 100K rows × ~3,200 bytes ≈ 960 MB — fits under the effective ~2 GB heap limit.
    // Truncating via .length = N releases array slots above N immediately (GC-eligible).
    if (filteredHealth.length      > MAX_DETAIL_ROWS) filteredHealth.length      = MAX_DETAIL_ROWS;
    if (filteredVariability.length > MAX_DETAIL_ROWS) filteredVariability.length = MAX_DETAIL_ROWS;
    if (filteredSummary.length     > MAX_DETAIL_ROWS) filteredSummary.length     = MAX_DETAIL_ROWS;

    // Helper: "100,000 of 342,819 (capped at 100,000)" or just "18,432"
    const fmtCapped = (shown, total) =>
      total > MAX_DETAIL_ROWS
        ? `${shown.toLocaleString()} of ${total.toLocaleString()} (capped at ${MAX_DETAIL_ROWS.toLocaleString()})`
        : shown.toLocaleString();

    log(`\nDetail tab filters (analysis ran on all ${numberSummaryArray.length.toLocaleString()} numbers):`);
    log(`  TN Health tab:          ${fmtCapped(filteredHealth.length, healthTotalCount)} numbers (Delivery Unlikely)`);
    log(`  Variability tab:        ${fmtCapped(filteredVariability.length, varTotalCount)} numbers (score < 60)`);
    log(`  Number Summary tab:     ${fmtCapped(filteredSummary.length, summaryTotalCount)} numbers (any flag)`);
  }

  // ============================================================================
  // PRE-COMPUTE ALL METRICS THAT NEED THE FULL numberSummaryArray, THEN FREE IT
  // ============================================================================
  // Do this in one single pass to avoid multiple scans over potentially 1M+ objects,
  // then immediately release the large array so GC can recover memory before ExcelJS
  // starts building the workbook (which can peak at 1-2 GB of XML/buffer data).
  log('Pre-computing aggregates from full number set...');
  const totalUniqueInSummary = numberSummaryArray.length;
  let _totalAttempts = 0, _totalSuccess = 0, _sumVariability = 0;
  let _backToBackIssueCount = 0, _lowDayVarietyCount = 0, _flaggedCount = 0;
  let _streak2 = 0, _streak3 = 0, _streak4 = 0, _streak5plus = 0;
  for (const ns of numberSummaryArray) {
    _totalAttempts        += ns.totalAttempts;
    _totalSuccess         += ns.successful;   // stored as 'successful' in the push, not 'successCount'
    _sumVariability       += ns.variabilityScore;
    if (ns.backToBackIdentical > 2) _backToBackIssueCount++;
    if (ns.dayEntropy < 0.3 && ns.totalAttempts > 2) _lowDayVarietyCount++;
    // Count flagged numbers for exec summary metrics (always needed, even when detail tabs are off)
    if (ns.tnHealth === 'Delivery Unlikely' ||
        (ns.totalAttempts > 1 && ns.variabilityScore < 60)) _flaggedCount++;
    // Streak buckets
    if (ns.maxSameStreak >= 2) _streak2++;
    if (ns.maxSameStreak >= 3) _streak3++;
    if (ns.maxSameStreak >= 4) _streak4++;
    if (ns.maxSameStreak >= 5) _streak5plus++;
  }
  const streak2 = _streak2, streak3 = _streak3, streak4 = _streak4, streak5plus = _streak5plus;
  const totalAttempts      = _totalAttempts;
  const avgVariability     = totalUniqueInSummary > 0 ? _sumVariability / totalUniqueInSummary : 0;
  const overallSuccessRate = _totalAttempts > 0 ? (_totalSuccess / _totalAttempts * 100) : 0;
  const backToBackIssues   = _backToBackIssueCount;
  const lowDayVariety      = _lowDayVarietyCount;
  const flaggedCount       = _flaggedCount;  // pre-cap — reflects all flagged numbers, not just the capped subset written to Excel
  const flaggedPct         = totalUniqueInSummary > 0 ? (flaggedCount / totalUniqueInSummary * 100) : 0;

  // Release unflagged ns objects — filteredX arrays keep flagged ones alive;
  // Healthy/high-variability entries not in any tab become GC-eligible here.
  numSummaryMap.clear();
  numberSummaryArray.length = 0;
  log(`  Released ${(totalUniqueInSummary - flaggedCount).toLocaleString()} unflagged number entries from memory`);

  // ── Pre-computed summary values used in both Excel and PPTX ─────────────────
  const agentHoursSaved = Math.round(_totalSuccess * 3 / 60);

  // Best Next Action — single highest-impact recommendation
  const _cadTotalBNA = cadenceSingleTouch + cadenceMultiTouchCount;
  const _stPctBNA    = _cadTotalBNA > 0 ? cadenceSingleTouch / _cadTotalBNA : 0;
  const _longCadPct  = cadenceMultiTouchCount > 0 ? (cadenceBucket_over30 + cadenceBucket_16to30) / cadenceMultiTouchCount : 0;
  let bestNextAction;
  if (_stPctBNA >= 0.4) {
    bestNextAction = `${cadenceSingleTouch.toLocaleString()} numbers (${(_stPctBNA * 100).toFixed(0)}% of the list) received only one DDVM during this period. Scheduling a follow-up campaign at a 3–10 day interval is the single highest-impact next step — consumers who didn't act on the first touch often engage on the second or third. This is existing-list volume with no new sourcing required.`;
  } else if (_stPctBNA >= 0.2) {
    bestNextAction = `${cadenceSingleTouch.toLocaleString()} numbers (${(_stPctBNA * 100).toFixed(0)}%) were contacted only once. Adding a follow-up campaign at a 3–10 day interval would put additional touches on a meaningful portion of the list — typically the easiest place to find incremental results without expanding the contact pool.`;
  } else if (_longCadPct > 0.3) {
    const _longCount = cadenceBucket_over30 + cadenceBucket_16to30;
    bestNextAction = `${_longCount.toLocaleString()} re-attempted numbers are on a 16+ day cadence. Tightening to the 3–10 day ideal keeps outreach timely and relevant — consumers are more likely to respond when the next touch arrives before their situation changes.`;
  } else if (overallSuccessRate >= 70) {
    bestNextAction = `Campaign is delivering at ${overallSuccessRate.toFixed(1)}% — well above the industry midpoint. This is a strong foundation for expanding outreach: adding more numbers or increasing frequency on the existing list should yield proportional returns with low risk.`;
  } else {
    bestNextAction = `Focus next on the ${healthyCount.toLocaleString()} numbers (${(healthyCount / Math.max(uniqueNumbers, 1) * 100).toFixed(0)}%) currently connecting well. Concentrating volume on this reachable segment while monitoring the rest will improve overall campaign efficiency and ROI.`;
  }

  // ============================================================================
  // CREATE WORKBOOK
  // ============================================================================

  const workbook = new ExcelJS.Workbook();
  workbook.creator = `VoApps Number Analysis and Delivery Intelligence Report v${VERSION}`;
  workbook.created = new Date();

  // VoApps-branded styles
  const headerStyle = {
    font: { bold: true, size: 14, color: { argb: 'FFFFFF' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: VOAPPS_DARK_NAVY } },
    alignment: { vertical: 'middle', horizontal: 'left', wrapText: true }
  };

  const sectionHeaderStyle = {
    font: { bold: true, size: 12, color: { argb: 'FFFFFF' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: VOAPPS_PURPLE } },
    alignment: { vertical: 'middle', horizontal: 'left' }
  };

  const tableHeaderStyle = {
    font: { bold: true, size: 11, color: { argb: 'FFFFFF' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: VOAPPS_DARK_NAVY } },
    alignment: { vertical: 'middle', horizontal: 'center' }
  };

  const contentStyle = {
    font: { size: 11, color: { argb: VOAPPS_CHARCOAL } },
    alignment: { vertical: 'top', horizontal: 'left', wrapText: true }
  };

  const warningStyle = {
    font: { size: 11, color: { argb: 'C65911' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: VOAPPS_BLUSH } }
  };

  const successStyle = {
    font: { size: 11, color: { argb: '375623' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'C6EFCE' } }
  };

  // ========================================
  // TAB 1: EXECUTIVE SUMMARY
  // ========================================

  log('Creating Executive Summary tab...');

  const execSheet = workbook.addWorksheet('Executive Summary');

  // Title
  execSheet.mergeCells('A1:C1');
  execSheet.getCell('A1').value = `Number Analysis and Delivery Intelligence Report v${VERSION}`;
  execSheet.getCell('A1').style = headerStyle;
  execSheet.getRow(1).height = 35;

  execSheet.mergeCells('A2:C2');
  execSheet.getCell('A2').value =
    'This report analyzes DDVM delivery patterns across campaigns in a selected date range to identify which phone numbers are ' +
    'receiving messages successfully, which are consistently failing, and where strategy changes — including ' +
    'message rotation, caller number diversity, and retry limits — can improve delivery outcomes and maximize effectiveness of DirectDrop Voicemail.';
  execSheet.getCell('A2').font = { italic: true, size: 10, color: { argb: 'FF555555' } };
  execSheet.getCell('A2').alignment = { wrapText: true };
  execSheet.getRow(2).height = 42;

  let row = 4;   // row 3 is blank spacer, row 4 starts Key Metrics

  // Key Metrics Box
  execSheet.mergeCells(`A${row}:C${row}`);
  execSheet.getCell(`A${row}`).value = 'Key Metrics';
  execSheet.getCell(`A${row}`).style = sectionHeaderStyle;
  row++;

  // Grade-specific advice for column C — value-focused framing
  const listGradeAdvice = listGrade === 'A'
    ? 'Strong list health — your numbers are connecting well. Maintaining this grade protects delivery speed and keeps cost-per-contact low. Monthly review of Delivery Unlikely numbers will keep performance at this level.'
    : listGrade === 'B'
    ? 'Good list health with clear room to improve. Removing numbers that are no longer reachable sharpens the active list, which typically raises overall delivery rate and reduces wasted attempts — a straightforward path to Grade A performance.'
    : listGrade === 'C'
    ? 'Moderate list health — there is meaningful opportunity here. Trimming numbers that have never delivered and suppressing persistent failures will concentrate your campaign spend on numbers that actually connect, improving ROI on every drop.'
    : 'The list has a high proportion of numbers that are not connecting. Focusing spend on the reachable segment — by removing numbers that have repeatedly failed — will dramatically increase delivered % and reduce cost-per-contact for the numbers that matter.';

  // All aggregates pre-computed and numberSummaryArray freed above
  const keyMetrics = [
    ['Total DDVM Attempts', totalAttempts.toLocaleString()],
    ['Unique Phone Numbers', uniqueNumbers.toLocaleString()],
    ['Delivered %', `${overallSuccessRate.toFixed(1)}%`,
      'The percentage of all DDVM delivery attempts that resulted in a successful voicemail drop (result code 200 | Successfully delivered). Each attempt on a number counts separately — a number attempted 3 times and delivered once counts as 1 success out of 3 attempts.'],
    ['Never Delivered %', `${neverDeliveredPct.toFixed(1)}%`,
      'The percentage of unique phone numbers that never received a successful voicemail delivery in this date range. Removing these from the active list redirects campaign spend entirely toward numbers that can and do connect.'],
    ['Average Variability Score', `${avgVariability.toFixed(0)}/100`,
      '0–100 composite score measuring call pattern diversity — message rotation, caller variety, time-of-day spread, and day-of-week distribution. Higher scores correlate with better deliverability and callback rates. Scores below 60 indicate patterns where increased variety is likely to improve results.'],
    ['List Quality Grade', listGrade, listGradeAdvice],
    ['Numbers Flagged in Detail Tabs', `${flaggedCount.toLocaleString()} of ${totalUniqueInSummary.toLocaleString()} (${flaggedPct.toFixed(1)}%) — Delivery Unlikely or variability < 60`,
      'Any number failing at least one threshold — classified Delivery Unlikely by TN Health, OR variability score below 60. A Healthy number with poor call diversity is still flagged. See TN Health and Variability Analysis tabs for the full breakdown (if enabled).'],
    ['Date Range', `${formatDate(minDate)} - ${formatDate(maxDate)}`],
    ['Timezone', detectedTimezone],
    ['Agent Hours Saved (est.)', `${agentHoursSaved.toLocaleString()} hrs`,
      `Estimated agent capacity freed by DDVM. Based on ${_totalSuccess.toLocaleString()} successful deliveries × 3 min avg manual voicemail handle time (dial + wait + message). Use the ROI Calculator (coming soon) to customize this assumption.`]
  ];

  for (const [label, value, desc] of keyMetrics) {
    execSheet.getCell(`A${row}`).value = label;
    execSheet.getCell(`A${row}`).font = { bold: true };
    execSheet.getCell(`B${row}`).value = value;
    if (label === 'List Quality Grade') {
      execSheet.getCell(`B${row}`).style = listGrade === 'A' || listGrade === 'B' ? successStyle : warningStyle;
    }
    if (desc) {
      execSheet.getCell(`C${row}`).value = desc;
      execSheet.getCell(`C${row}`).font = { italic: true, size: 9, color: { argb: 'FF555555' } };
      execSheet.getCell(`C${row}`).alignment = { wrapText: true };
    }
    row++;
  }

  row++; // Blank row

  // Message & Day Variability Insights
  execSheet.mergeCells(`A${row}:C${row}`);
  execSheet.getCell(`A${row}`).value = 'Message & Day Variability Insights';
  execSheet.getCell(`A${row}`).style = sectionHeaderStyle;
  row++;

  const streak2Pct  = totalUniqueInSummary > 0 ? (streak2  / totalUniqueInSummary * 100) : 0;
  const streak3Pct  = totalUniqueInSummary > 0 ? (streak3  / totalUniqueInSummary * 100) : 0;
  const streak4Pct  = totalUniqueInSummary > 0 ? (streak4  / totalUniqueInSummary * 100) : 0;
  const streak5Pct  = totalUniqueInSummary > 0 ? (streak5plus / totalUniqueInSummary * 100) : 0;
  const lowDayPct   = totalUniqueInSummary > 0 ? (lowDayVariety / totalUniqueInSummary * 100) : 0;

  const variabilityRows = [
    ['Same Message 2+ in a Row', `${streak2.toLocaleString()} numbers (${streak2Pct.toFixed(1)}%)`,
      'Numbers that received the same message consecutively at least twice. Consumers who hear the same voicemail repeatedly begin to tune it out.'],
    ['Same Message 3+ in a Row', `${streak3.toLocaleString()} numbers (${streak3Pct.toFixed(1)}%)`,
      'Receiving the same message three or more times raises the perceived robocall signature and reduces callback likelihood.'],
    ['Same Message 4+ in a Row', `${streak4.toLocaleString()} numbers (${streak4Pct.toFixed(1)}%)`,
      'Four or more consecutive identical messages suggests missing message rotation — consider adding a second or third message variant.'],
    ['Same Message 5+ in a Row', `${streak5plus.toLocaleString()} numbers (${streak5Pct.toFixed(1)}%)`,
      'High repetition. Listeners who recognize a repeated script often delete without listening.'],
    ['Low Day-of-Week Variety', `${lowDayVariety.toLocaleString()} numbers (${lowDayPct.toFixed(1)}%)`,
      'Contacted almost exclusively on the same day(s) of the week. Predictable timing allows consumers to categorize your calls as routine.'],
  ];

  for (const [label, value, desc] of variabilityRows) {
    execSheet.getCell(`A${row}`).value = label;
    execSheet.getCell(`A${row}`).font = { bold: true };
    execSheet.getCell(`B${row}`).value = value;
    execSheet.getCell(`C${row}`).value = desc;
    execSheet.getCell(`C${row}`).font = { italic: true, size: 9, color: { argb: 'FF555555' } };
    execSheet.getCell(`C${row}`).alignment = { wrapText: true };
    row++;
  }

  // Variability narrative
  execSheet.mergeCells(`A${row}:C${row}`);
  execSheet.getCell(`A${row}`).value =
    'Why variability drives callbacks: Consumers who receive the same message on the same day every week develop pattern recognition — they learn to dismiss or delete without listening. ' +
    'Varying both the message and the day of week creates unpredictability that feels relevant rather than automated. A consumer who usually gets your message on Tuesday but receives it on a ' +
    'Thursday is more likely to engage. Rotating two or three message variants also prevents voicemail fatigue and can meaningfully improve callback rates.';
  execSheet.getCell(`A${row}`).font = { italic: true, size: 9, color: { argb: 'FF444444' } };
  execSheet.getCell(`A${row}`).alignment = { wrapText: true };
  execSheet.getRow(row).height = 55;
  row++;

  row++; // Blank row

  // Delivery Cadence
  if (cadenceMultiTouchCount > 0) {
    execSheet.mergeCells(`A${row}:C${row}`);
    execSheet.getCell(`A${row}`).value = 'Delivery Cadence';
    execSheet.getCell(`A${row}`).style = sectionHeaderStyle;
    row++;

    const cadenceTotalNumbers = cadenceSingleTouch + cadenceMultiTouchCount;
    const cadenceRows = [
      ['Single Touch (1 attempt only)',
        `${cadenceSingleTouch.toLocaleString()} (${cadenceTotalNumbers > 0 ? (cadenceSingleTouch / cadenceTotalNumbers * 100).toFixed(1) : '0.0'}%)`,
        'Numbers reached with a single DDVM during this period. Consumers often need 2–3 touches before taking action — a callback, a payment, or a response rarely happens the first time a message is heard. These numbers represent real follow-up opportunity: adding a second or third DDVM attempt at the right cadence (3–10 days) typically produces meaningful incremental results without diminishing returns.'],
      ['Re-attempted (2+ attempts)',
        `${cadenceMultiTouchCount.toLocaleString()} (${cadenceTotalNumbers > 0 ? (cadenceMultiTouchCount / cadenceTotalNumbers * 100).toFixed(1) : '0.0'}%)`,
        'Numbers with multiple attempts. Cadence breakdown below is based on the median interval between consecutive attempts for each number.'],
      ['  Same-day re-attempt (< 1 day)',
        `${cadenceBucket_sameDay.toLocaleString()} (${(cadenceBucket_sameDay / cadenceMultiTouchCount * 100).toFixed(1)}% of re-attempted)`,
        'Numbers typically re-attempted on the same calendar day. Same-day re-attempts are rarely effective and may indicate a campaign configuration issue.'],
      ['  1–2 days',
        `${cadenceBucket_1to2.toLocaleString()} (${(cadenceBucket_1to2 / cadenceMultiTouchCount * 100).toFixed(1)}% of re-attempted)`,
        'Very short interval after a successful delivery. Re-attempting a number the day after a successful drop does not give the consumer time to respond and can accelerate list fatigue. Note: re-attempting a number the day after an unsuccessful delivery attempt is perfectly acceptable — this flag is relevant only when the prior attempt succeeded.'],
      ['  3–5 days ✓',
        `${cadenceBucket_3to5.toLocaleString()} (${(cadenceBucket_3to5 / cadenceMultiTouchCount * 100).toFixed(1)}% of re-attempted)`,
        'Ideal range. Frequent enough to maintain urgency while giving consumers time to respond.'],
      ['  6–10 days ✓',
        `${cadenceBucket_6to10.toLocaleString()} (${(cadenceBucket_6to10 / cadenceMultiTouchCount * 100).toFixed(1)}% of re-attempted)`,
        'Ideal range. Solid cadence that balances persistence with consumer experience.'],
      ['  11–15 days',
        `${cadenceBucket_11to15.toLocaleString()} (${(cadenceBucket_11to15 / cadenceMultiTouchCount * 100).toFixed(1)}% of re-attempted)`,
        'Longer than the 3–10 day ideal. Works well for lower-urgency or compliance-sensitive campaigns. For time-sensitive outreach, a tighter interval is likely to produce better engagement.'],
      ['  16–30 days',
        `${cadenceBucket_16to30.toLocaleString()} (${(cadenceBucket_16to30 / cadenceMultiTouchCount * 100).toFixed(1)}% of re-attempted)`,
        'Significantly longer than the 3–10 day ideal. For most campaign types, a tighter cadence will increase the likelihood of connecting while the consumer\'s situation is still active.'],
      ['  30+ days',
        `${cadenceBucket_over30.toLocaleString()} (${(cadenceBucket_over30 / cadenceMultiTouchCount * 100).toFixed(1)}% of re-attempted)`,
        'Infrequent. Re-attempt intervals this long are likely leaving significant outreach opportunity unrealized for most campaign types.'],
    ];
    if (cadenceOverallMedian !== null) {
      const cadenceMedianStr = cadenceOverallMedian < 1
        ? `${(cadenceOverallMedian * 24).toFixed(1)} hours`
        : `${cadenceOverallMedian.toFixed(1)} days`;
      cadenceRows.push(['Overall Median Re-attempt Interval', cadenceMedianStr,
        'The median of per-number median intervals across all re-attempted numbers. Half of re-attempted numbers were contacted at this frequency or faster.']);
    }

    for (const [label, value, desc] of cadenceRows) {
      execSheet.getCell(`A${row}`).value = label;
      execSheet.getCell(`A${row}`).font = label.startsWith('  ')
        ? { size: 10, color: { argb: 'FF333333' } }
        : { bold: true };
      execSheet.getCell(`B${row}`).value = value;
      if (desc) {
        execSheet.getCell(`C${row}`).value = desc;
        execSheet.getCell(`C${row}`).font = { italic: true, size: 9, color: { argb: 'FF555555' } };
        execSheet.getCell(`C${row}`).alignment = { wrapText: true };
      }
      row++;
    }
    row++; // Blank row
  }

  // Non-Deliverable Records
  const totalNonDeliverable = Object.values(nonDeliverableCounts).reduce((s, v) => s + v, 0);
  if (totalNonDeliverable > 0) {
    execSheet.mergeCells(`A${row}:C${row}`);
    execSheet.getCell(`A${row}`).value = 'Non-Deliverable Records (Excluded from Delivery Analysis)';
    execSheet.getCell(`A${row}`).style = sectionHeaderStyle;
    row++;

    const notUSTotal = nonDeliverableCounts['not a valid us number'];
    const notUSReal  = notUSTotal - notUSPlaceholderRows;

    const nonDelivRows = [
      ['Not a wireless number (401)',   nonDeliverableCounts['not a wireless number'],
        'Confirmed US numbers that are landlines or VoIP — cannot receive DDVM.'],
      ['Not a valid US number (403)',   notUSTotal,
        notUSTotal > 0
          ? `Includes ${notUSPlaceholderRows.toLocaleString()} placeholder/invalid entries (e.g. all-zero numbers) ` +
            `and ${notUSReal.toLocaleString()} real but non-US numbers (e.g. international contacts). ` +
            `Both are treated identically — VoApps only delivers to wireless US numbers.`
          : ''],
      ['Duplicate number (402)',        nonDeliverableCounts['duplicate number'],
        'Number appeared more than once in the submitted contact list.'],
      ['Undeliverable (404)',           nonDeliverableCounts['undeliverable'],
        'Number too short, too long, or contained an illegal NPA/NXX.'],
      ['Restricted (500–504)',          nonDeliverableCounts['restricted'],
        'Blocked by frequency, geographic, individual, or WebRecon restriction.'],
    ];

    for (const [label, count, note] of nonDelivRows) {
      if (count === 0) continue;
      execSheet.getCell(`A${row}`).value = label;
      execSheet.getCell(`A${row}`).font = { bold: true };
      execSheet.getCell(`B${row}`).value = count.toLocaleString();
      execSheet.getCell(`C${row}`).value = note;
      execSheet.getCell(`C${row}`).font = { italic: true, size: 9 };
      row++;
    }

    execSheet.getCell(`A${row}`).value = 'Total non-deliverable rows';
    execSheet.getCell(`A${row}`).font = { bold: true };
    execSheet.getCell(`B${row}`).value = totalNonDeliverable.toLocaleString();
    row++;
  }

  row++; // Blank row

  // Configuration Error Results (if any)
  if (hasConfigErrors) {
    execSheet.mergeCells(`A${row}:C${row}`);
    execSheet.getCell(`A${row}`).value = '⚠️ Configuration Error Results';
    execSheet.getCell(`A${row}`).style = sectionHeaderStyle;
    row++;

    const errorResultLabels = {
      'invalid message id':    'Invalid Message ID',
      'invalid caller number': 'Invalid Caller Number',
      'prohibited self call':  'Prohibited Self Call'
    };

    for (const [resultKey, label] of Object.entries(errorResultLabels)) {
      const entry = configErrors[resultKey];
      if (entry.total === 0) continue;

      // Summary line
      execSheet.getCell(`A${row}`).value = label;
      execSheet.getCell(`A${row}`).font = { bold: true };
      execSheet.getCell(`B${row}`).value = `${entry.total.toLocaleString()} occurrences`;
      execSheet.getCell(`B${row}`).style = warningStyle;
      row++;

      // By caller number (top 10)
      const callerEntries = Object.entries(entry.byCallerNumber)
        .sort((a, b) => b[1] - a[1]).slice(0, 10);
      if (callerEntries.length > 0) {
        execSheet.getCell(`A${row}`).value = '  By Caller Number:';
        execSheet.getCell(`A${row}`).font = { italic: true };
        row++;
        for (const [caller, count] of callerEntries) {
          execSheet.getCell(`A${row}`).value = `    ${caller}`;
          execSheet.getCell(`B${row}`).value = count.toLocaleString();
          row++;
        }
      }

      // By message ID (top 10)
      const msgEntries = Object.entries(entry.byMessageId)
        .sort((a, b) => b[1] - a[1]).slice(0, 10);
      if (msgEntries.length > 0 && resultKey !== 'invalid caller number') {
        execSheet.getCell(`A${row}`).value = '  By Message ID:';
        execSheet.getCell(`A${row}`).font = { italic: true };
        row++;
        for (const [msgId, count] of msgEntries) {
          execSheet.getCell(`A${row}`).value = `    ${msgId}`;
          execSheet.getCell(`B${row}`).value = count.toLocaleString();
          row++;
        }
      }

      row++; // spacing between error types
    }
    row++;
  }

  row++; // Blank row

  // TN Health Distribution
  execSheet.mergeCells(`A${row}:C${row}`);
  execSheet.getCell(`A${row}`).value = 'TN Health Distribution';
  execSheet.getCell(`A${row}`).style = sectionHeaderStyle;
  row++;

  const healthDist = [
    ['Healthy', `${healthyCount.toLocaleString()} (${healthyPct.toFixed(1)}%)`,
      'Acceptable success rate with no sustained consecutive-failure streak. Good deliverability — no immediate action required. Monitor variability score to avoid repetitive call patterns.'],
    ['Delivery Unlikely', `${toxicCount.toLocaleString()} (${toxicPct.toFixed(1)}%)`,
      'Success rate below 10% with 4+ consecutive failures; or 6+ consecutive failures regardless of rate; or 5+ attempts with zero successes. Successful DDVM delivery is highly unlikely. Suppression is recommended.'],
    ['Never Delivered', `${neverDeliveredCount.toLocaleString()} (${neverDeliveredPct.toFixed(1)}%)`,
      'Zero successful deliveries across all attempts in this dataset. Overlaps with all health categories — a number with only 1–2 attempts and no consecutive failures can be Healthy yet never have a successful delivery on record. % is of all unique numbers.']
  ];

  for (const [label, value, desc] of healthDist) {
    execSheet.getCell(`A${row}`).value = label;
    execSheet.getCell(`A${row}`).font = { bold: true };
    execSheet.getCell(`B${row}`).value = value;
    if (label === 'Delivery Unlikely' || label === 'Never Delivered') {
      execSheet.getCell(`B${row}`).style = warningStyle;
    } else if (label === 'Healthy') {
      execSheet.getCell(`B${row}`).style = successStyle;
    }
    if (desc) {
      execSheet.getCell(`C${row}`).value = desc;
      execSheet.getCell(`C${row}`).font = { italic: true, size: 9, color: { argb: 'FF555555' } };
      execSheet.getCell(`C${row}`).alignment = { wrapText: true };
    }
    row++;
  }

  row++; // Blank row

  // Retry Decay Summary
  execSheet.mergeCells(`A${row}:C${row}`);
  execSheet.getCell(`A${row}`).value = 'Success Probability by Attempt';
  execSheet.getCell(`A${row}`).style = sectionHeaderStyle;
  row++;

  for (const dc of decayCurve.slice(0, 6)) {
    execSheet.getCell(`A${row}`).value = `Attempt ${dc.attemptIndex}`;
    execSheet.getCell(`A${row}`).font = { bold: true };
    execSheet.getCell(`B${row}`).value = `${(dc.probability * 100).toFixed(1)}%`;
    execSheet.getCell(`C${row}`).value = `${dc.successful.toLocaleString()} successful of ${dc.total.toLocaleString()} attempts`;
    execSheet.getCell(`C${row}`).font = { italic: true, size: 9, color: { argb: 'FF555555' } };
    row++;
  }

  // Decay context note
  execSheet.mergeCells(`A${row}:C${row}`);
  execSheet.getCell(`A${row}`).value =
    'Success probability typically drops significantly after 4–6 consecutive unsuccessful attempts on a given number. ' +
    'Numbers that reach this threshold are strong candidates for suppression — continued retries consume capacity with diminishing return. ' +
    'Use the Suppression Candidates tab to identify numbers with extended consecutive failure runs.';
  execSheet.getCell(`A${row}`).font = { italic: true, size: 9, color: { argb: 'FF555555' } };
  execSheet.getCell(`A${row}`).alignment = { wrapText: true };
  execSheet.getRow(row).height = 40;
  row++;

  row++; // Blank row

  // Message Intelligence (AI) Section — only shown when transcripts are available
  const aiMessages = Object.values(messageStats).filter(m => m.transcript);
  if (aiMessages.length > 0) {
    execSheet.mergeCells(`A${row}:C${row}`);
    execSheet.getCell(`A${row}`).value = 'Message Intelligence (AI)';
    execSheet.getCell(`A${row}`).style = sectionHeaderStyle;
    row++;

    const totalMessages    = Object.keys(messageStats).length;
    const voiceAppendMsgs  = aiMessages.filter(m => m.voice_append);
    const urlMsgs          = aiMessages.filter(m => m.mentions_url);
    const callerMismatch   = aiMessages.filter(m => m.mentioned_phone); // simplified — all that mention a phone
    const intentCounts = {};
    for (const m of aiMessages) {
      const k = m.intent || 'unknown';
      intentCounts[k] = (intentCounts[k] || 0) + 1;
    }
    const intentDistStr = Object.entries(intentCounts)
      .sort((a, b) => b[1] - a[1])
      .map(([k, v]) => `${k}: ${v}`)
      .join(' | ');

    const aiMetrics = [
      ['Messages Analyzed', `${aiMessages.length} of ${totalMessages}`,
        'Messages with audio recordings transcribed and analyzed via AI.'],
      ['Voice Append Messages', voiceAppendMsgs.length > 0 ? `${voiceAppendMsgs.length} (${voiceAppendMsgs.map(m => m.message_name).join(', ')})` : 'None',
        'Messages used with VoApps Voice Append — detected via voapps_voice_append in campaign data.'],
      ['Messages Mentioning a Phone #', callerMismatch.length > 0 ? `${callerMismatch.length}` : 'None',
        'Messages where the transcript contains a spoken phone number. Review Caller # Match column in Message Insights tab.'],
      ['Messages with URLs', urlMsgs.length > 0 ? `${urlMsgs.length}` : 'None',
        'Messages that reference a website or URL in the transcript.'],
      ['Intent Distribution', intentDistStr || '—',
        'AI-inferred intent categories across analyzed messages.']
    ];

    for (const [label, value, desc] of aiMetrics) {
      execSheet.getCell(`A${row}`).value = label;
      execSheet.getCell(`A${row}`).font = { bold: true };
      execSheet.getCell(`B${row}`).value = value;
      if (desc) {
        execSheet.getCell(`C${row}`).value = desc;
        execSheet.getCell(`C${row}`).font = { italic: true, size: 9, color: { argb: 'FF555555' } };
        execSheet.getCell(`C${row}`).alignment = { wrapText: true };
      }
      row++;
    }

    // AI accuracy caveat
    execSheet.mergeCells(`A${row}:C${row}`);
    execSheet.getCell(`A${row}`).value =
      '⚠ AI transcription may contain inaccuracies. Review and correct transcripts in the AI settings panel before using for decision-making.';
    execSheet.getCell(`A${row}`).font = { italic: true, size: 9, color: { argb: 'FF7A5200' } };
    execSheet.getCell(`A${row}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFDE7' } };
    execSheet.getCell(`A${row}`).alignment = { wrapText: true };
    execSheet.getRow(row).height = 22;
    row++;

    row++; // Blank row
  }

  // Immediate Actions
  execSheet.mergeCells(`A${row}:C${row}`);
  execSheet.getCell(`A${row}`).value = 'Recommended Actions';
  execSheet.getCell(`A${row}`).style = sectionHeaderStyle;
  row++;

  const actions = [];

  if (toxicCount > 0) {
    actions.push('LIST QUALITY: Consider reviewing numbers with 4–6+ consecutive unsuccessful deliveries over a span of 30+ days. Numbers that have not connected after sustained attempts are unlikely to connect going forward — focusing future campaigns on the reachable segment concentrates spend where it produces results. The Suppression Candidates tab has this list ready to review.');
  }
  if (avgVariability < 40) {
    actions.push(`MESSAGE DIVERSITY: Variability score of ${avgVariability.toFixed(0)} suggests the same recordings are being used frequently on the same numbers. Introducing additional message recordings and rotating them across attempts tends to improve engagement — consumers are more likely to listen when each contact feels distinct. See Variability Analysis for numbers where this will have the most impact.`);
  }

  // Check for back-to-back issues (backToBackIssues pre-computed above)
  if (backToBackIssues > uniqueNumbers * 0.1) {
    actions.push(`MESSAGE ROTATION: ${backToBackIssues.toLocaleString()} numbers received the same message on consecutive attempts. Varying the recording each time a number is re-attempted keeps the outreach feeling fresh and increases the chance the consumer engages rather than dismisses it.`);
  }

  // Check day distribution (lowDayVariety pre-computed above)
  if (lowDayVariety > uniqueNumbers * 0.2) {
    actions.push(`DAY ROTATION: A portion of numbers are being contacted on the same day pattern each week. Consumers who receive outreach on a predictable day begin to mentally categorize it as routine, which reduces engagement. Spreading attempts across additional days keeps the contact feeling timely and relevant. See "Global Insights (Days)" for specifics.`);
  }

  // Add day-of-week recommendations for accounts/messages
  if (accountDayRecommendations.length > 0) {
    actions.push(`DAY ROTATION: ${accountDayRecommendations.length} account(s) are using a limited set of delivery days. Expanding to additional days of the week is a low-effort change that can meaningfully improve engagement rates. See "Global Insights (Days)" for details.`);
  }
  if (messageDayRecommendations.length > 0) {
    actions.push(`DAY ROTATION: ${messageDayRecommendations.length} message(s) are being used on a limited set of days. Rotating delivery days for these messages will help avoid the predictable-pattern effect. See "Global Insights (Days)" for details.`);
  }

  // Check for aggressive re-attempt cadence
  if (cadenceMultiTouchCount > 0) {
    if (cadenceBucket_sameDay > cadenceMultiTouchCount * 0.1) {
      actions.push(`CADENCE: ${cadenceBucket_sameDay.toLocaleString()} numbers (${(cadenceBucket_sameDay / cadenceMultiTouchCount * 100).toFixed(0)}% of re-attempted) are being contacted multiple times on the same calendar day. Same-day re-attempts rarely produce additional successful deliveries and may indicate a campaign configuration issue worth reviewing.`);
    }
    if (cadenceBucket_1to2 > cadenceMultiTouchCount * 0.2) {
      actions.push(`CADENCE: A significant share of re-attempts are occurring within 1–2 days of a prior successful delivery. Consumers need time to act on a message — re-contacting too quickly after a successful drop reduces the likelihood they respond and can diminish the impact of future attempts. A 3–10 day interval after a successful delivery gives the message time to work.`);
    }
  }

  // Add timezone discrepancy warning
  if (timezoneDiscrepancies.length > 0) {
    actions.push(`REPORTING ACCURACY: ${timezoneDiscrepancies.length} account(s) have timezone mismatches between account settings and results. Enabling "Use account timezone in results file" in DirectDrop Voicemail account settings will ensure timestamps and day-of-week analysis in this report reflect the correct local time for each account.`);
  }

  // Add config error warnings
  if (configErrors['invalid message id'].total > 0) {
    actions.push(`CONFIGURATION: ${configErrors['invalid message id'].total.toLocaleString()} attempts could not be delivered due to an invalid message ID — these represent campaign spend with no delivery outcome. Verifying message IDs in the campaign configuration will recover this capacity for productive attempts. See the Executive Summary config error section for a breakdown by caller number.`);
  }
  if (configErrors['invalid caller number'].total > 0) {
    actions.push(`CONFIGURATION: ${configErrors['invalid caller number'].total.toLocaleString()} attempts could not be delivered due to an invalid caller number. Confirming that all caller numbers are active and correctly configured in the account will ensure these attempts reach consumers on future campaigns.`);
  }
  if (configErrors['prohibited self call'].total > 0) {
    actions.push(`CONFIGURATION: ${configErrors['prohibited self call'].total.toLocaleString()} attempts were blocked because the caller number matched the destination number. Reviewing the campaign contact list for numbers that match any active caller IDs will prevent these from being skipped.`);
  }

  if (actions.length === 0) {
    actions.push('Campaign performance looks strong. Delivery rates, list health, and rotation patterns are all within recommended ranges — continue monitoring regularly to maintain this level of performance.');
  }

  for (const action of actions) {
    execSheet.mergeCells(`A${row}:C${row}`);
    execSheet.getCell(`A${row}`).value = action;
    execSheet.getCell(`A${row}`).style = contentStyle;
    execSheet.getRow(row).height = 25;
    row++;
  }

  row++; // Blank row

  // ── Best Next Action ──────────────────────────────────────────────────────
  {
    execSheet.mergeCells(`A${row}:C${row}`);
    execSheet.getCell(`A${row}`).value = 'Best Next Action';
    execSheet.getCell(`A${row}`).style = {
      ...sectionHeaderStyle,
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3F2FB8' } }
    };
    row++;
    execSheet.mergeCells(`A${row}:C${row}`);
    execSheet.getCell(`A${row}`).value = bestNextAction;
    execSheet.getCell(`A${row}`).style = {
      ...contentStyle,
      font: { size: 11, bold: true, color: { argb: 'FF0D053F' } },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E4FF' } }
    };
    execSheet.getRow(row).height = 55;
    row++;
    row++; // spacer
  }

  // Client Rationale
  execSheet.mergeCells(`A${row}:C${row}`);
  execSheet.getCell(`A${row}`).value = 'Why This Matters';
  execSheet.getCell(`A${row}`).style = sectionHeaderStyle;
  row++;

  const rationale = [
    'DDVM works best when it reaches numbers that are genuinely reachable. A clean, focused list means every drop has a higher probability of connecting — improving delivered %, reducing cost-per-contact, and generating more callbacks from the same campaign spend.',
    'Delivery success rate naturally declines as a number accumulates failed attempts. Numbers that haven\'t connected after several tries are unlikely to connect in the future — concentrating volume on the active segment unlocks the performance the channel is capable of.',
    'Message and caller rotation signal to both consumers and carriers that the outreach is varied and purposeful. Higher variability scores correlate with better deliverability and stronger callback rates over time.',
    'Cadence matters: the 3–10 day re-attempt window gives consumers enough time to act on a message before the next drop arrives. Campaigns operating in this window typically see meaningfully better engagement than those with very short or very long intervals.'
  ];

  for (const text of rationale) {
    execSheet.mergeCells(`A${row}:C${row}`);
    execSheet.getCell(`A${row}`).value = `• ${text}`;
    execSheet.getCell(`A${row}`).style = contentStyle;
    execSheet.getRow(row).height = 40;
    row++;
  }

  row++; // Blank row

  // Insights
  execSheet.mergeCells(`A${row}:C${row}`);
  execSheet.getCell(`A${row}`).value = 'Quick Insights';
  execSheet.getCell(`A${row}`).style = sectionHeaderStyle;
  row++;

  const insights = [
    '"TN Health" tab: identifies which numbers are connecting consistently and which have stalled — a direct indicator of where campaign spend is working hardest.',
    '"Suppression Candidates" tab: numbers with extended consecutive failure runs, ready to act on — removing these sharpens the active list and raises the overall delivery rate.',
    '"Variability Analysis" tab: shows where message and caller rotation is strong and where it has room to grow — higher diversity consistently correlates with better results.',
    '"Global Insights (Days)" tab: reveals day-of-week delivery patterns — varying the days consumers receive messages prevents the predictable-pattern effect and keeps outreach feeling timely.'
  ];

  for (const insight of insights) {
    execSheet.mergeCells(`A${row}:C${row}`);
    execSheet.getCell(`A${row}`).value = `• ${insight}`;
    execSheet.getCell(`A${row}`).style = contentStyle;
    row++;
  }

  execSheet.getColumn(1).width = 30;
  execSheet.getColumn(2).width = 50.83;  // renders as 50
  execSheet.getColumn(3).width = 130.83; // renders as 130

  // ========================================
  // TAB 2: TN HEALTH
  // ========================================

  log('Creating TN Health tab...');

  if (includeDetailTabs) {
    const healthSheet = workbook.addWorksheet('TN Health', {
      views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }]
    });

    const healthHeaders = ['Number', 'TN Health', 'Never Delivered', 'Success Rate', 'Total DDVM Attempts',
      'Consecutive Failures', 'Attempt Index', 'Last Success', 'Variability Score', 'Action'];
    healthSheet.getRow(1).values = healthHeaders;
    healthSheet.getRow(1).eachCell((cell) => {
      cell.style = tableHeaderStyle;
    });

    log(`  TN Health: ${filteredHealth.length.toLocaleString()} numbers shown (Delivery Unlikely)`);

    // Column-level numFmt (one call per column instead of N per-cell calls)
    healthSheet.getColumn(4).numFmt = '0.0%';  // D - Success Rate

    const healthRows = filteredHealth.map(ns => {
      const action = ns.tnHealth === 'Delivery Unlikely' ? 'Suppression recommended' : 'Monitor closely';
      // lastSuccessTimestamp is stored as ms-epoch number (or Date from older code paths)
      const lastSuccStr = ns.lastSuccessTimestamp
        ? (ns.lastSuccessTimestamp instanceof Date
            ? ns.lastSuccessTimestamp.toISOString().split('T')[0]
            : new Date(ns.lastSuccessTimestamp).toISOString().split('T')[0])
        : '';
      return [
        ns.number, ns.tnHealth, ns.neverDelivered ? 'Yes' : 'No',
        ns.successRate, ns.totalAttempts, ns.consecutiveFailures,
        ns.attemptIndex, lastSuccStr, ns.variabilityScore, action
      ];
    });
    healthSheet.addRows(healthRows);
    const healthLastRow = healthRows.length + 1;
    // Free source data — ExcelJS has its own internal copy now
    filteredHealth.length = 0; healthRows.length = 0;

    // Conditional formatting — one rule set for the whole column (no per-cell fill)
    healthSheet.addConditionalFormatting({
      ref: `B2:B${healthLastRow}`,
      rules: [
        { type: 'cellIs', operator: 'equal', formulae: ['"Delivery Unlikely"'], priority: 1,
          style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFC7CE' } } } }
      ]
    });
    healthSheet.addConditionalFormatting({
      ref: `C2:C${healthLastRow}`,
      rules: [
        { type: 'cellIs', operator: 'equal', formulae: ['"Yes"'], priority: 1,
          style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFC7CE' } } } }
      ]
    });

    autoFitColumns(healthSheet, 12, 55);

  }

  // ========================================
  // TAB 4: VARIABILITY ANALYSIS
  // ========================================

  log('Creating Variability Analysis tab...');

  if (includeDetailTabs) {
    const varSheet = workbook.addWorksheet('Variability Analysis', {
      views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }]
    });

    const varHeaders = ['Number', 'Variability Score', 'Top Msg %', 'Unique Messages', 'Top Caller Number %',
      'Unique Caller Numbers', 'Back-to-Back', 'Day Entropy', 'Hour Entropy', 'Day Distribution'];
    varSheet.getRow(1).values = varHeaders;
    varSheet.getRow(1).eachCell((cell) => {
      cell.style = tableHeaderStyle;
    });

    log(`  Variability Analysis: ${filteredVariability.length.toLocaleString()} numbers shown (score < 60)`);

    // Column-level numFmt — one call per column instead of N per-cell calls
    varSheet.getColumn(3).numFmt = '0.0%';   // C - Top Msg %
    varSheet.getColumn(5).numFmt = '0.0%';   // E - Top Caller %
    varSheet.getColumn(8).numFmt = '0.00';   // H - Day Entropy
    varSheet.getColumn(9).numFmt = '0.00';   // I - Hour Entropy

    const varRows = filteredVariability.map(ns => [
      ns.number, ns.variabilityScore, ns.topMsgPct, ns.uniqueMsgCount,
      ns.topCallerPct, ns.uniqueCallerCount, ns.backToBackIdentical,
      ns.dayEntropy, ns.hourEntropy, ns.dayDistribution
    ]);
    varSheet.addRows(varRows);
    const varLastRow = varRows.length + 1;
    // Free source data — ExcelJS has its own internal copy now
    filteredVariability.length = 0; varRows.length = 0;

    // Conditional formatting for variability score and back-to-back
    varSheet.addConditionalFormatting({
      ref: `B2:B${varLastRow}`,
      rules: [
        { type: 'cellIs', operator: 'lessThan',           formulae: ['30'], priority: 1,
          style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFC7CE' } } } },
        { type: 'cellIs', operator: 'lessThanOrEqual',    formulae: ['60'], priority: 2,
          style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFEB9C' } } } }
      ]
    });
    varSheet.addConditionalFormatting({
      ref: `G2:G${varLastRow}`,
      rules: [
        { type: 'cellIs', operator: 'greaterThan', formulae: ['2'], priority: 1,
          style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFEB9C' } } } }
      ]
    });

    autoFitColumns(varSheet, 12, 60);

  }

  // ========================================
  // TAB 5: NUMBER SUMMARY
  // ========================================

  log('Creating Number Summary tab...');

  if (includeDetailTabs) {
    const summarySheet = workbook.addWorksheet('Number Summary', {
      views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }]
    });

    const summaryHeaders = ['Number', 'Total DDVM Attempts', 'Successful', 'Unsuccessful', 'Success Rate',
      'TN Health', 'Variability', 'First Attempt', 'Last Attempt',
      'Top Msg ID', 'Top Msg Name', 'Top Msg %', 'Unique Msgs',
      'Top Caller', 'Top Caller Name', 'Top Caller Number %', 'Unique Caller Numbers',
      'Intent', 'Day Distribution'];
    summarySheet.getRow(1).values = summaryHeaders;
    summarySheet.getRow(1).eachCell((cell) => {
      cell.style = tableHeaderStyle;
    });

    log(`  Number Summary: ${filteredSummary.length.toLocaleString()} numbers shown (any flag)`);

    // Column-level numFmt — one call per column instead of N per-cell calls
    summarySheet.getColumn(5).numFmt  = '0.0%';  // E - Success Rate
    summarySheet.getColumn(12).numFmt = '0.0%';  // L - Top Msg %
    summarySheet.getColumn(16).numFmt = '0.0%';  // P - Top Caller %

    const summaryRows = filteredSummary.map(ns => [
      ns.number, ns.totalAttempts, ns.successful, ns.unsuccessful, ns.successRate,
      ns.tnHealth, ns.variabilityScore, ns.firstAttempt, ns.lastAttempt,
      Number(ns.topMsgId) || ns.topMsgId, ns.topMsgName, ns.topMsgPct, ns.uniqueMsgCount,
      ns.topCallerNum, ns.topCallerName, ns.topCallerPct, ns.uniqueCallerCount,
      ns.messageIntent, ns.dayDistribution
    ]);
    summarySheet.addRows(summaryRows);
    // Free source data — ExcelJS has its own internal copy now
    filteredSummary.length = 0; summaryRows.length = 0;

    // Set column widths
    autoFitColumns(summarySheet, 10, 60);

  }

  log(`  Number Summary: ${filteredSummary.length.toLocaleString()} flagged rows written`);

  // ========================================
  // TAB 6: MESSAGE INSIGHTS
  // ========================================

  log('Creating Message Insights tab...');

  const msgSheet = workbook.addWorksheet('Message Insights', {
    views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }]
  });

  const messageArray = Object.values(messageStats).map(m => ({
    ...m,
    uniqueNumbers: m.uniqueNumbers,
    success_rate: m.total > 0 ? m.successful / m.total : 0,
    dayPattern: getDayUsagePattern(m.dayOfWeekCounts)
  })).sort((a, b) => b.total - a.total);

  // AI columns are always present — populated when AI analysis has run, empty otherwise
  const hasAiData = messageArray.some(m => m.transcript);
  const msgHeaders = [
    'Message ID', 'Message Name', 'Intent', 'Total DDVM Attempts', 'Unique Numbers',
    'Successful', 'Unsuccessful', 'Success Rate', 'Day Usage', 'Recommendation',
    'Transcript', 'Mentioned Phone', 'Caller # Match', 'Contains URL', 'Voice Append'
  ];
  msgSheet.getRow(1).values = msgHeaders;
  msgSheet.getRow(1).eachCell((cell) => {
    cell.style = tableHeaderStyle;
  });

  let msgRow = 2;
  for (const msg of messageArray) {
    const mentionedPhone = msg.mentioned_phone || '';
    const transcript = msg.transcript || '';
    msgSheet.getRow(msgRow).values = [
      Number(msg.message_id) || msg.message_id, msg.message_name, msg.intent, msg.total, msg.uniqueNumbers,
      msg.successful, msg.unsuccessful, msg.success_rate,
      msg.dayPattern.days.join(', ') || 'All days',
      msg.dayPattern.limited ? msg.dayPattern.recommendation : '',
      // AI columns — blank when AI has not run; populated after transcription
      transcript,
      mentionedPhone,
      // Caller # Match: compare mentioned phone vs dominant caller number used for this message
      !transcript ? '' : getCallerMatchStatus(mentionedPhone, msg.callerNumbers),
      !transcript ? '' : (msg.mentions_url ? 'Yes' : 'No'),
      // Voice Append: only show 'Yes' when confirmed — blank when not detected (data may not be available)
      msg.voice_append ? 'Yes' : ''
    ];
    msgSheet.getCell(`H${msgRow}`).numFmt = '0.0%';

    // Hyperlink Message ID → VoApps platform message page
    const msgIdCell = msgSheet.getCell(`A${msgRow}`);
    if (msg.account_id && msg.account_id !== 'Unknown' && msg.message_id && msg.message_id !== 'Unknown') {
      msgIdCell.value = {
        text: String(msg.message_id),
        hyperlink: `https://directdropvoicemail.voapps.com/accounts/${msg.account_id}/messages/${msg.message_id}`
      };
      msgIdCell.font = { color: { argb: 'FF0563C1' }, underline: true };
    }

    if (msg.dayPattern.limited) {
      msgSheet.getCell(`J${msgRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEB9C' } };
    }
    if (msg.voice_append) {
      msgSheet.getCell(`P${msgRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E3F2FD' } };
    }
    msgRow++;
  }

  autoFitColumns(msgSheet, 12, 100);

  // ========================================
  // TAB 7: CALLER # INSIGHTS
  // ========================================

  log('Creating Caller # Insights tab...');

  const callerSheet = workbook.addWorksheet('Caller # Insights', {
    views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }]
  });

  const callerArray = Object.values(callerStats).map(c => ({
    ...c,
    uniqueNumbers: c.uniqueNumbers,
    success_rate: c.total > 0 ? c.successful / c.total : 0,
    dayPattern: getDayUsagePattern(c.dayOfWeekCounts)
  })).sort((a, b) => b.total - a.total);

  const callerHeaders = ['Caller Number', 'Caller Name', 'Total DDVM Attempts', 'Unique Numbers', 'Successful', 'Unsuccessful', 'Success Rate', 'Day Usage'];
  callerSheet.getRow(1).values = callerHeaders;
  callerSheet.getRow(1).eachCell((cell) => {
    cell.style = tableHeaderStyle;
  });

  let callerRow = 2;
  for (const caller of callerArray) {
    callerSheet.getRow(callerRow).values = [
      caller.caller_number, caller.caller_name, caller.total, caller.uniqueNumbers,
      caller.successful, caller.unsuccessful, caller.success_rate,
      caller.dayPattern.days.join(', ') || 'All days'
    ];
    callerSheet.getCell(`G${callerRow}`).numFmt = '0.0%';
    callerRow++;
  }

  autoFitColumns(callerSheet, 12, 100);

  // ========================================
  // TAB 8: GLOBAL INSIGHTS (TIME)
  // ========================================

  log('Creating Global Insights (Days) tab...');

  const timeSheet = workbook.addWorksheet('Global Insights (Days)', {
    views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }]
  });

  // Timezone notice at top
  let userTzDisplay;
  if (userTimezone === 'VoApps') {
    userTzDisplay = 'VoApps Time (UTC-7, constant)';
  } else if (userTimezone === 'UTC') {
    userTzDisplay = 'UTC';
  } else if (userTimezone.startsWith('America/')) {
    userTzDisplay = `${userTimezoneLabel} (DST-aware)`;
  } else if (userTimezone.startsWith('-') || userTimezone.startsWith('+')) {
    userTzDisplay = userTimezoneLabel ? `${userTimezoneLabel} (UTC${userTimezone})` : `UTC${userTimezone}`;
  } else {
    userTzDisplay = userTimezoneLabel || userTimezone;
  }
  timeSheet.getCell('A1').value = `Report Timezone: ${userTzDisplay}`;
  timeSheet.getCell('A1').font = { bold: true, size: 12, color: { argb: VOAPPS_PURPLE } };
  timeSheet.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: VOAPPS_PURPLE_PALE } };
  timeSheet.mergeCells('A1:E1');

  // Daily Success Patterns
  const dailyStats = globalDayStats;
  const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

  const dayHeaderRow = 3;
  timeSheet.getRow(dayHeaderRow).values = ['Day of Week', 'Total DDVM Attempts', 'Successful', 'Unsuccessful', 'Success Rate'];
  timeSheet.getRow(dayHeaderRow).eachCell((cell) => {
    cell.style = tableHeaderStyle;
  });

  let dayRow = dayHeaderRow + 1;
  for (let d = 0; d < 7; d++) {
    const stats = dailyStats[d];
    const successRate = stats.total > 0 ? stats.successful / stats.total : 0;
    timeSheet.getRow(dayRow).values = [dayNames[d], stats.total, stats.successful, stats.unsuccessful, successRate];
    timeSheet.getCell(`E${dayRow}`).numFmt = '0.0%';
    dayRow++;
  }

  // Day-of-week recommendations section
  if (accountDayRecommendations.length > 0 || messageDayRecommendations.length > 0) {
    const recoStartRow = dayRow + 2;

    // Preamble tip row
    const tipRow = recoStartRow - 1;
    timeSheet.mergeCells(`A${tipRow}:E${tipRow}`);
    timeSheet.getCell(`A${tipRow}`).value =
      'Consumers who receive your message on a predictable day pattern begin to recognize and mentally categorize it as routine — ' +
      'reducing the likelihood they listen or call back. Rotating across additional days of the week (in addition to using different ' +
      'message recordings) makes your outreach feel less automated and more timely. Keep in mind that if a specific message is part ' +
      'of a planned sequence and future messages will be different, the rotation benefit compounds — each touchpoint feels fresh ' +
      'rather than expected.';
    timeSheet.getCell(`A${tipRow}`).font = { italic: true, size: 10, color: { argb: 'FF555555' } };

    timeSheet.getCell(`A${recoStartRow}`).value = 'Day-of-Week Recommendations';
    timeSheet.getCell(`A${recoStartRow}`).font = { bold: true, size: 12 };
    timeSheet.mergeCells(`A${recoStartRow}:E${recoStartRow}`);

    const recoHeaderRow = recoStartRow + 1;
    timeSheet.getRow(recoHeaderRow).values = ['Type', 'ID', 'Name', 'Days Used', 'Recommendation'];
    timeSheet.getRow(recoHeaderRow).eachCell((cell) => {
      cell.style = tableHeaderStyle;
    });

    let recoRow = recoHeaderRow + 1;
    for (const reco of accountDayRecommendations) {
      timeSheet.getRow(recoRow).values = ['Account', Number(reco.id) || reco.id, '', reco.days, reco.recommendation];
      timeSheet.getCell(`E${recoRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEB9C' } };
      recoRow++;
    }
    for (const reco of messageDayRecommendations) {
      timeSheet.getRow(recoRow).values = ['Message', Number(reco.id) || reco.id, reco.name, reco.days, reco.recommendation];
      timeSheet.getCell(`E${recoRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEB9C' } };
      recoRow++;
    }
  }

  timeSheet.getColumn(1).width = 40;
  timeSheet.getColumn(2).width = 40;
  timeSheet.getColumn(3).width = 40;
  timeSheet.getColumn(4).width = 40;
  autoFitColumns(timeSheet, 12, 100, [1, 2, 3, 4]); // skip A–D, auto-fit remaining columns only

  // ========================================
  // TAB 9: CONSECUTIVE UNSUCCESSFUL
  // ========================================

  log('Creating Suppression Candidates tab...');

  const consecSheet = workbook.addWorksheet('Suppression Candidates', {
    views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }]
  });

  const consecHeaders = ['Number', 'Consecutive Failures', 'Run Start', 'Run End', 'Span (Days)', 'TN Health', 'Action'];
  consecSheet.getRow(1).values = consecHeaders;
  consecSheet.getRow(1).eachCell((cell) => {
    cell.style = tableHeaderStyle;
  });

  let consecRow = 2;
  for (const run of suppressionRuns) {
    consecSheet.getRow(consecRow).values = [
      Number(run.number), run.count,
      run.runStart && !isNaN(run.runStart.getTime()) ? run.runStart : null,
      run.runEnd && !isNaN(run.runEnd.getTime()) ? run.runEnd : null,
      Math.round(run.spanDays),
      run.tnHealth, 'Suppression recommended'
    ];

    if (run.runStart) consecSheet.getCell(`C${consecRow}`).numFmt = 'yyyy-mm-dd hh:mm';
    if (run.runEnd) consecSheet.getCell(`D${consecRow}`).numFmt = 'yyyy-mm-dd hh:mm';
    consecSheet.getCell(`F${consecRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC7CE' } };

    consecRow++;
  }

  log(`  Suppression Candidates: ${suppressionRuns.length.toLocaleString()} rows`);

  autoFitColumns(consecSheet, 12, 50);

  // ========================================
  // TAB 10: DELIVERY TREND
  // ========================================

  log('Creating Delivery Trend tab...');
  {
    const trendSheet = workbook.addWorksheet('Delivery Trend', {
      properties: { tabColor: { argb: 'FF16509B' } }
    });

    // ── Build sorted date list ──────────────────────────────────────────────
    const sortedDates = Object.keys(dailyStats).sort();
    const useWeekly   = sortedDates.length > 60; // collapse to weeks if wide range

    // Aggregate into buckets (daily or ISO-week Mon–Sun)
    const buckets = new Map(); // key → { label, attempts, successes, weekStart }
    for (const d of sortedDates) {
      let key, label;
      if (useWeekly) {
        // ISO week: find Monday of the week containing this date
        const dt   = new Date(d + 'T00:00:00Z');
        const dow  = (dt.getUTCDay() + 6) % 7; // Mon=0
        const mon  = new Date(dt.getTime() - dow * 86400000);
        const mon2 = new Date(mon.getTime() + 6 * 86400000);
        const fmt  = (d2) => {
          const mo = String(d2.getUTCMonth() + 1).padStart(2, '0');
          const dy = String(d2.getUTCDate()).padStart(2, '0');
          return `${d2.getUTCFullYear()}-${mo}-${dy}`;
        };
        key   = fmt(mon);
        label = `${fmt(mon)} – ${fmt(mon2)}`;
      } else {
        key = label = d;
      }
      const existing = buckets.get(key) || { label, attempts: 0, successes: 0 };
      existing.attempts  += dailyStats[d].attempts;
      existing.successes += dailyStats[d].successes;
      buckets.set(key, existing);
    }
    const bucketList = Array.from(buckets.entries()).sort((a, b) => a[0].localeCompare(b[0]));

    // ── Styles ──────────────────────────────────────────────────────────────
    const trHeaderStyle = {
      font: { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0D053F' } },
      alignment: { horizontal: 'center', vertical: 'middle' },
      border: { bottom: { style: 'thin', color: { argb: 'FFFF4B7D' } } }
    };
    const trEvenStyle = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFBF7F3' } } };
    const trOddStyle  = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } } };
    const trNumStyle  = { alignment: { horizontal: 'right' }, numFmt: '#,##0' };
    const trPctStyle  = { alignment: { horizontal: 'right' }, numFmt: '0.0%' };

    // ── Title ───────────────────────────────────────────────────────────────
    trendSheet.mergeCells('A1:F1');
    trendSheet.getCell('A1').value = `Delivery Trend — ${useWeekly ? 'Weekly' : 'Daily'} Breakdown`;
    trendSheet.getCell('A1').style = {
      font: { bold: true, size: 14, color: { argb: 'FF0D053F' } },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFBF7F3' } },
      alignment: { horizontal: 'left', vertical: 'middle' }
    };
    trendSheet.getRow(1).height = 30;

    // ── Column headers ───────────────────────────────────────────────────────
    const trHeaders = [
      useWeekly ? 'Week' : 'Date',
      'Attempts',
      'Successful',
      'Success Rate',
      'WoW Attempts',
      'WoW Successful'
    ];
    trendSheet.addRow([]); // row 2 blank
    const hdrRow = trendSheet.getRow(3);
    trHeaders.forEach((h, i) => {
      hdrRow.getCell(i + 1).value = h;
      hdrRow.getCell(i + 1).style = trHeaderStyle;
    });
    hdrRow.height = 24;

    // ── Data rows ────────────────────────────────────────────────────────────
    let dataStartRow = 4;
    bucketList.forEach(([, bkt], idx) => {
      const rowNum = dataStartRow + idx;
      const prev   = idx > 0 ? bucketList[idx - 1][1] : null;
      const wowAtt = prev ? bkt.attempts  - prev.attempts  : null;
      const wowSuc = prev ? bkt.successes - prev.successes : null;
      const rate   = bkt.attempts > 0 ? bkt.successes / bkt.attempts : 0;
      const baseStyle = idx % 2 === 0 ? trEvenStyle : trOddStyle;

      trendSheet.getRow(rowNum).height = 18;
      trendSheet.getCell(`A${rowNum}`).value = bkt.label;
      trendSheet.getCell(`A${rowNum}`).style = { ...baseStyle, font: { size: 10 }, alignment: { horizontal: 'left' } };
      trendSheet.getCell(`B${rowNum}`).value = bkt.attempts;
      trendSheet.getCell(`B${rowNum}`).style = { ...baseStyle, ...trNumStyle, font: { size: 10 } };
      trendSheet.getCell(`C${rowNum}`).value = bkt.successes;
      trendSheet.getCell(`C${rowNum}`).style = { ...baseStyle, ...trNumStyle, font: { size: 10 } };
      trendSheet.getCell(`D${rowNum}`).value = rate;
      trendSheet.getCell(`D${rowNum}`).style = { ...baseStyle, ...trPctStyle, font: { size: 10 } };
      if (wowAtt !== null) {
        trendSheet.getCell(`E${rowNum}`).value = wowAtt;
        trendSheet.getCell(`E${rowNum}`).style = {
          ...baseStyle, ...trNumStyle, font: { size: 10, color: { argb: wowAtt >= 0 ? 'FF1E7E34' : 'FFC0392B' } }
        };
        trendSheet.getCell(`F${rowNum}`).value = wowSuc;
        trendSheet.getCell(`F${rowNum}`).style = {
          ...baseStyle, ...trNumStyle, font: { size: 10, color: { argb: wowSuc >= 0 ? 'FF1E7E34' : 'FFC0392B' } }
        };
      }
    });

    // ── Data bars on Attempts (col B) and Successful (col C) ────────────────
    if (bucketList.length > 0) {
      const lastDataRow = dataStartRow + bucketList.length - 1;
      const attRange = `B${dataStartRow}:B${lastDataRow}`;
      const sucRange = `C${dataStartRow}:C${lastDataRow}`;
      trendSheet.addConditionalFormatting({
        ref: attRange,
        rules: [{ type: 'dataBar', priority: 1, minLength: 0, maxLength: 100,
          cfvo: [{ type: 'min' }, { type: 'max' }],
          color: { argb: 'FF0D053F' }  // navy
        }]
      });
      trendSheet.addConditionalFormatting({
        ref: sucRange,
        rules: [{ type: 'dataBar', priority: 1, minLength: 0, maxLength: 100,
          cfvo: [{ type: 'min' }, { type: 'max' }],
          color: { argb: 'FF1E7E34' }  // green
        }]
      });
    }

    // ── Totals row ───────────────────────────────────────────────────────────
    if (bucketList.length > 0) {
      const totRow = dataStartRow + bucketList.length + 1;
      trendSheet.getCell(`A${totRow}`).value = 'TOTAL';
      trendSheet.getCell(`A${totRow}`).style = { font: { bold: true, size: 10, color: { argb: 'FF0D053F' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFAD6D7' } } };
      trendSheet.getCell(`B${totRow}`).value = totalAttempts;
      trendSheet.getCell(`B${totRow}`).style = { font: { bold: true, size: 10 }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFAD6D7' } }, ...trNumStyle };
      trendSheet.getCell(`C${totRow}`).value = _totalSuccess;
      trendSheet.getCell(`C${totRow}`).style = { font: { bold: true, size: 10 }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFAD6D7' } }, ...trNumStyle };
      const totRate = totalAttempts > 0 ? _totalSuccess / totalAttempts : 0;
      trendSheet.getCell(`D${totRow}`).value = totRate;
      trendSheet.getCell(`D${totRow}`).style = { font: { bold: true, size: 10 }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFAD6D7' } }, ...trPctStyle };
    }

    trendSheet.getColumn('A').width = useWeekly ? 26 : 14;
    trendSheet.getColumn('B').width = 16;
    trendSheet.getColumn('C').width = 16;
    trendSheet.getColumn('D').width = 14;
    trendSheet.getColumn('E').width = 16;
    trendSheet.getColumn('F').width = 16;

    log(`  Delivery Trend: ${bucketList.length} ${useWeekly ? 'week' : 'day'} buckets`);
  }

  // ========================================
  // TAB 11: GLOSSARY
  // ========================================

  log('Creating Glossary tab...');

  const glossarySheet = workbook.addWorksheet('Glossary');

  glossarySheet.mergeCells('A1:B1');
  glossarySheet.getCell('A1').value = `Number Analysis and Delivery Intelligence Report v${VERSION} - Glossary`;
  glossarySheet.getCell('A1').style = headerStyle;
  glossarySheet.getRow(1).height = 35;

  let glossRow = 3;

  // Core Concepts
  glossarySheet.mergeCells(`A${glossRow}:B${glossRow}`);
  glossarySheet.getCell(`A${glossRow}`).value = 'Core Concepts';
  glossarySheet.getCell(`A${glossRow}`).style = sectionHeaderStyle;
  glossRow++;

  const coreConcepts = [
    ['DDVM', 'DirectDrop Voicemail - VoApps patented technology that delivers voicemail messages directly to mobile carrier voicemail platforms without ringing the phone.'],
    ['VoApps Time', 'A constant UTC-7 timezone (no DST adjustment) used by VoApps to slice days consistently for campaign scheduling. When VoApps Time is selected, timestamps always show UTC-7 regardless of season. US timezone options (ET, CT, MT, PT) are DST-aware and show the correct offset for each timestamp based on its date.'],
    ['Attempt Index', 'The number of DDVM delivery attempts to a phone number since the last successful delivery. Resets to 0 after each success. Higher values indicate declining deliverability.'],
    ['TN Health', 'Classification of phone number health based on delivery performance: Healthy (good deliverability) or Delivery Unlikely (successful delivery is highly unlikely based on consecutive failures and low success rate).'],
    ['Never Delivered', 'A phone number that has never received a successful DDVM delivery across all attempts.'],
    ['Variability Score', 'A 0-100 score measuring how much variety exists in messaging, caller numbers, and timing. Higher scores indicate better rotation practices.'],
    ['Success Probability by Attempt', `Shows how DDVM delivery success rate changes with each successive attempt on a phone number. "Attempt index" is how many times a given number has been tried within this date range.

How to read it: Each data point represents all numbers at that attempt count — not the same number over time. Attempt 1 numbers are those receiving their very first DDVM attempt; attempt 2 numbers have already had one failed delivery; and so on.

Example with real-looking numbers:
  • Attempt 1 (first try): 80% delivered — 1,000 numbers attempted; 800 succeed
  • Attempt 2 (second try): 30% delivered — The 200 that failed now get a second shot; 60 succeed
  • Attempt 3 (third try): 15% delivered — The remaining 140 get one more try; ~21 succeed

Why does overall Delivered % (76%) seem higher than attempt 2 (30%)? Because the vast majority of your volume is first-attempt numbers. If 90% of attempts are first tries (80% success) and 10% are second tries (30% success), the blended rate is (0.90 × 80%) + (0.10 × 30%) = 75% — close to your 76.1% overall.

The declining curve also reflects selection bias: numbers that succeed on attempt 1 are the easy-to-reach numbers. By attempt 2, the pool consists mostly of harder-to-reach numbers — full voicemails, non-wireless lines, carrier restrictions — which naturally have lower success rates regardless of how many times they are tried.

Use the data to set retry limits: when success probability drops below ~15–20%, additional retries produce very few new successful deliveries. Numbers with extended consecutive failure runs are listed in the Suppression Candidates tab.`],
    ['Back-to-Back Identical', 'Count of times the same message was delivered to a number in consecutive attempts. Should be minimized for natural delivery patterns.'],
    ['Day Entropy', 'Measure of how evenly distributed DDVM attempts are across days of the week. Higher entropy (closer to 1.0) means better day-of-week variety.'],
    ['Message Intent', 'Inferred purpose of a message based on its name or AI transcript (e.g., collections, reminder, appointment, callback, welcome, followup, loan servicing). When AI Message Analysis is enabled, intent is derived from the full transcript using a classification model for higher accuracy.'],
    ['List Quality Grade', 'Overall grade (A-D) for the phone number list based on TN health distribution. A: >80% Healthy, <5% Delivery Unlikely. B: >60% Healthy, <10% Delivery Unlikely. C: >40% Healthy, <20% Delivery Unlikely. D: All other cases.'],
    ['Message Transcript', 'Full spoken text of the DDVM voicemail recording, transcribed using Whisper (local or OpenAI). Populated when AI Message Analysis is enabled in settings. Stored permanently in the local DuckDB cache — each message is only transcribed once.'],
    ['Caller # Match', 'Indicates whether a phone number spoken aloud in the message matches the caller ID shown to the recipient. A mismatch means the recipient hears a different callback number than what their phone displays — which can cause confusion or reduce callback rates.'],
    ['Voice Append', 'Indicates the message was used with VoApps Voice Append — a feature that appends a personalized spoken element to the base recording. Detected via the voapps_voice_append field in campaign export data.'],
    ['Contains URL', 'Indicates the message transcript references a website URL or domain name (e.g., "visit us at acme.com" or "go to our website"). Detected via regex on the transcript when AI Message Analysis is enabled.']
  ];

  for (const [term, def] of coreConcepts) {
    glossarySheet.getCell(`A${glossRow}`).value = term;
    glossarySheet.getCell(`A${glossRow}`).font = { bold: true };
    glossarySheet.getCell(`B${glossRow}`).value = def;
    glossarySheet.getCell(`B${glossRow}`).style = contentStyle;
    glossarySheet.getRow(glossRow).height = term === 'Success Probability by Attempt' ? 240 : 40;
    glossRow++;
  }

  glossRow++;

  // TN Health Classifications
  glossarySheet.mergeCells(`A${glossRow}:B${glossRow}`);
  glossarySheet.getCell(`A${glossRow}`).value = 'TN Health Classifications';
  glossarySheet.getCell(`A${glossRow}`).style = sectionHeaderStyle;
  glossRow++;

  const healthDefs = [
    ['Healthy', 'Good delivery performance. No sustained consecutive-failure streak meeting the Delivery Unlikely thresholds. Continue normal operations.'],
    ['Delivery Unlikely', 'Very poor performance. Success rate below 10% with 4+ consecutive failures; or 6+ consecutive failures regardless of rate; or 5+ attempts with zero successes. Successful DDVM delivery is highly unlikely. Suppression is recommended to protect caller reputation and avoid wasted attempts.'],
    ['Never Delivered', 'Zero successful deliveries across all attempts in the date range. These numbers should be suppressed immediately — they consume budget with no return.'],
  ];

  const healthActions = {
    'Healthy': 'Continue monitoring',
    'Delivery Unlikely': 'Suppression recommended',
    'Never Delivered': 'Remove from list',
  };

  for (const [term, def] of healthDefs) {
    glossarySheet.getCell(`A${glossRow}`).value = term;
    glossarySheet.getCell(`A${glossRow}`).font = { bold: true };
    glossarySheet.getCell(`B${glossRow}`).value = `${def} Recommended action: ${healthActions[term]}.`;
    glossarySheet.getCell(`B${glossRow}`).style = contentStyle;
    glossarySheet.getRow(glossRow).height = 45;
    glossRow++;
  }

  glossRow++;

  // VoApps Results
  glossarySheet.mergeCells(`A${glossRow}:B${glossRow}`);
  glossarySheet.getCell(`A${glossRow}`).value = 'VoApps DDVM Result Codes';
  glossarySheet.getCell(`A${glossRow}`).style = sectionHeaderStyle;
  glossRow++;

  const resultDefs = [
    ['100 - Pending', 'The voicemail is not yet in progress. The message is waiting for campaign start time.'],
    ['101 - Running', 'The record is currently being processed. Pre-processing scrub or voicemail platform connection in progress.'],
    ['200 - Successfully delivered', 'The message was successfully delivered to the voicemail platform.'],
    ['300 - Expired', 'Not delivered - Out of time. Either campaign start time was outside state contact times or record could not be processed before window closed.'],
    ['301 - Canceled', 'Not delivered - Campaign was canceled by user before this record could be attempted.'],
    ['400 - Unsuccessful delivery attempt', 'Unable to connect to the voicemail platform to deliver the message.'],
    ['401 - Not a wireless number', 'DDVM can only deliver to mobile phones. The number provided is not identified as wireless.'],
    ['402 - Duplicate number', 'An identical phone number was in the submitted contact records.'],
    ['403 - Not a valid US number', 'The number cannot be identified as a wireless US number. This category covers two very different cases: (1) obviously invalid placeholders such as 0000000000 that are not real telephone numbers at all, and (2) real, valid phone numbers belonging to contacts outside the United States (e.g. a customer in Portugal with a legitimate Portuguese mobile number). VoApps can only deliver to wireless numbers within the US, so both cases are treated identically — neither will ever receive a message, and neither counts as a delivery attempt. See the Non-Deliverable Records section of the Executive Summary for a breakdown.'],
    ['404 - Undeliverable', 'Phone number unable to be attempted. Number likely too short, too long, or contained an illegal NPA_NXX.'],
    ['405 - Not in service', 'Phone number is not in service and unable to accept voicemails.'],
    ['406 - Voicemail not setup', 'The voicemail box for this phone number has not been setup.'],
    ['407 - Voicemail full', 'The voicemail box is full and cannot accept more messages at this time.'],
    ['408 - Invalid caller number', 'The caller number attempted does not exist in account.'],
    ['409 - Invalid message id', 'The message id attempted does not exist in account.'],
    ['410 - Prohibited self call', 'Caller number and phone number are identical, preventing delivery.'],
    ['500 - Restricted', 'Number was included in a restriction that has since been deleted.'],
    ['501 - Restricted for frequency', 'Number exceeds contact frequency allowed due to client-set frequency restriction.'],
    ['502 - Restricted geographical region', 'Area code is in a state with client-set geographical restriction.'],
    ['503 - Restricted individual number', 'Phone number is prohibited due to client-set individual restriction.'],
    ['504 - Restricted WebRecon', 'Phone number is prohibited due to inclusion on WebRecon Litigious Consumer database.']
  ];

  for (const [term, def] of resultDefs) {
    glossarySheet.getCell(`A${glossRow}`).value = term;
    glossarySheet.getCell(`A${glossRow}`).font = { bold: true };
    glossarySheet.getCell(`B${glossRow}`).value = def;
    glossarySheet.getCell(`B${glossRow}`).style = contentStyle;
    glossarySheet.getRow(glossRow).height = 28;
    glossRow++;
  }

  glossRow++;

  // Best Practices
  glossarySheet.mergeCells(`A${glossRow}:B${glossRow}`);
  glossarySheet.getCell(`A${glossRow}`).value = 'DDVM Delivery Best Practices';
  glossarySheet.getCell(`A${glossRow}`).style = sectionHeaderStyle;
  glossRow++;

  const bestPractices = [
    ['Message Rotation', 'Avoid sending the same message to a number repeatedly. Vary your messages to improve deliverability and engagement.'],
    ['Caller Number Diversity', 'Use multiple caller numbers to reduce the appearance of automated patterns and improve answer rates.'],
    ['Day-of-Week Variety', 'Consumers who receive your message on a predictable day pattern begin to recognize and mentally categorize it as routine — reducing the likelihood they listen or call back. Rotating across additional days of the week, in addition to varying message recordings, makes your outreach feel less automated and more timely. Note: most contact centers do not work Sundays, and some do not work Saturdays.'],
    ['Retry Limits', 'Stop retrying numbers after 4–6 consecutive failures. Success probability drops sharply with each additional attempt on a persistently failing number — continued retries waste carrier capacity and campaign budget with diminishing returns.'],
    ['List Hygiene', 'Regularly suppress numbers classified as Delivery Unlikely to protect caller reputation, improve delivery speed, and maintain clean analytics. Numbers classified as Never Delivered should be permanently removed — these are not reachable by DDVM and will never generate a return on your campaign spend.']
  ];

  for (const [term, def] of bestPractices) {
    glossarySheet.getCell(`A${glossRow}`).value = term;
    glossarySheet.getCell(`A${glossRow}`).font = { bold: true };
    glossarySheet.getCell(`B${glossRow}`).value = def;
    glossarySheet.getCell(`B${glossRow}`).style = contentStyle;
    glossarySheet.getRow(glossRow).height = 35;
    glossRow++;
  }

  glossarySheet.getColumn(1).width = 28;
  glossarySheet.getColumn(2).width = 200;

  // ========================================
  // SAVE WORKBOOK
  // ========================================

  log('Writing Excel file...');
  await workbook.xlsx.writeFile(outputPath);
  log(`Delivery Intelligence Analysis complete: ${path.basename(outputPath)}`);

  // ── Business Review Slides ───────────────────────────────────────────────────
  try {
    const pptxPath    = outputPath.replace(/\.xlsx$/i, '_Business_Review.pptx');
    const squareLogo  = path.join(__dirname, 'assets', 'logo_square.png');
    const circleLogo  = path.join(__dirname, 'assets', 'logo_circle.png');
    const slideAccountIds = Object.keys(accountStats).slice(0, 6);
    await generateBusinessReviewSlides(
      {
        uniqueNumbers,
        totalAttempts,
        totalSuccess:         _totalSuccess,
        overallSuccessRate,
        listGrade,
        healthyCount,
        toxicCount,
        neverDeliveredCount,
        avgVariability,
        decayCurve,
        cadence: {
          cadenceMultiTouchCount,
          cadenceSingleTouch,
          cadenceBucket_sameDay,
          cadenceBucket_1to2,
          cadenceBucket_3to5,
          cadenceBucket_6to10,
          cadenceBucket_11to15,
          cadenceBucket_16to30,
          cadenceBucket_over30,
          cadenceOverallMedian
        },
        actions,
        bestNextAction,
        agentHoursSaved,
        minDate,
        maxDate,
        accountIds: slideAccountIds
      },
      pptxPath,
      null,
      fs.existsSync(squareLogo) ? squareLogo : null,
      fs.existsSync(circleLogo) ? circleLogo : null
    );
    log(`Business review slides saved: ${path.basename(pptxPath)}`);
  } catch (slideErr) {
    log(`[Warning] Business review slides could not be generated: ${slideErr.message}`);
  }

  return {
    totalRecords: totalValidRows,
    uniqueNumbers,
    overallSuccessRate,
    listGrade,
    healthyCount,
    toxicCount,
    neverDeliveredCount,
    avgVariability,
    consecRunsCount: suppressionRuns.length,
    detectedTimezone
  };
}

module.exports = { generateTrendAnalysis, inferMessageIntent };
