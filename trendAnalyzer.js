// trendAnalyzer.js - Delivery Intelligence Report
// Analyzes phone numbers, caller numbers, and messages for delivery insights
// Generates comprehensive Excel analysis workbooks from campaign data
//
// Features:
// - Attempt Index tracking per TN (resets after success)
// - Success Probability by attempt number (decay curve)
// - TN Health Classification (Healthy/Degrading/Toxic)
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
  const name = messageName.toLowerCase();

  if (name.includes('collect') || name.includes('payment') || name.includes('past due') || name.includes('balance') || name.includes('debt')) {
    return 'collections';
  }
  if (name.includes('remind') || name.includes('reminder')) {
    return 'reminder';
  }
  if (name.includes('appt') || name.includes('appointment') || name.includes('schedule')) {
    return 'appointment';
  }
  if (name.includes('callback') || name.includes('call back') || name.includes('return call')) {
    return 'callback';
  }
  if (name.includes('confirm') || name.includes('verification')) {
    return 'confirmation';
  }
  if (name.includes('offer') || name.includes('promo') || name.includes('sale') || name.includes('discount')) {
    return 'marketing';
  }
  if (name.includes('urgent') || name.includes('important') || name.includes('immediate')) {
    return 'urgent';
  }
  if (name.includes('welcome') || name.includes('intro') || name.includes('onboard')) {
    return 'welcome';
  }
  if (name.includes('follow') || name.includes('followup')) {
    return 'followup';
  }
  if (name.includes('loan') || name.includes('servic')) {
    return 'loan servicing';
  }
  return 'general';
}

/**
 * Calculate TN Health classification
 */
function classifyTNHealth(successRate, consecutiveFailures, totalAttempts, recentSuccess14Days) {
  // Toxic: Very low success + high consecutive failures
  if (successRate < 0.1 && consecutiveFailures >= 4) return 'Toxic';
  if (consecutiveFailures >= 6) return 'Toxic';
  if (totalAttempts >= 5 && successRate === 0) return 'Toxic';

  // Degrading: Declining performance
  if (successRate < 0.25 && consecutiveFailures >= 2) return 'Degrading';
  if (successRate < 0.20 && !recentSuccess14Days) return 'Degrading';
  if (consecutiveFailures >= 3) return 'Degrading';

  // Healthy: Good performance
  return 'Healthy';
}

/**
 * Calculate List Quality Grade
 */
function calculateListGrade(healthyPct, degradingPct, toxicPct, neverDeliveredPct) {
  // A: >80% healthy, <5% toxic
  if (healthyPct >= 80 && toxicPct < 5 && neverDeliveredPct < 10) return 'A';
  // B: >60% healthy, <10% toxic
  if (healthyPct >= 60 && toxicPct < 10 && neverDeliveredPct < 20) return 'B';
  // C: >40% healthy, <20% toxic
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

  for (let i = 0; i < 7; i++) {
    if (dayOfWeekCounts[i] > total * 0.1) { // At least 10% on this day
      usedDays.push(dayNames[i]);
    }
  }

  // Check for problematic patterns (only 1-2 days, excluding Sunday)
  const workdays = usedDays.filter(d => d !== 'Sunday');
  if (workdays.length <= 2 && total > 0) {
    return {
      limited: true,
      days: workdays,
      recommendation: `Only used on ${workdays.join(' and ')}. Consider spreading DDVM attempts across more days of the week for better deliverability.`
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
async function generateTrendAnalysis(
  csvInput,
  outputPath,
  minConsecUnsuccessful = 4,
  minRunSpanDays = 30,
  messageMap = {},
  callerMap = {},
  accountTimezones = {},
  userTimezone = 'VoApps',
  userTimezoneLabel = 'VoApps'
) {
  log(`Starting Delivery Intelligence Analysis (v${VERSION})`);

  let csvRows = [];

  // Handle different input types
  if (Array.isArray(csvInput)) {
    if (csvInput.length > 0 && typeof csvInput[0] === 'string') {
      const files = csvInput;
      log(`Processing ${files.length} CSV file(s) with streaming...`);

      for (const csvFile of files) {
        log(`Reading: ${path.basename(csvFile)}`);
        const fileStream = fs.createReadStream(csvFile, { encoding: 'utf8' });

        await new Promise((resolve, reject) => {
          Papa.parse(fileStream, {
            header: true,
            skipEmptyLines: true,
            transformHeader: (h) => h.trim().toLowerCase(),
            chunk: (results) => {
              for (const row of results.data) {
                if (row.number) csvRows.push(row);
              }
            },
            complete: () => {
              log(`  -> Loaded ${csvRows.length.toLocaleString()} rows so far`);
              resolve();
            },
            error: (error) => reject(error)
          });
        });
      }

      log(`Combined ${csvRows.length.toLocaleString()} total records from ${files.length} file(s)`);
    } else {
      csvRows = csvInput;
      log(`Processing ${csvRows.length.toLocaleString()} row objects`);
    }
  } else if (typeof csvInput === 'string') {
    log(`Reading: ${path.basename(csvInput)}`);
    const fileStream = fs.createReadStream(csvInput, { encoding: 'utf8' });

    await new Promise((resolve, reject) => {
      Papa.parse(fileStream, {
        header: true,
        skipEmptyLines: true,
        transformHeader: (h) => h.trim().toLowerCase(),
        chunk: (results) => {
          for (const row of results.data) {
            if (row.number) csvRows.push(row);
          }
        },
        complete: () => resolve(),
        error: (error) => reject(error)
      });
    });

    log(`Loaded ${csvRows.length.toLocaleString()} records`);
  } else {
    throw new Error('Invalid input type for csvInput');
  }

  if (csvRows.length === 0) {
    throw new Error('No data found in CSV input');
  }

  // Normalize and validate rows
  const validRows = csvRows.filter(row => row.number && row.voapps_result && row.voapps_timestamp);
  log(`Valid rows: ${validRows.length.toLocaleString()}`);

  // Detect timezone from timestamps - both global and per-account
  let detectedTimezone = null;
  const timezoneCounts = {};
  const accountResultTimezones = {};  // Track most common offset per account

  for (const row of validRows) {
    const tz = extractTimezone(row.voapps_timestamp);
    if (tz) {
      const key = tz.offset;
      timezoneCounts[key] = (timezoneCounts[key] || 0) + 1;

      // Track per-account timezone from results
      const accountId = row.account_id || 'Unknown';
      if (!accountResultTimezones[accountId]) {
        accountResultTimezones[accountId] = {};
      }
      accountResultTimezones[accountId][key] = (accountResultTimezones[accountId][key] || 0) + 1;
    }
  }

  // Find most common timezone per account
  const accountMostCommonOffset = {};
  for (const [accountId, offsets] of Object.entries(accountResultTimezones)) {
    let maxAccountCount = 0;
    let mostCommon = null;
    for (const [offset, count] of Object.entries(offsets)) {
      if (count > maxAccountCount) {
        maxAccountCount = count;
        mostCommon = offset;
      }
    }
    accountMostCommonOffset[accountId] = mostCommon;
  }

  // Find most common timezone globally
  let maxCount = 0;
  let mostCommonOffset = null;
  for (const [offset, count] of Object.entries(timezoneCounts)) {
    if (count > maxCount) {
      maxCount = count;
      mostCommonOffset = offset;
    }
  }

  if (mostCommonOffset) {
    detectedTimezone = getTimezoneDisplayName({ offset: mostCommonOffset });
  } else {
    detectedTimezone = 'Unknown (timestamps may be UTC)';
  }

  log(`Detected timezone: ${detectedTimezone} (from ${maxCount.toLocaleString()} timestamps)`);

  // Detect timezone discrepancies between account settings and results
  const timezoneDiscrepancies = detectTimezoneDiscrepancies(accountTimezones, accountMostCommonOffset);
  if (timezoneDiscrepancies.length > 0) {
    log(`⚠️  Found ${timezoneDiscrepancies.length} timezone discrepancy(ies) between account settings and results`);
  }

  // Parse timestamps and enrich data
  // VoApps result categories:
  // - DELIVERY ATTEMPTS: Have a timestamp and actually attempted delivery
  //   - "Successfully delivered" = success
  //   - "Unsuccessful delivery attempt" = failed attempt
  //   - "Voicemail not setup" = failed attempt
  //   - "Voicemail full" = failed attempt
  // - NON-ATTEMPTS: No timestamp - filtered out before delivery was attempted
  //   - "Not a wireless number", "Restricted geographical region", "Restricted individual number",
  //   - "Duplicate number", "Expired", "Not a valid US number", "Restricted for frequency", "Undeliverable"
  log('Parsing timestamps and enriching data...');

  for (const row of validRows) {
    row.parsedDate = new Date(row.voapps_timestamp);
    row.voapps_result_normalized = String(row.voapps_result || '').trim().toLowerCase();
    row.isSuccess = row.voapps_result_normalized === 'successfully delivered';
    // A delivery attempt has a timestamp - no timestamp means it was filtered before delivery
    row.isDeliveryAttempt = !!(row.voapps_timestamp && row.voapps_timestamp.trim());
  }

  // Sort by timestamp
  log('Sorting rows by date...');
  validRows.sort((a, b) => a.parsedDate - b.parsedDate);

  // Calculate date range
  let minDate = null, maxDate = null;
  for (const row of validRows) {
    const d = row.parsedDate;
    if (!d || isNaN(d.getTime())) continue;
    if (minDate === null || d < minDate) minDate = d;
    if (maxDate === null || d > maxDate) maxDate = d;
  }
  const fourteenDaysAgo = maxDate ? new Date(maxDate.getTime() - 14 * 24 * 60 * 60 * 1000) : null;

  // ============================================================================
  // BUILD NUMBER-LEVEL DATA WITH ATTEMPT INDEX
  // ============================================================================

  log('Building number-level data with attempt indexing...');
  const numberData = {};

  for (const row of validRows) {
    const num = row.number;

    if (!numberData[num]) {
      numberData[num] = {
        number: num,
        attempts: [],
        attemptIndex: 0,  // Resets after success
        consecutiveFailures: 0,
        totalAttempts: 0,
        successCount: 0,
        unsuccessfulCount: 0,
        lastSuccessTimestamp: null,
        messageIds: {},
        callerNumbers: {},
        accountIds: {},
        hourCounts: new Array(24).fill(0),
        dayOfWeekCounts: new Array(7).fill(0),
        backToBackIdentical: 0,
        lastMessageId: null
      };
    }

    const nd = numberData[num];

    // Only count delivery attempts (not filtered-out results like "Not a wireless number")
    // for success rate calculations
    if (row.isDeliveryAttempt) {
      nd.totalAttempts++;
      nd.attemptIndex++;  // Increment attempt index (resets after success)
    }

    const attempt = {
      timestamp: row.parsedDate,
      result: row.voapps_result_normalized,
      isSuccess: row.isSuccess,
      isDeliveryAttempt: row.isDeliveryAttempt,
      hour: row.parsedDate.getHours(),
      dayOfWeek: row.parsedDate.getDay(),
      message_id: row.message_id || '',
      message_name: row.message_name || '',
      caller_number: row.caller_number || '',
      caller_number_name: row.caller_number_name || '',
      account_id: row.account_id || '',
      campaign_id: row.campaign_id || '',
      campaign_name: row.campaign_name || '',
      attemptIndex: row.isDeliveryAttempt ? nd.attemptIndex : 0
    };

    nd.attempts.push(attempt);

    // Track message usage
    const msgId = attempt.message_id || 'Unknown';
    nd.messageIds[msgId] = (nd.messageIds[msgId] || 0) + 1;

    // Check for back-to-back identical messages
    if (nd.lastMessageId && nd.lastMessageId === msgId) {
      nd.backToBackIdentical++;
    }
    nd.lastMessageId = msgId;

    // Track caller usage
    const callerNum = attempt.caller_number || 'Unknown';
    nd.callerNumbers[callerNum] = (nd.callerNumbers[callerNum] || 0) + 1;

    // Track account usage
    const accountId = attempt.account_id || 'Unknown';
    nd.accountIds[accountId] = (nd.accountIds[accountId] || 0) + 1;

    // Track time distribution (only for delivery attempts)
    if (row.isDeliveryAttempt) {
      nd.hourCounts[attempt.hour]++;
      nd.dayOfWeekCounts[attempt.dayOfWeek]++;
    }

    if (row.isSuccess) {
      nd.successCount++;
      nd.consecutiveFailures = 0;
      nd.attemptIndex = 0;  // Reset attempt index after success
      nd.lastSuccessTimestamp = row.parsedDate;
    } else if (row.isDeliveryAttempt) {
      // Only count as unsuccessful if it was an actual delivery attempt
      nd.unsuccessfulCount++;
      nd.consecutiveFailures++;
    }
  }

  const uniqueNumbers = Object.keys(numberData).length;
  log(`Analyzed ${uniqueNumbers.toLocaleString()} unique numbers`);

  // ============================================================================
  // CALCULATE SUCCESS PROBABILITY BY ATTEMPT INDEX
  // ============================================================================

  log('Calculating success probability by attempt index...');
  const attemptStats = {};  // attemptIndex -> { successful, total }

  for (const num in numberData) {
    for (const attempt of numberData[num].attempts) {
      // Only count actual delivery attempts for success probability
      if (!attempt.isDeliveryAttempt) continue;

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
  // BUILD ACCOUNT AND MESSAGE LEVEL STATS WITH DAY-OF-WEEK ANALYSIS
  // ============================================================================

  log('Building account and message level stats...');

  // Global account stats
  const accountStats = {};
  const messageStats = {};
  const callerStats = {};

  // Global hourly stats
  const globalHourlyStats = {};
  for (let h = 0; h < 24; h++) globalHourlyStats[h] = { successful: 0, unsuccessful: 0, total: 0 };

  for (const num in numberData) {
    for (const attempt of numberData[num].attempts) {
      // Only count delivery attempts for success rate statistics
      if (!attempt.isDeliveryAttempt) continue;

      // Account stats
      const accountId = attempt.account_id || 'Unknown';
      if (!accountStats[accountId]) {
        accountStats[accountId] = {
          account_id: accountId,
          successful: 0,
          unsuccessful: 0,
          total: 0,
          uniqueNumbers: new Set(),
          dayOfWeekCounts: new Array(7).fill(0)
        };
      }
      accountStats[accountId].total++;
      accountStats[accountId].uniqueNumbers.add(num);
      accountStats[accountId].dayOfWeekCounts[attempt.dayOfWeek]++;
      if (attempt.isSuccess) accountStats[accountId].successful++;
      else accountStats[accountId].unsuccessful++;

      // Message stats
      const msgId = attempt.message_id || 'Unknown';
      const msgName = attempt.message_name || messageMap[msgId]?.name || '';
      if (!messageStats[msgId]) {
        messageStats[msgId] = {
          message_id: msgId,
          message_name: msgName,
          intent: inferMessageIntent(msgName),
          successful: 0,
          unsuccessful: 0,
          total: 0,
          uniqueNumbers: new Set(),
          dayOfWeekCounts: new Array(7).fill(0)
        };
      }
      messageStats[msgId].total++;
      messageStats[msgId].uniqueNumbers.add(num);
      messageStats[msgId].dayOfWeekCounts[attempt.dayOfWeek]++;
      if (attempt.isSuccess) messageStats[msgId].successful++;
      else messageStats[msgId].unsuccessful++;

      // Caller stats
      const callerNum = attempt.caller_number || 'Unknown';
      const callerName = attempt.caller_number_name || callerMap[callerNum] || '';
      if (!callerStats[callerNum]) {
        callerStats[callerNum] = {
          caller_number: callerNum,
          caller_name: callerName,
          successful: 0,
          unsuccessful: 0,
          total: 0,
          uniqueNumbers: new Set(),
          dayOfWeekCounts: new Array(7).fill(0)
        };
      }
      callerStats[callerNum].total++;
      callerStats[callerNum].uniqueNumbers.add(num);
      callerStats[callerNum].dayOfWeekCounts[attempt.dayOfWeek]++;
      if (attempt.isSuccess) callerStats[callerNum].successful++;
      else callerStats[callerNum].unsuccessful++;

      // Global hourly stats
      const h = attempt.hour;
      globalHourlyStats[h].total++;
      if (attempt.isSuccess) globalHourlyStats[h].successful++;
      else globalHourlyStats[h].unsuccessful++;
    }
  }

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

  let healthyCount = 0, degradingCount = 0, toxicCount = 0, neverDeliveredCount = 0;
  const numberSummaryArray = [];

  for (const num in numberData) {
    const nd = numberData[num];
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
    else if (tnHealth === 'Degrading') degradingCount++;
    else if (tnHealth === 'Toxic') toxicCount++;

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

    // First and last attempt
    const firstAttempt = nd.attempts[0]?.timestamp;
    const lastAttempt = nd.attempts[nd.attempts.length - 1]?.timestamp;
    const validFirstAttempt = firstAttempt && !isNaN(firstAttempt.getTime()) ? firstAttempt : null;
    const validLastAttempt = lastAttempt && !isNaN(lastAttempt.getTime()) ? lastAttempt : null;

    // Infer intent from top message
    const messageIntent = inferMessageIntent(topMsgName);

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
  const degradingPct = (degradingCount / uniqueNumbers) * 100;
  const toxicPct = (toxicCount / uniqueNumbers) * 100;
  const neverDeliveredPct = (neverDeliveredCount / uniqueNumbers) * 100;
  const listGrade = calculateListGrade(healthyPct, degradingPct, toxicPct, neverDeliveredPct);

  log(`TN Health: Healthy=${healthyCount.toLocaleString()}, Degrading=${degradingCount.toLocaleString()}, Toxic=${toxicCount.toLocaleString()}, Never Delivered=${neverDeliveredCount.toLocaleString()}`);
  log(`List Grade: ${listGrade}`);

  // ============================================================================
  // BUILD CONSECUTIVE UNSUCCESSFUL RUNS
  // ============================================================================

  log('Building consecutive unsuccessful runs...');
  const consecRuns = [];

  for (const num in numberData) {
    const nd = numberData[num];
    const attempts = nd.attempts;
    let currentRun = [];

    for (const attempt of attempts) {
      if (!attempt.isSuccess) {
        currentRun.push(attempt);
      } else {
        // Check if the run meets criteria
        if (currentRun.length >= minConsecUnsuccessful) {
          const runStart = currentRun[0].timestamp;
          const runEnd = currentRun[currentRun.length - 1].timestamp;
          const spanDays = runStart && runEnd ? (runEnd - runStart) / (1000 * 60 * 60 * 24) : 0;

          // Include if span is >= minRunSpanDays OR if we have many consecutive failures
          if (spanDays >= minRunSpanDays || currentRun.length >= 6) {
            const ns = numberSummaryArray.find(n => n.number === num);
            consecRuns.push({
              number: num,
              count: currentRun.length,
              runStart, runEnd, spanDays,
              tnHealth: ns?.tnHealth || 'Unknown'
            });
          }
        }
        currentRun = [];
      }
    }

    // Check final run (if the number ends with consecutive failures)
    if (currentRun.length >= minConsecUnsuccessful) {
      const runStart = currentRun[0].timestamp;
      const runEnd = currentRun[currentRun.length - 1].timestamp;
      const spanDays = runStart && runEnd ? (runEnd - runStart) / (1000 * 60 * 60 * 24) : 0;

      // Include regardless of span if consecutive failures are high
      const ns = numberSummaryArray.find(n => n.number === num);
      consecRuns.push({
        number: num,
        count: currentRun.length,
        runStart, runEnd, spanDays,
        tnHealth: ns?.tnHealth || 'Unknown'
      });
    }
  }

  // Also add all numbers with current consecutive failures >= threshold (that aren't already included)
  for (const ns of numberSummaryArray) {
    if (ns.consecutiveFailures >= minConsecUnsuccessful) {
      // Check if already in consecRuns
      const existing = consecRuns.find(r => r.number === ns.number);
      if (!existing) {
        const nd = numberData[ns.number];
        const recentFailures = nd.attempts.slice(-ns.consecutiveFailures);
        const runStart = recentFailures[0]?.timestamp;
        const runEnd = recentFailures[recentFailures.length - 1]?.timestamp;
        const spanDays = runStart && runEnd ? (runEnd - runStart) / (1000 * 60 * 60 * 24) : 0;

        consecRuns.push({
          number: ns.number,
          count: ns.consecutiveFailures,
          runStart, runEnd, spanDays,
          tnHealth: ns.tnHealth
        });
      }
    }
  }

  consecRuns.sort((a, b) => b.count - a.count);
  log(`  Found ${consecRuns.length.toLocaleString()} consecutive unsuccessful patterns`);

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

  let row = 3;

  // Key Metrics Box
  execSheet.mergeCells(`A${row}:C${row}`);
  execSheet.getCell(`A${row}`).value = 'Key Metrics';
  execSheet.getCell(`A${row}`).style = sectionHeaderStyle;
  row++;

  // Calculate overall success rate (only counting delivery attempts)
  let totalSuccess = 0, totalAttempts = 0;
  for (const r of validRows) {
    if (r.isDeliveryAttempt) {
      totalAttempts++;
      if (r.isSuccess) totalSuccess++;
    }
  }
  const overallSuccessRate = totalAttempts > 0 ? (totalSuccess / totalAttempts * 100) : 0;

  // Calculate average variability score
  const avgVariability = numberSummaryArray.reduce((sum, n) => sum + n.variabilityScore, 0) / numberSummaryArray.length;

  const keyMetrics = [
    ['Total DDVM Attempts', totalAttempts.toLocaleString()],
    ['Unique Phone Numbers', uniqueNumbers.toLocaleString()],
    ['Delivered %', `${overallSuccessRate.toFixed(1)}%`],
    ['Never Delivered %', `${neverDeliveredPct.toFixed(1)}%`],
    ['Average Variability Score', `${avgVariability.toFixed(0)}/100`],
    ['List Quality Grade', listGrade],
    ['Date Range', `${formatDate(minDate)} - ${formatDate(maxDate)}`],
    ['Timezone', detectedTimezone]
  ];

  for (const [label, value] of keyMetrics) {
    execSheet.getCell(`A${row}`).value = label;
    execSheet.getCell(`A${row}`).font = { bold: true };
    execSheet.getCell(`B${row}`).value = value;
    if (label === 'List Quality Grade') {
      execSheet.getCell(`B${row}`).style = listGrade === 'A' || listGrade === 'B' ? successStyle : warningStyle;
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
    ['Healthy', `${healthyCount.toLocaleString()} (${healthyPct.toFixed(1)}%)`],
    ['Degrading', `${degradingCount.toLocaleString()} (${degradingPct.toFixed(1)}%)`],
    ['Toxic', `${toxicCount.toLocaleString()} (${toxicPct.toFixed(1)}%)`],
    ['Never Delivered', `${neverDeliveredCount.toLocaleString()} (${neverDeliveredPct.toFixed(1)}%)`]
  ];

  for (const [label, value] of healthDist) {
    execSheet.getCell(`A${row}`).value = label;
    execSheet.getCell(`A${row}`).font = { bold: true };
    execSheet.getCell(`B${row}`).value = value;
    if (label === 'Toxic' || label === 'Never Delivered') {
      execSheet.getCell(`B${row}`).style = warningStyle;
    } else if (label === 'Healthy') {
      execSheet.getCell(`B${row}`).style = successStyle;
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
    execSheet.getCell(`C${row}`).value = `(${dc.total.toLocaleString()} attempts)`;
    execSheet.getCell(`C${row}`).font = { color: { argb: '666666' } };
    row++;
  }

  row++; // Blank row

  // Immediate Actions
  execSheet.mergeCells(`A${row}:C${row}`);
  execSheet.getCell(`A${row}`).value = 'Recommended Actions';
  execSheet.getCell(`A${row}`).style = sectionHeaderStyle;
  row++;

  const actions = [];

  if (toxicCount > 0) {
    actions.push(`SUPPRESS: ${toxicCount.toLocaleString()} toxic TNs should be removed from DDVM lists immediately.`);
  }
  if (neverDeliveredCount > uniqueNumbers * 0.1) {
    actions.push(`REVIEW: ${neverDeliveredCount.toLocaleString()} numbers have never been successfully delivered. Consider HLR lookup or removal.`);
  }
  if (avgVariability < 40) {
    actions.push(`IMPROVE ROTATION: Average variability score of ${avgVariability.toFixed(0)} is low. Increase message and caller diversity.`);
  }

  // Check for back-to-back issues
  const backToBackIssues = numberSummaryArray.filter(n => n.backToBackIdentical > 2).length;
  if (backToBackIssues > uniqueNumbers * 0.1) {
    actions.push(`MESSAGE FATIGUE: ${backToBackIssues.toLocaleString()} numbers received back-to-back identical messages. Vary your messaging.`);
  }

  // Check day distribution
  const lowDayVariety = numberSummaryArray.filter(n => n.dayEntropy < 0.3 && n.totalAttempts > 2).length;
  if (lowDayVariety > uniqueNumbers * 0.2) {
    actions.push(`WEEKDAY VARIANCE: ${lowDayVariety.toLocaleString()} numbers have DDVM attempts on the same days. Spread attempts across the week.`);
  }

  // Add day-of-week recommendations for accounts/messages
  if (accountDayRecommendations.length > 0) {
    actions.push(`DAY DISTRIBUTION: ${accountDayRecommendations.length} account(s) only send DDVM on limited days. See "Global Insights (Time)" for details.`);
  }
  if (messageDayRecommendations.length > 0) {
    actions.push(`DAY DISTRIBUTION: ${messageDayRecommendations.length} message(s) only used on limited days. See "Global Insights (Time)" for details.`);
  }

  // Add timezone discrepancy warning
  if (timezoneDiscrepancies.length > 0) {
    actions.push(`TIMEZONE SETTINGS: ${timezoneDiscrepancies.length} account(s) have timezone mismatches between account settings and results. Check DirectDrop Voicemail account settings to ensure the "Use account timezone in results file" checkbox is enabled for consistent reporting.`);
  }

  if (actions.length === 0) {
    actions.push('No critical issues detected. Continue monitoring delivery performance.');
  }

  for (const action of actions) {
    execSheet.mergeCells(`A${row}:C${row}`);
    execSheet.getCell(`A${row}`).value = action;
    execSheet.getCell(`A${row}`).style = contentStyle;
    execSheet.getRow(row).height = 25;
    row++;
  }

  row++; // Blank row

  // Client Rationale
  execSheet.mergeCells(`A${row}:C${row}`);
  execSheet.getCell(`A${row}`).value = 'Why This Matters';
  execSheet.getCell(`A${row}`).style = sectionHeaderStyle;
  row++;

  const rationale = [
    'Even though unsuccessful deliveries are not billed, repeated retries consume carrier capacity, delay successful drops, distort analytics, and can degrade caller reputation.',
    'After several consecutive failures, success probability typically drops below 20%. Continuing to retry mostly creates noise rather than results.',
    'Suppressing persistently failing numbers improves delivery speed for good numbers, protects your caller reputation with carriers, and produces cleaner performance data.',
    'Varying messages, caller numbers, and delivery timing creates a more natural pattern that improves deliverability and increases the likelihood that messages are received and heard.'
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
    'Review "TN Health" tab to identify numbers that should be suppressed.',
    'Review "Retry Decay Curve" to understand when retries become unproductive.',
    'Review "Global Insights (Time)" to identify your DDVM delivery patterns to consider a change in strategy.',
    'Review "Variability Analysis" to find numbers that may benefit from more diverse messaging.',
    'Review "Consecutive Unsuccessful" to find long-running failure patterns.'
  ];

  for (const insight of insights) {
    execSheet.mergeCells(`A${row}:C${row}`);
    execSheet.getCell(`A${row}`).value = `• ${insight}`;
    execSheet.getCell(`A${row}`).style = contentStyle;
    row++;
  }

  execSheet.getColumn(1).width = 30;
  execSheet.getColumn(2).width = 25;
  execSheet.getColumn(3).width = 30;

  // ========================================
  // TAB 2: RETRY DECAY CURVE
  // ========================================

  log('Creating Retry Decay Curve tab...');

  const decaySheet = workbook.addWorksheet('Retry Decay Curve', {
    views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }]
  });

  // Add table headers with VoApps styling
  const decayHeaders = ['Attempt Index', 'Total DDVM Attempts', 'Successful', 'Success Probability', 'Insight'];
  decaySheet.getRow(1).values = decayHeaders;
  decaySheet.getRow(1).eachCell((cell) => {
    cell.style = tableHeaderStyle;
  });

  let decayRow = 2;
  for (const dc of decayCurve) {
    let insight = '';
    if (dc.probability >= 0.5) insight = 'Good - Continue';
    else if (dc.probability >= 0.25) insight = 'Declining - Monitor';
    else if (dc.probability >= 0.15) insight = 'Low - Consider suppression';
    else insight = 'Very Low - Suppress';

    decaySheet.getRow(decayRow).values = [
      dc.attemptIndex,
      dc.total,
      dc.successful,
      dc.probability,
      insight
    ];
    decaySheet.getCell(`D${decayRow}`).numFmt = '0.0%';

    if (dc.probability < 0.15) {
      decaySheet.getCell(`E${decayRow}`).style = warningStyle;
    }

    decayRow++;
  }

  decaySheet.getColumn(1).width = 15;
  decaySheet.getColumn(2).width = 20;
  decaySheet.getColumn(3).width = 12;
  decaySheet.getColumn(4).width = 20;
  decaySheet.getColumn(5).width = 25;

  // ========================================
  // TAB 3: TN HEALTH
  // ========================================

  log('Creating TN Health tab...');

  const healthSheet = workbook.addWorksheet('TN Health', {
    views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }]
  });

  const healthHeaders = ['Number', 'TN Health', 'Never Delivered', 'Success Rate', 'Total DDVM Attempts',
    'Consecutive Failures', 'Attempt Index', 'Last Success', 'Variability Score', 'Action'];
  healthSheet.getRow(1).values = healthHeaders;
  healthSheet.getRow(1).eachCell((cell) => {
    cell.style = tableHeaderStyle;
  });

  // Sort by health (Toxic first, then Degrading, then Healthy)
  const healthSorted = [...numberSummaryArray].sort((a, b) => {
    const healthOrder = { 'Toxic': 0, 'Degrading': 1, 'Healthy': 2 };
    return healthOrder[a.tnHealth] - healthOrder[b.tnHealth];
  });

  let healthRow = 2;
  for (const ns of healthSorted) {
    let action = '';
    if (ns.tnHealth === 'Toxic') action = 'Suppress immediately';
    else if (ns.tnHealth === 'Degrading') action = 'Monitor / Consider removal';
    else action = 'Continue';

    healthSheet.getRow(healthRow).values = [
      Number(ns.number),
      ns.tnHealth,
      ns.neverDelivered ? 'Yes' : 'No',
      ns.successRate,
      ns.totalAttempts,
      ns.consecutiveFailures,
      ns.attemptIndex,
      ns.lastSuccessTimestamp,
      ns.variabilityScore,
      action
    ];

    healthSheet.getCell(`D${healthRow}`).numFmt = '0.0%';
    if (ns.lastSuccessTimestamp) {
      healthSheet.getCell(`H${healthRow}`).numFmt = 'yyyy-mm-dd';
    }

    // Color coding
    if (ns.tnHealth === 'Toxic') {
      healthSheet.getCell(`B${healthRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC7CE' } };
    } else if (ns.tnHealth === 'Degrading') {
      healthSheet.getCell(`B${healthRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEB9C' } };
    } else {
      healthSheet.getCell(`B${healthRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'C6EFCE' } };
    }

    if (ns.neverDelivered) {
      healthSheet.getCell(`C${healthRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC7CE' } };
    }

    healthRow++;
  }

  healthSheet.getColumn(1).width = 15;
  healthSheet.getColumn(2).width = 12;
  healthSheet.getColumn(3).width = 14;
  healthSheet.getColumn(4).width = 12;
  healthSheet.getColumn(5).width = 18;
  healthSheet.getColumn(6).width = 18;
  healthSheet.getColumn(7).width = 14;
  healthSheet.getColumn(8).width = 14;
  healthSheet.getColumn(9).width = 16;
  healthSheet.getColumn(10).width = 22;

  // ========================================
  // TAB 4: VARIABILITY ANALYSIS
  // ========================================

  log('Creating Variability Analysis tab...');

  const varSheet = workbook.addWorksheet('Variability Analysis', {
    views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }]
  });

  const varHeaders = ['Number', 'Variability Score', 'Top Msg %', 'Unique Messages', 'Top Caller %',
    'Unique Callers', 'Back-to-Back', 'Day Entropy', 'Hour Entropy', 'Day Distribution'];
  varSheet.getRow(1).values = varHeaders;
  varSheet.getRow(1).eachCell((cell) => {
    cell.style = tableHeaderStyle;
  });

  // Sort by variability score (lowest first to highlight problems)
  const varSorted = [...numberSummaryArray].sort((a, b) => a.variabilityScore - b.variabilityScore);

  let varRow = 2;
  for (const ns of varSorted) {
    varSheet.getRow(varRow).values = [
      Number(ns.number),
      ns.variabilityScore,
      ns.topMsgPct,
      ns.uniqueMsgCount,
      ns.topCallerPct,
      ns.uniqueCallerCount,
      ns.backToBackIdentical,
      ns.dayEntropy,
      ns.hourEntropy,
      ns.dayDistribution
    ];

    varSheet.getCell(`C${varRow}`).numFmt = '0.0%';
    varSheet.getCell(`E${varRow}`).numFmt = '0.0%';
    varSheet.getCell(`H${varRow}`).numFmt = '0.00';
    varSheet.getCell(`I${varRow}`).numFmt = '0.00';

    // Color code variability score
    if (ns.variabilityScore < 30) {
      varSheet.getCell(`B${varRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC7CE' } };
    } else if (ns.variabilityScore < 50) {
      varSheet.getCell(`B${varRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEB9C' } };
    } else {
      varSheet.getCell(`B${varRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'C6EFCE' } };
    }

    // Highlight back-to-back issues
    if (ns.backToBackIdentical > 2) {
      varSheet.getCell(`G${varRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEB9C' } };
    }

    varRow++;
  }

  varSheet.getColumn(1).width = 15;
  varSheet.getColumn(2).width = 16;
  varSheet.getColumn(3).width = 12;
  varSheet.getColumn(4).width = 15;
  varSheet.getColumn(5).width = 13;
  varSheet.getColumn(6).width = 14;
  varSheet.getColumn(7).width = 13;
  varSheet.getColumn(8).width = 12;
  varSheet.getColumn(9).width = 12;
  varSheet.getColumn(10).width = 55;

  // ========================================
  // TAB 5: NUMBER SUMMARY
  // ========================================

  log('Creating Number Summary tab...');

  const summarySheet = workbook.addWorksheet('Number Summary', {
    views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }]
  });

  const summaryHeaders = ['Number', 'Total DDVM Attempts', 'Successful', 'Unsuccessful', 'Success Rate',
    'TN Health', 'Variability', 'First Attempt', 'Last Attempt',
    'Top Msg ID', 'Top Msg Name', 'Top Msg %', 'Unique Msgs',
    'Top Caller', 'Top Caller Name', 'Top Caller %', 'Unique Callers',
    'Intent', 'Day Distribution'];
  summarySheet.getRow(1).values = summaryHeaders;
  summarySheet.getRow(1).eachCell((cell) => {
    cell.style = tableHeaderStyle;
  });

  let sumRow = 2;
  for (const ns of numberSummaryArray) {
    summarySheet.getRow(sumRow).values = [
      Number(ns.number), ns.totalAttempts, ns.successful, ns.unsuccessful, ns.successRate,
      ns.tnHealth, ns.variabilityScore, ns.firstAttempt, ns.lastAttempt,
      Number(ns.topMsgId) || ns.topMsgId, ns.topMsgName, ns.topMsgPct, ns.uniqueMsgCount,
      ns.topCallerNum, ns.topCallerName, ns.topCallerPct, ns.uniqueCallerCount,
      ns.messageIntent, ns.dayDistribution
    ];

    summarySheet.getCell(`E${sumRow}`).numFmt = '0.0%';
    if (ns.firstAttempt) summarySheet.getCell(`H${sumRow}`).numFmt = 'yyyy-mm-dd hh:mm';
    if (ns.lastAttempt) summarySheet.getCell(`I${sumRow}`).numFmt = 'yyyy-mm-dd hh:mm';
    summarySheet.getCell(`L${sumRow}`).numFmt = '0.0%';
    summarySheet.getCell(`P${sumRow}`).numFmt = '0.0%';

    sumRow++;
  }

  // Set column widths
  const widths = [15, 18, 10, 12, 11, 10, 10, 16, 16, 12, 22, 10, 10, 14, 22, 11, 12, 12, 55];
  widths.forEach((w, idx) => summarySheet.getColumn(idx + 1).width = w);

  log(`  Number Summary: ${numberSummaryArray.length.toLocaleString()} rows`);

  // ========================================
  // TAB 6: MESSAGE INSIGHTS
  // ========================================

  log('Creating Message Insights tab...');

  const msgSheet = workbook.addWorksheet('Message Insights', {
    views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }]
  });

  const messageArray = Object.values(messageStats).map(m => ({
    ...m,
    uniqueNumbers: m.uniqueNumbers.size,
    success_rate: m.total > 0 ? m.successful / m.total : 0,
    dayPattern: getDayUsagePattern(m.dayOfWeekCounts)
  })).sort((a, b) => b.total - a.total);

  const msgHeaders = ['Message ID', 'Message Name', 'Intent', 'Total DDVM Attempts', 'Unique Numbers', 'Successful', 'Unsuccessful', 'Success Rate', 'Day Usage', 'Recommendation'];
  msgSheet.getRow(1).values = msgHeaders;
  msgSheet.getRow(1).eachCell((cell) => {
    cell.style = tableHeaderStyle;
  });

  let msgRow = 2;
  for (const msg of messageArray) {
    msgSheet.getRow(msgRow).values = [
      Number(msg.message_id) || msg.message_id, msg.message_name, msg.intent, msg.total, msg.uniqueNumbers,
      msg.successful, msg.unsuccessful, msg.success_rate,
      msg.dayPattern.days.join(', ') || 'All days',
      msg.dayPattern.limited ? msg.dayPattern.recommendation : ''
    ];
    msgSheet.getCell(`H${msgRow}`).numFmt = '0.0%';

    if (msg.dayPattern.limited) {
      msgSheet.getCell(`J${msgRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEB9C' } };
    }
    msgRow++;
  }

  msgSheet.getColumn(1).width = 15;
  msgSheet.getColumn(2).width = 35;
  msgSheet.getColumn(3).width = 14;
  msgSheet.getColumn(4).width = 18;
  msgSheet.getColumn(5).width = 15;
  msgSheet.getColumn(6).width = 12;
  msgSheet.getColumn(7).width = 14;
  msgSheet.getColumn(8).width = 14;
  msgSheet.getColumn(9).width = 20;
  msgSheet.getColumn(10).width = 50;

  // ========================================
  // TAB 7: CALLER # INSIGHTS
  // ========================================

  log('Creating Caller # Insights tab...');

  const callerSheet = workbook.addWorksheet('Caller # Insights', {
    views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }]
  });

  const callerArray = Object.values(callerStats).map(c => ({
    ...c,
    uniqueNumbers: c.uniqueNumbers.size,
    success_rate: c.total > 0 ? c.successful / c.total : 0,
    dayPattern: getDayUsagePattern(c.dayOfWeekCounts)
  })).sort((a, b) => b.total - a.total);

  const callerHeaders = ['Caller Number', 'Caller Name', 'Total DDVM Attempts', 'Unique Numbers', 'Successful', 'Unsuccessful', 'Success Rate', 'Day Usage', 'Recommendation'];
  callerSheet.getRow(1).values = callerHeaders;
  callerSheet.getRow(1).eachCell((cell) => {
    cell.style = tableHeaderStyle;
  });

  let callerRow = 2;
  for (const caller of callerArray) {
    callerSheet.getRow(callerRow).values = [
      caller.caller_number, caller.caller_name, caller.total, caller.uniqueNumbers,
      caller.successful, caller.unsuccessful, caller.success_rate,
      caller.dayPattern.days.join(', ') || 'All days',
      caller.dayPattern.limited ? caller.dayPattern.recommendation : ''
    ];
    callerSheet.getCell(`G${callerRow}`).numFmt = '0.0%';

    if (caller.dayPattern.limited) {
      callerSheet.getCell(`I${callerRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEB9C' } };
    }
    callerRow++;
  }

  callerSheet.getColumn(1).width = 18;
  callerSheet.getColumn(2).width = 30;
  callerSheet.getColumn(3).width = 18;
  callerSheet.getColumn(4).width = 15;
  callerSheet.getColumn(5).width = 12;
  callerSheet.getColumn(6).width = 14;
  callerSheet.getColumn(7).width = 14;
  callerSheet.getColumn(8).width = 20;
  callerSheet.getColumn(9).width = 50;

  // ========================================
  // TAB 8: GLOBAL INSIGHTS (TIME)
  // ========================================

  log('Creating Global Insights (Time) tab...');

  const timeSheet = workbook.addWorksheet('Global Insights (Time)', {
    views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }]
  });

  // Timezone notice at top - show user's selected timezone with label
  // Handle IANA timezone names vs legacy offset format
  let userTzDisplay;
  if (userTimezone === 'VoApps') {
    userTzDisplay = 'VoApps Time (UTC-7, constant)';
  } else if (userTimezone === 'UTC') {
    userTzDisplay = 'UTC';
  } else if (userTimezone.startsWith('America/')) {
    // IANA timezone name - show label with DST note
    userTzDisplay = `${userTimezoneLabel} (DST-aware)`;
  } else if (userTimezone.startsWith('-') || userTimezone.startsWith('+')) {
    // Legacy offset format
    userTzDisplay = userTimezoneLabel ? `${userTimezoneLabel} (UTC${userTimezone})` : `UTC${userTimezone}`;
  } else {
    userTzDisplay = userTimezoneLabel || userTimezone;
  }
  timeSheet.getCell('A1').value = `Report Timezone: ${userTzDisplay}`;
  timeSheet.getCell('A1').font = { bold: true, size: 12, color: { argb: VOAPPS_PURPLE } };
  timeSheet.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: VOAPPS_PURPLE_PALE } };
  timeSheet.mergeCells('A1:E1');

  // Hourly stats
  timeSheet.getCell('A3').value = 'Hourly Success Patterns';
  timeSheet.getCell('A3').font = { bold: true, size: 12 };

  const timeHeaders = ['Hour', 'Total DDVM Attempts', 'Successful', 'Unsuccessful', 'Success Rate'];
  timeSheet.getRow(4).values = timeHeaders;
  timeSheet.getRow(4).eachCell((cell) => {
    cell.style = tableHeaderStyle;
  });

  let timeRow = 5;
  for (let h = 0; h < 24; h++) {
    const stats = globalHourlyStats[h];
    const successRate = stats.total > 0 ? stats.successful / stats.total : 0;
    timeSheet.getRow(timeRow).values = [
      `${String(h).padStart(2, '0')}:00`, stats.total, stats.successful, stats.unsuccessful, successRate
    ];
    timeSheet.getCell(`E${timeRow}`).numFmt = '0.0%';
    timeRow++;
  }

  // Daily stats
  const dayStartRow = timeRow + 2;
  timeSheet.getCell(`A${dayStartRow}`).value = 'Daily Success Patterns';
  timeSheet.getCell(`A${dayStartRow}`).font = { bold: true, size: 12 };

  const dailyStats = {};
  const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  for (let d = 0; d < 7; d++) dailyStats[d] = { successful: 0, unsuccessful: 0, total: 0 };

  for (const num in numberData) {
    for (const attempt of numberData[num].attempts) {
      // Only count delivery attempts for success rate statistics
      if (!attempt.isDeliveryAttempt) continue;

      const d = attempt.dayOfWeek;
      dailyStats[d].total++;
      if (attempt.isSuccess) dailyStats[d].successful++;
      else dailyStats[d].unsuccessful++;
    }
  }

  const dayHeaderRow = dayStartRow + 1;
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

  timeSheet.getColumn(1).width = 15;
  timeSheet.getColumn(2).width = 18;
  timeSheet.getColumn(3).width = 12;
  timeSheet.getColumn(4).width = 14;
  timeSheet.getColumn(5).width = 50;

  // ========================================
  // TAB 9: CONSECUTIVE UNSUCCESSFUL
  // ========================================

  log('Creating Consecutive Unsuccessful tab...');

  const consecSheet = workbook.addWorksheet('Consecutive Unsuccessful', {
    views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }]
  });

  const consecHeaders = ['Number', 'Consecutive Unsuccessful', 'Run Start', 'Run End', 'Span (Days)', 'TN Health', 'Action'];
  consecSheet.getRow(1).values = consecHeaders;
  consecSheet.getRow(1).eachCell((cell) => {
    cell.style = tableHeaderStyle;
  });

  let consecRow = 2;
  for (const run of consecRuns) {
    const action = run.tnHealth === 'Toxic' ? 'Suppress immediately' :
                   run.count >= 6 ? 'Strongly consider removal' : 'Review and consider removal';

    consecSheet.getRow(consecRow).values = [
      Number(run.number), run.count,
      run.runStart && !isNaN(run.runStart.getTime()) ? run.runStart : null,
      run.runEnd && !isNaN(run.runEnd.getTime()) ? run.runEnd : null,
      Math.round(run.spanDays),
      run.tnHealth, action
    ];

    if (run.runStart) consecSheet.getCell(`C${consecRow}`).numFmt = 'yyyy-mm-dd hh:mm';
    if (run.runEnd) consecSheet.getCell(`D${consecRow}`).numFmt = 'yyyy-mm-dd hh:mm';

    if (run.tnHealth === 'Toxic') {
      consecSheet.getCell(`F${consecRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC7CE' } };
    }

    consecRow++;
  }

  log(`  Consecutive Unsuccessful: ${consecRuns.length.toLocaleString()} rows`);

  consecSheet.getColumn(1).width = 15;
  consecSheet.getColumn(2).width = 22;
  consecSheet.getColumn(3).width = 18;
  consecSheet.getColumn(4).width = 18;
  consecSheet.getColumn(5).width = 12;
  consecSheet.getColumn(6).width = 12;
  consecSheet.getColumn(7).width = 25;

  // ========================================
  // TAB 10: GLOSSARY
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
    ['TN Health', 'Classification of phone number health: Healthy (good deliverability), Degrading (declining performance), or Toxic (should be suppressed).'],
    ['Never Delivered', 'A phone number that has never received a successful DDVM delivery across all attempts.'],
    ['Variability Score', 'A 0-100 score measuring how much variety exists in messaging, caller numbers, and timing. Higher scores indicate better rotation practices.'],
    ['Retry Decay Curve', 'Graph showing how success probability decreases with each retry attempt. Used to identify when retries become statistically unproductive.'],
    ['Back-to-Back Identical', 'Count of times the same message was delivered to a number in consecutive attempts. Should be minimized for natural delivery patterns.'],
    ['Day Entropy', 'Measure of how evenly distributed DDVM attempts are across days of the week. Higher entropy (closer to 1.0) means better day-of-week variety.'],
    ['Message Intent', 'Inferred purpose of a message based on its name (e.g., collections, reminder, appointment, callback, welcome, followup, loan servicing).'],
    ['List Quality Grade', 'Overall grade (A-D) for the phone number list based on TN health distribution.']
  ];

  for (const [term, def] of coreConcepts) {
    glossarySheet.getCell(`A${glossRow}`).value = term;
    glossarySheet.getCell(`A${glossRow}`).font = { bold: true };
    glossarySheet.getCell(`B${glossRow}`).value = def;
    glossarySheet.getCell(`B${glossRow}`).style = contentStyle;
    glossarySheet.getRow(glossRow).height = 40;
    glossRow++;
  }

  glossRow++;

  // TN Health Classifications
  glossarySheet.mergeCells(`A${glossRow}:B${glossRow}`);
  glossarySheet.getCell(`A${glossRow}`).value = 'TN Health Classifications';
  glossarySheet.getCell(`A${glossRow}`).style = sectionHeaderStyle;
  glossRow++;

  const healthDefs = [
    ['Healthy', 'Good delivery performance. Success rate typically above 25% with recent successful deliveries. Continue normal operations.'],
    ['Degrading', 'Declining performance. Multiple consecutive failures or low recent success. Monitor closely and consider reducing retry frequency.'],
    ['Toxic', 'Very poor performance. High consecutive failures, very low success rate, or no recent successes. Should be suppressed to protect caller reputation.']
  ];

  for (const [term, def] of healthDefs) {
    glossarySheet.getCell(`A${glossRow}`).value = term;
    glossarySheet.getCell(`A${glossRow}`).font = { bold: true };
    glossarySheet.getCell(`B${glossRow}`).value = def;
    glossarySheet.getCell(`B${glossRow}`).style = contentStyle;
    glossarySheet.getRow(glossRow).height = 35;
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
    ['403 - Not a valid US number', 'Phone number is outside the United States or formatted incorrectly.'],
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
    ['Day-of-Week Variety', 'Distribute DDVM attempts across different days of the week. Most call centers do not work Sundays, and some do not work Saturdays.'],
    ['Hour-of-Day Optimization', 'Analyze success rates by hour to identify when your audience is most likely to receive voicemails.'],
    ['Retry Limits', 'Stop retrying numbers after 4-6 consecutive failures. Success probability drops below 20% and continued retries waste resources.'],
    ['List Hygiene', 'Regularly remove toxic numbers to protect caller reputation, improve delivery speed, and maintain clean analytics.'],
    ['Speech-to-Text (Future)', 'Planned feature: Analyze message content for personalization, urgency, and call-to-action effectiveness.']
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
  glossarySheet.getColumn(2).width = 100;

  // ========================================
  // SAVE WORKBOOK
  // ========================================

  log('Writing Excel file...');
  await workbook.xlsx.writeFile(outputPath);
  log(`Delivery Intelligence Analysis complete: ${path.basename(outputPath)}`);

  return {
    totalRecords: validRows.length,
    uniqueNumbers,
    overallSuccessRate,
    listGrade,
    healthyCount,
    degradingCount,
    toxicCount,
    neverDeliveredCount,
    avgVariability,
    consecRunsCount: consecRuns.length,
    detectedTimezone
  };
}

module.exports = { generateTrendAnalysis };
