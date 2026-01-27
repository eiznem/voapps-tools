/**
 * VoApps — Number History Trend Analyzer (Node.js Port)
 * Version: 1.0.4
 * Ported from Excel Office Script to Node.js with exceljs
 */

const ExcelJS = require('exceljs');

// Configuration
const BRAND = {
  fontName: "Montserrat",
  fontSize: 11,
  textColor: "333333",
  headerBg: "3F2FB8",
  headerText: "FFFFFF"
};

/**
 * Main function to generate trend analysis Excel file from CSV data
 * @param {Array} csvRows - Array of row objects from CSV
 * @param {string} outputPath - Path where Excel file should be saved
 * @param {number} minConsecUnsuccessful - Minimum consecutive unsuccessful attempts (default: 4)
 * @param {number} minRunSpanDays - Minimum span in days for the run (default: 30)
 */
async function generateTrendAnalysis(csvRows, outputPath, minConsecUnsuccessful = 4, minRunSpanDays = 30) {
  const MIN_BUCKET_ATTEMPTS = 3;
  const MIN_CONSEC_UNSUCCESS = minConsecUnsuccessful;
  const MIN_RUN_SPAN_DAYS = minRunSpanDays;

  const logs = [];
  const log = (m) => logs.push(`[${new Date().toLocaleString()}] ${m}`);
  
  log("===== VoApps — Number History Trend Analyzer v1.0.4 =====");
  log(`Consecutive Unsuccessful rule: >= ${MIN_CONSEC_UNSUCCESS} consecutive "Unsuccessful delivery attempt" spanning >= ${MIN_RUN_SPAN_DAYS} days`);
  log(`Processing ${csvRows.length} rows`);

  // Parse and build records
  const recs = [];
  for (const row of csvRows) {
    const numNorm = normalizePhone(safeToString(row.number || row.phone_number || row.phone));
    if (!numNorm) continue;

    const dt = coerceDate(row.voapps_timestamp || row.timestamp || row.date || row.datetime);
    if (!dt) continue;

    const res = safeToString(row.voapps_result || row.result || row.status).trim();
    const code = safeToString(row.voapps_code || row.code || row.result_code || row.status_code).trim();
    const msg = safeToString(row.message_id).trim();
    const callRaw = safeToString(row.voapps_caller_number || row.caller_number).trim();
    const callNorm = normalizePhone(callRaw);
    const call = callNorm || callRaw;

    recs.push({
      number: numNorm,
      timestamp: dt,
      resultRaw: res,
      codeRaw: code,
      messageId: msg,
      callerNumber: call
    });
  }

  if (recs.length === 0) {
    log("No timestamped rows found");
    throw new Error("No timestamped rows found for analysis");
  }

  log(`Parsed rows (timestamp present): ${recs.length}`);

  // Group by number
  const byNumber = new Map();
  for (const r of recs) {
    const existing = byNumber.get(r.number);
    if (existing) existing.push(r);
    else byNumber.set(r.number, [r]);
  }

  // Sort each number's rows by timestamp
  byNumber.forEach((arr) => arr.sort((a, b) => a.timestamp.getTime() - b.timestamp.getTime()));

  // Per-number summary + consecutive unsuccessful detection
  const numberSummaryRows = [];
  const unsuccessfulBest = new Map();

  byNumber.forEach((arr, num) => {
    const attempts = arr.length;
    const firstDt = arr[0].timestamp;
    const lastDt = arr[arr.length - 1].timestamp;

    // Cadence gaps
    const gaps = [];
    for (let i = 1; i < arr.length; i++) {
      const gapDays = Math.max(0, Math.round((arr[i].timestamp.getTime() - arr[i - 1].timestamp.getTime()) / 86400000));
      gaps.push(gapDays);
    }
    const cadenceStr = buildCadenceString(firstDt, gaps);

    // Gap stats
    const gapCount = gaps.length;
    const gapMin = gapCount ? Math.min(...gaps) : 0;
    const gapMax = gapCount ? Math.max(...gaps) : 0;
    const gapAvg = gapCount ? round2(gaps.reduce((sum, v) => sum + v, 0) / gapCount) : 0;
    const gapMed = gapCount ? round2(median(gaps)) : 0;

    // Success (strict)
    const deliveredCount = arr.filter(r => isDeliveredExact(r.resultRaw)).length;
    const successRatePct = attempts > 0 ? round2((deliveredCount / attempts) * 100) : 0;

    // Attempts per week/month
    const weeksSpan = Math.max(1, diffWeeks(firstDt, lastDt));
    const monthsSpan = Math.max(1, diffMonths(firstDt, lastDt));
    const attemptsPerWeek = round2(attempts / weeksSpan);
    const attemptsPerMonth = round2(attempts / monthsSpan);

    // Top message / caller / combo
    const msgCounts = new Map();
    const callerCounts = new Map();
    const comboCounts = new Map();
    const hourBuckets = new Map();
    const dowBuckets = new Map();

    for (const r of arr) {
      const delivered01 = isDeliveredExact(r.resultRaw) ? 1 : 0;
      const mid = r.messageId || "(blank)";
      const cnum = r.callerNumber || "(blank)";
      const comboKey = `${mid} ⟂ ${cnum}`;

      bumpCounts(msgCounts, mid, delivered01);
      bumpCounts(callerCounts, cnum, delivered01);
      bumpCounts(comboCounts, comboKey, delivered01);
      bumpCounts(hourBuckets, r.timestamp.getHours(), delivered01);
      bumpCounts(dowBuckets, r.timestamp.getDay(), delivered01);
    }

    const topMsg = pickTopKey(msgCounts);
    const topCaller = pickTopKey(callerCounts);
    const topCombo = pickTopKey(comboCounts);
    const topComboRatePct = topCombo ? round2((comboCounts.get(topCombo).delivered / comboCounts.get(topCombo).tot) * 100) : 0;

    // Best/Worst hour and Best DOW
    const bestHour = pickBestBucketByRate(hourBuckets, MIN_BUCKET_ATTEMPTS, true);
    const worstHour = pickBestBucketByRate(hourBuckets, MIN_BUCKET_ATTEMPTS, false);
    const bestDow = pickBestBucketByRate(dowBuckets, MIN_BUCKET_ATTEMPTS, true);

    const bestHourLabel = bestHour ? `${bestHour.key} (${bestHour.ratePct}% on ${bestHour.attempts})` : "(insufficient data)";
    const worstHourLabel = worstHour ? `${worstHour.key} (${worstHour.ratePct}% on ${worstHour.attempts})` : "(insufficient data)";
    const dowName = (d) => ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"][d] || String(d);
    const bestDowLabel = bestDow ? `${dowName(bestDow.key)} (${bestDow.ratePct}% on ${bestDow.attempts})` : "(insufficient data)";

    // Consecutive unsuccessful detection
    const bestRunForNum = findBestConsecutiveUnsuccessfulRun(arr, MIN_CONSEC_UNSUCCESS, MIN_RUN_SPAN_DAYS);
    if (bestRunForNum) unsuccessfulBest.set(num, bestRunForNum);

    numberSummaryRows.push([
      num,
      attempts,
      firstDt,
      lastDt,
      cadenceStr,
      gapCount,
      gapMin,
      gapAvg,
      gapMed,
      gapMax,
      successRatePct,
      attemptsPerWeek,
      attemptsPerMonth,
      topMsg || "",
      topCaller || "",
      topCombo || "",
      topComboRatePct,
      bestHourLabel,
      worstHourLabel,
      bestDowLabel
    ]);
  });

  // Build consecutive unsuccessful table
  const unsuccessfulRows = [];
  const unsuccessfulAll = Array.from(unsuccessfulBest.entries()).map(([num, run]) => ({ num, run }));
  unsuccessfulAll.sort((a, b) => {
    if (b.run.runLength !== a.run.runLength) return b.run.runLength - a.run.runLength;
    if (b.run.spanDays !== a.run.spanDays) return b.run.spanDays - a.run.spanDays;
    return b.run.end.getTime() - a.run.end.getTime();
  });

  for (const item of unsuccessfulAll) {
    const num = item.num;
    const run = item.run;
    const allAttemptsForNum = byNumber.get(num)?.length ?? run.runLength;
    const overallDelivered = (byNumber.get(num) || []).filter(r => isDeliveredExact(r.resultRaw)).length;
    const overallSuccessPct = allAttemptsForNum > 0 ? round2((overallDelivered / allAttemptsForNum) * 100) : 0;

    unsuccessfulRows.push([
      num,
      run.runLength,
      run.start,
      run.end,
      run.spanDays,
      allAttemptsForNum,
      overallSuccessPct,
      run.lastMessageId || "(blank)",
      run.lastCallerNumber || "(blank)"
    ]);
  }

  // Global insights
  const hourMap = new Map();
  const dowMap = new Map();
  const msgMap = new Map();
  const callerMap = new Map();

  for (const r of recs) {
    const delivered01 = isDeliveredExact(r.resultRaw) ? 1 : 0;
    bumpCounts(hourMap, r.timestamp.getHours(), delivered01);
    bumpCounts(dowMap, r.timestamp.getDay(), delivered01);
    bumpCounts(msgMap, r.messageId || "(blank)", delivered01);
    bumpCounts(callerMap, r.callerNumber || "(blank)", delivered01);
  }

  const hourRows = Array.from(hourMap.keys()).sort((a, b) => a - b).map(h => {
    const v = hourMap.get(h);
    return [h, v.tot, v.delivered, v.tot ? round2((v.delivered / v.tot) * 100) : 0];
  });

  const dowRows = Array.from(dowMap.keys()).sort((a, b) => a - b).map(d => {
    const v = dowMap.get(d);
    const name = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"][d] || String(d);
    return [`${d} (${name})`, v.tot, v.delivered, v.tot ? round2((v.delivered / v.tot) * 100) : 0];
  });

  const msgRows = Array.from(msgMap.entries()).sort((a, b) => b[1].tot - a[1].tot).map(([k, v]) => {
    return [k, v.tot, v.delivered, v.tot ? round2((v.delivered / v.tot) * 100) : 0];
  });

  const callerRows = Array.from(callerMap.entries()).sort((a, b) => b[1].tot - a[1].tot).map(([k, v]) => {
    return [k, v.tot, v.delivered, v.tot ? round2((v.delivered / v.tot) * 100) : 0];
  });

  // Create Excel workbook
  const workbook = new ExcelJS.Workbook();
  
  log("Creating Excel workbook with 5 sheets...");

  // Sheet 1: Number Summary
  createNumberSummarySheet(workbook, numberSummaryRows);

  // Sheet 2: Consecutive Unsuccessful
  createConsecutiveUnsuccessfulSheet(workbook, unsuccessfulRows, MIN_CONSEC_UNSUCCESS, MIN_RUN_SPAN_DAYS);

  // Sheet 3: Global Insights (Time)
  createGlobalInsightsTimeSheet(workbook, hourRows, dowRows);

  // Sheet 4: Global Insights (Msg & Caller)
  createGlobalInsightsMsgCallerSheet(workbook, msgRows, callerRows);

  // Sheet 5: Analyzer Log
  createLogSheet(workbook, logs);

  // Save workbook
  await workbook.xlsx.writeFile(outputPath);
  log(`Excel file saved: ${outputPath}`);
  
  return { success: true, logs };
}

// Helper functions
function safeToString(v) {
  return v === null || v === undefined ? "" : String(v);
}

function normalizePhone(s) {
  const digits = (s || "").replace(/\D+/g, "");
  if (digits.length === 11 && digits.startsWith("1")) return digits.slice(1);
  if (digits.length === 10) return digits;
  return digits.length >= 7 ? digits : "";
}

function coerceDate(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (typeof v === "string" && v.trim()) {
    const d = new Date(v);
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

function isDeliveredExact(resultRaw) {
  return (resultRaw || "").trim() === "Successfully delivered";
}

function isUnsuccessfulExact(resultRaw) {
  return (resultRaw || "").trim() === "Unsuccessful delivery attempt";
}

function buildCadenceString(first, gaps) {
  const base = mdy(first);
  if (!gaps || gaps.length === 0) return base;
  return `${base} | ${gaps.map(g => `+${g}`).join(" | ")}`;
}

function mdy(d) {
  const m = d.getMonth() + 1;
  const day = d.getDate();
  const y = d.getFullYear();
  return `${m}/${day}/${y}`;
}

function median(arr) {
  const a = [...arr].sort((x, y) => x - y);
  const n = a.length;
  if (n === 0) return 0;
  const mid = Math.floor(n / 2);
  return n % 2 ? a[mid] : (a[mid - 1] + a[mid]) / 2;
}

function round2(n) {
  return Math.round(n * 100) / 100;
}

function diffWeeks(a, b) {
  const ms = Math.abs(b.getTime() - a.getTime());
  const weeks = ms / (7 * 86400000);
  return Math.ceil(weeks);
}

function diffMonths(a, b) {
  const start = new Date(a.getFullYear(), a.getMonth(), 1);
  const end = new Date(b.getFullYear(), b.getMonth(), 1);
  const months = (end.getFullYear() - start.getFullYear()) * 12 + (end.getMonth() - start.getMonth());
  return Math.max(1, months + 1);
}

function spanDaysFloor(a, b) {
  const ms = b.getTime() - a.getTime();
  return Math.max(0, Math.floor(ms / 86400000));
}

function bumpCounts(m, key, delivered01) {
  const cur = m.get(key);
  if (cur) {
    cur.tot += 1;
    cur.delivered += delivered01;
  } else {
    m.set(key, { tot: 1, delivered: delivered01 });
  }
}

function pickTopKey(m) {
  let bestKey = null;
  let bestTot = -1;
  let bestDelivered = -1;

  for (const [k, v] of m.entries()) {
    if (v.tot > bestTot) {
      bestKey = k; bestTot = v.tot; bestDelivered = v.delivered;
    } else if (v.tot === bestTot) {
      if (v.delivered > bestDelivered) {
        bestKey = k; bestDelivered = v.delivered;
      } else if (v.delivered === bestDelivered && bestKey !== null) {
        if (String(k).toLowerCase() < String(bestKey).toLowerCase()) bestKey = k;
      }
    }
  }
  return bestKey;
}

function pickBestBucketByRate(buckets, minAttempts, best) {
  const keys = Array.from(buckets.keys()).sort((a, b) => a - b);
  let chosen = null;

  for (const k of keys) {
    const v = buckets.get(k);
    if (v.tot < minAttempts) continue;
    const ratePct = round2((v.delivered / v.tot) * 100);

    if (!chosen) {
      chosen = { key: k, attempts: v.tot, ratePct };
      continue;
    }

    if (best) {
      if (ratePct > chosen.ratePct) chosen = { key: k, attempts: v.tot, ratePct };
      else if (ratePct === chosen.ratePct && v.tot > chosen.attempts) chosen = { key: k, attempts: v.tot, ratePct };
    } else {
      if (ratePct < chosen.ratePct) chosen = { key: k, attempts: v.tot, ratePct };
      else if (ratePct === chosen.ratePct && v.tot > chosen.attempts) chosen = { key: k, attempts: v.tot, ratePct };
    }
  }

  return chosen;
}

function findBestConsecutiveUnsuccessfulRun(arr, minConsec, minSpanDays) {
  let runStartIdx = -1;
  let runLen = 0;
  let best = null;

  const considerRun = (startIdx, endIdx) => {
    const start = arr[startIdx].timestamp;
    const end = arr[endIdx].timestamp;
    const spanDays = spanDaysFloor(start, end);
    const length = endIdx - startIdx + 1;

    if (length < minConsec) return;
    if (spanDays < minSpanDays) return;

    const candidate = {
      runLength: length,
      start,
      end,
      spanDays,
      lastMessageId: arr[endIdx].messageId || "",
      lastCallerNumber: arr[endIdx].callerNumber || ""
    };

    if (!best) {
      best = candidate;
      return;
    }
    if (candidate.runLength > best.runLength) { best = candidate; return; }
    if (candidate.runLength === best.runLength && candidate.spanDays > best.spanDays) { best = candidate; return; }
    if (candidate.runLength === best.runLength && candidate.spanDays === best.spanDays && candidate.end.getTime() > best.end.getTime()) {
      best = candidate; return;
    }
  };

  for (let i = 0; i < arr.length; i++) {
    const isUnsucc = isUnsuccessfulExact(arr[i].resultRaw);

    if (isUnsucc) {
      if (runLen === 0) runStartIdx = i;
      runLen += 1;
    } else {
      if (runLen > 0 && runStartIdx >= 0) {
        considerRun(runStartIdx, i - 1);
      }
      runStartIdx = -1;
      runLen = 0;
    }
  }

  if (runLen > 0 && runStartIdx >= 0) {
    considerRun(runStartIdx, arr.length - 1);
  }

  return best;
}

// Excel sheet creation functions
function createNumberSummarySheet(workbook, rows) {
  const sheet = workbook.addWorksheet("Number Summary");
  const headers = [
    "Number", "Attempts", "First Attempt", "Last Attempt", "Cadence (First | +days ...)",
    "Gap Count", "Gap Min (days)", "Gap Avg (days)", "Gap Median (days)", "Gap Max (days)",
    "Success Rate (%)", "Attempts / Week", "Attempts / Month", "Top Message ID", "Top Caller Number",
    "Top (Message ⟂ Caller) Combo", "Top Combo Delivered Rate (%)", "Best Hour (rate% on N, min 3)",
    "Worst Hour (rate% on N, min 3)", "Best Day-of-Week (rate% on N, min 3)"
  ];

  sheet.addRow(headers);
  rows.forEach(row => sheet.addRow(row));

  styleTable(sheet, headers.length, rows.length);
  
  // Format columns
  sheet.getColumn(3).numFmt = 'm/d/yyyy'; // First Attempt
  sheet.getColumn(4).numFmt = 'm/d/yyyy'; // Last Attempt
  sheet.getColumn(5).numFmt = '@'; // Cadence as text
  sheet.getColumn(11).numFmt = '0.0%'; // Success Rate
  sheet.getColumn(17).numFmt = '0.0%'; // Top Combo Rate
  
  // Convert percentages from 0-100 to 0-1
  for (let i = 2; i <= rows.length + 1; i++) {
    const successRate = sheet.getCell(i, 11).value;
    if (typeof successRate === 'number') {
      sheet.getCell(i, 11).value = successRate / 100;
    }
    const comboRate = sheet.getCell(i, 17).value;
    if (typeof comboRate === 'number') {
      sheet.getCell(i, 17).value = comboRate / 100;
    }
  }

  sheet.columns.forEach(col => col.width = 15);
  sheet.getRow(1).height = 20;
}

function createConsecutiveUnsuccessfulSheet(workbook, rows, minConsec, minSpanDays) {
  const sheet = workbook.addWorksheet("Consecutive Unsuccessful");
  const headers = [
    "Number",
    `Consecutive "Unsuccessful delivery attempt" (>=${minConsec})`,
    "Run Start",
    "Run End",
    `Run Span (days) (>=${minSpanDays})`,
    "Total Attempts (All)",
    "Overall Success Rate (%)",
    "Last Message ID in Run",
    "Last Caller # in Run"
  ];

  sheet.addRow(headers);
  rows.forEach(row => sheet.addRow(row));

  styleTable(sheet, headers.length, rows.length);
  
  sheet.getColumn(3).numFmt = 'm/d/yyyy'; // Run Start
  sheet.getColumn(4).numFmt = 'm/d/yyyy'; // Run End
  sheet.getColumn(7).numFmt = '0.0%'; // Success Rate
  
  // Convert percentages
  for (let i = 2; i <= rows.length + 1; i++) {
    const successRate = sheet.getCell(i, 7).value;
    if (typeof successRate === 'number') {
      sheet.getCell(i, 7).value = successRate / 100;
    }
  }

  sheet.columns.forEach(col => col.width = 15);
  sheet.getRow(1).height = 20;
}

function createGlobalInsightsTimeSheet(workbook, hourRows, dowRows) {
  const sheet = workbook.addWorksheet("Global Insights (Time)");
  
  const hourHeaders = ["Hour (0-23)", "Attempts", "Delivered", "Success Rate (%)"];
  sheet.addRow(hourHeaders);
  hourRows.forEach(row => sheet.addRow(row));
  
  const hourEndRow = hourRows.length + 1;
  styleTable(sheet, hourHeaders.length, hourRows.length, 1);
  sheet.getColumn(4).numFmt = '0.0%';
  
  // Convert percentages
  for (let i = 2; i <= hourEndRow; i++) {
    const rate = sheet.getCell(i, 4).value;
    if (typeof rate === 'number') {
      sheet.getCell(i, 4).value = rate / 100;
    }
  }
  
  // Add spacer row
  const spacerRow = hourEndRow + 2;
  
  // Day of week table
  const dowHeaders = ["DayOfWeek", "Attempts", "Delivered", "Success Rate (%)"];
  sheet.addRow([]);
  sheet.addRow(dowHeaders);
  dowRows.forEach(row => sheet.addRow(row));
  
  styleTable(sheet, dowHeaders.length, dowRows.length, spacerRow);
  sheet.getColumn(4).numFmt = '0.0%';
  
  // Convert percentages
  for (let i = spacerRow + 1; i <= spacerRow + dowRows.length; i++) {
    const rate = sheet.getCell(i, 4).value;
    if (typeof rate === 'number') {
      sheet.getCell(i, 4).value = rate / 100;
    }
  }

  sheet.columns.forEach(col => col.width = 15);
}

function createGlobalInsightsMsgCallerSheet(workbook, msgRows, callerRows) {
  const sheet = workbook.addWorksheet("Global Insights (Msg & Caller)");
  
  const msgHeaders = ["Message ID", "Attempts", "Delivered", "Success Rate (%)"];
  sheet.addRow(msgHeaders);
  msgRows.forEach(row => sheet.addRow(row));
  
  const msgEndRow = msgRows.length + 1;
  styleTable(sheet, msgHeaders.length, msgRows.length, 1);
  sheet.getColumn(4).numFmt = '0.0%';
  
  // Convert percentages
  for (let i = 2; i <= msgEndRow; i++) {
    const rate = sheet.getCell(i, 4).value;
    if (typeof rate === 'number') {
      sheet.getCell(i, 4).value = rate / 100;
    }
  }
  
  // Add spacer row
  const spacerRow = msgEndRow + 2;
  
  // Caller table
  const callerHeaders = ["Caller Number", "Attempts", "Delivered", "Success Rate (%)"];
  sheet.addRow([]);
  sheet.addRow(callerHeaders);
  callerRows.forEach(row => sheet.addRow(row));
  
  styleTable(sheet, callerHeaders.length, callerRows.length, spacerRow);
  sheet.getColumn(4).numFmt = '0.0%';
  
  // Convert percentages
  for (let i = spacerRow + 1; i <= spacerRow + callerRows.length; i++) {
    const rate = sheet.getCell(i, 4).value;
    if (typeof rate === 'number') {
      sheet.getCell(i, 4).value = rate / 100;
    }
  }

  sheet.columns.forEach(col => col.width = 15);
}

function createLogSheet(workbook, logs) {
  const sheet = workbook.addWorksheet("Analyzer Log");
  sheet.addRow(["Log"]);
  logs.forEach(line => sheet.addRow([line]));
  
  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: BRAND.headerText }, name: BRAND.fontName, size: BRAND.fontSize };
  headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: BRAND.headerBg } };
  
  sheet.getColumn(1).width = 100;
}

function styleTable(sheet, colCount, rowCount, startRow = 1) {
  const headerRow = sheet.getRow(startRow);
  
  // Style header
  headerRow.font = { bold: true, color: { argb: BRAND.headerText }, name: BRAND.fontName, size: BRAND.fontSize };
  headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: BRAND.headerBg } };
  headerRow.alignment = { vertical: 'middle', horizontal: 'left' };
  
  // Style data rows
  for (let i = startRow + 1; i <= startRow + rowCount; i++) {
    const row = sheet.getRow(i);
    row.font = { name: BRAND.fontName, size: BRAND.fontSize, color: { argb: BRAND.textColor } };
    
    // Alternating row colors
    if (i % 2 === 0) {
      row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F2F2F2' } };
    }
  }
  
  // Freeze header row
  sheet.views = [{ state: 'frozen', xSplit: 0, ySplit: startRow }];
}

module.exports = { generateTrendAnalysis };