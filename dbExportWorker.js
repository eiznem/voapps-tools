// dbExportWorker.js
// Runs the DuckDB scan + CSV file writing entirely off the main thread.
// Opens its own DuckDB connection so the main connection is never blocked.

const { workerData, parentPort } = require('worker_threads');
const duckdb = require('duckdb');
const fs = require('fs');
const path = require('path');

async function run() {
  const {
    dbPath, startDate, endDate,
    csvHeaders, outputDir, filePrefix, suffix, maxRowsPerFile
  } = workerData;

  const escapeCsvVal = (val) => {
    if (val === null || val === undefined) return '';
    const str = String(val);
    return str.includes(',') || str.includes('"') || str.includes('\n')
      ? `"${str.replace(/"/g, '""')}"` : str;
  };

  // Open our own connection — read-only so it doesn't conflict with main writer
  const db = await new Promise((resolve, reject) => {
    const instance = new duckdb.Database(dbPath, duckdb.OPEN_READONLY, (err) => {
      if (err) reject(err);
      else resolve(instance);
    });
  });

  try {
    // Count first for progress reporting
    const countRows = await new Promise((resolve, reject) => {
      db.all(
        `SELECT COUNT(*) as cnt FROM campaign_results WHERE target_date >= '${startDate}' AND target_date <= '${endDate}'`,
        (err, rows) => err ? reject(err) : resolve(rows)
      );
    });
    const expectedRows = Number(countRows[0]?.cnt || 0);

    if (expectedRows === 0) {
      db.close();
      parentPort.postMessage({ type: 'error', error: `No data found for ${startDate} to ${endDate}` });
      return;
    }

    parentPort.postMessage({ type: 'progress', text: `\n💾 Streaming ${expectedRows.toLocaleString()} rows from database...` });

    // Set up split CSV output files
    const files = [];
    let currentStream = null;
    let currentFileIndex = 1;
    let currentRowCount = 0;
    let totalRows = 0;
    let lastLoggedPct = 0;

    const openNewFile = () => {
      if (currentStream) currentStream.end();
      const filePath = path.join(outputDir, `${filePrefix}db_analysis_temp_${suffix}_part${currentFileIndex}.csv`);
      currentStream = fs.createWriteStream(filePath, { encoding: 'utf8' });
      currentStream.write(csvHeaders.join(',') + '\n');
      files.push(filePath);
      currentFileIndex++;
      currentRowCount = 0;
    };

    openNewFile();

    const query = `
      SELECT
        number, account_id, account_name,
        campaign_id, campaign_name,
        caller_number, caller_number_name, message_id, message_name, message_description,
        voapps_result, voapps_code, voapps_timestamp, campaign_url
      FROM campaign_results
      WHERE target_date >= '${startDate}' AND target_date <= '${endDate}'
    `;

    await new Promise((resolve, reject) => {
      db.each(query, (err, row) => {
        if (err) return reject(err);
        const line = csvHeaders.map(h => escapeCsvVal(row[h])).join(',') + '\n';
        currentStream.write(line);
        totalRows++;
        currentRowCount++;
        if (currentRowCount >= maxRowsPerFile) openNewFile();

        const pct = Math.floor((totalRows / expectedRows) * 100);
        if (pct >= lastLoggedPct + 10) {
          lastLoggedPct = pct;
          parentPort.postMessage({
            type: 'progress',
            text: `  Streamed ${totalRows.toLocaleString()} / ${expectedRows.toLocaleString()} rows (${pct}%)`
          });
        }
      }, (err) => {
        if (err) return reject(err);
        if (currentStream) currentStream.end();
        resolve();
      });
    });

    db.close();
    parentPort.postMessage({ type: 'done', totalRows, files });

  } catch (err) {
    try { db.close(); } catch (_) {}
    parentPort.postMessage({ type: 'error', error: err.message });
  }
}

run();
