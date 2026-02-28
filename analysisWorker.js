// analysisWorker.js - Runs generateTrendAnalysis in a worker thread
// so the Electron main/renderer process stays responsive during heavy Excel writes.

const { workerData, parentPort } = require('worker_threads');
const { generateTrendAnalysis } = require('./trendAnalyzer');

async function run() {
  const { inputData, outputPath, minConsec, minSpan, messageMap, callerMap, accountMap, userTz, userTzLabel } = workerData;

  try {
    await generateTrendAnalysis(
      inputData,
      outputPath,
      minConsec,
      minSpan,
      messageMap || {},
      callerMap || {},
      accountMap || {},
      userTz,
      userTzLabel
    );
    parentPort.postMessage({ ok: true });
  } catch (err) {
    parentPort.postMessage({ ok: false, error: err.message, stack: err.stack });
  }
}

run();
