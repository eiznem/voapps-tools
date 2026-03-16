// analysisWorker.js - Runs generateTrendAnalysis in a worker thread
// so the Electron main/renderer process stays responsive during heavy Excel writes.

const { workerData, parentPort } = require('worker_threads');
const { generateTrendAnalysis } = require('./trendAnalyzer');

async function run() {
  const { inputData, outputPath, minConsec, minSpan, messageMap, callerMap, accountMap, userTz, userTzLabel, includeDetailTabs = false, transcriptMap = {}, includeReAttemptTabs = false, pptxOptions = {}, includeSuppressionCandidates = true, jobId = null } = workerData;

  // Forward named progress stages back to the main thread so server.js can relay them via SSE
  const progressCallback = jobId
    ? (message) => parentPort.postMessage({ type: 'progress', message })
    : null;

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
      userTzLabel,
      includeDetailTabs,
      transcriptMap,
      includeReAttemptTabs,
      includeSuppressionCandidates,
      pptxOptions,
      progressCallback
    );
    parentPort.postMessage({ ok: true });
  } catch (err) {
    parentPort.postMessage({ ok: false, error: err.message, stack: err.stack });
  }
}

run();
