'use strict';
/**
 * aiTranscribeChild.js – isolated child process for Whisper inference.
 *
 * Spawned by server.js via child_process.fork() so that if ONNX Runtime
 * crashes (SIGTRAP, OOM, etc.) it only kills this process — not Electron.
 *
 * Protocol:
 *   Parent → child:  { audioPath, variant }
 *   Child  → parent: { ok: true,  transcript, durationSec }
 *                 or { ok: false, error: '<message>' }
 *   Log lines:        { log: '<message>' }
 */

const fs   = require('fs');
const path = require('path');
const fsp  = fs.promises;

// ── ESM import shim (same pattern as server.js) ──────────────────────────────
const esImport = new Function('p', 'return import(p)');

// ── Module-level Xenova reference (loaded once per process) ──────────────────
let _xenovaMod = null;

// ── Helpers ───────────────────────────────────────────────────────────────────

function xenovaPkgDir() {
  const inAsar = __dirname.includes('app.asar');
  if (inAsar) {
    const unpackedDir = path.join(process.resourcesPath, 'app.asar.unpacked',
      'node_modules', '@xenova', 'transformers');
    if (fs.existsSync(path.join(unpackedDir, 'package.json'))) return unpackedDir;
    const resourcesDir = path.join(process.resourcesPath,
      'node_modules', '@xenova', 'transformers');
    if (fs.existsSync(path.join(resourcesDir, 'package.json'))) return resourcesDir;
  }
  return path.join(__dirname, 'node_modules', '@xenova', 'transformers');
}

function xenovaCacheDir() {
  if (__dirname.includes('app.asar')) {
    try {
      const { app } = require('electron');
      return path.join(app.getPath('userData'), 'models') + path.sep;
    } catch (_) {}
  }
  return path.join(xenovaPkgDir(), '.cache') + path.sep;
}

function xenovaInstallDir() {
  return __dirname.includes('app.asar') ? process.resourcesPath : __dirname;
}

function sendLog(msg) {
  if (process.send) process.send({ log: msg });
  else console.log(msg);
}

function installNpmPackage(pkgName) {
  return new Promise((resolve, reject) => {
    const { spawn } = require('child_process');
    const npmBin = process.platform === 'win32' ? 'npm.cmd' : 'npm';
    sendLog(`[AI] Installing ${pkgName}…`);
    const child = spawn(npmBin, ['install', pkgName, '--no-audit', '--no-fund'], {
      cwd: xenovaInstallDir(), stdio: ['ignore', 'pipe', 'pipe']
    });
    child.stdout.on('data', d => d.toString().split('\n').filter(l => l.trim()).forEach(l => sendLog(`[npm] ${l}`)));
    child.stderr.on('data', d => d.toString().split('\n').filter(l => l.trim()).forEach(l => sendLog(`[npm] ${l}`)));
    child.on('close', code => code === 0 ? resolve() : reject(new Error(`npm install exited with code ${code}`)));
    child.on('error', reject);
  });
}

async function getXenovaMod() {
  if (_xenovaMod) return _xenovaMod;
  const { pathToFileURL } = require('url');
  let mod;
  try {
    mod = await esImport('@xenova/transformers');
  } catch (e) {
    const pkgDir = xenovaPkgDir();
    const pkgJson = JSON.parse(await fsp.readFile(path.join(pkgDir, 'package.json'), 'utf8'));
    const exp = pkgJson.exports?.['.'];
    const mainFile = (typeof exp === 'string' ? exp
      : exp?.import || exp?.default || exp?.require
        || pkgJson.module || pkgJson.main || 'src/transformers.js');
    const entryPath = path.join(pkgDir, mainFile.replace(/^\.\//, ''));
    mod = await esImport(pathToFileURL(entryPath).href);
  }
  if (mod.env?.onnx !== undefined) mod.env.onnx = { logSeverityLevel: 3 };
  else if (mod.env) mod.env.onnx = { logSeverityLevel: 3 };
  if (mod.env) {
    mod.env.cacheDir = xenovaCacheDir();
    sendLog(`[AI] Cache dir: ${mod.env.cacheDir}`);
  }
  _xenovaMod = mod;
  return mod;
}

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

function decodeWavPcm(buffer) {
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
      numChannels   = buffer.readUInt16LE(offset + 2);
      sampleRate    = buffer.readUInt32LE(offset + 4);
      bitsPerSample = buffer.readUInt16LE(offset + 14);
      if (audioFormat !== 1) throw new Error(`WAV audio format ${audioFormat} not supported (only PCM=1).`);
    } else if (chunkId === 'data') {
      dataOffset = offset;
      dataSize   = chunkSize;
      break;
    }
    offset += chunkSize + (chunkSize & 1);
  }
  if (!sampleRate || !dataOffset) throw new Error('WAV file missing fmt or data chunk');
  const bytesPerSample    = bitsPerSample >> 3;
  const frameSize         = numChannels * bytesPerSample;
  const samplesPerChannel = Math.floor(dataSize / frameSize);
  sendLog(`[AI]   🔍 Audio decoded: ${sampleRate} Hz, ${samplesPerChannel} samples, ${(samplesPerChannel / sampleRate).toFixed(1)}s (WAV PCM ${bitsPerSample}-bit, ${numChannels}ch)`);
  const scale = bitsPerSample === 8 ? 128 : Math.pow(2, bitsPerSample - 1);
  const out   = new Float32Array(samplesPerChannel);
  for (let i = 0; i < samplesPerChannel; i++) {
    const basePos = dataOffset + i * frameSize;
    let sum = 0;
    for (let ch = 0; ch < numChannels; ch++) {
      const pos = basePos + ch * bytesPerSample;
      let val;
      if (bitsPerSample === 8)       { val = buffer.readUInt8(pos) - 128; }
      else if (bitsPerSample === 16) { val = buffer.readInt16LE(pos); }
      else if (bitsPerSample === 24) {
        let u = buffer.readUInt8(pos) | (buffer.readUInt8(pos + 1) << 8) | (buffer.readUInt8(pos + 2) << 16);
        if (u & 0x800000) u |= 0xFF000000;
        val = u | 0;
      } else if (bitsPerSample === 32) { val = buffer.readInt32LE(pos); }
      else { val = 0; }
      sum += val;
    }
    out[i] = (sum / numChannels) / scale;
  }
  return resampleTo16kHz(out, sampleRate);
}

async function decodeAudio(audioPath) {
  const fileBuffer = fs.readFileSync(audioPath);
  const magic = fileBuffer.slice(0, 4).toString('ascii');
  if (magic === 'RIFF') {
    sendLog(`[AI]   🔍 Detected WAV format — using native PCM decoder`);
    return decodeWavPcm(fileBuffer);
  }
  // MP3 path
  const { pathToFileURL } = require('url');
  let MPEGDecoder;
  try {
    ({ MPEGDecoder } = await esImport('mpg123-decoder'));
  } catch (e) {
    if (e.code === 'ERR_MODULE_NOT_FOUND' || e.message?.includes('Cannot find')) {
      await installNpmPackage('mpg123-decoder');
      const pkgDir  = path.join(__dirname, 'node_modules', 'mpg123-decoder');
      const pkgJson = JSON.parse(await fsp.readFile(path.join(pkgDir, 'package.json'), 'utf8'));
      const exp = pkgJson.exports?.['.'];
      const mainFile = (typeof exp === 'string' ? exp : exp?.import || exp?.default || exp?.require || pkgJson.module || pkgJson.main || 'src/mpg123-decoder.js');
      const entryPath = path.join(pkgDir, mainFile.replace(/^\.\//, ''));
      ({ MPEGDecoder } = await esImport(pathToFileURL(entryPath).href));
    } else { throw e; }
  }
  const decoder = new MPEGDecoder();
  await decoder.ready;
  const { channelData, samplesDecoded, sampleRate } = decoder.decode(fileBuffer);
  decoder.free();
  sendLog(`[AI]   🔍 Audio decoded: ${sampleRate} Hz, ${samplesDecoded} samples, ${(samplesDecoded / (sampleRate || 1)).toFixed(1)}s (MP3)`);
  if (!sampleRate || samplesDecoded === 0) throw new Error('MP3 decode produced no samples');
  let samples;
  if (Array.isArray(channelData) && channelData.length > 1) {
    const L = channelData[0], R = channelData[1];
    samples = new Float32Array(L.length);
    for (let i = 0; i < L.length; i++) samples[i] = (L[i] + R[i]) * 0.5;
  } else {
    samples = Array.isArray(channelData) ? channelData[0] : channelData;
  }
  return resampleTo16kHz(samples, sampleRate);
}

function isWhisperHallucination(text) {
  if (!text || !text.trim()) return true;
  const stripped = text
    .replace(/\[[\w\s]+\]/gi, '')
    .replace(/\([\w\s]+\)/gi, '')
    .replace(/\s+/g, ' ').trim();
  return !stripped || stripped.length < 3;
}

function stitchTranscriptSegments(texts) {
  if (!texts.length) return '';
  let result = texts[0];
  for (let i = 1; i < texts.length; i++) {
    const next = (texts[i] || '').trim();
    if (!next) continue;
    const rWords = result.trim().split(/\s+/);
    const nWords = next.split(/\s+/);
    const norm = (w) => w.toLowerCase().replace(/[^a-z0-9]/g, '');
    const rNorm = rWords.map(norm);
    const nNorm = nWords.map(norm);
    let stitchAt = 0;
    const maxSearch = Math.min(rWords.length, nWords.length, 15);
    outer: for (let len = maxSearch; len >= 2; len--) {
      const rTail = rNorm.slice(-len);
      for (let offset = 0; offset <= Math.min(nNorm.length - len, 5); offset++) {
        const nChunk = nNorm.slice(offset, offset + len);
        if (rTail.every((w, j) => w === nChunk[j])) { stitchAt = offset + len; break outer; }
      }
    }
    result = stitchAt > 0
      ? result + ' ' + nWords.slice(stitchAt).join(' ')
      : result + ' ' + next;
  }
  return result.trim();
}

// ── Main transcription logic ──────────────────────────────────────────────────

async function run(audioPath, variant) {
  const audioData = await decodeAudio(audioPath);
  const audioDurationSec = audioData.length / 16000;
  const peakAmp = audioData.reduce((m, v) => Math.max(m, Math.abs(v)), 0);
  sendLog(`[AI]   🔍 Peak amplitude (full resampled audio): ${peakAmp.toFixed(4)}`);
  sendLog(`[AI]   🔍 Audio duration: ${audioDurationSec.toFixed(1)}s (${audioData.length} samples @ 16kHz)`);

  const { pipeline: pipelineFn } = await getXenovaMod();
  const sttLabels = { base: 'whisper-base (Lite)', small: 'whisper-small (Standard)' };
  sendLog(`[AI]   Using ${sttLabels[variant] || `whisper-${variant}`}`);

  const transcriber = await pipelineFn('automatic-speech-recognition', `Xenova/whisper-${variant}`);
  const WHISPER_SR       = 16000;
  const MANUAL_CHUNK_S   = 30;
  const MANUAL_OVERLAP_S = 8;
  const passOpts = { no_repeat_ngram_size: 3 };

  let rawText;
  if (audioDurationSec > 30) {
    const segTexts = [];
    let segStart = 0, segIdx = 0;
    while (segStart < audioData.length) {
      const segEnd   = Math.min(segStart + MANUAL_CHUNK_S * WHISPER_SR, audioData.length);
      const segAudio = audioData.slice(segStart, segEnd);
      const segSec   = segAudio.length / WHISPER_SR;
      const segRes   = await transcriber(segAudio, passOpts);
      const segText  = (segRes.text || '')
        .replace(/(\s*\[S\])+/g, '').replace(/\[BLANK_AUDIO\]/gi, '').trim();
      sendLog(`[AI]   🔍 Segment ${++segIdx} (${(segStart/WHISPER_SR)|0}–${(segEnd/WHISPER_SR)|0}s, ${segSec.toFixed(1)}s): ${JSON.stringify(segText.slice(0, 120))}`);
      segTexts.push(segText);
      if (segEnd >= audioData.length) break;
      segStart += (MANUAL_CHUNK_S - MANUAL_OVERLAP_S) * WHISPER_SR;
    }
    rawText = stitchTranscriptSegments(segTexts);
  } else {
    const result = await transcriber(audioData, passOpts);
    rawText = result.text || '';
  }

  sendLog(`[AI]   🔍 Raw Whisper output: ${JSON.stringify(rawText.slice(0, 200))}`);
  const cleanText = rawText
    .replace(/(\s*\[S\])+/g, '').replace(/\[BLANK_AUDIO\]/gi, '').trim();

  if (isWhisperHallucination(cleanText)) {
    sendLog(`[AI]   ⚠️  Whisper produced only non-speech tokens ("${rawText.slice(0, 80).trim()}") — discarding`);
    return { transcript: '', durationSec: audioDurationSec };
  }
  return { transcript: cleanText, durationSec: audioDurationSec };
}

// ── IPC entry point ───────────────────────────────────────────────────────────
process.on('message', async ({ audioPath, variant }) => {
  try {
    const result = await run(audioPath, variant || 'base');
    process.send({ ok: true, ...result });
  } catch (e) {
    process.send({ ok: false, error: e.message || String(e) });
  } finally {
    process.exit(0);
  }
});
