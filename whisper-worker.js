// whisper-worker.js: speech to text with Whisper, running entirely in the browser.
// The model is downloaded once from Hugging Face and cached by the browser;
// audio never leaves the computer.

import { pipeline, env } from 'https://cdn.jsdelivr.net/npm/@huggingface/transformers@3.0.2';

env.allowLocalModels = false;

let asr = null;
let loadedModel = null;

async function load(model) {
  if (asr && loadedModel === model) return asr;
  asr = null;
  const files = {};
  asr = await pipeline('automatic-speech-recognition', model, {
    dtype: 'q8',
    device: 'wasm',
    progress_callback: (p) => {
      if (p.status === 'progress' && p.total) {
        files[p.file] = { loaded: p.loaded, total: p.total };
        const loaded = Object.values(files).reduce((a, f) => a + f.loaded, 0);
        const total = Object.values(files).reduce((a, f) => a + f.total, 0);
        self.postMessage({ type: 'progress', loaded, total });
      }
    },
  });
  loadedModel = model;
  return asr;
}

self.onmessage = async (event) => {
  const { type, audio, language, task, model, requestId } = event.data;
  try {
    if (type === 'load') {
      await load(model);
      self.postMessage({ type: 'ready', requestId });
      return;
    }
    if (type === 'transcribe') {
      const transcriber = await load(model);
      self.postMessage({ type: 'working', requestId });
      const out = await transcriber(audio, {
        language,
        task: task || 'transcribe',
        chunk_length_s: 30,
        stride_length_s: 5,
      });
      self.postMessage({ type: 'result', text: (out.text || '').trim(), requestId });
    }
  } catch (err) {
    self.postMessage({ type: 'error', message: String(err && err.message ? err.message : err), requestId });
  }
};
