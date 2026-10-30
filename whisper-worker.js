// whisper-worker.js: speech to text with Whisper, running entirely in the browser.
// The model is downloaded once from Hugging Face and cached by the browser;
// audio never leaves the computer.

import { pipeline, env } from 'https://cdn.jsdelivr.net/npm/@huggingface/transformers@3.8.1';

env.allowLocalModels = false;

let asr = null;
let loadedKey = null;

// model = { id, device: 'webgpu' | 'wasm', dtype }
async function load(model) {
  const key = JSON.stringify(model);
  if (asr && loadedKey === key) return asr;
  if (asr && asr.dispose) await asr.dispose();
  asr = null;
  loadedKey = null;
  const files = {};
  asr = await pipeline('automatic-speech-recognition', model.id, {
    dtype: model.dtype,
    device: model.device,
    progress_callback: (p) => {
      if (p.status === 'progress' && p.total) {
        files[p.file] = { loaded: p.loaded, total: p.total };
        const loaded = Object.values(files).reduce((a, f) => a + f.loaded, 0);
        const total = Object.values(files).reduce((a, f) => a + f.total, 0);
        self.postMessage({ type: 'progress', loaded, total });
      }
    },
  });
  loadedKey = key;
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
      const started = performance.now();
      const out = await transcriber(audio, {
        language,
        task: task || 'transcribe',
        chunk_length_s: 30,
        stride_length_s: 5,
      });
      self.postMessage({ type: 'result', text: (out.text || '').trim(), requestId, seconds: (performance.now() - started) / 1000 });
    }
  } catch (err) {
    if (type === 'load') loadedKey = null;
    self.postMessage({ type: 'error', message: String(err && err.message ? err.message : err), requestId });
  }
};
