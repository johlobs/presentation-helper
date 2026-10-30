// thumbs.js: render a PDF export of the deck to small images and match pages to slides.
// PowerPoint often leaves hidden slides out of PDF exports, so pages are matched
// by text similarity when the page count differs from the slide count.

import * as pdfjsLib from 'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.4.168/build/pdf.min.mjs';
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.4.168/build/pdf.worker.min.mjs';

const THUMB_WIDTH = 480;

function words(text) {
  return new Set(
    text.toLowerCase().normalize('NFKD').replace(/[^\p{L}\p{N}\s]/gu, ' ').split(/\s+/).filter((w) => w.length > 2)
  );
}

function similarity(a, b) {
  if (a.size === 0 || b.size === 0) return 0;
  let common = 0;
  for (const w of a) if (b.has(w)) common++;
  return common / Math.min(a.size, b.size);
}

export async function renderPdfThumbs(bytes, slides, onProgress) {
  const pdf = await pdfjsLib.getDocument({ data: bytes }).promise;
  const pages = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const base = page.getViewport({ scale: 1 });
    const viewport = page.getViewport({ scale: THUMB_WIDTH / base.width });
    const canvas = document.createElement('canvas');
    canvas.width = Math.round(viewport.width);
    canvas.height = Math.round(viewport.height);
    await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
    const content = await page.getTextContent();
    pages.push({
      image: canvas.toDataURL('image/jpeg', 0.8),
      words: words(content.items.map((it) => it.str).join(' ')),
    });
    if (onProgress) onProgress(i, pdf.numPages);
  }

  // Match pages to slides by text, best pairs first. A small penalty for distance
  // in position breaks ties between slides with similar text. The PDF may come
  // from a reordered version of the deck, so order is not required.
  const n = slides.length;
  const m = pages.length;
  const slideWords = slides.map((s) => words([s.title, ...s.body].join(' ')));
  const pairs = [];
  for (let i = 0; i < n; i++) {
    for (let j = 0; j < m; j++) {
      const sim = similarity(slideWords[i], pages[j].words);
      if (sim >= 0.35 && common(slideWords[i], pages[j].words) >= 2) {
        pairs.push({ i, j, score: sim - 0.05 * Math.abs(i / n - j / m) });
      }
    }
  }
  pairs.sort((a, b) => b.score - a.score);
  const thumbs = {};
  const usedSlides = new Set();
  const usedPages = new Set();
  for (const { i, j } of pairs) {
    if (usedSlides.has(i) || usedPages.has(j)) continue;
    thumbs[slides[i].id] = pages[j].image;
    usedSlides.add(i);
    usedPages.add(j);
  }
  // Slides without text (e.g. only a picture): use the page at the same position
  // when the PDF has exactly one page per slide.
  if (m === n) {
    for (let i = 0; i < n; i++) {
      if (!usedSlides.has(i) && !usedPages.has(i)) {
        thumbs[slides[i].id] = pages[i].image;
        usedSlides.add(i);
        usedPages.add(i);
      }
    }
  }
  return { thumbs, matched: usedSlides.size, pages: m };
}

function common(a, b) {
  let c = 0;
  for (const w of a) if (b.has(w)) c++;
  return c;
}
