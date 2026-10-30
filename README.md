# Presentation helper

A browser tool for preparing a talk from a PowerPoint file: decide the order and content, budget the time and work on the speaker notes.

Live: https://johlobs.github.io/presentation-helper/

## What it does

- Opens any `.pptx` and groups the slides (existing PowerPoint sections, otherwise section header slides start a group). Groups can be renamed, split, merged, added and reordered. Slides can be dragged within and between groups.
- Time budget per group, or per slide, against a time slot, with start times and a timeline that shows overruns.
- Keep / maybe / cut per slide with a short reason (decision log).
- Speaker notes editor with dictation (Swedish or English) using Whisper, which runs locally in the browser via transformers.js. The model is downloaded once and cached.
- Optional PDF export of the same deck for thumbnails. Pages are matched to slides by text, so hidden slides or a reordered export still work.
- **Save & update pptx** rewrites the file: new slide order, groups as PowerPoint sections, cut slides hidden (not deleted), edited notes as plain text. Notes that are not edited keep their formatting. **Download pptx** downloads the result.

## Privacy

Everything stays in the browser (IndexedDB). The pptx is read and rewritten locally and never uploaded. Use **Backup** to save the plan and the original file as one JSON file, for example before clearing browser data or to move to another computer.

## Files

- `index.html`, `style.css`, `app.js`: interface and state
- `pptx.js`: reading and writing the pptx XML (JSZip)
- `thumbs.js`: PDF thumbnails (pdf.js) and page to slide matching
- `whisper-worker.js`: speech to text in a web worker
- `store.js`: IndexedDB storage

No build step. Serve the folder with any static server, for example `python -m http.server`, and open `index.html`. Dictation needs `https` or `localhost` for microphone access.
