import { readPptx, writePptx, suggestGroups } from './pptx.js';
import * as store from './store.js';
import { googleTranslate, joinSegments, myMemoryTranslate, externalLinks } from './translate.js';

const $ = (id) => document.getElementById(id);
const GROUP_COLORS = ['#2563eb', '#0d9488', '#7c3aed', '#db2777', '#ea580c', '#65a30d', '#0891b2', '#9333ea', '#ca8a04', '#475569'];
const MODELS = {
  'Xenova/whisper-small': 'Accurate (about 250 MB download)',
  'Xenova/whisper-base': 'Fast (about 80 MB download)',
};

let project = null;
let originalBytes = null;
let byId = {};
let selectedId = null;
let filter = 'all';
let saveTimer = null;
let sortables = [];
let notesCursorTouched = false;

// ---------- small helpers ----------

const uid = () => Math.random().toString(36).slice(2, 10);

function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function toast(message, ms = 3500) {
  const el = $('toast');
  el.textContent = message;
  el.hidden = false;
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => { el.hidden = true; }, ms);
}

function clock(minutes) {
  const totalSeconds = Math.round(minutes * 60);
  return Math.floor(totalSeconds / 60) + ':' + String(totalSeconds % 60).padStart(2, '0');
}

function minutesLabel(m) {
  return (Math.round(m * 10) / 10).toString().replace('.', ',') + ' min';
}

function formatTime(ts) {
  const d = new Date(ts);
  return d.toLocaleDateString('sv-SE') + ' ' + d.toLocaleTimeString('sv-SE', { hour: '2-digit', minute: '2-digit' });
}

const st = (id) => project.slideState[id];
const decisionOf = (id) => st(id).decision;
const notesOf = (id) => (st(id).notes != null ? st(id).notes : byId[id].notes);
const notesEdited = (id) => st(id).notes != null && st(id).notes !== byId[id].notes;

function groupOfSlide(id) {
  return project.groups.find((g) => g.slideIds.includes(id));
}

function orderedSlideIds() {
  return project.groups.flatMap((g) => g.slideIds);
}

// Show number = position among slides that are not cut.
function showNumbers() {
  const nums = {};
  let n = 0;
  for (const id of orderedSlideIds()) {
    if (decisionOf(id) !== 'cut') nums[id] = ++n;
  }
  return nums;
}

// A group's time is the sum of its slides' own minutes when any are set,
// otherwise the minutes set on the group. Cut slides never count.
function groupTime(g) {
  const active = g.slideIds.filter((id) => decisionOf(id) !== 'cut');
  const withMinutes = active.filter((id) => st(id).minutes != null);
  if (withMinutes.length > 0) {
    let total = 0;
    let maybe = 0;
    for (const id of withMinutes) {
      total += st(id).minutes;
      if (decisionOf(id) === 'maybe') maybe += st(id).minutes;
    }
    return { total, maybe, fromSlides: true };
  }
  if (active.length === 0) return { total: 0, maybe: 0, fromSlides: false };
  const allMaybe = active.every((id) => decisionOf(id) === 'maybe');
  const minutes = g.minutes || 0;
  return { total: minutes, maybe: allMaybe ? minutes : 0, fromSlides: false };
}

// ---------- persistence ----------

function setStatus() {
  const el = $('saveStatus');
  if (!project) { el.textContent = ''; return; }
  el.textContent = project.pptxDirty
    ? 'Plan saved in browser · pptx not updated yet'
    : project.pptxSavedAt ? 'pptx updated ' + formatTime(project.pptxSavedAt) : 'Plan saved in browser';
}

function changed({ pptx = true } = {}) {
  if (pptx) project.pptxDirty = true;
  clearTimeout(saveTimer);
  saveTimer = setTimeout(async () => {
    project.updatedAt = Date.now();
    await store.putProject(project);
    setStatus();
  }, 400);
  setStatus();
}

async function flushSave() {
  clearTimeout(saveTimer);
  project.updatedAt = Date.now();
  await store.putProject(project);
  setStatus();
}

// ---------- home ----------

async function showHome() {
  if (project) await flushSave();
  project = null;
  stopRecordingIfAny();
  $('planner').hidden = true;
  $('home').hidden = false;
  $('projectActions').hidden = true;
  $('projectName').textContent = '';
  const list = $('projectList');
  const projects = (await store.listProjects()).sort((a, b) => b.updatedAt - a.updatedAt);
  if (projects.length === 0) {
    list.innerHTML = '<li class="small">No presentations yet. Open a .pptx file above.</li>';
    return;
  }
  list.innerHTML = projects.map((p) => `
    <li data-id="${esc(p.id)}">
      <div><strong>${esc(p.name)}</strong><br>
        <span class="meta">${p.info.slides.length} slides · ${p.groups.length} groups · edited ${formatTime(p.updatedAt)}</span></div>
      <div><button class="btn small open">Open</button> <button class="btn ghost small delete">Delete</button></div>
    </li>`).join('');
}

$('projectList').addEventListener('click', async (e) => {
  const li = e.target.closest('li[data-id]');
  if (!li) return;
  const id = li.dataset.id;
  if (e.target.classList.contains('open')) openProject(id);
  if (e.target.classList.contains('delete')) {
    if (confirm('Delete this presentation and its plan from this browser? Your original file on disk is not affected.')) {
      await store.deleteProject(id);
      showHome();
    }
  }
});

async function createProject(file) {
  if (!/\.pptx$/i.test(file.name)) { toast('Please choose a .pptx file.'); return; }
  try {
    const bytes = new Uint8Array(await file.arrayBuffer());
    const info = await readPptx(JSZip, bytes);
    if (info.slides.length === 0) { toast('The file has no slides.'); return; }
    const targetMinutes = 25;
    const slideState = {};
    for (const s of info.slides) {
      slideState[s.id] = { decision: s.hidden ? 'cut' : 'keep', reason: '', minutes: null, notes: null };
    }
    const visibleCount = info.slides.filter((s) => !s.hidden).length || 1;
    const groups = suggestGroups(info).map((g) => {
      const visible = g.slideIds.filter((id) => slideState[id].decision !== 'cut').length;
      return {
        id: uid(),
        name: g.name,
        slideIds: g.slideIds,
        collapsed: true,
        minutes: Math.round((targetMinutes * visible / visibleCount) * 2) / 2,
      };
    });
    const now = Date.now();
    const p = {
      id: uid(),
      name: file.name.replace(/\.pptx$/i, ''),
      fileName: file.name,
      createdAt: now,
      updatedAt: now,
      targetMinutes,
      info,
      groups,
      slideState,
      thumbs: {},
      pptxDirty: false,
      pptxSavedAt: null,
    };
    await store.putFile(p.id + ':original', bytes);
    await store.putProject(p);
    await openProject(p.id);
    toast(`Loaded ${info.slides.length} slides in ${groups.length} suggested groups. Time per group is a starting suggestion.`, 6000);
  } catch (err) {
    console.error(err);
    toast('Could not read the file: ' + err.message, 6000);
  }
}

$('pptxInput').addEventListener('change', (e) => {
  if (e.target.files[0]) createProject(e.target.files[0]);
  e.target.value = '';
});
const dz = $('dropzone');
dz.addEventListener('dragover', (e) => { e.preventDefault(); dz.classList.add('over'); });
dz.addEventListener('dragleave', () => dz.classList.remove('over'));
dz.addEventListener('drop', (e) => {
  e.preventDefault();
  dz.classList.remove('over');
  if (e.dataTransfer.files[0]) createProject(e.dataTransfer.files[0]);
});
$('homeBtn').addEventListener('click', showHome);

// ---------- open project ----------

async function openProject(id) {
  project = await store.getProject(id);
  originalBytes = await store.getFile(id + ':original');
  if (project.info.slides.some((s) => !s.shapes)) {
    const fresh = await readPptx(JSZip, originalBytes);
    const freshById = Object.fromEntries(fresh.slides.map((s) => [s.id, s]));
    for (const s of project.info.slides) s.shapes = freshById[s.id] ? freshById[s.id].shapes : [];
    await store.putProject(project);
  }
  byId = Object.fromEntries(project.info.slides.map((s) => [s.id, s]));
  selectedId = orderedSlideIds()[0];
  filter = 'all';
  $('filterSelect').value = 'all';
  $('home').hidden = true;
  $('planner').hidden = false;
  $('projectActions').hidden = false;
  $('projectName').textContent = project.fileName;
  $('targetMinutes').value = project.targetMinutes;
  renderAll();
  setStatus();
}

// ---------- rendering ----------

function renderAll() {
  renderTiming();
  renderOutline();
  renderDetail();
}

function renderTiming() {
  const times = project.groups.map(groupTime);
  const total = times.reduce((a, t) => a + t.total, 0);
  const maybe = times.reduce((a, t) => a + t.maybe, 0);
  const target = project.targetMinutes;
  const nums = showNumbers();
  const inShow = Object.keys(nums).length;
  const cut = project.info.slides.length - inShow;

  let text = `Planned ${minutesLabel(total)} of ${minutesLabel(target)}`;
  if (maybe > 0) text += ` (${minutesLabel(maybe)} of it is “maybe”)`;
  const diff = total - target;
  if (diff > 0.05) text += ` · ${minutesLabel(diff)} over`;
  else if (diff < -0.05) text += ` · ${minutesLabel(-diff)} left`;
  text += ` · ${inShow} slides in the show, ${cut} cut`;
  const summary = $('timingSummary');
  summary.textContent = text;
  summary.className = diff > 0.05 ? 'over' : diff < -0.05 ? 'under' : '';

  const scale = Math.max(target, total) || 1;
  const segs = project.groups.map((g, i) => {
    const t = times[i];
    if (t.total <= 0) return '';
    const isMaybe = t.maybe > 0 && t.maybe >= t.total;
    return `<div class="seg ${isMaybe ? 'maybe' : ''}" data-gid="${g.id}" style="width:${(t.total / scale) * 100}%;background:${GROUP_COLORS[i % GROUP_COLORS.length]}" title="${esc(g.name)}: ${minutesLabel(t.total)}">${esc(g.name)}</div>`;
  }).join('');
  const mark = total > target ? `<div class="target-mark" style="left:${(target / scale) * 100}%" title="End of time slot"></div>` : '';
  $('timeline').innerHTML = segs + mark;
}

function slideRowHtml(id, indexInGroup, nums) {
  const s = byId[id];
  const state = st(id);
  const d = state.decision;
  const thumb = project.thumbs[id];
  const hiddenByFilter = filter !== 'all' && filter !== d;
  return `
    <li class="slide-row ${d === 'cut' ? 'cut' : ''} ${id === selectedId ? 'selected' : ''}" data-sid="${id}" ${hiddenByFilter ? 'hidden' : ''}>
      <span class="handle slide-handle" title="Drag to move">⠿</span>
      <span class="num" title="Number in the show (original ${s.originalIndex + 1})">${nums[id] || '–'}</span>
      ${thumb ? `<img class="thumb" src="${thumb}" alt="">` : ''}
      <span class="s-title" title="${esc(titleOf(id))}">${esc(titleOf(id))}</span>
      <span class="badges">
        ${notesEdited(id) ? '<span class="dot" title="Notes edited">✎</span>' : ''}
        ${state.minutes != null ? `<span class="dot">${minutesLabel(state.minutes)}</span>` : ''}
        ${d !== 'keep' ? `<span class="chip ${d}">${d}</span>` : ''}
        ${indexInGroup > 0 ? '<button class="icon-btn split-btn" title="Start a new group from this slide">⤓</button>' : ''}
      </span>
    </li>`;
}

function renderOutline() {
  for (const s of sortables) s.destroy();
  sortables = [];
  const nums = showNumbers();
  let start = 0;
  const html = project.groups.map((g, gi) => {
    const t = groupTime(g);
    const color = GROUP_COLORS[gi % GROUP_COLORS.length];
    const startLabel = clock(start);
    start += t.total;
    const rows = g.slideIds.map((id, i) => slideRowHtml(id, i, nums)).join('');
    return `
      <div class="group ${g.collapsed ? 'collapsed' : ''}" data-gid="${g.id}" style="--gcolor:${color}">
        <div class="group-head">
          <span class="handle group-handle" title="Drag to move the whole group">⠿</span>
          <button class="toggle" title="Show or hide slides">${g.collapsed ? '▸' : '▾'}</button>
          <input class="group-name" spellcheck="false" value="${esc(g.name)}" title="Group name (becomes the section name in PowerPoint)">
          <span class="group-meta">${g.slideIds.length} sl. · starts ${startLabel}</span>
          <input type="number" class="group-min" min="0" step="0.5" value="${t.fromSlides ? Math.round(t.total * 10) / 10 : (g.minutes || 0)}"
            ${t.fromSlides ? 'disabled title="Sum of the minutes set on individual slides"' : 'title="Minutes for this group"'}>
          <span class="group-meta">min</span>
          ${gi > 0 ? '<button class="icon-btn merge" title="Merge into the group above">⤒</button>' : ''}
          ${g.slideIds.length === 0 ? '<button class="icon-btn del" title="Delete this empty group">✕</button>' : ''}
        </div>
        <ul class="slide-list" data-gid="${g.id}">${rows}</ul>
      </div>`;
  }).join('');
  const container = $('groups');
  container.innerHTML = html;

  sortables.push(new Sortable(container, {
    handle: '.group-handle',
    draggable: '.group',
    animation: 150,
    onEnd: syncFromDom,
  }));
  for (const ul of container.querySelectorAll('.slide-list')) {
    sortables.push(new Sortable(ul, {
      group: 'slides',
      handle: '.slide-handle',
      animation: 150,
      onEnd: syncFromDom,
    }));
  }
}

function syncFromDom() {
  const byGroupId = Object.fromEntries(project.groups.map((g) => [g.id, g]));
  const groups = [];
  for (const gEl of $('groups').querySelectorAll('.group')) {
    const g = byGroupId[gEl.dataset.gid];
    g.slideIds = Array.from(gEl.querySelectorAll('.slide-row')).map((li) => li.dataset.sid);
    groups.push(g);
  }
  project.groups = groups;
  changed();
  renderTiming();
  renderOutline();
  renderDetail();
}

$('groups').addEventListener('click', (e) => {
  const gEl = e.target.closest('.group');
  if (!gEl) return;
  const gi = project.groups.findIndex((g) => g.id === gEl.dataset.gid);
  const g = project.groups[gi];

  if (e.target.closest('.toggle')) {
    g.collapsed = !g.collapsed;
    changed({ pptx: false });
    renderOutline();
    return;
  }
  if (e.target.closest('.merge')) {
    const prev = project.groups[gi - 1];
    prev.slideIds.push(...g.slideIds);
    prev.minutes = (prev.minutes || 0) + (g.minutes || 0);
    prev.collapsed = false;
    project.groups.splice(gi, 1);
    changed();
    renderAll();
    return;
  }
  if (e.target.closest('.del')) {
    project.groups.splice(gi, 1);
    changed();
    renderAll();
    return;
  }
  const row = e.target.closest('.slide-row');
  if (!row) return;
  const sid = row.dataset.sid;
  if (e.target.closest('.split-btn')) {
    const at = g.slideIds.indexOf(sid);
    const moved = g.slideIds.splice(at);
    const share = g.minutes ? Math.round((g.minutes * moved.length / (moved.length + g.slideIds.length)) * 2) / 2 : 0;
    g.minutes = Math.max(0, (g.minutes || 0) - share);
    project.groups.splice(gi + 1, 0, { id: uid(), name: titleOf(sid), slideIds: moved, collapsed: false, minutes: share });
    changed();
    renderAll();
    return;
  }
  if (e.target.closest('.handle')) return;
  select(sid);
});

$('groups').addEventListener('input', (e) => {
  if (!e.target.classList.contains('group-name')) return;
  const g = project.groups.find((x) => x.id === e.target.closest('.group').dataset.gid);
  g.name = e.target.value;
  changed();
  renderTiming();
});

$('groups').addEventListener('change', (e) => {
  if (!e.target.classList.contains('group-min')) return;
  const g = project.groups.find((x) => x.id === e.target.closest('.group').dataset.gid);
  g.minutes = Math.max(0, parseFloat(e.target.value.replace(',', '.')) || 0);
  changed({ pptx: false });
  renderTiming();
  renderOutline();
});

$('timeline').addEventListener('click', (e) => {
  const seg = e.target.closest('.seg');
  if (!seg) return;
  const g = project.groups.find((x) => x.id === seg.dataset.gid);
  g.collapsed = false;
  renderOutline();
  if (g.slideIds[0]) select(g.slideIds[0]);
  document.querySelector(`.group[data-gid="${g.id}"]`).scrollIntoView({ behavior: 'smooth', block: 'start' });
});

$('addGroupBtn').addEventListener('click', () => {
  project.groups.push({ id: uid(), name: 'New group', slideIds: [], collapsed: false, minutes: 0 });
  changed();
  renderOutline();
  const inputs = $('groups').querySelectorAll('.group-name');
  inputs[inputs.length - 1].select();
});
$('expandBtn').addEventListener('click', () => {
  for (const g of project.groups) g.collapsed = false;
  changed({ pptx: false });
  renderOutline();
});
$('collapseBtn').addEventListener('click', () => {
  for (const g of project.groups) g.collapsed = true;
  changed({ pptx: false });
  renderOutline();
});
$('filterSelect').addEventListener('change', (e) => {
  filter = e.target.value;
  if (filter !== 'all') for (const g of project.groups) g.collapsed = false;
  renderOutline();
});
$('targetMinutes').addEventListener('change', (e) => {
  project.targetMinutes = Math.max(1, parseFloat(e.target.value.replace(',', '.')) || project.targetMinutes);
  changed({ pptx: false });
  renderTiming();
});

function select(id) {
  selectedId = id;
  const g = groupOfSlide(id);
  if (g && g.collapsed) {
    g.collapsed = false;
    renderOutline();
  } else {
    for (const row of $('groups').querySelectorAll('.slide-row')) {
      row.classList.toggle('selected', row.dataset.sid === id);
    }
  }
  renderDetail();
  const row = document.querySelector(`.slide-row[data-sid="${id}"]`);
  if (row) row.scrollIntoView({ block: 'nearest' });
}

function renderDetail() {
  const el = $('detail');
  if (!selectedId || !byId[selectedId]) {
    el.innerHTML = '<div class="empty">Select a slide to see its content, decide on it and edit its notes.</div>';
    return;
  }
  const id = selectedId;
  const s = byId[id];
  const state = st(id);
  const nums = showNumbers();
  const g = groupOfSlide(id);
  const order = orderedSlideIds();
  const pos = order.indexOf(id);
  const lang = localStorage.getItem('asrLang') || 'sv';
  const model = localStorage.getItem('asrModel') || 'Xenova/whisper-small';
  const recordingHere = recorder && recorder.slideId === id;

  el.innerHTML = `
    <div class="inline" style="justify-content:space-between">
      <button class="btn ghost small" id="prevBtn" ${pos <= 0 ? 'disabled' : ''}>← Previous</button>
      <span class="small">${pos + 1} of ${order.length}</span>
      <button class="btn ghost small" id="nextBtn" ${pos >= order.length - 1 ? 'disabled' : ''}>Next →</button>
    </div>
    <h2>${esc(titleOf(id))}</h2>
    <div class="sub">${nums[id] ? 'Slide ' + nums[id] + ' in the show' : 'Cut (hidden in the show)'} · originally slide ${s.originalIndex + 1} · group “${esc(g ? g.name : '')}”</div>
    ${project.thumbs[id] ? `<img class="big-thumb" src="${project.thumbs[id]}" alt="Slide preview">` : ''}
    ${slideTextHtml(id)}

    <div class="field">
      <span class="field-title">Decision</span>
      <div class="decision-btns">
        <button class="keep ${state.decision === 'keep' ? 'on' : ''}" data-d="keep">Keep</button>
        <button class="maybe ${state.decision === 'maybe' ? 'on' : ''}" data-d="maybe">Maybe</button>
        <button class="cut ${state.decision === 'cut' ? 'on' : ''}" data-d="cut">Cut</button>
      </div>
    </div>
    <div class="field">
      <label for="reasonBox">Why? <span class="small">(your decision log)</span></label>
      <textarea id="reasonBox" rows="2" lang="${lang === 'sv' ? 'sv' : 'en'}" placeholder="E.g. strongest example for the policy point, or too detailed for this audience">${esc(state.reason)}</textarea>
    </div>
    <div class="field inline">
      <label for="slideMin" style="font-weight:600">Minutes for this slide</label>
      <input type="number" id="slideMin" min="0" step="0.5" value="${state.minutes != null ? state.minutes : ''}" placeholder="–">
      <span class="small">Optional. When set on any slide in a group, the group's time is the sum of its slides.</span>
    </div>

    <div class="field">
      <div class="notes-bar">
        <span class="field-title" style="margin:0">Speaker notes</span>
        <span class="changed-flag" id="notesFlag">${notesEdited(id) ? 'edited, written to the pptx on save' : ''}</span>
        <span class="spacer"></span>
        <button class="btn small" id="translateBtn" title="Translate the selected text, or all notes if nothing is selected, from Swedish to English">Translate SV→EN</button>
        <span class="lang-toggle" title="Dictation language. SV→EN: speak Swedish, get English text">
          <button data-lang="sv" class="${lang === 'sv' ? 'on' : ''}">SV</button><button data-lang="en" class="${lang === 'en' ? 'on' : ''}">EN</button><button data-lang="svEn" class="${lang === 'svEn' ? 'on' : ''}">SV→EN</button>
        </span>
        <button class="btn small mic-btn ${recordingHere ? 'recording' : ''}" id="micBtn">${recordingHere ? '■ Stop' : '🎙 Dictate'}</button>
      </div>
      <textarea id="notesBox" class="notes" lang="${lang === 'sv' ? 'sv' : 'en'}" placeholder="No notes yet.">${esc(notesOf(id))}</textarea>
      <div class="asr-status" id="asrStatus">${esc(asrStatusText)}</div>
      <div class="translate-panel" id="translatePanel" hidden></div>
      <div class="inline small" style="justify-content:space-between;margin-top:0.3rem">
        <label>Speech model <select id="modelSelect" class="small">
          ${Object.entries(MODELS).map(([k, v]) => `<option value="${k}" ${k === model ? 'selected' : ''}>${v}</option>`).join('')}
        </select></label>
        ${notesEdited(id) ? '<button class="link-btn" id="revertNotes">Revert to the notes in the file</button>' : ''}
      </div>
      <p class="small">Dictated text is inserted where the cursor is in the notes, or at the end if you have not clicked in them. The speech model runs in this browser and is downloaded the first time you dictate. Translation sends the selected text (or all the notes) to Google Translate and MyMemory. Edited notes are saved as plain text in the pptx; notes you do not edit keep their formatting.</p>
    </div>`;

  notesCursorTouched = false;
  $('prevBtn').onclick = () => select(order[pos - 1]);
  $('nextBtn').onclick = () => select(order[pos + 1]);
  for (const b of el.querySelectorAll('.decision-btns button')) {
    b.onclick = () => {
      state.decision = b.dataset.d;
      changed();
      renderAll();
    };
  }
  $('reasonBox').oninput = (e) => { state.reason = e.target.value; changed({ pptx: false }); };
  $('slideMin').onchange = (e) => {
    const v = e.target.value.trim();
    state.minutes = v === '' ? null : Math.max(0, parseFloat(v.replace(',', '.')) || 0);
    changed({ pptx: false });
    renderTiming();
    renderOutline();
  };
  const box = $('notesBox');
  const touch = () => { notesCursorTouched = true; };
  box.addEventListener('click', touch);
  box.addEventListener('keyup', touch);
  box.addEventListener('focus', touch);
  box.oninput = () => {
    setNotes(id, box.value);
    $('notesFlag').textContent = notesEdited(id) ? 'edited, written to the pptx on save' : '';
  };
  box.onchange = () => renderOutline();
  for (const b of el.querySelectorAll('.lang-toggle button')) {
    b.onclick = () => { localStorage.setItem('asrLang', b.dataset.lang); renderDetail(); };
  }
  $('modelSelect').onchange = (e) => localStorage.setItem('asrModel', e.target.value);
  $('micBtn').onclick = () => toggleRecording(id);
  $('translateBtn').onclick = () => startTranslation(id);
  if (translation && translation.slideId === id) renderTranslation();
  bindSlideText(id);
  const revert = $('revertNotes');
  if (revert) revert.onclick = () => { state.notes = null; changed(); renderOutline(); renderDetail(); };
}

// ---------- slide title and text editing ----------

function shapeParagraphs(id, shape) {
  const edits = st(id).text || {};
  return edits[shape.index] ? edits[shape.index] : shape.paragraphs;
}

function slideTextEdited(id) {
  return Object.keys(st(id).text || {}).length > 0;
}

function titleOf(id) {
  const s = byId[id];
  const titleShape = (s.shapes || []).find((sh) => sh.isTitle);
  if (!titleShape) return s.title;
  const text = shapeParagraphs(id, titleShape).join(' ').replace(/\s+/g, ' ').trim();
  return text || '(untitled)';
}

function slideTextHtml(id) {
  const s = byId[id];
  const shapes = s.shapes || [];
  if (shapes.length === 0) {
    return '<div class="field"><span class="field-title">Slide text</span><div class="small">This slide has no editable text (pictures, charts or tables only).</div></div>';
  }
  let bodyCount = 0;
  const boxes = shapes.map((shape) => {
    const label = shape.isTitle ? 'Title' : `Text box ${++bodyCount}`;
    const paras = shapeParagraphs(id, shape);
    const rows = paras.map((text, pi) => `
      <div class="para-row">
        <textarea class="para" rows="1" lang="en" data-shape="${shape.index}" data-p="${pi}">${esc(text)}</textarea>
        ${shape.isTitle ? '' : `<button class="icon-btn" data-act="add" data-shape="${shape.index}" data-p="${pi}" title="Add a paragraph after this one">+</button>
        <button class="icon-btn" data-act="del" data-shape="${shape.index}" data-p="${pi}" title="Remove this paragraph" ${paras.length < 2 ? 'disabled' : ''}>✕</button>`}
      </div>`).join('');
    return `<div class="shape-box ${shape.isTitle ? 'title-box' : ''}"><div class="shape-label">${label}</div>${rows}</div>`;
  }).join('');
  return `
    <div class="field" id="slideTextField">
      <div class="notes-bar">
        <span class="field-title" style="margin:0">Slide text</span>
        <span class="changed-flag">${slideTextEdited(id) ? 'edited, written to the pptx on save' : ''}</span>
        <span class="spacer"></span>
        ${slideTextEdited(id) ? '<button class="link-btn small" data-act="revert-text">Revert to the text in the file</button>' : ''}
      </div>
      ${boxes}
      <p class="small">Each box is one paragraph (one bullet). Edited paragraphs keep the formatting of their first word; mixed formatting inside a paragraph (e.g. one bold word) is lost when you edit it.</p>
    </div>`;
}

function autosize(el) {
  el.style.height = 'auto';
  el.style.height = el.scrollHeight + 2 + 'px';
}

function readShapeFromDom(id, shapeIndex) {
  const field = $('slideTextField');
  const values = Array.from(field.querySelectorAll(`textarea.para[data-shape="${shapeIndex}"]`)).map((t) => t.value);
  setShapeText(id, shapeIndex, values);
}

function setShapeText(id, shapeIndex, values) {
  const shape = byId[id].shapes.find((sh) => sh.index === Number(shapeIndex));
  const state = st(id);
  state.text = state.text || {};
  const same = values.length === shape.paragraphs.length && values.every((v, i) => v === shape.paragraphs[i]);
  if (same) delete state.text[shape.index];
  else state.text[shape.index] = values;
  changed();
}

function bindSlideText(id) {
  const field = $('slideTextField');
  if (!field) return;
  for (const t of field.querySelectorAll('textarea.para')) autosize(t);
  field.addEventListener('input', (e) => {
    const t = e.target.closest('textarea.para');
    if (!t) return;
    autosize(t);
    readShapeFromDom(id, t.dataset.shape);
    field.querySelector('.changed-flag').textContent = slideTextEdited(id) ? 'edited, written to the pptx on save' : '';
  });
  field.addEventListener('change', (e) => {
    if (e.target.closest('textarea.para')) {
      renderOutline();
      document.querySelector('#detail h2').textContent = titleOf(id);
    }
  });
  field.addEventListener('click', (e) => {
    const b = e.target.closest('button[data-act]');
    if (!b) return;
    const act = b.dataset.act;
    if (act === 'revert-text') {
      st(id).text = {};
      changed();
      renderOutline();
      renderDetail();
      return;
    }
    const shape = byId[id].shapes.find((sh) => sh.index === Number(b.dataset.shape));
    const values = shapeParagraphs(id, shape).slice();
    const pi = Number(b.dataset.p);
    if (act === 'add') values.splice(pi + 1, 0, '');
    if (act === 'del' && values.length > 1) values.splice(pi, 1);
    setShapeText(id, shape.index, values);
    renderDetail();
    if (act === 'add') {
      const next = $('slideTextField').querySelector(`textarea.para[data-shape="${shape.index}"][data-p="${pi + 1}"]`);
      if (next) next.focus();
    }
  });
}

function setNotes(id, text) {
  st(id).notes = text === byId[id].notes ? null : text;
  changed();
}

document.addEventListener('keydown', (e) => {
  if (!project || $('planner').hidden) return;
  if (['INPUT', 'TEXTAREA', 'SELECT'].includes(document.activeElement.tagName)) return;
  if (e.key !== 'ArrowDown' && e.key !== 'ArrowUp') return;
  const order = orderedSlideIds().filter((id) => filter === 'all' || decisionOf(id) === filter);
  const pos = order.indexOf(selectedId);
  const next = order[pos + (e.key === 'ArrowDown' ? 1 : -1)];
  if (next) { e.preventDefault(); select(next); }
});

// ---------- dictation (Whisper in a web worker) ----------

let worker = null;
let recorder = null;
let asrStatusText = '';
let requestCounter = 0;
const pendingInserts = new Map();

function setAsrStatus(text, isError = false) {
  asrStatusText = text;
  const el = $('asrStatus');
  if (el) {
    el.textContent = text;
    el.classList.toggle('error', isError);
  }
}

function getWorker() {
  if (!worker) {
    worker = new Worker(new URL('./whisper-worker.js', import.meta.url), { type: 'module' });
    worker.onmessage = (e) => {
      const msg = e.data;
      if (msg.type === 'progress') {
        setAsrStatus(`Downloading the speech model (first time only): ${Math.round(msg.loaded / 1e6)} of ${Math.round(msg.total / 1e6)} MB`);
      } else if (msg.type === 'ready') {
        if (!recorder) setAsrStatus('Speech model ready.');
      } else if (msg.type === 'working') {
        setAsrStatus('Transcribing…');
      } else if (msg.type === 'result') {
        insertDictation(msg.requestId, msg.text);
      } else if (msg.type === 'error') {
        pendingInserts.delete(msg.requestId);
        setAsrStatus('Dictation failed: ' + msg.message, true);
      }
    };
  }
  return worker;
}

async function toggleRecording(slideId) {
  if (recorder) {
    const sameSlide = recorder.slideId === slideId;
    recorder.media.stop();
    if (sameSlide) return;
  }
  let stream;
  try {
    stream = await navigator.mediaDevices.getUserMedia({ audio: true });
  } catch (err) {
    setAsrStatus('No microphone access: ' + err.message, true);
    return;
  }
  const box = $('notesBox');
  const current = notesOf(slideId);
  const position = notesCursorTouched && box ? box.selectionEnd : current.length;
  const media = new MediaRecorder(stream);
  const chunks = [];
  const lang = localStorage.getItem('asrLang') || 'sv';
  const model = localStorage.getItem('asrModel') || 'Xenova/whisper-small';
  recorder = { media, slideId, position, startedAt: Date.now() };
  media.ondataavailable = (e) => { if (e.data.size > 0) chunks.push(e.data); };
  media.onstop = async () => {
    stream.getTracks().forEach((t) => t.stop());
    const info = recorder;
    recorder = null;
    clearInterval(recordTimer);
    renderDetailMicOnly();
    try {
      const blob = new Blob(chunks, { type: media.mimeType });
      const ctx = new AudioContext({ sampleRate: 16000 });
      const buffer = await ctx.decodeAudioData(await blob.arrayBuffer());
      ctx.close();
      let audio = buffer.getChannelData(0);
      if (buffer.numberOfChannels > 1) {
        const other = buffer.getChannelData(1);
        const mixed = new Float32Array(audio.length);
        for (let i = 0; i < audio.length; i++) mixed[i] = (audio[i] + other[i]) / 2;
        audio = mixed;
      }
      const requestId = ++requestCounter;
      pendingInserts.set(requestId, info);
      setAsrStatus('Transcribing…');
      getWorker().postMessage({
        type: 'transcribe', audio, requestId, model,
        language: lang === 'en' ? 'english' : 'swedish',
        task: lang === 'svEn' ? 'translate' : 'transcribe',
      }, [audio.buffer]);
    } catch (err) {
      setAsrStatus('Could not process the recording: ' + err.message, true);
    }
  };
  media.start();
  getWorker().postMessage({ type: 'load', model });
  const recordTimer = setInterval(() => {
    if (recorder) setAsrStatus(`Recording ${clock((Date.now() - recorder.startedAt) / 60000)}. Press Stop when you are done.`);
  }, 500);
  setAsrStatus('Recording 0:00. Press Stop when you are done.');
  renderDetailMicOnly();
}

function renderDetailMicOnly() {
  const btn = $('micBtn');
  if (!btn) return;
  const here = recorder && recorder.slideId === selectedId;
  btn.classList.toggle('recording', Boolean(here));
  btn.textContent = here ? '■ Stop' : '🎙 Dictate';
}

function stopRecordingIfAny() {
  if (recorder) recorder.media.stop();
}

function insertDictation(requestId, text) {
  const info = pendingInserts.get(requestId);
  pendingInserts.delete(requestId);
  if (!info || !project || !byId[info.slideId]) return;
  if (!text) { setAsrStatus('No speech was recognised.'); return; }
  const current = notesOf(info.slideId);
  const pos = Math.min(info.position, current.length);
  const before = current.slice(0, pos);
  const after = current.slice(pos);
  const lead = before && !/\s$/.test(before) ? ' ' : '';
  const trail = after && !/^\s/.test(after) ? ' ' : '';
  const updated = before + lead + text + trail + after;
  setNotes(info.slideId, updated);
  setAsrStatus('Dictation added to the notes.');
  if (selectedId === info.slideId) {
    const box = $('notesBox');
    box.value = updated;
    const caret = (before + lead + text).length;
    box.focus();
    box.setSelectionRange(caret, caret);
    $('notesFlag').textContent = notesEdited(info.slideId) ? 'edited, written to the pptx on save' : '';
  }
  renderOutline();
}

// ---------- translation (Swedish to English, several alternatives) ----------

let translation = null;

async function startTranslation(slideId) {
  const box = $('notesBox');
  const notes = notesOf(slideId);
  let start = box.selectionStart;
  let end = box.selectionEnd;
  if (start === end) { start = 0; end = notes.length; }
  const source = notes.slice(start, end).trim();
  if (!source) { setAsrStatus('Nothing to translate. Write or select some Swedish text first.'); return; }
  // Keep the trimmed range so Replace does not remove surrounding spaces.
  start = notes.indexOf(source, start);
  end = start + source.length;
  const isWord = !/\s/.test(source);
  translation = { slideId, start, end, source, isWord, google: null, googleError: null, chosen: [], mm: null, mmError: null, loading: true };
  renderTranslation();
  const [g, m] = await Promise.allSettled([googleTranslate(source), myMemoryTranslate(source)]);
  if (!translation || translation.source !== source) return;
  translation.loading = false;
  if (g.status === 'fulfilled') translation.google = g.value; else translation.googleError = g.reason.message;
  if (m.status === 'fulfilled') translation.mm = m.value; else translation.mmError = m.reason.message;
  renderTranslation();
}

function translationCandidates() {
  const t = translation;
  const list = [];
  if (t.google) {
    const composed = t.google.segments.length ? joinSegments(t.google.segments, t.chosen) : t.google.main;
    list.push({ text: composed, kind: 'main' });
    for (const d of t.google.dictionary) {
      for (const e of d.entries) list.push({ text: e.word, kind: 'dict', pos: d.pos, back: e.back });
    }
  }
  for (const m of t.mm || []) list.push({ text: m.text, kind: 'mm', source: m.source });
  return list;
}

function renderTranslation() {
  const panel = $('translatePanel');
  if (!panel || !translation || translation.slideId !== selectedId) return;
  const t = translation;
  panel.hidden = false;
  const short = t.source.length > 160 ? t.source.slice(0, 160) + '…' : t.source;
  let html = `<div class="tp-head"><span><strong>Translating:</strong> “${esc(short)}”</span><button class="icon-btn" data-act="close" title="Close">✕</button></div>`;
  if (t.loading) {
    panel.innerHTML = html + '<div class="small">Asking Google Translate and MyMemory…</div>';
    panel.onclick = (e) => { if (e.target.closest('[data-act=close]')) closeTranslation(); };
    return;
  }
  const cands = translationCandidates();
  const actions = (i) => `<span class="tp-actions"><button class="btn small" data-act="replace" data-i="${i}">Replace</button><button class="btn ghost small" data-act="insert" data-i="${i}">Insert after</button><button class="btn ghost small" data-act="copy" data-i="${i}">Copy</button></span>`;

  html += '<div class="tp-section"><div class="tp-title">Google Translate</div>';
  if (t.googleError) html += `<div class="small error-text">Not available: ${esc(t.googleError)}</div>`;
  cands.forEach((c, i) => {
    if (c.kind === 'main') html += `<div class="tp-cand main"><span class="tp-text">${esc(c.text)}</span>${actions(i)}</div>`;
  });
  if (t.google) {
    if (t.google.segments.some((seg) => seg.options.length > 1)) {
      html += '<div class="small">Other wordings. Click one to use it in the translation above:</div>';
      t.google.segments.forEach((seg, si) => {
        if (seg.options.length < 2) return;
        html += '<div class="tp-alts">' + seg.options.map((o, oi) =>
          `<button class="alt-chip ${(t.chosen[si] || 0) === oi ? 'on' : ''}" data-act="alt" data-seg="${si}" data-opt="${oi}">${esc(o)}</button>`).join('') + '</div>';
      });
    }
    const dict = cands.map((c, i) => [c, i]).filter(([c]) => c.kind === 'dict');
    if (dict.length) {
      html += '<div class="small" style="margin-top:0.4rem">Dictionary meanings (← the Swedish words each one can translate):</div>';
      for (const [c, i] of dict) {
        html += `<div class="tp-cand"><span class="tp-text"><strong>${esc(c.text)}</strong> <span class="small">${esc(c.pos)} ← ${esc(c.back.slice(0, 6).join(', '))}</span></span>${actions(i)}</div>`;
      }
    }
  }
  html += '</div><div class="tp-section"><div class="tp-title">MyMemory <span class="small">(translation memories, quality varies)</span></div>';
  if (t.mmError) html += `<div class="small error-text">Not available: ${esc(t.mmError)}</div>`;
  else if (!t.mm || t.mm.length === 0) html += '<div class="small">No suggestions.</div>';
  cands.forEach((c, i) => {
    if (c.kind === 'mm') html += `<div class="tp-cand"><span class="tp-text">${esc(c.text)} <span class="small">· ${esc(c.source)}</span></span>${actions(i)}</div>`;
  });
  html += '</div><div class="tp-links small">Compare in: ' +
    externalLinks(t.source, t.isWord).map((l) => `<a href="${l.url}" target="_blank" rel="noopener">${esc(l.label)}</a>`).join(' · ') + '</div>';
  panel.innerHTML = html;

  panel.onclick = async (e) => {
    const b = e.target.closest('button[data-act]');
    if (!b) return;
    const act = b.dataset.act;
    if (act === 'close') { closeTranslation(); return; }
    if (act === 'alt') {
      t.chosen[Number(b.dataset.seg)] = Number(b.dataset.opt);
      renderTranslation();
      return;
    }
    const text = translationCandidates()[Number(b.dataset.i)].text;
    if (act === 'copy') {
      await navigator.clipboard.writeText(text);
      setAsrStatus('Copied.');
      return;
    }
    applyTranslation(text, act === 'replace');
  };
}

function applyTranslation(text, replace) {
  const t = translation;
  const notes = notesOf(t.slideId);
  let start = t.start;
  let end = t.end;
  if (notes.slice(start, end) !== t.source) {
    // The notes changed since translating: find the source text again, or add at the end.
    const at = notes.indexOf(t.source);
    if (at >= 0) { start = at; end = at + t.source.length; } else { start = notes.length; end = notes.length; replace = false; }
  }
  let updated;
  let caret;
  if (replace) {
    updated = notes.slice(0, start) + text + notes.slice(end);
    caret = start + text.length;
  } else {
    const sep = start === 0 && end === notes.length ? '\n\n' : ' ';
    updated = notes.slice(0, end) + sep + text + notes.slice(end);
    caret = end + sep.length + text.length;
  }
  setNotes(t.slideId, updated);
  const box = $('notesBox');
  box.value = updated;
  box.focus();
  box.setSelectionRange(caret, caret);
  $('notesFlag').textContent = notesEdited(t.slideId) ? 'edited, written to the pptx on save' : '';
  renderOutline();
  if (replace) closeTranslation();
}

function closeTranslation() {
  translation = null;
  const panel = $('translatePanel');
  if (panel) { panel.hidden = true; panel.innerHTML = ''; }
}

// ---------- save, download, backup ----------

async function buildPptx() {
  const plan = {
    groups: project.groups.map((g) => ({ name: g.name, slideIds: g.slideIds })),
    cut: new Set(project.info.slides.filter((s) => decisionOf(s.id) === 'cut').map((s) => s.id)),
    notes: {},
    texts: {},
  };
  for (const s of project.info.slides) {
    if (notesEdited(s.id)) plan.notes[s.id] = st(s.id).notes;
    if (slideTextEdited(s.id)) plan.texts[s.id] = st(s.id).text;
  }
  const { bytes, skippedNotes } = await writePptx(JSZip, originalBytes, project.info, plan);
  await store.putFile(project.id + ':saved', bytes);
  project.pptxDirty = false;
  project.pptxSavedAt = Date.now();
  await flushSave();
  if (skippedNotes.length) {
    toast(`Saved, but notes could not be written for ${skippedNotes.length} slide(s): the file has no notes layout.`, 6000);
  }
  return bytes;
}

$('saveBtn').addEventListener('click', async () => {
  try {
    $('saveBtn').disabled = true;
    await buildPptx();
    toast('Saved. The pptx now has the new order, sections, hidden cut slides and edited notes.');
  } catch (err) {
    console.error(err);
    toast('Could not update the pptx: ' + err.message, 6000);
  } finally {
    $('saveBtn').disabled = false;
  }
});

function downloadBytes(bytes, name, type) {
  const url = URL.createObjectURL(new Blob([bytes], { type }));
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

$('downloadBtn').addEventListener('click', async () => {
  try {
    let bytes = await store.getFile(project.id + ':saved');
    if (!bytes || project.pptxDirty) bytes = await buildPptx();
    const stamp = new Date().toLocaleString('sv-SE').slice(0, 16).replace(/[-: ]/g, '').replace(/^(\d{8})(\d{4})$/, '$1-$2');
    const name = project.fileName.replace(/\.pptx$/i, '') + ` (planned ${stamp}).pptx`;
    downloadBytes(bytes, name, 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
  } catch (err) {
    console.error(err);
    toast('Could not create the pptx: ' + err.message, 6000);
  }
});

function toBase64(bytes) {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

function fromBase64(text) {
  const s = atob(text);
  const bytes = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) bytes[i] = s.charCodeAt(i);
  return bytes;
}

$('backupBtn').addEventListener('click', async () => {
  await flushSave();
  const backup = { format: 'presentation-helper-backup', version: 1, project, original: toBase64(originalBytes) };
  downloadBytes(JSON.stringify(backup), project.name + ' (plan backup).json', 'application/json');
});

$('backupInput').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  e.target.value = '';
  if (!file) return;
  try {
    const backup = JSON.parse(await file.text());
    if (backup.format !== 'presentation-helper-backup') throw new Error('not a Presentation helper backup');
    const p = backup.project;
    const existing = await store.getProject(p.id);
    if (existing && !confirm('This presentation already exists in this browser. Replace it with the backup?')) return;
    await store.putFile(p.id + ':original', fromBase64(backup.original));
    p.pptxDirty = true;
    await store.putProject(p);
    openProject(p.id);
    toast('Backup restored.');
  } catch (err) {
    toast('Could not restore the backup: ' + err.message, 6000);
  }
});

// ---------- PDF thumbnails ----------

$('pdfBtn').addEventListener('click', () => $('pdfInput').click());
$('pdfInput').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  e.target.value = '';
  if (!file) return;
  try {
    const { renderPdfThumbs } = await import('./thumbs.js');
    const bytes = new Uint8Array(await file.arrayBuffer());
    const result = await renderPdfThumbs(bytes, project.info.slides, (i, n) => toast(`Reading PDF page ${i} of ${n}…`, 60000));
    project.thumbs = result.thumbs;
    changed({ pptx: false });
    renderAll();
    const total = project.info.slides.length;
    let message = `Thumbnails for ${result.matched} of ${total} slides`;
    if (result.matched < total / 2) message += `. Few pages matched the slide text. Is the PDF an export of this presentation?`;
    else if (result.matched < total) message += ` (the PDF has ${result.pages} pages; hidden or changed slides may be missing).`;
    else message += '.';
    toast(message, 8000);
  } catch (err) {
    console.error(err);
    toast('Could not read the PDF: ' + err.message, 6000);
  }
});

window.addEventListener('beforeunload', () => {
  if (project) store.putProject(project);
});

showHome();
