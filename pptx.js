// pptx.js: read and rewrite .pptx files in the browser.
// A .pptx is a zip of XML parts. We only touch what the planner changes:
// slide order (p:sldIdLst), PowerPoint sections (p14:sectionLst),
// hidden flag on cut slides (p:sld/@show) and speaker notes (notesSlide body).

const NS = {
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  p14: 'http://schemas.microsoft.com/office/powerpoint/2010/main',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
};
const REL_SLIDE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
const REL_NOTES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide';
const REL_LAYOUT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout';
const REL_NOTES_MASTER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster';
const SECTION_EXT_URI = '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}';
const NOTES_CT = 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml';

function parseXml(text) {
  return new DOMParser().parseFromString(text, 'application/xml');
}

function serializeXml(doc) {
  let out = new XMLSerializer().serializeToString(doc);
  if (!out.startsWith('<?xml')) {
    out = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + out;
  }
  return out;
}

async function readXml(zip, path) {
  const file = zip.file(path);
  if (!file) return null;
  return parseXml(await file.async('string'));
}

// Resolve a relationship target relative to the part that owns the .rels file.
function resolvePath(basePart, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = basePart.split('/');
  parts.pop();
  for (const piece of target.split('/')) {
    if (piece === '..') parts.pop();
    else if (piece !== '.') parts.push(piece);
  }
  return parts.join('/');
}

function relsPathFor(part) {
  const parts = part.split('/');
  const name = parts.pop();
  return [...parts, '_rels', name + '.rels'].join('/');
}

async function readRels(zip, part) {
  const doc = await readXml(zip, relsPathFor(part));
  const rels = [];
  if (!doc) return { doc: null, rels };
  for (const el of Array.from(doc.getElementsByTagNameNS(NS.rel, 'Relationship'))) {
    rels.push({
      id: el.getAttribute('Id'),
      type: el.getAttribute('Type'),
      target: el.getAttribute('Target'),
      path: el.getAttribute('TargetMode') === 'External' ? null : resolvePath(part, el.getAttribute('Target')),
    });
  }
  return { doc, rels };
}

function textOfParagraph(pEl) {
  let s = '';
  for (const node of Array.from(pEl.childNodes)) {
    if (node.localName === 'r' || node.localName === 'fld') {
      const t = node.getElementsByTagNameNS(NS.a, 't')[0];
      if (t) s += t.textContent;
    } else if (node.localName === 'br') {
      s += '\n';
    }
  }
  return s;
}

function paragraphsOf(el) {
  return Array.from(el.getElementsByTagNameNS(NS.a, 'p')).map(textOfParagraph);
}

function placeholderType(sp) {
  const ph = sp.getElementsByTagNameNS(NS.p, 'ph')[0];
  if (!ph) return null;
  return ph.getAttribute('type') || 'body';
}

function notesBodyShape(notesDoc) {
  for (const sp of Array.from(notesDoc.getElementsByTagNameNS(NS.p, 'sp'))) {
    if (placeholderType(sp) === 'body') return sp;
  }
  return null;
}

// Read everything the planner needs from a .pptx (ArrayBuffer/Uint8Array).
export async function readPptx(JSZip, data) {
  const zip = await JSZip.loadAsync(data);
  const presPath = 'ppt/presentation.xml';
  const pres = await readXml(zip, presPath);
  if (!pres) throw new Error('Not a PowerPoint file (ppt/presentation.xml missing).');
  const { rels: presRels } = await readRels(zip, presPath);
  const relById = Object.fromEntries(presRels.map((r) => [r.id, r]));

  const slides = [];
  const sldIdLst = pres.getElementsByTagNameNS(NS.p, 'sldIdLst')[0];
  const sldIds = sldIdLst ? Array.from(sldIdLst.getElementsByTagNameNS(NS.p, 'sldId')) : [];
  for (const [index, el] of sldIds.entries()) {
    const id = el.getAttribute('id');
    const rel = relById[el.getAttributeNS(NS.r, 'id')];
    const path = rel.path;
    const doc = await readXml(zip, path);
    const { rels } = await readRels(zip, path);

    let title = '';
    const body = [];
    for (const sp of Array.from(doc.getElementsByTagNameNS(NS.p, 'sp'))) {
      const txBody = sp.getElementsByTagNameNS(NS.p, 'txBody')[0];
      if (!txBody) continue;
      const lines = paragraphsOf(txBody).map((l) => l.trim()).filter(Boolean);
      const type = placeholderType(sp);
      if ((type === 'title' || type === 'ctrTitle') && !title) title = lines.join(' ');
      else body.push(...lines);
    }
    // Tables and grouped shapes also carry text in a:t; pick it up for search and PDF matching.
    for (const tbl of Array.from(doc.getElementsByTagNameNS(NS.a, 'tbl'))) {
      body.push(...paragraphsOf(tbl).map((l) => l.trim()).filter(Boolean));
    }

    let layoutType = null;
    const layoutRel = rels.find((r) => r.type === REL_LAYOUT);
    if (layoutRel) {
      const layoutDoc = await readXml(zip, layoutRel.path);
      const cSld = layoutDoc && layoutDoc.documentElement;
      layoutType = cSld ? cSld.getAttribute('type') : null;
    }

    let notes = '';
    let notesPath = null;
    const notesRel = rels.find((r) => r.type === REL_NOTES);
    if (notesRel) {
      notesPath = notesRel.path;
      const notesDoc = await readXml(zip, notesPath);
      const sp = notesDoc && notesBodyShape(notesDoc);
      if (sp) notes = paragraphsOf(sp).join('\n');
    }

    slides.push({
      id,
      originalIndex: index,
      path,
      title: title || '(untitled)',
      hasTitle: Boolean(title),
      body,
      layoutType,
      hidden: doc.documentElement.getAttribute('show') === '0',
      notes,
      notesPath,
    });
  }

  // Existing PowerPoint sections, if any.
  const sections = [];
  for (const sec of Array.from(pres.getElementsByTagNameNS(NS.p14, 'section'))) {
    sections.push({
      name: sec.getAttribute('name'),
      slideIds: Array.from(sec.getElementsByTagNameNS(NS.p14, 'sldId')).map((e) => e.getAttribute('id')),
    });
  }

  const hasNotesMaster = presRels.some((r) => r.type === REL_NOTES_MASTER);
  return { slides, sections, hasNotesMaster };
}

// Suggest groups: existing sections first, otherwise start a new group at
// section header slides (or title-only slides when the deck has none).
export function suggestGroups(info) {
  const { slides, sections } = info;
  if (sections.length > 0) {
    const groups = sections.map((s) => ({ name: s.name || 'Section', slideIds: s.slideIds.slice() }));
    const seen = new Set(groups.flatMap((g) => g.slideIds));
    const rest = slides.filter((s) => !seen.has(s.id)).map((s) => s.id);
    if (rest.length) groups.push({ name: 'Other slides', slideIds: rest });
    return groups;
  }

  let isStart = (s) => s.layoutType === 'secHead';
  if (slides.filter(isStart).length < 2) {
    isStart = (s) => s.layoutType === 'secHead' || (s.hasTitle && s.body.length === 0);
  }
  const groups = [];
  for (const [i, s] of slides.entries()) {
    if (i === 0 || isStart(s)) groups.push({ name: s.title, slideIds: [] });
    groups[groups.length - 1].slideIds.push(s.id);
  }
  return groups;
}

function newGuid() {
  const hex = () => Math.floor(Math.random() * 16).toString(16).toUpperCase();
  const pattern = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx';
  return '{' + pattern.replace(/[xy]/g, (c) => (c === 'x' ? hex() : '89AB'[Math.floor(Math.random() * 4)])) + '}';
}

function fillNotesBody(doc, txBody, text) {
  // Keep bodyPr and lstStyle, replace the paragraphs with plain text lines.
  for (const p of Array.from(txBody.getElementsByTagNameNS(NS.a, 'p'))) {
    if (p.parentNode === txBody) txBody.removeChild(p);
  }
  const lines = text.replace(/\r\n/g, '\n').split('\n');
  for (const line of lines) {
    const p = doc.createElementNS(NS.a, 'a:p');
    if (line.length > 0) {
      const r = doc.createElementNS(NS.a, 'a:r');
      const rPr = doc.createElementNS(NS.a, 'a:rPr');
      rPr.setAttribute('lang', 'en-US');
      rPr.setAttribute('dirty', '0');
      const t = doc.createElementNS(NS.a, 'a:t');
      t.textContent = line;
      r.appendChild(rPr);
      r.appendChild(t);
      p.appendChild(r);
    } else {
      const end = doc.createElementNS(NS.a, 'a:endParaRPr');
      end.setAttribute('lang', 'en-US');
      end.setAttribute('dirty', '0');
      p.appendChild(end);
    }
    txBody.appendChild(p);
  }
}

function escapeXml(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function nextRelId(relsDoc) {
  let max = 0;
  for (const el of Array.from(relsDoc.getElementsByTagNameNS(NS.rel, 'Relationship'))) {
    const n = parseInt((el.getAttribute('Id') || '').replace(/\D/g, ''), 10);
    if (n > max) max = n;
  }
  return 'rId' + (max + 1);
}

function addRelationship(relsDoc, id, type, target) {
  const el = relsDoc.createElementNS(NS.rel, 'Relationship');
  el.setAttribute('Id', id);
  el.setAttribute('Type', type);
  el.setAttribute('Target', target);
  relsDoc.documentElement.appendChild(el);
}

async function createNotesSlide(zip, slidePath, notesMasterPath, text) {
  let n = 1;
  while (zip.file(`ppt/notesSlides/notesSlide${n}.xml`)) n++;
  const notesPath = `ppt/notesSlides/notesSlide${n}.xml`;
  const slideFile = slidePath.split('/').pop();
  const masterFile = notesMasterPath.split('/').pop();

  const paragraphs = text.replace(/\r\n/g, '\n').split('\n').map((line) =>
    line.length > 0
      ? `<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${escapeXml(line)}</a:t></a:r></a:p>`
      : '<a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p>'
  ).join('');
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<p:notes xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld><p:spTree>` +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>' +
    `<p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody></p:sp>` +
    '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>';
  zip.file(notesPath, xml);

  const notesRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<Relationships xmlns="${NS.rel}">` +
    `<Relationship Id="rId1" Type="${REL_NOTES_MASTER}" Target="../notesMasters/${masterFile}"/>` +
    `<Relationship Id="rId2" Type="${REL_SLIDE}" Target="../slides/${slideFile}"/>` +
    '</Relationships>';
  zip.file(relsPathFor(notesPath), notesRels);

  const slideRelsPath = relsPathFor(slidePath);
  let slideRels = await readXml(zip, slideRelsPath);
  if (!slideRels) slideRels = parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS.rel}"/>`);
  addRelationship(slideRels, nextRelId(slideRels), REL_NOTES, `../notesSlides/notesSlide${n}.xml`);
  zip.file(slideRelsPath, serializeXml(slideRels));

  const ct = await readXml(zip, '[Content_Types].xml');
  const override = ct.createElementNS(NS.ct, 'Override');
  override.setAttribute('PartName', '/' + notesPath);
  override.setAttribute('ContentType', NOTES_CT);
  ct.documentElement.appendChild(override);
  zip.file('[Content_Types].xml', serializeXml(ct));
}

// Write a new .pptx from the original bytes and the plan.
// plan = { groups: [{ name, slideIds }], cut: Set of slide ids, notes: { [slideId]: text } }
// Only notes listed in plan.notes are rewritten; others keep their original formatting.
export async function writePptx(JSZip, data, info, plan) {
  const zip = await JSZip.loadAsync(data);
  const presPath = 'ppt/presentation.xml';
  const pres = await readXml(zip, presPath);
  const { rels: presRels } = await readRels(zip, presPath);
  const byId = Object.fromEntries(info.slides.map((s) => [s.id, s]));

  // 1. Slide order.
  const order = plan.groups.flatMap((g) => g.slideIds);
  const sldIdLst = pres.getElementsByTagNameNS(NS.p, 'sldIdLst')[0];
  const elById = {};
  for (const el of Array.from(sldIdLst.getElementsByTagNameNS(NS.p, 'sldId'))) {
    elById[el.getAttribute('id')] = el;
  }
  const missing = Object.keys(elById).filter((id) => !order.includes(id));
  for (const id of [...order, ...missing]) sldIdLst.appendChild(elById[id]);

  // 2. Sections: replace any existing section list with the planner's groups.
  let extLst = null;
  for (const child of Array.from(pres.documentElement.childNodes)) {
    if (child.localName === 'extLst') extLst = child;
  }
  if (!extLst) {
    extLst = pres.createElementNS(NS.p, 'p:extLst');
    pres.documentElement.appendChild(extLst);
  }
  for (const ext of Array.from(extLst.childNodes)) {
    if (ext.localName === 'ext' && ext.getAttribute('uri') === SECTION_EXT_URI) extLst.removeChild(ext);
  }
  const ext = pres.createElementNS(NS.p, 'p:ext');
  ext.setAttribute('uri', SECTION_EXT_URI);
  const secLst = pres.createElementNS(NS.p14, 'p14:sectionLst');
  const sectionGroups = plan.groups.filter((g) => g.slideIds.length > 0);
  if (missing.length) sectionGroups.push({ name: 'Other slides', slideIds: missing });
  for (const g of sectionGroups) {
    const sec = pres.createElementNS(NS.p14, 'p14:section');
    sec.setAttribute('name', g.name || 'Section');
    sec.setAttribute('id', newGuid());
    const lst = pres.createElementNS(NS.p14, 'p14:sldIdLst');
    for (const id of g.slideIds) {
      const s = pres.createElementNS(NS.p14, 'p14:sldId');
      s.setAttribute('id', id);
      lst.appendChild(s);
    }
    sec.appendChild(lst);
    secLst.appendChild(sec);
  }
  ext.appendChild(secLst);
  extLst.insertBefore(ext, extLst.firstChild);
  zip.file(presPath, serializeXml(pres));

  // 3. Hidden flag for cut slides (kept in the file, skipped in the slide show).
  for (const s of info.slides) {
    const shouldHide = plan.cut.has(s.id);
    if (shouldHide === s.hidden) continue;
    const doc = await readXml(zip, s.path);
    if (shouldHide) doc.documentElement.setAttribute('show', '0');
    else doc.documentElement.removeAttribute('show');
    zip.file(s.path, serializeXml(doc));
  }

  // 4. Speaker notes.
  const notesMaster = presRels.find((r) => r.type === REL_NOTES_MASTER);
  const skippedNotes = [];
  for (const [id, text] of Object.entries(plan.notes)) {
    const s = byId[id];
    if (!s) continue;
    if (s.notesPath) {
      const doc = await readXml(zip, s.notesPath);
      const sp = notesBodyShape(doc);
      if (!sp) { skippedNotes.push(id); continue; }
      let txBody = sp.getElementsByTagNameNS(NS.p, 'txBody')[0];
      if (!txBody) {
        txBody = doc.createElementNS(NS.p, 'p:txBody');
        txBody.appendChild(doc.createElementNS(NS.a, 'a:bodyPr'));
        txBody.appendChild(doc.createElementNS(NS.a, 'a:lstStyle'));
        sp.appendChild(txBody);
      }
      fillNotesBody(doc, txBody, text);
      zip.file(s.notesPath, serializeXml(doc));
    } else if (notesMaster) {
      await createNotesSlide(zip, s.path, notesMaster.path, text);
    } else {
      skippedNotes.push(id);
    }
  }

  const blob = await zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  });
  return { bytes: blob, skippedNotes };
}
