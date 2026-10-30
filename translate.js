// translate.js: Swedish to English with several alternatives, from two free online sources.
// The text you translate is sent to Google Translate and MyMemory (nothing else is).

const GOOGLE = 'https://translate.googleapis.com/translate_a/single';
const MYMEMORY = 'https://api.mymemory.translated.net/get';

// Google: one translation per sentence, alternative wordings per sentence,
// and for single words a dictionary with meanings and their Swedish back-translations.
export async function googleTranslate(text, from = 'sv', to = 'en') {
  const params = new URLSearchParams({ client: 'gtx', sl: from, tl: to, dj: '1', q: text });
  const url = GOOGLE + '?' + params + '&dt=t&dt=at&dt=bd';
  const res = await fetch(url);
  if (!res.ok) throw new Error('Google answered ' + res.status);
  const data = await res.json();

  // Each segment is a sentence (or the whole text). Keep the leading space flag so
  // the translation can be put back together after the user picks alternatives.
  const segments = (data.alternative_translations || []).map((seg) => ({
    source: seg.src_phrase,
    options: dedupe((seg.alternative || []).map((a) => a.word_postproc)),
    space: seg.alternative && seg.alternative[0] ? seg.alternative[0].has_preceding_space : true,
  })).filter((s) => s.options.length > 0);

  const main = (data.sentences || []).map((s) => s.trans || '').join('');
  const dictionary = (data.dict || []).map((d) => ({
    pos: d.pos,
    entries: (d.entry || []).map((e) => ({ word: e.word, back: e.reverse_translation || [] })),
  }));
  return { main, segments, dictionary };
}

export function joinSegments(segments, chosen) {
  return segments.map((s, i) => (i > 0 && s.space ? ' ' : '') + s.options[chosen[i] || 0]).join('');
}

// MyMemory: translation memories (often human translations from EU and other sources)
// plus machine translation. Quality varies, so each match shows where it comes from.
export async function myMemoryTranslate(text, from = 'sv', to = 'en') {
  const params = new URLSearchParams({ q: text.slice(0, 500), langpair: `${from}|${to}` });
  const res = await fetch(MYMEMORY + '?' + params);
  if (!res.ok) throw new Error('MyMemory answered ' + res.status);
  const data = await res.json();
  if (data.responseStatus && Number(data.responseStatus) !== 200) {
    throw new Error(data.responseDetails || 'MyMemory error ' + data.responseStatus);
  }
  const seen = new Set();
  const matches = [];
  for (const m of data.matches || []) {
    const t = (m.translation || '').trim();
    if (!t || seen.has(t.toLowerCase())) continue;
    seen.add(t.toLowerCase());
    const machine = m['created-by'] === 'MT!' || String(m['created-by']).startsWith('MT');
    matches.push({ text: t, source: machine ? 'machine translation' : (m['created-by'] || 'translation memory'), match: Number(m.match) || 0 });
  }
  return matches.slice(0, 6);
}

export function externalLinks(text, isWord) {
  const enc = encodeURIComponent(text);
  const links = [
    { label: 'Google Translate', url: `https://translate.google.com/?sl=sv&tl=en&text=${enc}&op=translate` },
    { label: 'DeepL', url: `https://www.deepl.com/translator#sv/en/${enc}` },
  ];
  if (isWord) {
    links.push({ label: 'Glosbe (examples in context)', url: `https://glosbe.com/sv/en/${enc}` });
  }
  return links;
}

function dedupe(list) {
  const seen = new Set();
  return list.filter((x) => {
    const k = (x || '').trim().toLowerCase();
    if (!k || seen.has(k)) return false;
    seen.add(k);
    return true;
  });
}
