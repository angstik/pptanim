// Découpage d'une image RGBA de texte (fond transparent) en lettres, mots et lignes.
// Aucune dépendance : fonctionne sur { width, height, data } (data = RGBA 8 bits).

// Caractères dessinés en plusieurs morceaux côte à côte (un seul caractère, plusieurs blocs).
const MULTI = { '"': 2, '“': 2, '”': 2, '„': 2, '«': 2, '»': 2, '…': 3 };

const box = (b) => ({ x0: b.x0, y0: b.y0, x1: b.x1, y1: b.y1 });
const union = (a, b) => ({
  x0: Math.min(a.x0, b.x0), y0: Math.min(a.y0, b.y0),
  x1: Math.max(a.x1, b.x1), y1: Math.max(a.y1, b.y1),
});

function graphemes(text) {
  const s = (text || '').normalize('NFC');
  if (typeof Intl !== 'undefined' && Intl.Segmenter) {
    return Array.from(new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(s), (g) => g.segment);
  }
  return Array.from(s);
}

// Composantes connexes (8-voisinage) du masque alpha.
function label(mask, W, H) {
  const labels = new Int32Array(W * H);
  const stack = new Int32Array(W * H);
  const comps = [];
  for (let start = 0; start < W * H; start++) {
    if (!mask[start] || labels[start]) continue;
    const id = comps.length + 1;
    const c = { id, x0: W, y0: H, x1: 0, y1: 0, area: 0 };
    let top = 0;
    stack[top++] = start;
    labels[start] = id;
    while (top) {
      const p = stack[--top];
      const x = p % W, y = (p - x) / W;
      c.area++;
      if (x < c.x0) c.x0 = x;
      if (x > c.x1) c.x1 = x;
      if (y < c.y0) c.y0 = y;
      if (y > c.y1) c.y1 = y;
      for (let dy = -1; dy <= 1; dy++) {
        const yy = y + dy;
        if (yy < 0 || yy >= H) continue;
        for (let dx = -1; dx <= 1; dx++) {
          const xx = x + dx;
          if (xx < 0 || xx >= W) continue;
          const q = yy * W + xx;
          if (mask[q] && !labels[q]) { labels[q] = id; stack[top++] = q; }
        }
      }
    }
    comps.push(c);
  }
  return { labels, comps };
}

// Bandes horizontales (lignes de texte). Les petites bandes proches (accents, points) sont rattachées.
function bands(comps) {
  const sorted = [...comps].sort((a, b) => a.y0 - b.y0);
  let out = [];
  for (const c of sorted) {
    const last = out[out.length - 1];
    if (last && c.y0 <= last.y1) last.y1 = Math.max(last.y1, c.y1);
    else out.push({ y0: c.y0, y1: c.y1 });
  }
  const h = (b) => b.y1 - b.y0 + 1;
  const maxH = Math.max(...out.map(h));
  let changed = true;
  while (changed && out.length > 1) {
    changed = false;
    for (let i = 0; i < out.length; i++) {
      if (h(out[i]) >= 0.45 * maxH) continue;
      const below = out[i + 1] ? out[i + 1].y0 - out[i].y1 : Infinity;
      const above = out[i - 1] ? out[i].y0 - out[i - 1].y1 : Infinity;
      const j = below <= above ? i + 1 : i - 1;
      if (Math.min(below, above) > 0.35 * maxH) continue;
      out[j] = { y0: Math.min(out[i].y0, out[j].y0), y1: Math.max(out[i].y1, out[j].y1) };
      out.splice(i, 1);
      changed = true;
      break;
    }
  }
  return out;
}

// Mots d'un texte : pour chaque mot, ses caractères et le nombre de blocs que chacun dessine.
function wordsOf(text) {
  const words = [];
  let open = false;
  for (const g of graphemes(text)) {
    if (/^\s+$/u.test(g)) { open = false; continue; }
    if (!open) { words.push([]); open = true; }
    words[words.length - 1].push({ ch: g, n: MULTI[g] || 1 });
  }
  return words;
}

/**
 * Ressemblance entre une image découpée et un texte, pour retrouver à quelle zone de texte l'image
 * correspond. @param seenCounts nombre de blocs de chaque mot vu dans l'image (seg.seenCounts)
 * @returns 0 à 1 : 1 si l'image montre les mêmes mots, avec le même nombre de lettres chacun
 */
export function matchScore(seenCounts, text) {
  const wanted = wordsOf(text).map((w) => w.reduce((sum, t) => sum + t.n, 0));
  if (!wanted.length || !seenCounts.length) return 0;
  if (wanted.length !== seenCounts.length) {
    // Pas le même nombre de mots : au mieux une ressemblance lointaine
    return 0.3 * Math.max(0, 1 - Math.abs(wanted.length - seenCounts.length) / Math.max(wanted.length, seenCounts.length));
  }
  return 0.5 + 0.5 * (wanted.filter((n, i) => n === seenCounts[i]).length / wanted.length);
}

/**
 * @param img  { width, height, data } RGBA
 * @param text texte de la forme (optionnel) : sert à fiabiliser lettres et mots
 * @returns { letters, words, lines, matched, aligned, fused, expected, found, labels, maxLineH }
 *   aligned : l'image montre autant de mots que le texte ; fused : mots restés entiers faute de découpage fiable
 *   Chaque unité : { x0, y0, x1, y1, ids:[composantes], line, text }
 */
export function segment(img, text, opts = {}) {
  const { width: W, height: H, data } = img;
  const alphaMin = opts.alphaMin ?? 24;
  const mask = new Uint8Array(W * H);
  let on = 0;
  for (let i = 0, p = 3; i < W * H; i++, p += 4) if (data[p] > alphaMin) { mask[i] = 1; on++; }
  if (!on) throw new Error("Image vide : aucun pixel visible.");
  if (on > 0.6 * W * H) {
    throw new Error("Le fond de l'image n'est pas transparent : la zone de texte a sans doute un remplissage. Seul le texte sur fond transparent est pris en charge.");
  }

  let { labels, comps } = label(mask, W, H);
  const minArea = opts.minArea ?? 6;
  const dropped = new Set(comps.filter((c) => c.area < minArea).map((c) => c.id));
  if (dropped.size) {
    for (let i = 0; i < labels.length; i++) if (dropped.has(labels[i])) labels[i] = -1; // poussière : ignorée partout
    comps = comps.filter((c) => !dropped.has(c.id));
  }
  if (!comps.length) throw new Error("Aucune lettre détectée.");

  const lineBands = bands(comps);
  const maxLineH = Math.max(...lineBands.map((b) => b.y1 - b.y0 + 1));

  // Regroupe les composantes d'une même ligne qui se superposent horizontalement (i + point, e + accent, :, =, %…).
  const perLine = lineBands.map(() => []);
  for (const c of comps) {
    const cy = (c.y0 + c.y1) / 2;
    let li = lineBands.findIndex((b) => cy >= b.y0 && cy <= b.y1);
    if (li < 0) li = 0;
    perLine[li].push(c);
  }
  const blobs = []; // dans l'ordre de lecture
  perLine.forEach((list, li) => {
    list.sort((a, b) => a.x0 - b.x0);
    const groups = [];
    for (const c of list) {
      let merged = false;
      for (let k = groups.length - 1; k >= 0 && k >= groups.length - 3; k--) {
        const g = groups[k];
        const ov = Math.min(g.x1, c.x1) - Math.max(g.x0, c.x0) + 1;
        // Morceaux empilés (point du i, accent, deux-points) : un faible recouvrement suffit, même en italique.
        const stacked = c.y0 > g.y1 || c.y1 < g.y0;
        if (ov > (stacked ? 0.2 : 0.5) * Math.min(g.x1 - g.x0 + 1, c.x1 - c.x0 + 1)) {
          Object.assign(g, union(g, c));
          g.ids.push(c.id);
          merged = true;
          break;
        }
      }
      if (!merged) groups.push({ ...box(c), ids: [c.id], line: li });
    }
    groups.sort((a, b) => a.x0 - b.x0);
    blobs.push(...groups);
  });

  // Mots vus dans l'image : séparés par les grands espaces de chaque ligne.
  const thr = lineBands.map((_, li) => {
    const row = blobs.filter((b) => b.line === li);
    const gaps = row.slice(1).map((b, i) => b.x0 - row[i].x1).sort((p, q) => p - q);
    return Math.max(0.2 * maxLineH, 2 * (gaps.length ? gaps[gaps.length >> 1] : 0));
  });
  const seen = [];
  blobs.forEach((b, i) => {
    const prev = blobs[i - 1];
    if (!prev || prev.line !== b.line || b.x0 - prev.x1 > thr[b.line]) seen.push([]);
    seen[seen.length - 1].push(b);
  });

  // Mots du texte : un caractère non blanc = un bloc (ou plusieurs pour « " … »).
  const wanted = wordsOf(text);
  const expected = wanted.flat().reduce((sum, t) => sum + t.n, 0);
  const merge = (list, extra) => list.reduce((u, b) => ({ ...u, ...union(u, b), ids: u.ids.concat(b.ids) }), { ...box(list[0]), ids: [], line: list[0].line, ...extra });

  // Le texte n'est exploité que si l'image montre le même nombre de mots ; la vérification se fait
  // ensuite mot par mot, pour qu'une erreur ne décale pas les lettres des mots suivants.
  const aligned = wanted.length > 0 && wanted.length === seen.length;
  const letters = [], words = [];
  let fused = 0;
  seen.forEach((group, w) => {
    const tokens = aligned ? wanted[w] : null;
    const label = tokens ? tokens.map((t) => t.ch).join('') : '';
    words.push(merge(group, { text: label }));
    if (tokens && tokens.reduce((sum, t) => sum + t.n, 0) === group.length) {
      let k = 0;
      for (const t of tokens) { letters.push(merge(group.slice(k, k + t.n), { text: t.ch, word: w })); k += t.n; }
    } else if (tokens) {
      // Lettres liées ou qui se touchent : le mot reste entier.
      fused++;
      letters.push(merge(group, { text: label, word: w, fused: true }));
    } else {
      for (const b of group) letters.push(merge([b], { text: '', word: w }));
    }
  });
  const matched = aligned && fused === 0;

  const lines = [];
  for (const l of letters) {
    const ln = lines[l.line];
    if (!ln) lines[l.line] = { ...box(l), ids: [...l.ids], line: l.line, text: l.text };
    else { Object.assign(ln, union(ln, l)); ln.ids.push(...l.ids); ln.text += l.text; }
  }

  return {
    width: W, height: H, labels, maxLineH,
    letters, words, lines: lines.filter(Boolean),
    lineBands, matched, aligned, fused, expected, found: blobs.length,
    seenCounts: seen.map((group) => group.length),
  };
}

/** Unités selon le choix de l'utilisateur. */
export function unitsOf(seg, kind) {
  return kind === 'line' ? seg.lines : kind === 'word' ? seg.words : seg.letters;
}

/**
 * Extrait une unité : rectangle englobant + marge, pixels des autres lettres effacés.
 * @returns { x, y, w, h, data } (RGBA)
 */
export function cropUnit(img, seg, unit, pad = 2) {
  const W = img.width, H = img.height;
  const x = Math.max(0, unit.x0 - pad), y = Math.max(0, unit.y0 - pad);
  const x1 = Math.min(W - 1, unit.x1 + pad), y1 = Math.min(H - 1, unit.y1 + pad);
  const w = x1 - x + 1, h = y1 - y + 1;
  const mine = new Set(unit.ids);
  const out = new Uint8ClampedArray(w * h * 4);
  for (let j = 0; j < h; j++) {
    for (let i = 0; i < w; i++) {
      const p = (y + j) * W + (x + i);
      const lab = seg.labels[p];
      // lab = 0 : pixel très pâle du lissage, gardé ; lab > 0 : seulement s'il appartient à l'unité.
      if (lab === 0 || mine.has(lab)) {
        const s = p * 4, d = (j * w + i) * 4;
        out[d] = img.data[s]; out[d + 1] = img.data[s + 1]; out[d + 2] = img.data[s + 2]; out[d + 3] = img.data[s + 3];
      }
    }
  }
  return { x, y, w, h, data: out };
}
