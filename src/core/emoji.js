// Effet « emoji qui pousse » : à l'emplacement de chaque lettre, un emoji part de très petit et grandit ;
// les emoji d'une série se succèdent en fondu, puis le dernier se fond dans la lettre.
// Tout est en pixels de l'image source et en millisecondes ; la conversion vers la diapo se fait ailleurs.

import { rng, shuffled } from './effects.js';

export const START_SCALE = 0.05; // taille de départ, en proportion de la taille finale
const GROW = 0.75;               // part de la durée consacrée à la croissance ; le reste est le fondu vers la lettre
const CROSS = 0.4;               // largeur d'un fondu entre deux emoji, en proportion d'une étape
const MIN_DUR = 100;

function graphemes(text) {
  const s = (text || '').normalize('NFC');
  if (typeof Intl !== 'undefined' && Intl.Segmenter) {
    return Array.from(new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(s), (g) => g.segment);
  }
  return Array.from(s);
}

/** Séries saisies par l'utilisateur : jusqu'à 3 séries de 1 à 3 emoji, les séries vides sont ignorées. */
export function parseSeries(texts) {
  return texts.slice(0, 3)
    .map((t) => graphemes(t).filter((g) => !/^\s+$/u.test(g)).slice(0, 3))
    .filter((s) => s.length);
}

/** Ordre de passage des lettres : lignes, puis mots de chaque ligne, puis lettres de chaque mot. */
export function readingOrder(seg, o, rand) {
  const pick = (list, mode) => (mode === 'random' ? shuffled(list.length, rand).map((i) => list[i]) : list);
  const uniq = (list) => [...new Set(list)];
  const letters = seg.letters.map((l, index) => ({ index, line: l.line, word: l.word }));
  const out = [];
  for (const line of pick(uniq(letters.map((l) => l.line)), o.orderLine)) {
    const inLine = letters.filter((l) => l.line === line);
    for (const word of pick(uniq(inLine.map((l) => l.word)), o.orderWord)) {
      for (const l of pick(inLine.filter((x) => x.word === word), o.orderLetter)) out.push(l.index);
    }
  }
  return out;
}

// Taille à l'instant t (ms depuis le début de la lettre) : croissance continue, ralentie à la fin.
export function scaleAt(it, t) {
  const u = Math.min(1, Math.max(0, t / it.grow));
  return START_SCALE + (1 - START_SCALE) * (1 - (1 - u) * (1 - u));
}

/**
 * @param seg résultat de segment()
 * @param o   { series: [[emoji…]…], mean, sd, total (ms), orderLine, orderWord, orderLetter ('seq' | 'random'), seed }
 * @returns { items, total, parallel, objects }
 *   items[i] (une par lettre, dans l'ordre de seg.letters) :
 *     { index, start, dur, grow, box: { cx, bottom, size }, emojis: [{ ch, a, fadeIn, b, fadeOut, end, kf: [{ t, s }] }] }
 *   Les temps a, b, end et kf[].t sont comptés depuis le début de la lettre.
 */
export function emojiPlan(seg, o) {
  if (!o.series || !o.series.length) throw new Error('indiquez au moins un emoji dans une série.');
  const rand = rng(o.seed ?? 1);
  const n = seg.letters.length;
  const total = Math.max(MIN_DUR, o.total);

  // Durées : loi normale (moyenne, écart-type), bornées entre un plancher et la durée totale
  const normal = () => Math.sqrt(-2 * Math.log(1 - rand())) * Math.cos(2 * Math.PI * rand());
  const durs = seg.letters.map(() => Math.min(total, Math.max(MIN_DUR, o.mean + o.sd * normal())));

  // Départs régulièrement espacés dans l'ordre choisi, calés pour que la dernière fin tombe sur la durée totale
  const order = readingOrder(seg, o, rand);
  const rank = new Array(n);
  order.forEach((index, r) => { rank[index] = r; });
  let step = Infinity;
  for (let i = 0; i < n; i++) if (rank[i] > 0) step = Math.min(step, (total - durs[i]) / rank[i]);
  if (!Number.isFinite(step)) { step = 0; durs[0] = total; } // une seule lettre : elle occupe toute la durée

  const items = seg.letters.map((l, i) => {
    const dur = durs[i];
    const series = o.series[Math.floor(rand() * o.series.length)];
    const k = series.length;
    const grow = dur * GROW;
    const stage = grow / k;          // durée d'une étape (un emoji)
    const cross = stage * CROSS;
    const band = seg.lineBands[l.line];
    const it = {
      index: i, start: rank[i] * step, dur, grow,
      box: { cx: (l.x0 + l.x1 + 1) / 2, bottom: band.y1 + 1, size: band.y1 - band.y0 + 1 }, // même taille pour toute la ligne
    };
    it.emojis = series.map((ch, j) => {
      const a = j === 0 ? 0 : j * stage - cross / 2;
      const fadeIn = j === 0 ? Math.min(cross, 0.15 * stage) : cross;
      const last = j === k - 1;
      const b = last ? grow : (j + 1) * stage - cross / 2;
      const fadeOut = last ? dur - grow : cross;
      const end = b + fadeOut;
      const stop = Math.min(end, grow); // la taille ne change plus après la croissance
      const kf = [];
      for (let s = 0; s <= 5; s++) { const t = a + ((stop - a) * s) / 5; kf.push({ t, s: scaleAt(it, t) }); }
      if (end > stop + 1) kf.push({ t: end, s: 1 });
      return { ch, a, fadeIn, b, fadeOut, end, kf };
    });
    return it;
  });

  const end = Math.max(...items.map((it) => it.start + it.dur));
  return {
    items, total: end,
    parallel: durs.reduce((s, d) => s + d, 0) / end,
    objects: items.reduce((s, it) => s + it.emojis.length + 1, 0),
  };
}
