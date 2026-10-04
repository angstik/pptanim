// Génération du XML d'animation PowerPoint (p:timing).
// Les décalages sont en fractions de la diapo (x : largeur, y : hauteur), comme ppt_x / ppt_y.

export const NS = {
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
};

const num = (v) => {
  const s = Math.abs(v).toFixed(5).replace(/\.?0+$/, '');
  return s === '' ? '0' : s;
};
const signed = (v) => (Math.abs(v) < 5e-6 ? '0' : (v < 0 ? '-' : '') + num(v));
const formula = (attr, v) => (Math.abs(v) < 5e-6 ? `#${attr}` : `#${attr}${v < 0 ? '-' : '+'}${num(v)}`);

const target = (spid) => `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>`;

/**
 * Entrée « ordonner » : la forme apparaît à sa position mélangée, attend, puis rejoint sa place.
 * Même structure que l'effet d'entrée « Entrée brusque » de PowerPoint, avec nos images clés.
 * @param kf [{ t (0..1), x, y }] décalages en fractions de diapo
 */
export function orderEffect({ nextId, spid, kf, total, nodeType }) {
  const tav = (attr, key) =>
    kf.map((k) => `<p:tav tm="${Math.round(k.t * 100000)}"><p:val><p:strVal val="${formula(attr, k[key])}"/></p:val></p:tav>`).join('');
  const anim = (attr, key) =>
    `<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="${nextId()}" dur="${total}" fill="hold"/>${target(spid)}` +
    `<p:attrNameLst><p:attrName>${attr}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst>${tav(attr, key)}</p:tavLst></p:anim>`;
  const outer = nextId();
  return (
    `<p:par><p:cTn id="${outer}" presetID="2" presetClass="entr" presetSubtype="4" fill="hold" nodeType="${nodeType}">` +
    `<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    `<p:set><p:cBhvr><p:cTn id="${nextId()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>${target(spid)}` +
    `<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>` +
    anim('ppt_x', 'x') + anim('ppt_y', 'y') +
    `</p:childTnLst></p:cTn></p:par>`
  );
}

/**
 * « Mélanger » : trajectoire personnalisée de la place normale vers la position mélangée.
 * @param off, c1, c2 en fractions de diapo
 */
export function shuffleEffect({ nextId, spid, off, c1, c2, curved, delay, dur, nodeType }) {
  const path = curved
    ? `M 0 0 C ${signed(c1.x)} ${signed(c1.y)} ${signed(c2.x)} ${signed(c2.y)} ${signed(off.x)} ${signed(off.y)}`
    : `M 0 0 L ${signed(off.x)} ${signed(off.y)}`;
  const outer = nextId();
  return (
    `<p:par><p:cTn id="${outer}" presetID="0" presetClass="path" presetSubtype="0" accel="50000" decel="50000" fill="hold" nodeType="${nodeType}">` +
    `<p:stCondLst><p:cond delay="${delay}"/></p:stCondLst><p:childTnLst>` +
    `<p:animMotion origin="layout" path="${path}" pathEditMode="relative"${curved ? '' : ' ptsTypes="AA"'}>` +
    `<p:cBhvr><p:cTn id="${nextId()}" dur="${dur}" fill="hold"/>${target(spid)}` +
    `<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr>` +
    `</p:animMotion></p:childTnLst></p:cTn></p:par>`
  );
}

/** Groupe « au clic » : le premier effet démarre au clic, les suivants avec lui. */
export function clickGroup({ nextId, effects }) {
  const g1 = nextId(), g2 = nextId();
  return (
    `<p:par><p:cTn id="${g1}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="${g2}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    effects(nextId) +
    `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`
  );
}

/** Séquence principale vide. */
export function mainSeq(id = 2) {
  return (
    `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${id}" dur="indefinite" nodeType="mainSeq"><p:childTnLst/></p:cTn>` +
    `<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>` +
    `<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>`
  );
}

export function emptyTiming() {
  return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>${mainSeq()}</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
}

// ---------- Chronologies complètes ----------
// ctx = { nextId, spids, fx, fy, ax, ay } : identifiants, numéros de formes dans l'ordre des images,
// conversions pixels → fractions de diapo (fx, fy : longueurs ; ax, ay : positions absolues).

/** Scrambler : un effet par unité, le premier au clic, les autres avec lui. */
export function scrambleTimeline(plan, keyframesOf, curved) {
  return ({ nextId, spids, fx, fy }) => {
    const frac = (p) => ({ x: fx(p.x), y: fy(p.y) });
    const order = [...plan.items].sort((p, q) => p.delay - q.delay || p.index - q.index);
    return order.map((it, k) => {
      const nodeType = k === 0 ? 'clickEffect' : 'withEffect';
      const spid = spids[it.index];
      if (plan.mode === 'order') {
        const kf = keyframesOf(it).map((p) => ({ t: p.t, x: fx(p.x), y: fy(p.y) }));
        return orderEffect({ nextId, spid, kf, total: Math.round(it.delay + it.dur), nodeType });
      }
      return shuffleEffect({ nextId, spid, off: frac(it.off), c1: frac(it.c1), c2: frac(it.c2), curved, delay: Math.round(it.delay), dur: Math.round(it.dur), nodeType });
    }).join('');
  };
}

const ms = (v) => Math.max(0, Math.round(v));
const ctn = (id, dur, delay, hold = true) =>
  `<p:cTn id="${id}" dur="${Math.max(1, ms(dur))}"${hold ? ' fill="hold"' : ''}>` +
  (delay > 0 || dur === 1 ? `<p:stCondLst><p:cond delay="${ms(delay)}"/></p:stCondLst>` : '') + `</p:cTn>`;
const setVisibility = (nextId, spid, value, delay) =>
  `<p:set><p:cBhvr>${ctn(nextId(), 1, delay)}${target(spid)}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>` +
  `<p:to><p:strVal val="${value}"/></p:to></p:set>`;
const fade = (nextId, spid, way, dur, delay) =>
  `<p:animEffect transition="${way}" filter="fade"><p:cBhvr>${ctn(nextId(), dur, delay, false)}${target(spid)}</p:cBhvr></p:animEffect>`;
// Propriété animée par images clés : frames = [{ tm (0..100000), val (texte) }]
const keyed = (nextId, spid, attr, frames, dur, delay) =>
  // Position : additive="base" comme dans les entrées natives ; taille : rien, comme dans le zoom natif
  `<p:anim calcmode="lin" valueType="num"><p:cBhvr${/^ppt_[xy]$/.test(attr) ? ' additive="base"' : ''}>${ctn(nextId(), dur, delay)}${target(spid)}` +
  `<p:attrNameLst><p:attrName>${attr}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst>` +
  frames.map((f) => `<p:tav tm="${f.tm}"><p:val><p:strVal val="${f.val}"/></p:val></p:tav>`).join('') + `</p:tavLst></p:anim>`;
// Un effet (une ligne du volet Animations) ; son identifiant est pris avant ceux de ses mouvements.
const effect = (nextId, attrs, nodeType, delay, body) => {
  const id = nextId();
  return `<p:par><p:cTn id="${id}" ${attrs} fill="hold" nodeType="${nodeType}"><p:stCondLst><p:cond delay="${ms(delay)}"/></p:stCondLst>` +
    `<p:childTnLst>${body(nextId)}</p:childTnLst></p:cTn></p:par>`;
};

const ZOOM_IN = 'presetID="53" presetClass="entr" presetSubtype="16"'; // zoom avec fondu
const FADE_IN = 'presetID="10" presetClass="entr" presetSubtype="0"';
const FADE_OUT = 'presetID="10" presetClass="exit" presetSubtype="0"';

/**
 * Emoji qui pousse. Chaque emoji est une image rangée hors de la diapo ; l'animation la place à
 * l'emplacement de la lettre, la fait grandir depuis le bas, puis la fond dans la suivante ou dans la lettre.
 * @param plan      résultat de emojiPlan ; chaque emoji porte pic (indice d'image), chaque lettre aussi
 * @param grouping  'image' : une animation par image ; 'split' : entrée et sortie séparées ;
 *                  'single' : une seule animation contenant tous les mouvements (expérimental)
 */
export function emojiTimeline(plan, grouping = 'image') {
  return ({ nextId, spids, fy, ax, ay }) => {
    const blocks = []; // { at, enter(id, t0) → mouvements d'entrée, leave(id, t0) → mouvements de sortie, exitAt }
    for (const it of plan.items) {
      const size = fy(it.box.size), cx = ax(it.box.cx), bottom = ay(it.box.bottom);
      for (const e of it.emojis) {
        const spid = spids[e.pic];
        const len = e.end - e.a;
        const frames = (val) => {
          const out = [];
          for (const k of e.kf) {
            const tm = Math.round((100000 * (k.t - e.a)) / len);
            if (out.length && tm <= out[out.length - 1].tm) continue;
            out.push({ tm, val: val(k.s) });
          }
          out[0].tm = 0;
          out[out.length - 1].tm = 100000;
          return out;
        };
        const here = signed(cx);
        blocks.push({
          at: it.start + e.a, exitAt: it.start + e.b, spid,
          enter: (id, t0) =>
            setVisibility(id, spid, 'visible', t0) +
            keyed(id, spid, 'ppt_x', [{ tm: 0, val: here }, { tm: 100000, val: here }], len, t0) +
            keyed(id, spid, 'ppt_y', frames((s) => signed(bottom - (s * size) / 2)), len, t0) +
            keyed(id, spid, 'ppt_w', frames((s) => (s > 0.99999 ? '#ppt_w' : `#ppt_w*${num(s)}`)), len, t0) +
            keyed(id, spid, 'ppt_h', frames((s) => (s > 0.99999 ? '#ppt_h' : `#ppt_h*${num(s)}`)), len, t0) +
            fade(id, spid, 'in', e.fadeIn, t0),
          leave: (id, t0) =>
            fade(id, spid, 'out', e.fadeOut, t0) +
            setVisibility(id, spid, 'hidden', t0 + e.fadeOut - 1),
          leaveOffset: e.b - e.a,
        });
      }
      const spid = spids[it.pic];
      blocks.push({
        at: it.start + it.grow, spid, letter: true,
        enter: (id, t0) => setVisibility(id, spid, 'visible', t0) + fade(id, spid, 'in', it.dur - it.grow, t0),
      });
    }
    blocks.sort((p, q) => p.at - q.at);

    if (grouping === 'single') {
      return effect(nextId, ZOOM_IN, 'clickEffect', 0,
        (id) => blocks.map((b) => b.enter(id, b.at) + (b.leave ? b.leave(id, b.at + b.leaveOffset) : '')).join(''));
    }
    const effects = [];
    for (const b of blocks) {
      if (b.letter) effects.push({ at: b.at, attrs: FADE_IN, body: (id) => b.enter(id, 0) });
      else if (grouping === 'split') {
        effects.push({ at: b.at, attrs: ZOOM_IN, body: (id) => b.enter(id, 0) });
        effects.push({ at: b.exitAt, attrs: FADE_OUT, body: (id) => b.leave(id, 0) });
      } else effects.push({ at: b.at, attrs: ZOOM_IN, body: (id) => b.enter(id, 0) + b.leave(id, b.leaveOffset) });
    }
    effects.sort((p, q) => p.at - q.at);
    return effects.map((e, k) => effect(nextId, e.attrs, k === 0 ? 'clickEffect' : 'withEffect', e.at, e.body)).join('');
  };
}

/** Enveloppe un fragment pour pouvoir l'analyser avec les bons espaces de noms. */
export const wrap = (xml) => `<root xmlns:p="${NS.p}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">${xml}</root>`;
