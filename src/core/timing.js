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

/** Enveloppe un fragment pour pouvoir l'analyser avec les bons espaces de noms. */
export const wrap = (xml) => `<root xmlns:p="${NS.p}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">${xml}</root>`;
