// Effet « scrambler » : calcule, pour chaque unité, d'où elle part et où elle arrive.
// Tout est exprimé en pixels de l'image source ; la conversion vers la diapo se fait ailleurs.

function rng(seed) {
  let a = (seed >>> 0) || 1;
  return () => {
    a |= 0; a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function shuffled(n, rand) {
  const p = Array.from({ length: n }, (_, i) => i);
  for (let i = n - 1; i > 0; i--) {
    const j = Math.floor(rand() * (i + 1));
    [p[i], p[j]] = [p[j], p[i]];
  }
  return p;
}

// Permutation sans point fixe (chaque unité change de place) dès que n > 1.
function derangement(n, rand) {
  if (n < 2) return [0];
  for (let tries = 0; tries < 50; tries++) {
    const p = shuffled(n, rand);
    if (p.every((v, i) => v !== i)) return p;
  }
  return Array.from({ length: n }, (_, i) => (i + 1) % n);
}

export const ease = (u) => (u < 0.5 ? 4 * u * u * u : 1 - Math.pow(-2 * u + 2, 3) / 2);

/**
 * @param units  unités { x0, y0, x1, y1, line }
 * @param seg    résultat de segment() (pour les bandes de lignes)
 * @param o      { mode: 'order' | 'shuffle', order: 'seq' | 'random', duration, stagger, hold, curved, seed }
 * @returns { items: [{ index, off:{x,y}, c1:{x,y}, c2:{x,y}, delay, dur }], total, mode }
 *   off = position mélangée relative à la position normale ; c1, c2 = points de contrôle de la courbe
 *   allant de la position normale (0,0) à la position mélangée (off).
 */
export function scramblePlan(units, seg, o = {}) {
  const mode = o.mode || 'order';
  const duration = o.duration ?? 600, stagger = o.stagger ?? 40;
  const hold = mode === 'order' ? (o.hold ?? 400) : 0;
  const rand = rng(o.seed ?? 12345);
  const n = units.length;
  const slot = derangement(n, rand);
  const rank = o.order === 'random' ? shuffled(n, rand) : Array.from({ length: n }, (_, i) => i);
  const cx = (u) => (u.x0 + u.x1) / 2;
  const top = (u) => seg.lineBands[u.line].y0;

  const items = units.map((u, i) => {
    const t = units[slot[i]];
    // On garde la hauteur naturelle de la lettre sur sa ligne : seul le changement de ligne décale en y.
    const off = { x: cx(t) - cx(u), y: top(t) - top(u) };
    const d = Math.hypot(off.x, off.y);
    let c1 = { x: off.x / 3, y: off.y / 3 }, c2 = { x: (2 * off.x) / 3, y: (2 * off.y) / 3 };
    if (o.curved !== false && d > 0) {
      const amp = Math.min(0.35 * d, 1.2 * seg.maxLineH) * (i % 2 ? 1 : -1);
      const nx = (-off.y / d) * amp, ny = (off.x / d) * amp;
      c1 = { x: c1.x + nx, y: c1.y + ny };
      c2 = { x: c2.x + nx, y: c2.y + ny };
    }
    return { index: i, off, c1, c2, delay: hold + rank[i] * stagger, dur: duration };
  });
  return { items, mode, hold, total: hold + Math.max(0, n - 1) * stagger + duration };
}

// Point de la courbe (0,0) → off, u dans [0,1].
export function curvePoint(it, u) {
  const v = 1 - u;
  const b1 = 3 * v * v * u, b2 = 3 * v * u * u, b3 = u * u * u;
  return { x: b1 * it.c1.x + b2 * it.c2.x + b3 * it.off.x, y: b1 * it.c1.y + b2 * it.c2.y + b3 * it.off.y };
}

/**
 * Images clés d'une unité sur sa durée propre (delay + dur), pour le mode « ordonner » :
 * elle attend à sa position mélangée, puis rejoint sa place.
 * @returns [{ t (0..1), x, y }] décalages par rapport à la position normale
 */
export function orderKeyframes(it, steps = 10) {
  const T = it.delay + it.dur;
  const t0 = it.delay / T;
  const kf = [{ t: 0, ...curvePoint(it, 1) }];
  if (t0 > 0.0005) kf.push({ t: t0, ...curvePoint(it, 1) });
  for (let k = 1; k <= steps; k++) {
    const u = k / steps;
    kf.push({ t: t0 + (1 - t0) * u, ...curvePoint(it, 1 - ease(u)) });
  }
  kf[kf.length - 1] = { t: 1, x: 0, y: 0 };
  return kf;
}
