// PPT Anim : application web installable (pptx + image de la zone de texte, pptx animé téléchargé)
// et, via taskpane.html, complément PowerPoint (zone de texte sélectionnée, diapo animée réinsérée).

import { segment, unitsOf, cropUnit, matchScore } from './core/segment.js';
import { scramblePlan, orderKeyframes, curvePoint, ease } from './core/effects.js';
import { emojiPlan, parseSeries } from './core/emoji.js';
import { scrambleTimeline, emojiTimeline } from './core/timing.js';
import { createEngine } from './core/pptx.js';
import { initPwa } from './pwa.js';

const $ = (id) => document.getElementById(id);
const engine = createEngine({ DOMParser, XMLSerializer });
const EMU = 12700; // par point
// state.file est toujours le même objet : les deux fichiers peuvent arriver en même temps, dans n'importe quel ordre.
const state = { addin: false, img: null, seg: null, text: '', src: null, file: {}, seed: 1, color: '#ffffff' };

function log(msg) {
  const t = new Date().toLocaleTimeString('fr-FR');
  $('log').textContent += `${t}  ${msg}\n`;
  $('log').scrollTop = $('log').scrollHeight;
}
function status(msg, kind = 'ok') {
  $('status').hidden = !msg;
  $('status').textContent = msg || '';
  $('status').dataset.kind = kind;
}
async function guard(label, fn, btn) {
  if (btn) btn.disabled = true;
  try {
    status('');
    await fn();
  } catch (e) {
    const detail = e && e.debugInfo ? ` [${e.code} @ ${e.debugInfo.errorLocation || '?'}]` : '';
    status(`${label} : ${e.message || e}${detail}`, 'err');
    log(`ERREUR ${label} : ${e.message || e}${detail}`);
  } finally {
    if (btn) btn.disabled = false;
  }
}

// ---------- Images ----------
async function toImageData(blob) {
  const bmp = await createImageBitmap(blob);
  const c = document.createElement('canvas');
  c.width = bmp.width; c.height = bmp.height;
  const g = c.getContext('2d', { willReadFrequently: true });
  g.drawImage(bmp, 0, 0);
  return g.getImageData(0, 0, c.width, c.height);
}
function cropCanvas(crop) {
  const c = document.createElement('canvas');
  c.width = crop.w; c.height = crop.h;
  c.getContext('2d').putImageData(new ImageData(crop.data, crop.w, crop.h), 0, 0);
  return c;
}
async function toPng(crop) {
  const blob = await new Promise((r) => cropCanvas(crop).toBlob(r, 'image/png'));
  return new Uint8Array(await blob.arrayBuffer());
}
const b64ToBlob = (b64, type) => new Blob([Uint8Array.from(atob(b64), (ch) => ch.charCodeAt(0))], { type });

// ---------- Options ----------
const seconds = (id, fallback) => { const v = parseFloat($(id).value); return (Number.isFinite(v) ? v : fallback) * 1000; };
const opts = () => ({
  mode: $('o-mode').value,
  unit: $('o-unit').value,
  order: $('o-order').value,
  curved: $('o-curved').value === '1',
  duration: Math.max(100, +$('o-dur').value || 600),
  stagger: Math.max(0, +$('o-stagger').value || 0),
  // emoji qui pousse
  series: parseSeries([$('e-s1').value, $('e-s2').value, $('e-s3').value]),
  mean: Math.max(100, seconds('e-mean', 1.2)),
  sd: Math.max(0, seconds('e-sd', 0)),
  total: Math.max(100, seconds('e-total', 6)),
  orderLine: $('e-line').value, orderWord: $('e-word').value, orderLetter: $('e-letter').value,
  grouping: $('e-group').value,
  seed: state.seed,
});
// Travail à faire selon l'effet choisi : unités animées et plan d'animation.
function current() {
  const o = opts();
  if (o.mode === 'emoji') return { kind: 'emoji', o, units: state.seg.letters, plan: emojiPlan(state.seg, o) };
  const units = unitsOf(state.seg, o.unit);
  return { kind: 'scramble', o, units, plan: scramblePlan(units, state.seg, o) };
}
const timelineOf = (job) => (job.kind === 'emoji'
  ? emojiTimeline(job.plan, job.o.grouping)
  : scrambleTimeline(job.plan, orderKeyframes, job.o.curved));

// ---------- Emoji ----------
// Chaque emoji est dessiné une fois par l'appareil, recadré au plus juste et posé en bas d'un carré.
const glyphs = new Map();
function glyph(ch) {
  const key = `${state.color}|${ch}`;
  if (glyphs.has(key)) return glyphs.get(key);
  const R = 256;
  const big = document.createElement('canvas');
  big.width = big.height = 2 * R;
  const g = big.getContext('2d', { willReadFrequently: true });
  g.font = `${R}px "Segoe UI Emoji", "Apple Color Emoji", "Noto Color Emoji", "Segoe UI Symbol", sans-serif`;
  g.textAlign = 'center';
  g.textBaseline = 'middle';
  g.fillStyle = state.color; // pour les symboles sans couleur propre (★, ✿…) : la couleur du texte
  g.fillText(ch, R, R);
  const d = g.getImageData(0, 0, 2 * R, 2 * R).data;
  let x0 = 2 * R, y0 = 2 * R, x1 = -1, y1 = -1;
  for (let y = 0; y < 2 * R; y++) for (let x = 0; x < 2 * R; x++) {
    if (d[(y * 2 * R + x) * 4 + 3] > 8) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
  }
  if (x1 < 0) throw new Error(`« ${ch} » ne peut pas être dessiné sur cet appareil.`);
  const w = x1 - x0 + 1, h = y1 - y0 + 1, k = R / Math.max(w, h);
  const canvas = document.createElement('canvas');
  canvas.width = canvas.height = R;
  const c = canvas.getContext('2d');
  c.imageSmoothingQuality = 'high';
  c.drawImage(big, x0, y0, w, h, (R - w * k) / 2, R - h * k, w * k, h * k);
  const out = { canvas, url: canvas.toDataURL('image/png'), png: null };
  glyphs.set(key, out);
  return out;
}
// Couleur dominante du texte, pour les symboles monochromes.
function textColor(img) {
  let r = 0, g = 0, b = 0, n = 0;
  for (let p = 0; p < img.data.length; p += 16) {
    if (img.data[p + 3] > 200) { r += img.data[p]; g += img.data[p + 1]; b += img.data[p + 2]; n++; }
  }
  return n ? `rgb(${Math.round(r / n)}, ${Math.round(g / n)}, ${Math.round(b / n)})` : '#ffffff';
}

// ---------- Découpage ----------
function analyze(img, text) {
  const t0 = performance.now();
  const seg = segment(img, text);
  state.img = img; state.seg = seg; state.text = text;
  state.color = textColor(img);
  const ms = Math.round(performance.now() - t0);
  const counts = `${seg.letters.length} ${seg.aligned ? 'lettres' : 'blocs'}, ${seg.words.length} mots, ${seg.lines.length} ligne${seg.lines.length > 1 ? 's' : ''}`;
  if (seg.matched) {
    $('seg-info').textContent = `${counts}. Le découpage correspond au texte.`;
  } else if (seg.aligned) {
    $('seg-info').textContent = `${counts}. ${seg.fused} mot${seg.fused > 1 ? 's restent entiers' : ' reste entier'} : ses lettres se touchent ou sont liées.`;
  } else {
    $('seg-info').textContent = `${counts}. Le découpage ne correspond pas au texte : l'unité « mot » est choisie, plus fiable ici.`;
    $('o-unit').value = 'word';
  }
  log(`Découpage ${img.width}×${img.height} px en ${ms} ms : ${counts}, texte ${seg.matched ? 'reconnu' : seg.aligned ? `reconnu sauf ${seg.fused} mot(s)` : 'non reconnu'} (${seg.found} blocs, ${seg.expected} attendus)`);
  for (const id of ['seg', 'fx', 'out']) $(id).hidden = false;
  refresh();
}

function drawSegmentation() {
  const { img, seg } = state;
  const cv = $('seg-canvas');
  const wCss = cv.parentElement.clientWidth || 600;
  const k = wCss / img.width;
  const dpr = window.devicePixelRatio || 1;
  cv.width = Math.round(wCss * dpr); cv.height = Math.round(img.height * k * dpr);
  cv.style.width = `${wCss}px`; cv.style.height = `${img.height * k}px`;
  const g = cv.getContext('2d');
  const src = document.createElement('canvas');
  src.width = img.width; src.height = img.height;
  src.getContext('2d').putImageData(img, 0, 0);
  g.drawImage(src, 0, 0, cv.width, cv.height);
  g.lineWidth = Math.max(1, dpr);
  const s = k * dpr;
  unitsOf(seg, $('o-mode').value === 'emoji' ? 'letter' : $('o-unit').value).forEach((u, i) => {
    g.strokeStyle = i % 2 ? '#ffd166' : '#06d6a0';
    g.strokeRect(u.x0 * s, u.y0 * s, (u.x1 - u.x0 + 1) * s, (u.y1 - u.y0 + 1) * s);
  });
}

// ---------- Aperçu ----------
function buildPreview(play) {
  const { img, seg } = state;
  const job = current();
  const { units, plan } = job;
  const stage = $('preview');
  stage.textContent = '';
  const pad = 0.8 * seg.maxLineH;
  const k = (stage.clientWidth || 600) / (img.width + 2 * pad);
  stage.style.height = `${(img.height + 2 * pad) * k}px`;
  const tr = (p) => `translate(${p.x * k}px, ${p.y * k}px)`;
  const place = (el, x, y, w, h) => {
    el.style.left = `${(x + pad) * k}px`; el.style.top = `${(y + pad) * k}px`;
    el.style.width = `${w * k}px`; el.style.height = `${h * k}px`;
    stage.appendChild(el);
  };
  units.forEach((u, i) => {
    const crop = cropUnit(img, seg, u);
    const c = cropCanvas(crop);
    place(c, crop.x, crop.y, crop.w, crop.h);
    if (!play) return;
    const it = plan.items[i];
    if (job.kind === 'emoji') {
      c.animate([{ opacity: 0 }, { opacity: 1 }], { duration: it.dur - it.grow, delay: it.start + it.grow, fill: 'both', easing: 'linear' });
      for (const e of it.emojis) {
        const el = document.createElement('img');
        el.src = glyph(e.ch).url;
        el.alt = '';
        place(el, it.box.cx - it.box.size / 2, it.box.bottom - it.box.size, it.box.size, it.box.size);
        const len = e.end - e.a, timing = { duration: len, delay: it.start + e.a, fill: 'both', easing: 'linear' };
        el.animate(e.kf.map((f) => ({ offset: Math.min(1, (f.t - e.a) / len), transform: `scale(${f.s})` })), timing);
        el.animate([
          { offset: 0, opacity: 0 }, { offset: e.fadeIn / len, opacity: 1 },
          { offset: (e.b - e.a) / len, opacity: 1 }, { offset: 1, opacity: 0 },
        ], timing);
      }
    } else if (plan.mode === 'order') {
      c.animate(orderKeyframes(it).map((p) => ({ offset: p.t, transform: tr(p) })),
        { duration: it.delay + it.dur, fill: 'both', easing: 'linear' });
    } else {
      const kf = Array.from({ length: 11 }, (_, j) => ({ offset: j / 10, transform: tr(curvePoint(it, ease(j / 10))) }));
      c.animate(kf, { duration: it.dur, delay: it.delay, fill: 'both', easing: 'linear' });
    }
  });
  if (job.kind === 'emoji') {
    const lines = { single: 1, image: plan.objects, split: 2 * plan.objects - units.length }[job.o.grouping];
    $('fx-info').textContent = `${units.length} lettres, ${(plan.total / 1000).toFixed(2)} s au total, environ ${plan.parallel.toFixed(1)} animations de front, ` +
      `${lines} ligne${lines > 1 ? 's' : ''} dans le volet Animations`;
  } else {
    $('fx-info').textContent = `${units.length} animations, ${(plan.total / 1000).toFixed(1)} s au total`;
  }
  return job;
}
function refresh() {
  const emoji = $('o-mode').value === 'emoji';
  $('fx-scramble').hidden = emoji;
  $('fx-emoji').hidden = !emoji;
  $('reroll').hidden = !emoji;
  if (!state.seg) return;
  try {
    drawSegmentation();
    buildPreview(false);
  } catch (e) {
    $('fx-info').textContent = `Réglage incomplet : ${e.message}`;
  }
}

// Images à placer dans la diapo : les lettres (ou mots, lignes), puis les emoji, rangés hors de la diapo.
async function buildPictures(job) {
  const pics = [];
  for (const u of job.units) {
    const c = cropUnit(state.img, state.seg, u);
    pics.push({ x: c.x, y: c.y, w: c.w, h: c.h, png: await toPng(c), label: (u.text || '').slice(0, 20) });
  }
  if (job.kind === 'emoji') {
    for (const it of job.plan.items) {
      it.pic = it.index;
      for (const e of it.emojis) {
        const g = glyph(e.ch);
        if (!g.png) g.png = new Uint8Array(await (await new Promise((r) => g.canvas.toBlob(r, 'image/png'))).arrayBuffer());
        e.pic = pics.length;
        pics.push({
          x: it.box.cx - it.box.size / 2, y: it.box.bottom - it.box.size, w: it.box.size, h: it.box.size,
          png: g.png, mediaKey: `emoji:${e.ch}`, park: true, label: e.ch,
        });
      }
    }
  }
  return pics;
}

// ---------- Mode complément ----------
async function analyzeSelection() {
  await PowerPoint.run(async (ctx) => {
    const sel = ctx.presentation.getSelectedShapes();
    const n = sel.getCount();
    await ctx.sync();
    if (n.value !== 1) throw new Error('sélectionnez une seule zone de texte.');
    const shape = sel.getItemAt(0);
    shape.load('id,name,left,top,width,height,rotation');
    const slide = shape.getParentSlide();
    slide.load('id');
    const tf = shape.getTextFrameOrNullObject();
    tf.load('hasText');
    await ctx.sync();
    if (tf.isNullObject || !tf.hasText) throw new Error('l\'objet sélectionné ne contient pas de texte.');
    if (shape.rotation) log(`Attention : zone de texte pivotée de ${shape.rotation}°, ce qui n'est pas géré.`);
    tf.textRange.load('text');
    const width = Math.min(4096, Math.max(64, Math.round(shape.width * 4)));
    const png = shape.getImageAsBase64({ width });
    await ctx.sync();
    state.src = {
      slideId: slide.id,
      ref: {
        id: shape.id, name: shape.name,
        box: { x: Math.round(shape.left * EMU), y: Math.round(shape.top * EMU), cx: Math.round(shape.width * EMU), cy: Math.round(shape.height * EMU) },
      },
    };
    log(`Zone de texte « ${shape.name} » (id ${shape.id}) sur la diapo ${slide.id} : ${shape.width.toFixed(1)}×${shape.height.toFixed(1)} pt`);
    const img = await toImageData(b64ToBlob(png.value, 'image/png'));
    const ratio = (img.width / img.height) / (shape.width / shape.height);
    if (Math.abs(ratio - 1) > 0.03) log(`Attention : l'image (${img.width}×${img.height}) n'a pas les proportions de la zone de texte ; le placement des lettres peut être décalé.`);
    analyze(img, tf.textRange.text);
  });
}

async function generateAddin() {
  const job = current();
  const pictures = await buildPictures(job);
  await PowerPoint.run(async (ctx) => {
    const src = state.src;
    const slide = ctx.presentation.slides.getItem(src.slideId);
    const exp = slide.exportAsBase64();
    await ctx.sync();
    log(`Diapo exportée : ${Math.round(exp.value.length * 0.75 / 1024)} Ko`);
    const zip = await JSZip.loadAsync(exp.value, { base64: true });
    const deck = await engine.readDeck(zip);
    const sl = await engine.readSlide(zip, deck.slides[0].path);
    const shape = engine.findShape(sl, src.ref);
    if (!shape) throw new Error(`zone de texte « ${src.ref.name} » introuvable dans la diapo exportée (zones présentes : ${sl.shapes.map((s) => `${s.id}:${s.name}`).join(', ')}).`);
    const res = await engine.animate(zip, { deck, slide: sl, shape, box: shape.box || src.ref.box, img: state.img, pictures, timeline: timelineOf(job) });
    const out = await zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });
    ctx.presentation.insertSlidesFromBase64(out, { formatting: $('o-format').value, targetSlideId: src.slideId });
    await ctx.sync();
    log(`Diapo insérée : ${res.pictures} images, ${res.existingEffects} animations existantes conservées, chronologie ${res.timingCreated ? 'créée' : 'complétée'}.`);
    status('Diapo animée insérée après l\'originale. Lancez le diaporama pour vérifier.', 'ok');
  });
}

// ---------- Mode fichier ----------
const snippet = (t, n = 40) => t.replace(/\s+/g, ' ').trim().slice(0, n);

async function loadPptx(file) {
  const bytes = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(bytes);
  const deck = await engine.readDeck(zip);
  // Toutes les diapos sont lues : la zone de texte à animer peut être sur n'importe laquelle.
  const slides = [];
  for (const s of deck.slides) slides.push(await engine.readSlide(zip, s.path));
  Object.assign(state.file, { name: file.name, bytes, deck, zip, slides });
  const sel = $('s-slide');
  sel.textContent = '';
  slides.forEach((sl, i) => {
    const op = document.createElement('option');
    op.value = i;
    op.textContent = `Diapo ${i + 1} : ${sl.shapes.length ? snippet(sl.shapes[0].text, 30) : 'aucune zone de texte'}`;
    sel.appendChild(op);
  });
  sel.disabled = false;
  const zones = slides.reduce((n, sl) => n + sl.shapes.length, 0);
  log(`Présentation chargée : ${slides.length} diapo(s), ${zones} zone(s) de texte, ${deck.size.cx}×${deck.size.cy} EMU`);
  showSlide(0);
  locate();
}

// Affiche les zones de texte d'une diapo et en présélectionne une.
function showSlide(index, shapeIndex = 0) {
  const f = state.file;
  $('s-slide').value = String(index);
  f.slide = f.slides[index];
  const sel = $('s-shape');
  sel.textContent = '';
  f.slide.shapes.forEach((s, i) => {
    const op = document.createElement('option');
    op.value = i;
    op.textContent = `${s.name} : ${snippet(s.text)}`;
    sel.appendChild(op);
  });
  sel.disabled = !f.slide.shapes.length;
  if (f.slide.shapes.length) sel.value = String(Math.min(shapeIndex, f.slide.shapes.length - 1));
  readyFile();
}

function tell(msg, kind) {
  $('locate').hidden = !msg;
  $('locate').textContent = msg || '';
  $('locate').dataset.kind = kind || '';
}

// Classe les zones de texte de la présentation selon leur ressemblance avec l'image.
function candidates() {
  const f = state.file;
  const ratio = f.img.width / f.img.height;
  const list = [];
  f.slides.forEach((sl, slide) => sl.shapes.forEach((sh, shape) => {
    const score = matchScore(f.seen, sh.text);
    // À ressemblance égale, on préfère la zone dont les proportions sont celles de l'image.
    const fit = sh.box ? Math.abs(Math.log(ratio / (sh.box.cx / sh.box.cy))) : 1;
    list.push({ slide, shape, score, fit, name: sh.name });
  }));
  return list.sort((p, q) => q.score - p.score || p.fit - q.fit);
}

// Retrouve la diapo et la zone de texte à partir de l'image, dès que les deux fichiers sont là.
function locate() {
  const f = state.file;
  if (!f.slides || !f.img) return;
  if (!f.seen) {
    try {
      f.seen = segment(f.img, '').seenCounts;
    } catch (e) {
      tell(`Image inexploitable : ${e.message}`, 'warn');
      return;
    }
  }
  const list = candidates();
  if (!list.length) { tell('Cette présentation ne contient aucune zone de texte utilisable.', 'warn'); return; }
  const best = list[0];
  if (best.score < 0.75) {
    tell('Aucune zone de texte ne correspond nettement à l\'image : choisissez la diapo et la zone de texte.', 'warn');
    log(`Recherche de la zone : aucune correspondance nette (meilleure : diapo ${best.slide + 1}, « ${best.name} », ${Math.round(best.score * 100)} %).`);
    return;
  }
  showSlide(best.slide, best.shape);
  const same = list.filter((c) => c.score === best.score);
  if (same.length > 1) {
    const where = [...new Set(same.map((c) => c.slide + 1))].join(', ');
    tell(`Plusieurs zones de texte correspondent à l'image (diapo${where.includes(',') ? 's' : ''} ${where}). Choix proposé : diapo ${best.slide + 1}, « ${best.name} ». Vérifiez-le.`, 'warn');
  } else {
    tell(`Zone de texte retrouvée d'après l'image : diapo ${best.slide + 1}, « ${best.name} ».`, 'ok');
  }
  log(`Recherche de la zone : diapo ${best.slide + 1}, « ${best.name} » (${Math.round(best.score * 100)} %, ${same.length} candidate(s) à égalité).`);
}

// Choix manuel d'une diapo : on y présélectionne la zone qui ressemble le plus à l'image.
function pickSlide() {
  const f = state.file;
  const index = +$('s-slide').value;
  const here = f.seen ? candidates().filter((c) => c.slide === index) : [];
  showSlide(index, here.length ? here[0].shape : 0);
  tell(f.slide.shapes.length ? '' : 'Aucune zone de texte utilisable sur cette diapo (les zones placées dans un groupe ne sont pas prises en compte).', 'warn');
}

async function loadPng(file) {
  const f = state.file;
  f.png = file;
  f.img = await toImageData(file);
  f.seen = null;
  log(`Image de la zone de texte : ${file.name || 'collée'} (${f.img.width}×${f.img.height} px, ${Math.round(file.size / 1024)} Ko)`);
  readyFile();
  locate();
}

function readyFile() {
  const f = state.file;
  $('analyze-file').disabled = !(f.slide && f.slide.shapes.length && f.img);
}
async function analyzeFile() {
  const f = state.file;
  const shape = f.slide.shapes[+$('s-shape').value];
  if (shape.rotated) log('Attention : zone de texte pivotée ou retournée, ce qui n\'est pas géré.');
  const img = f.img;
  if (shape.box) {
    const ratio = (img.width / img.height) / (shape.box.cx / shape.box.cy);
    if (Math.abs(ratio - 1) > 0.03) log(`Attention : l'image (${img.width}×${img.height}) n'a pas les proportions de la zone de texte ; le placement des lettres peut être décalé.`);
  }
  f.shapeIndex = +$('s-shape').value;
  f.slideIndex = +$('s-slide').value;
  log(`Analyse : diapo ${f.slideIndex + 1}, zone de texte « ${shape.name} ».`);
  analyze(img, shape.text);
}
async function generateFile() {
  const f = state.file;
  const job = current();
  const pictures = await buildPictures(job);
  const zip = await JSZip.loadAsync(f.bytes); // on repart toujours du fichier d'origine
  const deck = await engine.readDeck(zip);
  const sl = await engine.readSlide(zip, deck.slides[f.slideIndex].path);
  const shape = sl.shapes[f.shapeIndex];
  const res = await engine.animate(zip, { deck, slide: sl, shape, img: state.img, pictures, timeline: timelineOf(job) });
  const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = f.name.replace(/\.pptx$/i, '') + '-anim.pptx';
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 10000);
  log(`Fichier généré (diapo ${f.slideIndex + 1}) : ${res.pictures} images, ${res.existingEffects} animations existantes conservées, chronologie ${res.timingCreated ? 'créée' : 'complétée'}.`);
  status(`${a.download} téléchargé. Ouvrez-le dans PowerPoint et lancez le diaporama.`, 'ok');
}

// ---------- Démarrage ----------
// Fichiers reçus par collage ou par dépôt : on les range dans les champs, comme un choix manuel.
function takeFiles(files) {
  let n = 0;
  for (const f of files) {
    const id = /\.pptx$/i.test(f.name) ? 'f-pptx' : f.type === 'image/png' ? 'f-png' : null;
    if (!id) continue;
    const dt = new DataTransfer();
    dt.items.add(f);
    $(id).files = dt.files;
    $(id).dispatchEvent(new Event('change'));
    n++;
  }
  return n;
}

async function start() {
  let info = null;
  // office.js n'est chargé que par taskpane.html (page du complément) ; l'application web ne l'utilise pas.
  if (window.Office && Office.onReady) {
    info = await Promise.race([Office.onReady(), new Promise((r) => setTimeout(() => r(null), 4000))]);
  }
  state.addin = !!(info && info.host && String(info.host) === 'PowerPoint');
  if (state.addin) {
    const ok = Office.context.requirements.isSetSupported('PowerPointApi', '1.10');
    $('mode').textContent = 'complément PowerPoint';
    $('src-addin').hidden = false;
    $('out-addin').hidden = false;
    log(`Complément PowerPoint, plateforme ${info.platform}, PowerPointApi 1.10 : ${ok ? 'oui' : 'non'}`);
    if (!ok) {
      $('analyze-addin').disabled = true;
      status('Cette version de PowerPoint ne fournit pas PowerPointApi 1.10 (image d\'une zone de texte). Utilisez l\'application web dans un navigateur.', 'warn');
    }
  } else {
    $('mode').textContent = '';
    $('src-file').hidden = false;
    $('out-file').hidden = false;
    log('Application prête.');
    document.addEventListener('paste', (e) => {
      const files = e.clipboardData ? Array.from(e.clipboardData.files) : [];
      if (takeFiles(files)) e.preventDefault();
      else if (files.length) status('Le presse-papiers ne contient pas d\'image PNG ni de fichier .pptx.', 'warn');
    });
    document.addEventListener('dragover', (e) => { e.preventDefault(); document.body.classList.add('drag'); });
    document.addEventListener('dragleave', () => document.body.classList.remove('drag'));
    document.addEventListener('drop', (e) => {
      e.preventDefault();
      document.body.classList.remove('drag');
      takeFiles(Array.from(e.dataTransfer.files));
    });
    initPwa({ log });
  }

  // Aide : fenêtre modale, contenu selon le mode
  $('help-app').hidden = state.addin;
  $('help-addin').hidden = !state.addin;
  $('help-open').onclick = () => $('help').showModal();
  $('help-close').onclick = () => $('help').close();
  $('help').addEventListener('click', (e) => { if (e.target === $('help')) $('help').close(); }); // clic hors de la fenêtre

  $('analyze-addin').onclick = (e) => guard('Analyse', analyzeSelection, e.currentTarget);
  $('go-addin').onclick = (e) => guard('Génération', generateAddin, e.currentTarget);
  $('f-pptx').onchange = (e) => { const f = e.target.files[0]; if (f) guard('Lecture du pptx', () => loadPptx(f)); };
  $('f-png').onchange = (e) => { const f = e.target.files[0]; if (f) guard('Lecture de l\'image', () => loadPng(f)); };
  $('s-slide').onchange = () => guard('Lecture de la diapo', async () => pickSlide());
  $('s-shape').onchange = () => tell('');
  $('analyze-file').onclick = (e) => guard('Analyse', analyzeFile, e.currentTarget);
  $('go-file').onclick = (e) => guard('Génération', generateFile, e.currentTarget);
  $('play').onclick = () => guard('Aperçu', async () => buildPreview(true));
  for (const id of ['o-mode', 'o-unit', 'o-order', 'o-curved', 'o-dur', 'o-stagger',
    'e-s1', 'e-s2', 'e-s3', 'e-mean', 'e-sd', 'e-total', 'e-line', 'e-word', 'e-letter', 'e-group']) $(id).onchange = refresh;
  $('reroll').onclick = () => { state.seed = (state.seed * 7919 + 13) % 1000003; guard('Aperçu', async () => buildPreview(true)); };
  refresh();
  let lastW = window.innerWidth;
  window.addEventListener('resize', () => { if (window.innerWidth !== lastW) { lastW = window.innerWidth; refresh(); } });
}
start();
