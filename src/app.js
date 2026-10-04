// PPT Anim : application web installable (pptx + image de la forme, pptx animé téléchargé)
// et, via taskpane.html, complément PowerPoint (forme sélectionnée, diapo animée réinsérée).

import { segment, unitsOf, cropUnit } from './core/segment.js';
import { scramblePlan, orderKeyframes, curvePoint, ease } from './core/effects.js';
import { createEngine } from './core/pptx.js';
import { initPwa } from './pwa.js';

const $ = (id) => document.getElementById(id);
const engine = createEngine({ DOMParser, XMLSerializer });
const EMU = 12700; // par point
const state = { addin: false, img: null, seg: null, text: '', src: null, file: null };

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
const opts = () => ({
  mode: $('o-mode').value,
  unit: $('o-unit').value,
  order: $('o-order').value,
  curved: $('o-curved').value === '1',
  duration: Math.max(100, +$('o-dur').value || 600),
  stagger: Math.max(0, +$('o-stagger').value || 0),
});
function current() {
  const o = opts();
  const units = unitsOf(state.seg, o.unit);
  const plan = scramblePlan(units, state.seg, o);
  return { o, units, plan };
}

// ---------- Découpage ----------
function analyze(img, text) {
  const t0 = performance.now();
  const seg = segment(img, text);
  state.img = img; state.seg = seg; state.text = text;
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
  unitsOf(seg, $('o-unit').value).forEach((u, i) => {
    g.strokeStyle = i % 2 ? '#ffd166' : '#06d6a0';
    g.strokeRect(u.x0 * s, u.y0 * s, (u.x1 - u.x0 + 1) * s, (u.y1 - u.y0 + 1) * s);
  });
}

// ---------- Aperçu ----------
function buildPreview(play) {
  const { img, seg } = state;
  const { o, units, plan } = current();
  const stage = $('preview');
  stage.textContent = '';
  const pad = 0.8 * seg.maxLineH;
  const k = (stage.clientWidth || 600) / (img.width + 2 * pad);
  stage.style.height = `${(img.height + 2 * pad) * k}px`;
  const tr = (p) => `translate(${p.x * k}px, ${p.y * k}px)`;
  units.forEach((u, i) => {
    const crop = cropUnit(img, seg, u);
    const c = cropCanvas(crop);
    c.style.left = `${(crop.x + pad) * k}px`; c.style.top = `${(crop.y + pad) * k}px`;
    c.style.width = `${crop.w * k}px`; c.style.height = `${crop.h * k}px`;
    stage.appendChild(c);
    if (!play) return;
    const it = plan.items[i];
    if (plan.mode === 'order') {
      c.animate(orderKeyframes(it).map((p) => ({ offset: p.t, transform: tr(p) })),
        { duration: it.delay + it.dur, fill: 'both', easing: 'linear' });
    } else {
      const kf = Array.from({ length: 11 }, (_, j) => ({ offset: j / 10, transform: tr(curvePoint(it, ease(j / 10))) }));
      c.animate(kf, { duration: it.dur, delay: it.delay, fill: 'both', easing: 'linear' });
    }
  });
  $('fx-info').textContent = `${units.length} animations, ${(plan.total / 1000).toFixed(1)} s au total`;
  return { o, units, plan };
}
function refresh() {
  if (!state.seg) return;
  drawSegmentation();
  buildPreview(false);
}

async function buildCrops(units) {
  const crops = [];
  for (const u of units) {
    const c = cropUnit(state.img, state.seg, u);
    crops.push({ x: c.x, y: c.y, w: c.w, h: c.h, png: await toPng(c), label: (u.text || '').slice(0, 20) });
  }
  return crops;
}

// ---------- Mode complément ----------
async function analyzeSelection() {
  await PowerPoint.run(async (ctx) => {
    const sel = ctx.presentation.getSelectedShapes();
    const n = sel.getCount();
    await ctx.sync();
    if (n.value !== 1) throw new Error('sélectionnez une seule forme de texte.');
    const shape = sel.getItemAt(0);
    shape.load('id,name,left,top,width,height,rotation');
    const slide = shape.getParentSlide();
    slide.load('id');
    const tf = shape.getTextFrameOrNullObject();
    tf.load('hasText');
    await ctx.sync();
    if (tf.isNullObject || !tf.hasText) throw new Error('la forme sélectionnée ne contient pas de texte.');
    if (shape.rotation) log(`Attention : forme pivotée de ${shape.rotation}°, non géré par ce PoC.`);
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
    log(`Forme « ${shape.name} » (id ${shape.id}) sur la diapo ${slide.id} : ${shape.width.toFixed(1)}×${shape.height.toFixed(1)} pt`);
    const img = await toImageData(b64ToBlob(png.value, 'image/png'));
    const ratio = (img.width / img.height) / (shape.width / shape.height);
    if (Math.abs(ratio - 1) > 0.03) log(`Attention : l'image (${img.width}×${img.height}) n'a pas les proportions de la forme ; le placement des lettres peut être décalé.`);
    analyze(img, tf.textRange.text);
  });
}

async function generateAddin() {
  const { o, units, plan } = current();
  const crops = await buildCrops(units);
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
    if (!shape) throw new Error(`forme « ${src.ref.name} » introuvable dans la diapo exportée (formes : ${sl.shapes.map((s) => `${s.id}:${s.name}`).join(', ')}).`);
    const res = await engine.animate(zip, { deck, slide: sl, shape, box: shape.box || src.ref.box, img: state.img, crops, plan, curved: o.curved });
    const out = await zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });
    ctx.presentation.insertSlidesFromBase64(out, { formatting: $('o-format').value, targetSlideId: src.slideId });
    await ctx.sync();
    log(`Diapo insérée : ${res.pictures} images, ${res.existingEffects} animations existantes conservées, chronologie ${res.timingCreated ? 'créée' : 'complétée'}.`);
    status('Diapo animée insérée après l\'originale. Lancez le diaporama pour vérifier.', 'ok');
  });
}

// ---------- Mode fichier ----------
async function loadPptx(file) {
  const bytes = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(bytes);
  const deck = await engine.readDeck(zip);
  state.file = { ...(state.file || {}), name: file.name, bytes, deck, zip };
  $('s-slide').innerHTML = deck.slides.map((s, i) => `<option value="${i}">Diapo ${i + 1}</option>`).join('');
  $('s-slide').disabled = false;
  log(`Présentation chargée : ${deck.slides.length} diapo(s), ${deck.size.cx}×${deck.size.cy} EMU`);
  await loadSlide();
}
async function loadSlide() {
  const f = state.file;
  f.slide = await engine.readSlide(f.zip, f.deck.slides[+$('s-slide').value].path);
  const sel = $('s-shape');
  sel.textContent = '';
  f.slide.shapes.forEach((s, i) => {
    const op = document.createElement('option');
    op.value = i;
    op.textContent = `${s.name} : ${s.text.replace(/\s+/g, ' ').slice(0, 40)}`;
    sel.appendChild(op);
  });
  sel.disabled = !f.slide.shapes.length;
  if (!f.slide.shapes.length) status('Aucune forme de texte de premier niveau sur cette diapo.', 'warn');
  readyFile();
}
function readyFile() {
  $('analyze-file').disabled = !(state.file && state.file.slide && state.file.slide.shapes.length && state.file.png);
}
async function analyzeFile() {
  const f = state.file;
  const shape = f.slide.shapes[+$('s-shape').value];
  if (shape.rotated) log('Attention : forme pivotée ou retournée, non géré par ce PoC.');
  const img = await toImageData(f.png);
  if (shape.box) {
    const ratio = (img.width / img.height) / (shape.box.cx / shape.box.cy);
    if (Math.abs(ratio - 1) > 0.03) log(`Attention : l'image (${img.width}×${img.height}) n'a pas les proportions de la forme ; le placement des lettres peut être décalé.`);
  }
  f.shapeIndex = +$('s-shape').value;
  f.slideIndex = +$('s-slide').value;
  analyze(img, shape.text);
}
async function generateFile() {
  const f = state.file;
  const { o, units, plan } = current();
  const crops = await buildCrops(units);
  const zip = await JSZip.loadAsync(f.bytes); // on repart toujours du fichier d'origine
  const deck = await engine.readDeck(zip);
  const sl = await engine.readSlide(zip, deck.slides[f.slideIndex].path);
  const shape = sl.shapes[f.shapeIndex];
  const res = await engine.animate(zip, { deck, slide: sl, shape, img: state.img, crops, plan, curved: o.curved });
  const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = f.name.replace(/\.pptx$/i, '') + '-anim.pptx';
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 10000);
  log(`Fichier généré : ${res.pictures} images, ${res.existingEffects} animations existantes conservées, chronologie ${res.timingCreated ? 'créée' : 'complétée'}.`);
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
      status('Cette version de PowerPoint ne fournit pas PowerPointApi 1.10 (image d\'une forme). Utilisez l\'application web dans un navigateur.', 'warn');
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

  $('analyze-addin').onclick = (e) => guard('Analyse', analyzeSelection, e.currentTarget);
  $('go-addin').onclick = (e) => guard('Génération', generateAddin, e.currentTarget);
  $('f-pptx').onchange = (e) => { const f = e.target.files[0]; if (f) guard('Lecture du pptx', () => loadPptx(f)); };
  $('f-png').onchange = (e) => {
    const f = e.target.files[0];
    if (!f) return;
    state.file = state.file || {};
    state.file.png = f;
    log(`Image de la forme : ${f.name || 'collée'} (${Math.round(f.size / 1024)} Ko)`);
    readyFile();
  };
  $('s-slide').onchange = () => guard('Lecture de la diapo', loadSlide);
  $('analyze-file').onclick = (e) => guard('Analyse', analyzeFile, e.currentTarget);
  $('go-file').onclick = (e) => guard('Génération', generateFile, e.currentTarget);
  $('play').onclick = () => guard('Aperçu', async () => buildPreview(true));
  for (const id of ['o-mode', 'o-unit', 'o-order', 'o-curved', 'o-dur', 'o-stagger']) $(id).onchange = refresh;
  let lastW = window.innerWidth;
  window.addEventListener('resize', () => { if (window.innerWidth !== lastW) { lastW = window.innerWidth; refresh(); } });
}
start();
