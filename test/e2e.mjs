// Essai de bout en bout dans Chromium sur l'application construite : analyse, aperçu, génération des pptx.
// Usage : node scripts/build.mjs && node test/e2e.mjs
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { serve } from '../scripts/serve.mjs';

const require = createRequire(import.meta.url);
let chromium;
try { ({ chromium } = require('playwright')); }
catch { ({ chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright-core')); } // repli : module installé ailleurs

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const fx = path.join(root, 'test', 'fixtures');
const out = path.join(root, 'test', 'out');
fs.mkdirSync(out, { recursive: true });

const server = await serve(); // dist/ sous /pptanim/, comme sur GitHub Pages
const url = server.url;

const cases = [
  { name: 'a-ordonner-lettre', deck: 'a.pptx', mode: 'order', unit: 'letter', order: 'seq', curved: '1' },
  { name: 'a-melanger-mot', deck: 'a.pptx', mode: 'shuffle', unit: 'word', order: 'random', curved: '1' },
  { name: 'b-ordonner-lettre', deck: 'b.pptx', mode: 'order', unit: 'letter', order: 'random', curved: '1' },
  { name: 'b-melanger-lettre-droit', deck: 'b.pptx', mode: 'shuffle', unit: 'letter', order: 'seq', curved: '0' },
  { name: 'b-ordonner-ligne', deck: 'b.pptx', mode: 'order', unit: 'line', order: 'seq', curved: '0' },
];

const browser = await chromium.launch();
const errors = [];
for (const c of cases) {
  const page = await browser.newPage({ viewport: { width: 760, height: 1100 }, acceptDownloads: true });
  page.on('pageerror', (e) => errors.push(`${c.name}: ${e.message}`));
  page.on('console', (m) => { if (m.type() === 'error') errors.push(`${c.name}: ${m.text()}`); });
  page.on('request', (r) => { if (/appsforoffice/.test(r.url())) errors.push(`${c.name}: l'application a demandé office.js`); });
  await page.goto(url);
  await page.waitForSelector('#src-file:not([hidden])');
  await page.setInputFiles('#f-pptx', path.join(fx, c.deck));
  await page.setInputFiles('#f-png', path.join(fx, 'titre.png'));
  await page.waitForFunction(() => !document.getElementById('analyze-file').disabled);
  await page.selectOption('#s-shape', { index: 0 });
  await page.click('#analyze-file');
  await page.waitForSelector('#fx:not([hidden])');
  await page.selectOption('#o-mode', c.mode);
  await page.selectOption('#o-unit', c.unit);
  await page.selectOption('#o-order', c.order);
  await page.selectOption('#o-curved', c.curved);
  const info = await page.textContent('#seg-info');
  await page.click('#play');
  await page.waitForTimeout(c.mode === 'order' ? 900 : 500);
  await page.screenshot({ path: path.join(out, `${c.name}-volet.png`), fullPage: true });
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#go-file')]);
  await dl.saveAs(path.join(out, `${c.name}.pptx`));
  const status = await page.textContent('#status');
  console.log(`${c.name}: ${info} | ${await page.textContent('#fx-info')} | ${status}`);
  await page.close();
}
// Mode complément avec une fausse API Office (simulation : ne prouve rien sur PowerPoint lui-même,
// mais exerce le chemin base64, la recherche de la forme et l'appel d'insertion).
{
  const png = fs.readFileSync(path.join(fx, 'titre.png')).toString('base64');
  const pptx = fs.readFileSync(path.join(fx, 'b.pptx')).toString('base64');
  const mock = `
    window.Office = { onReady: () => Promise.resolve({ host: 'PowerPoint', platform: 'PC' }),
      context: { requirements: { isSetSupported: () => true } } };
    const noop = () => {};
    const shape = { load: noop, id: '4#1234', name: 'Titre', left: 86.4, top: 165.6, width: 792, height: 244.8, rotation: 0,
      getParentSlide: () => ({ load: noop, id: '256#' }),
      getTextFrameOrNullObject: () => ({ load: noop, isNullObject: false, hasText: true,
        textRange: { load: noop, text: 'Bonjour à tous,\\rvoici le "scrambler" d\\'été !' } }),
      getImageAsBase64: (o) => { window.__imgOpts = o; return { value: '${png}' }; } };
    window.PowerPoint = { run: (fn) => fn({ sync: async () => {}, presentation: {
      getSelectedShapes: () => ({ getCount: () => ({ value: 1 }), getItemAt: () => shape }),
      slides: { getItem: () => ({ exportAsBase64: () => ({ value: '${pptx}' }) }) },
      insertSlidesFromBase64: (b64, opts) => { window.__inserted = { b64, opts }; } } }) };`;
  const page = await browser.newPage({ viewport: { width: 340, height: 900 } });
  await page.route('**/appsforoffice.microsoft.com/**', (r) => r.fulfill({ contentType: 'text/javascript', body: mock }));
  page.on('pageerror', (e) => errors.push(`complement: ${e.message}`));
  await page.goto(url + 'taskpane.html');
  await page.waitForSelector('#src-addin:not([hidden])');
  await page.click('#analyze-addin');
  await page.waitForSelector('#fx:not([hidden])');
  await page.click('#go-addin');
  await page.waitForFunction(() => window.__inserted || document.getElementById('status').dataset.kind === 'err');
  const r = await page.evaluate(() => ({ ins: window.__inserted, img: window.__imgOpts, status: document.getElementById('status').textContent, log: document.getElementById('log').textContent }));
  await page.screenshot({ path: path.join(out, 'b-complement-simule-volet.png'), fullPage: true });
  if (!r.ins) errors.push(`complement: ${r.status}`);
  else fs.writeFileSync(path.join(out, 'b-complement-simule.pptx'), Buffer.from(r.ins.b64, 'base64'));
  console.log(`complement simulé: image demandée ${JSON.stringify(r.img)}, options d'insertion ${JSON.stringify(r.ins && r.ins.opts)} | ${r.status}`);
  await page.close();
}
await browser.close();
await server.close();
if (errors.length) { console.error('ERREURS\n' + errors.join('\n')); process.exit(1); }
