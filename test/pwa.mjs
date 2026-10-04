// Contrôle de la PWA construite, servie comme sur GitHub Pages (sous /pptanim/) :
// manifeste, service worker, cache, installabilité, fonctionnement hors ligne, absence de requête hors du chemin.
// Usage : node scripts/build.mjs && node test/pwa.mjs
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
const BASE = '/pptanim/';
const fails = [];
const check = (ok, msg) => { console.log(`${ok ? 'ok  ' : 'ÉCHEC'} ${msg}`); if (!ok) fails.push(msg); };

const server = await serve({ base: BASE });
const browser = await chromium.launch();
const context = await browser.newContext({ viewport: { width: 760, height: 1000 }, acceptDownloads: true });
const page = await context.newPage();
const requests = [];
page.on('request', (r) => requests.push(r.url()));
const pageErrors = [];
page.on('pageerror', (e) => pageErrors.push(e.message));

await page.goto(server.url);
await page.waitForSelector('#src-file:not([hidden])');

// Service worker : enregistré sous /pptanim/, actif, et il contrôle la page
const sw = await page.evaluate(async () => {
  const reg = await navigator.serviceWorker.ready;
  if (!navigator.serviceWorker.controller) await new Promise((r) => navigator.serviceWorker.addEventListener('controllerchange', r, { once: true }));
  return { scope: reg.scope, script: reg.active.scriptURL, controlled: !!navigator.serviceWorker.controller };
});
check(new URL(sw.scope).pathname === BASE, `portée du service worker : ${new URL(sw.scope).pathname}`);
check(new URL(sw.script).pathname === `${BASE}sw.js`, `script du service worker : ${new URL(sw.script).pathname}`);
check(sw.controlled, 'la page est contrôlée par le service worker');

// Cache : tous les fichiers annoncés y sont, sous /pptanim/
const swText = fs.readFileSync(path.join(root, 'dist', 'sw.js'), 'utf8');
const assets = JSON.parse(swText.match(/const ASSETS = (\[[\s\S]*?\]);/)[1]);
const cached = await page.evaluate(async () => {
  const out = [];
  for (const k of await caches.keys()) for (const r of await (await caches.open(k)).keys()) out.push(new URL(r.url).pathname);
  return out;
});
check(assets.every((a) => cached.includes(BASE + a)), `cache complet : ${cached.length} entrées pour ${assets.length} fichiers`);
check(cached.every((p) => p.startsWith(BASE)), 'toutes les entrées du cache sont sous /pptanim/');

// Manifeste et installabilité (critères de Chromium)
const cdp = await context.newCDPSession(page);
const man = await cdp.send('Page.getAppManifest');
const parsed = JSON.parse(man.data);
check(new URL(man.url).pathname === `${BASE}manifest.webmanifest`, `manifeste chargé depuis ${new URL(man.url).pathname}`);
check((man.errors || []).length === 0, `manifeste sans erreur d'analyse ${JSON.stringify(man.errors || [])}`);
check(parsed.start_url === BASE && parsed.scope === BASE && parsed.id === BASE, `start_url, scope et id valent ${BASE}`);
const inst = await cdp.send('Page.getInstallabilityErrors');
check(inst.installabilityErrors.length === 0, `installable ${JSON.stringify(inst.installabilityErrors.map((e) => e.errorId))}`);
for (const icon of parsed.icons) {
  const res = await page.request.get(new URL(icon.src, man.url).href);
  check(res.ok() && new URL(res.url()).pathname.startsWith(BASE), `icône ${icon.src} (${icon.sizes}, ${icon.purpose}) : ${res.status()}`);
}
check(/^\d+\.\d+\.\d+-[0-9a-f]{8}$/.test(await page.textContent('#version')), `version affichée : ${await page.textContent('#version')}`);
check((await page.textContent('#offline')) === 'disponible hors ligne', `état hors ligne : ${await page.textContent('#offline')}`);

// Hors ligne : la page se recharge et l'application produit un pptx sans réseau
await context.setOffline(true);
await page.reload();
await page.waitForSelector('#src-file:not([hidden])');
check(true, 'rechargement hors ligne');
if (fs.existsSync(path.join(fx, 'b.pptx'))) {
  await page.setInputFiles('#f-pptx', path.join(fx, 'b.pptx'));
  await page.setInputFiles('#f-png', path.join(fx, 'titre.png'));
  await page.waitForFunction(() => !document.getElementById('analyze-file').disabled);
  await page.click('#analyze-file');
  await page.waitForSelector('#fx:not([hidden])');
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#go-file')]);
  const file = await dl.path();
  check(fs.statSync(file).size > 50000, `génération hors ligne : ${dl.suggestedFilename()} (${fs.statSync(file).size} octets)`);
} else {
  console.log('     (fixtures absentes : génération hors ligne non essayée, lancer test/make_fixtures.py)');
}
await context.setOffline(false);

// Collage d'une image PNG et dépôt de fichiers : ils alimentent les champs
if (fs.existsSync(path.join(fx, 'titre.png'))) {
  await page.reload();
  await page.waitForSelector('#src-file:not([hidden])');
  const png = fs.readFileSync(path.join(fx, 'titre.png')).toString('base64');
  const pptx = fs.readFileSync(path.join(fx, 'a.pptx')).toString('base64');
  await page.evaluate(({ png, pptx }) => {
    const file = (b64, name, type) => new File([Uint8Array.from(atob(b64), (c) => c.charCodeAt(0))], name, { type });
    const paste = new DataTransfer();
    paste.items.add(file(png, 'image.png', 'image/png'));
    document.dispatchEvent(new ClipboardEvent('paste', { clipboardData: paste, bubbles: true, cancelable: true }));
    const drop = new DataTransfer();
    drop.items.add(file(pptx, 'a.pptx', 'application/vnd.openxmlformats-officedocument.presentationml.presentation'));
    document.dispatchEvent(new DragEvent('drop', { dataTransfer: drop, bubbles: true, cancelable: true }));
  }, { png, pptx });
  await page.waitForFunction(() => !document.getElementById('analyze-file').disabled, null, { timeout: 5000 }).then(
    () => check(true, 'image collée et pptx déposé : prêt à analyser'),
    () => check(false, 'image collée et pptx déposé : prêt à analyser'));
}

// Aucune requête hors de /pptanim/, aucune erreur de script
const origin = new URL(server.url).origin;
const stray = requests.filter((u) => u.startsWith('http') && !(u.startsWith(origin + BASE)));
check(stray.length === 0, `requêtes hors de ${BASE} côté page : ${JSON.stringify(stray)}`);
check(server.outside.length === 0, `requêtes hors de ${BASE} reçues par le serveur : ${JSON.stringify(server.outside)}`);
check(pageErrors.length === 0, `erreurs de script : ${JSON.stringify(pageErrors)}`);

await page.screenshot({ path: path.join(root, 'test', 'out', 'pwa.png'), fullPage: true }).catch(() => {});
await browser.close();
await server.close();
console.log(fails.length ? `\n${fails.length} contrôle(s) en échec` : '\nPWA : tous les contrôles passent.');
process.exit(fails.length ? 1 : 0);
