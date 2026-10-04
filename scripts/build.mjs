// Construit dist/ à partir de src/ : copie, numéro de version, service worker avec la liste des
// fichiers à mettre en cache, page du complément PowerPoint. Puis contrôle que rien ne suppose
// une publication à la racine du domaine. Aucune dépendance : `node scripts/build.mjs`.

import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const src = path.join(root, 'src');
const dist = path.join(root, 'dist');
const BASE = process.env.BASE_PATH || '/pptanim/'; // chemin de publication sur GitHub Pages
const OFFICE_JS = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js';
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));

const walk = (dir, base = dir) =>
  fs.readdirSync(dir, { withFileTypes: true }).flatMap((e) =>
    e.isDirectory() ? walk(path.join(dir, e.name), base) : [path.relative(base, path.join(dir, e.name)).split(path.sep).join('/')]);
const read = (rel) => fs.readFileSync(path.join(dist, rel), 'utf8');
const write = (rel, text) => fs.writeFileSync(path.join(dist, rel), text);
const errors = [];
const check = (ok, msg) => { if (!ok) errors.push(msg); };

// 1. Copie
fs.rmSync(dist, { recursive: true, force: true });
fs.cpSync(src, dist, { recursive: true });
const assets = walk(dist).sort();

// 2. Version : celle de package.json + empreinte du contenu, pour que le cache change à chaque modification
const hash = crypto.createHash('sha256');
for (const f of assets) hash.update(f).update(fs.readFileSync(path.join(dist, f)));
const version = `${pkg.version}-${hash.digest('hex').slice(0, 8)}`;

// 3. Pages : l'application (sans office.js) et la page du complément (avec)
const html = read('index.html');
check(html.includes('<!-- office.js -->') && html.includes('__APP_VERSION__'), 'index.html : repères de build absents');
const page = html.replaceAll('__APP_VERSION__', version);
write('index.html', page.replace('<!-- office.js -->\n', ''));
write('taskpane.html', page
  .replace('<!-- office.js -->', `<script src="${OFFICE_JS}"></script>`)
  .replace(/<link rel="manifest"[^>]*>\n/, ''));

// 4. Service worker
const sw = fs.readFileSync(path.join(root, 'scripts', 'sw.template.js'), 'utf8')
  .replace('__VERSION__', version)
  .replace('__ASSETS__', JSON.stringify(assets, null, 2));
write('sw.js', sw);

// 5. Contrôles
const exists = (rel) => fs.existsSync(path.join(dist, rel));
const pngSize = (rel) => { const b = fs.readFileSync(path.join(dist, rel)); return `${b.readUInt32BE(16)}x${b.readUInt32BE(20)}`; };

// 5a. Manifeste : seuls id, start_url et scope sont absolus, et ils valent le chemin de publication
check(BASE.startsWith('/') && BASE.endsWith('/'), `BASE_PATH doit commencer et finir par « / » : ${BASE}`);
const manifest = JSON.parse(read('manifest.webmanifest'));
for (const key of ['id', 'start_url', 'scope']) check(manifest[key] === BASE, `manifest.${key} = ${manifest[key]}, attendu ${BASE}`);
check(manifest.name && manifest.short_name && manifest.display === 'standalone', 'manifest : name, short_name ou display manquant');
for (const icon of manifest.icons || []) {
  check(!/^([a-z]+:)?\//i.test(icon.src), `manifest : icône en chemin absolu ${icon.src}`);
  check(exists(icon.src), `manifest : icône absente ${icon.src}`);
  if (exists(icon.src)) check(pngSize(icon.src) === icon.sizes, `manifest : ${icon.src} fait ${pngSize(icon.src)}, déclaré ${icon.sizes}`);
}
check(['192x192', '512x512'].every((s) => (manifest.icons || []).some((i) => i.sizes === s)), 'manifest : icônes 192 et 512 requises');
check((manifest.icons || []).some((i) => /maskable/.test(i.purpose || '')), 'manifest : icône maskable absente');

// 5b. Pages et scripts : toutes les références locales sont relatives et pointent sur un fichier présent
const local = (ref) => !/^([a-z]+:|\/\/|#|data:)/i.test(ref);
for (const file of ['index.html', 'taskpane.html']) {
  const text = read(file);
  for (const m of text.matchAll(/\b(?:href|src)="([^"]+)"/g)) {
    const ref = m[1];
    if (!local(ref)) { check(file === 'taskpane.html' && ref === OFFICE_JS, `${file} : référence externe ${ref}`); continue; }
    check(!ref.startsWith('/'), `${file} : chemin absolu ${ref}`);
    check(exists(ref.split(/[?#]/)[0]), `${file} : fichier absent ${ref}`);
  }
  check(!/url\(\s*['"]?\//.test(text), `${file} : url(/…) absolu dans le style`);
}
check(!read('index.html').includes('office.js'), "index.html : l'application ne doit pas charger office.js");
for (const file of [...assets.filter((f) => f.endsWith('.js') && !f.startsWith('vendor/')), 'sw.js']) {
  const text = read(file);
  for (const m of text.matchAll(/\bfrom\s+'([^']+)'|\bimport\s+'([^']+)'/g)) {
    const ref = m[1] || m[2];
    check(ref.startsWith('./') || ref.startsWith('../'), `${file} : import non relatif ${ref}`);
    check(exists(path.posix.join(path.posix.dirname(file), ref)), `${file} : import introuvable ${ref}`);
  }
  for (const m of text.matchAll(/(?:register|fetch|importScripts|new Worker)\(\s*['"`](\/[^'"`]*)/g)) {
    check(false, `${file} : chemin absolu ${m[1]}`);
  }
}

// 5c. Service worker : enregistré en relatif, liste de cache complète et sans intrus
check(/register\('\.\/sw\.js',\s*\{\s*scope:\s*'\.\/'\s*\}\)/.test(read('pwa.js')), 'pwa.js : enregistrement du service worker inattendu');
check(assets.every(exists), 'sw.js : fichier de cache absent');
check(assets.includes('index.html') && assets.includes('manifest.webmanifest'), 'sw.js : index.html ou manifeste hors cache');
check(!assets.includes('sw.js') && !assets.includes('taskpane.html'), 'sw.js : sw.js ou taskpane.html ne doivent pas être mis en cache');
check(!sw.includes('__'), 'sw.js : repère de build non remplacé');

if (errors.length) {
  console.error(`Build en échec :\n- ${errors.join('\n- ')}`);
  process.exit(1);
}
const size = walk(dist).reduce((n, f) => n + fs.statSync(path.join(dist, f)).size, 0);
console.log(`dist/ prêt : version ${version}, ${assets.length} fichiers en cache, ${Math.round(size / 1024)} Ko, publication sous ${BASE}`);
