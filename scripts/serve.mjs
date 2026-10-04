// Sert dist/ sous le même chemin que GitHub Pages (/pptanim/), pour essayer l'application en local.
// Usage : node scripts/serve.mjs [port]   puis http://localhost:8080/pptanim/

import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const TYPES = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.png': 'image/png',
  '.webmanifest': 'application/manifest+json', '.json': 'application/json', '.css': 'text/css; charset=utf-8',
};

/** @returns {Promise<{ url, close, outside }>} outside = requêtes reçues hors du chemin de publication */
export function serve({ port = 0, base = '/pptanim/', dir = path.join(root, 'dist') } = {}) {
  const outside = [];
  const server = http.createServer((req, res) => {
    const pathname = decodeURIComponent(new URL(req.url, 'http://x').pathname);
    if (!pathname.startsWith(base)) {
      outside.push(pathname);
      res.writeHead(404).end('hors du chemin de publication');
      return;
    }
    let file = path.join(dir, pathname.slice(base.length));
    if (pathname.endsWith('/')) file = path.join(file, 'index.html');
    if (!file.startsWith(dir) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) { res.writeHead(404).end(); return; }
    res.writeHead(200, { 'content-type': TYPES[path.extname(file)] || 'application/octet-stream', 'cache-control': 'no-cache' });
    fs.createReadStream(file).pipe(res);
  });
  return new Promise((resolve) => server.listen(port, () => resolve({
    url: `http://localhost:${server.address().port}${base}`,
    outside,
    close: () => new Promise((r) => { server.closeAllConnections(); server.close(r); }),
  })));
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const { url } = await serve({ port: +process.argv[2] || 8080 });
  console.log(`PPT Anim : ${url}`);
}
