// Modification d'un fichier .pptx (ouvert avec JSZip) : remplace une forme de texte par des images
// de lettres et ajoute leurs animations à la chronologie de la diapo, sans toucher au reste.

import { NS, wrap, orderEffect, shuffleEffect, clickGroup, mainSeq, emptyTiming } from './timing.js';
import { orderKeyframes } from './effects.js';

const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const IMG_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const PROLOG = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]))
  .replace(/[\u0000-\u001F]/g, ' ');
const kids = (el, ns, name) => Array.from(el.childNodes).filter((n) => n.nodeType === 1 && n.namespaceURI === ns && n.localName === name);
const kid = (el, ns, name) => kids(el, ns, name)[0] || null;
const all = (el, ns, name) => Array.from(el.getElementsByTagNameNS(ns, name));

export function createEngine({ DOMParser, XMLSerializer }) {
  const parse = (xml, what) => {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error(`XML illisible : ${what}`);
    return doc;
  };
  const serialize = (doc) => PROLOG + new XMLSerializer().serializeToString(doc).replace(/^<\?xml[^>]*\?>\s*/, '');
  const fragment = (doc, xml) => {
    const tmp = parse(wrap(xml), 'fragment généré');
    return Array.from(tmp.documentElement.childNodes).map((n) => doc.importNode(n, true));
  };
  const text = async (zip, path) => {
    const f = zip.file(path);
    if (!f) throw new Error(`Fichier absent du pptx : ${path}`);
    return f.async('string');
  };
  const resolve = (base, target) => {
    if (target.startsWith('/')) return target.slice(1);
    const parts = base.split('/').slice(0, -1);
    for (const seg of target.split('/')) {
      if (seg === '..') parts.pop();
      else if (seg !== '.') parts.push(seg);
    }
    return parts.join('/');
  };
  const relsPathOf = (part) => part.replace(/([^/]+)$/, '_rels/$1.rels');

  /** Taille de la diapo et liste des diapos, dans l'ordre de la présentation. */
  async function readDeck(zip) {
    const pres = parse(await text(zip, 'ppt/presentation.xml'), 'presentation.xml');
    const rels = parse(await text(zip, 'ppt/_rels/presentation.xml.rels'), 'presentation.xml.rels');
    const targets = {};
    for (const r of all(rels, REL_NS, 'Relationship')) targets[r.getAttribute('Id')] = r.getAttribute('Target');
    const sz = all(pres, NS.p, 'sldSz')[0];
    if (!sz) throw new Error('Taille de diapo introuvable.');
    const slides = all(pres, NS.p, 'sldId').map((s) => {
      const rid = s.getAttributeNS(NS.r, 'id');
      return { path: resolve('ppt/presentation.xml', targets[rid]) };
    });
    return { size: { cx: +sz.getAttribute('cx'), cy: +sz.getAttribute('cy') }, slides };
  }

  function shapeInfo(sp) {
    const nv = all(sp, NS.p, 'cNvPr')[0];
    const paras = all(sp, NS.a, 'p').map((p) =>
      Array.from(p.getElementsByTagName('*'))
        .filter((n) => n.namespaceURI === NS.a && (n.localName === 't' || n.localName === 'br'))
        .map((n) => (n.localName === 'br' ? '\n' : n.textContent)).join(''));
    const spPr = kid(sp, NS.p, 'spPr');
    const xfrm = spPr && kid(spPr, NS.a, 'xfrm');
    const off = xfrm && kid(xfrm, NS.a, 'off'), ext = xfrm && kid(xfrm, NS.a, 'ext');
    return {
      el: sp,
      id: nv ? nv.getAttribute('id') : '',
      name: nv ? nv.getAttribute('name') : '',
      text: paras.join('\n'),
      box: off && ext ? { x: +off.getAttribute('x'), y: +off.getAttribute('y'), cx: +ext.getAttribute('cx'), cy: +ext.getAttribute('cy') } : null,
      rotated: !!xfrm && (!!+xfrm.getAttribute('rot') || xfrm.getAttribute('flipH') === '1' || xfrm.getAttribute('flipV') === '1'),
    };
  }

  /** Formes de texte de premier niveau d'une diapo. */
  async function readSlide(zip, path) {
    const doc = parse(await text(zip, path), path);
    const tree = all(doc, NS.p, 'spTree')[0];
    if (!tree) throw new Error('Diapo sans arbre de formes.');
    const shapes = kids(tree, NS.p, 'sp').map(shapeInfo).filter((s) => s.text.trim());
    return { path, doc, shapes };
  }

  /** Retrouve dans le XML la forme désignée par l'API (identifiant, nom, position). */
  function findShape(slide, ref) {
    let best = null, bestScore = 0;
    for (const s of slide.shapes) {
      let score = 0;
      if (ref.name && s.name === ref.name) score += 2;
      if (ref.id && String(s.id) === String(ref.id)) score += 2;
      if (ref.box && s.box && Math.abs(s.box.x - ref.box.x) + Math.abs(s.box.y - ref.box.y) < 25400) score += 1;
      if (score > bestScore) { best = s; bestScore = score; }
    }
    return bestScore >= 2 ? best : null;
  }

  function ensureMainSeqs(doc) {
    let timings = all(doc, NS.p, 'timing');
    let created = false;
    if (!timings.length) {
      const sld = doc.documentElement;
      const t = fragment(doc, emptyTiming())[0];
      sld.insertBefore(t, kid(sld, NS.p, 'extLst'));
      timings = [t];
      created = true;
    }
    return {
      created,
      targets: timings.map((t) => {
        if (!kid(t, NS.p, 'tnLst')) {
          const fresh = fragment(doc, emptyTiming())[0];
          t.insertBefore(kid(fresh, NS.p, 'tnLst'), t.firstChild);
        }
        const ctns = all(t, NS.p, 'cTn');
        let max = Math.max(0, ...ctns.map((c) => +c.getAttribute('id') || 0));
        const nextId = () => ++max;
        let main = ctns.find((c) => c.getAttribute('nodeType') === 'mainSeq');
        if (!main) {
          const root = ctns.find((c) => c.getAttribute('nodeType') === 'tmRoot') || ctns[0];
          let list = kid(root, NS.p, 'childTnLst');
          if (!list) list = root.appendChild(doc.createElementNS(NS.p, 'p:childTnLst'));
          const seq = fragment(doc, mainSeq(nextId()))[0];
          list.insertBefore(seq, list.firstChild);
          main = kid(seq, NS.p, 'cTn');
        }
        let list = kid(main, NS.p, 'childTnLst');
        if (!list) list = main.appendChild(doc.createElementNS(NS.p, 'p:childTnLst'));
        return { list, nextId, existing: ctns.filter((c) => c.hasAttribute('presetClass')).length };
      }),
    };
  }

  /**
   * @param zip      JSZip du pptx
   * @param o.deck   résultat de readDeck
   * @param o.slide  résultat de readSlide
   * @param o.shape  forme source (élément de slide.shapes)
   * @param o.box    cadre de la forme en EMU (par défaut celui du XML)
   * @param o.img    { width, height } de l'image source en pixels
   * @param o.crops  [{ x, y, w, h (pixels), png: Uint8Array, label }] dans l'ordre des unités
   * @param o.plan   résultat de scramblePlan
   */
  async function animate(zip, o) {
    const { deck, slide, shape, img, crops, plan } = o;
    const doc = slide.doc;
    const b = o.box || shape.box;
    if (!b) throw new Error("Position de la forme inconnue (espace réservé hérité du masque). Déplacez-la légèrement dans PowerPoint puis recommencez.");
    const kx = b.cx / img.width, ky = b.cy / img.height; // EMU par pixel
    const fx = (px) => (px * kx) / deck.size.cx, fy = (px) => (px * ky) / deck.size.cy;
    const frac = (p) => ({ x: fx(p.x), y: fy(p.y) });

    // 1. Images dans le paquet + relations
    const relsPath = relsPathOf(slide.path);
    const rels = parse(await text(zip, relsPath), relsPath);
    const used = new Set(all(rels, REL_NS, 'Relationship').map((r) => r.getAttribute('Id')));
    let rn = 0;
    const nextRid = () => { do rn++; while (used.has(`rId${rn}`)); used.add(`rId${rn}`); return `rId${rn}`; };
    const tag = o.tag || `${Date.now().toString(36)}${Math.floor(Math.random() * 1296).toString(36)}`;
    const mediaDir = resolve(slide.path, '../media/x').replace(/x$/, '');
    const rids = crops.map((c, i) => {
      const name = `pptanim-${tag}-${String(i + 1).padStart(3, '0')}.png`;
      zip.file(mediaDir + name, c.png);
      const rid = nextRid();
      const r = rels.createElementNS(REL_NS, 'Relationship');
      r.setAttribute('Id', rid);
      r.setAttribute('Type', IMG_REL);
      r.setAttribute('Target', `../media/${name}`);
      rels.documentElement.appendChild(r);
      return rid;
    });
    zip.file(relsPath, serialize(rels));

    const ctPath = '[Content_Types].xml';
    const ct = parse(await text(zip, ctPath), ctPath);
    if (!all(ct, CT_NS, 'Default').some((d) => (d.getAttribute('Extension') || '').toLowerCase() === 'png')) {
      const d = ct.createElementNS(CT_NS, 'Default');
      d.setAttribute('Extension', 'png');
      d.setAttribute('ContentType', 'image/png');
      ct.documentElement.insertBefore(d, ct.documentElement.firstChild);
      zip.file(ctPath, serialize(ct));
    }

    // 2. Images de lettres juste au-dessus de la forme d'origine, qui est masquée
    let maxId = Math.max(0, ...all(doc, NS.p, 'cNvPr').map((n) => +n.getAttribute('id') || 0));
    const spids = [];
    let anchor = shape.el;
    crops.forEach((c, i) => {
      const id = ++maxId;
      spids.push(id);
      const name = `PPTAnim ${shape.name} ${String(i + 1).padStart(3, '0')}${c.label ? ' ' + c.label : ''}`;
      const xml =
        `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
        `<p:blipFill><a:blip r:embed="${rids[i]}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
        `<p:spPr><a:xfrm><a:off x="${Math.round(b.x + c.x * kx)}" y="${Math.round(b.y + c.y * ky)}"/>` +
        `<a:ext cx="${Math.round(c.w * kx)}" cy="${Math.round(c.h * ky)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
      const pic = fragment(doc, xml)[0];
      anchor.parentNode.insertBefore(pic, anchor.nextSibling);
      anchor = pic;
    });
    if (o.hideOriginal !== false) all(shape.el, NS.p, 'cNvPr')[0].setAttribute('hidden', '1');

    // 3. Animations ajoutées à la fin de la séquence principale
    const order = [...plan.items].sort((p, q) => p.delay - q.delay || p.index - q.index);
    const effects = (nextId) =>
      order.map((it, k) => {
        const nodeType = k === 0 ? 'clickEffect' : 'withEffect';
        const spid = spids[it.index];
        if (plan.mode === 'order') {
          const kf = orderKeyframes(it).map((p) => ({ t: p.t, x: fx(p.x), y: fy(p.y) }));
          return orderEffect({ nextId, spid, kf, total: Math.round(it.delay + it.dur), nodeType });
        }
        return shuffleEffect({
          nextId, spid, off: frac(it.off), c1: frac(it.c1), c2: frac(it.c2),
          curved: o.curved !== false, delay: Math.round(it.delay), dur: Math.round(it.dur), nodeType,
        });
      }).join('');
    const seqs = ensureMainSeqs(doc);
    for (const s of seqs.targets) s.list.appendChild(fragment(doc, clickGroup({ nextId: s.nextId, effects }))[0]);

    zip.file(slide.path, serialize(doc));
    return { pictures: crops.length, timingCreated: seqs.created, existingEffects: seqs.targets[0].existing };
  }

  return { readDeck, readSlide, findShape, animate };
}
