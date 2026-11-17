/* 2X Selection Fix: makes every box on 2X Interim slides clickable, movable and deletable in PowerPoint.
   JavaScript port of fix_selection.py (2X Deck Kit). Runs in a browser or in Node (pass JSZip, DOMParser,
   XMLSerializer). Usage: const out = await TwoXFix.fixPackage(arrayBuffer, env) -> {data: Uint8Array, report}. */
(function (root) {
  'use strict';
  const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
  const ART_TAG = '2X layout artwork', TEXTBOX_TAG = '2X text box', NUMBER_TAG = '2X number';
  const GAP = 45720, SHORT = Math.trunc(0.55 * 914400), TALL = Math.trunc(0.62 * 914400);
  const LVLS = []; for (let i = 1; i <= 9; i++) LVLS.push('lvl' + i + 'pPr');
  const SHAPE_TAGS = ['sp', 'pic', 'grpSp', 'cxnSp', 'graphicFrame', 'contentPart'];
  // placeholder type codes as python-pptx reports them
  const PH_TYPE = { title: 1, body: 2, ctrTitle: 3, subTitle: 4, obj: 7, chart: 8, clipArt: 9, media: 10, dgm: 11,
    tbl: 12, sldNum: 13, hdr: 14, ftr: 15, dt: 16, pic: 18, sldImg: 101 };
  const BASE_TYPE = { 2: 2, 8: 2, 9: 2, 3: 1, 11: 2, 16: 16, 15: 15, 10: 2, 7: 2, 18: 2, 13: 13, 4: 2, 12: 2, 1: 1 };
  const NO_GROW = [13, 14, 18];

  // ---------------------------------------------------------------- DOM helpers
  const kids = (el) => { const o = []; if (!el) return o; for (let c = el.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) o.push(c); return o; };
  const ln = (el) => el.localName || (el.nodeName.indexOf(':') >= 0 ? el.nodeName.split(':')[1] : el.nodeName);
  const isA = (el, ns, name) => el && el.nodeType === 1 && el.namespaceURI === ns && ln(el) === name;
  const child = (el, ns, name) => { if (!el) return null; for (const c of kids(el)) if (isA(c, ns, name)) return c; return null; };
  const path = (el, steps) => { let cur = el; for (const [ns, n] of steps) { cur = child(cur, ns, n); if (!cur) return null; } return cur; };
  const desc = (el, ns, name) => { // first descendant (document order), not el itself
    if (!el) return null;
    for (const c of kids(el)) { if (isA(c, ns, name)) return c; const d = desc(c, ns, name); if (d) return d; }
    return null;
  };
  const allDesc = (el, ns, name, out = []) => { for (const c of kids(el)) { if (isA(c, ns, name)) out.push(c); allDesc(c, ns, name, out); } return out; };
  const iterAll = (el, out = []) => { out.push(el); for (const c of kids(el)) iterAll(c, out); return out; };
  const mk = (doc, ns, q) => doc.createElementNS(ns, q);
  const A = (doc, name, attrs) => { const e = mk(doc, NS_A, 'a:' + name); if (attrs) for (const k in attrs) e.setAttribute(k, String(attrs[k])); return e; };
  const Pn = (doc, name, attrs) => { const e = mk(doc, NS_P, 'p:' + name); if (attrs) for (const k in attrs) e.setAttribute(k, String(attrs[k])); return e; };
  const remove = (el) => { if (el && el.parentNode) el.parentNode.removeChild(el); };
  const before = (ref, el) => ref.parentNode.insertBefore(el, ref);
  const after = (ref, el) => ref.parentNode.insertBefore(el, ref.nextSibling);
  const clone = (doc, el) => (el.ownerDocument === doc ? el.cloneNode(true) : doc.importNode(el, true));
  const attr = (el, k) => (el && el.hasAttribute(k) ? el.getAttribute(k) : null);
  const index = (el) => kids(el.parentNode).indexOf(el);
  const replaceEl = (old, neu) => { old.parentNode.insertBefore(neu, old); remove(old); };

  // ---------------------------------------------------------------- package model
  function dirOf(p) { return p.substring(0, p.lastIndexOf('/') + 1); }
  function resolve(base, target) {
    if (target.startsWith('/')) return target.substring(1);
    const parts = (dirOf(base) + target).split('/'); const out = [];
    for (const s of parts) { if (s === '..') out.pop(); else if (s !== '.' && s !== '') out.push(s); }
    return out.join('/');
  }
  function relPath(fromPart, toPart) {
    const a = dirOf(fromPart).split('/').filter(Boolean), b = toPart.split('/');
    let i = 0; while (i < a.length && i < b.length - 1 && a[i] === b[i]) i++;
    return '../'.repeat(a.length - i) + b.slice(i).join('/');
  }
  function relsName(p) { return dirOf(p) + '_rels/' + p.substring(p.lastIndexOf('/') + 1) + '.rels'; }

  class Pkg {
    constructor(zip, env) { this.zip = zip; this.env = env; this.docs = {}; this.dirty = new Set(); this.decl = {}; }
    async load(name) {
      if (this.docs[name]) return this.docs[name];
      const f = this.zip.file(name); if (!f) return null;
      const txt = await f.async('string');
      const m = txt.match(/^\s*<\?xml[^>]*\?>/); this.decl[name] = m ? m[0].trim() : null;
      const doc = new this.env.DOMParser().parseFromString(txt, 'application/xml');
      this.docs[name] = doc; return doc;
    }
    async rels(part) {
      const name = relsName(part); let doc = await this.load(name);
      if (!doc) {
        doc = new this.env.DOMParser().parseFromString('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="' + NS_REL + '"/>', 'application/xml');
        this.docs[name] = doc; this.decl[name] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
      }
      const map = {};
      for (const r of kids(doc.documentElement)) map[r.getAttribute('Id')] = { id: r.getAttribute('Id'), type: r.getAttribute('Type'), target: r.getAttribute('Target'), external: r.getAttribute('TargetMode') === 'External', abs: r.getAttribute('TargetMode') === 'External' ? null : resolve(part, r.getAttribute('Target')) };
      return { doc, map, name };
    }
    async relateTo(part, abs, type, external) {
      const R = await this.rels(part);
      for (const id in R.map) { const r = R.map[id]; if (r.type === type && r.external === !!external && (external ? r.target === abs : r.abs === abs)) return id; }
      let n = 1; const used = new Set(Object.keys(R.map)); while (used.has('rId' + n)) n++;
      const id = 'rId' + n, el = R.doc.createElementNS(NS_REL, 'Relationship');
      el.setAttribute('Id', id); el.setAttribute('Type', type); el.setAttribute('Target', external ? abs : relPath(part, abs));
      if (external) el.setAttribute('TargetMode', 'External');
      R.doc.documentElement.appendChild(el); this.dirty.add(R.name); return id;
    }
    async save() {
      const ser = new this.env.XMLSerializer();
      for (const name of this.dirty) {
        let x = ser.serializeToString(this.docs[name]).replace(/^\s*<\?xml[^>]*\?>\s*/, '');
        x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + x;
        this.zip.file(name, x, { createFolders: false });
      }
      for (const n of Object.keys(this.zip.files)) if (this.zip.files[n].dir) delete this.zip.files[n];
      return this.zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
    }
  }

  // ---------------------------------------------------------------- shapes
  // A "shape" wraps an spTree child element and knows its part (slide / layout / master).
  function nv(el) { // the nvXxPr element
    for (const c of kids(el)) if (ln(c).startsWith('nv') && c.namespaceURI === NS_P) return c;
    return null;
  }
  function cNvPr(el) { const n = nv(el); return n ? child(n, NS_P, 'cNvPr') : null; }
  function phEl(el) { const n = nv(el); const p = n ? child(n, NS_P, 'nvPr') : null; return p ? child(p, NS_P, 'ph') : null; }
  function shapeName(el) { const c = cNvPr(el); return c ? (c.getAttribute('name') || '') : ''; }
  function phType(el) { const p = phEl(el); const t = p ? attr(p, 'type') : null; return t ? (PH_TYPE[t] || 0) : 7; }
  function phIdx(el) { const p = phEl(el); const i = p ? attr(p, 'idx') : null; return i ? parseInt(i, 10) : 0; }
  function isSp(el) { return isA(el, NS_P, 'sp'); }
  function isAutoShape(el) {
    if (!isSp(el) || phEl(el)) return false;
    const sp = child(el, NS_P, 'spPr'); if (!sp || !child(sp, NS_A, 'prstGeom')) return false;
    const c = child(nv(el), NS_P, 'cNvSpPr'); return !(c && attr(c, 'txBox') === '1');
  }
  function textOf(el) {
    const tb = child(el, NS_P, 'txBody'); if (!tb) return '';
    return kids(tb).filter((p) => isA(p, NS_A, 'p')).map((p) => {
      let s = ''; for (const n of iterAll(p)) { if (isA(n, NS_A, 't')) s += n.textContent; else if (isA(n, NS_A, 'br')) s += '\v'; }
      return s;
    }).join('\n');
  }
  function shapesOf(tree) { return kids(tree).filter((c) => c.namespaceURI === NS_P && SHAPE_TAGS.includes(ln(c))); }
  function ownXfrm(el) {
    if (isA(el, NS_P, 'graphicFrame')) return child(el, NS_P, 'xfrm');
    if (isA(el, NS_P, 'grpSp')) return path(el, [[NS_P, 'grpSpPr'], [NS_A, 'xfrm']]);
    return path(el, [[NS_P, 'spPr'], [NS_A, 'xfrm']]);
  }

  class Model {
    constructor(pkg) { this.pkg = pkg; }
    async init() {
      const pkg = this.pkg;
      const root = await pkg.rels('');
      const main = Object.values(root.map).find((r) => r.type === RT + 'officeDocument');
      this.presName = main.abs; this.pres = await pkg.load(this.presName);
      const pr = await pkg.rels(this.presName);
      const sz = desc(this.pres.documentElement, NS_P, 'sldSz'); this.slideH = parseInt(sz.getAttribute('cy'), 10); this.slideW = parseInt(sz.getAttribute('cx'), 10);
      this.masters = []; this.layouts = []; this.slides = [];
      const mlist = child(this.pres.documentElement, NS_P, 'sldMasterIdLst');
      for (const m of kids(mlist)) {
        const name = pr.map[m.getAttributeNS(NS_R, 'id')].abs; const doc = await pkg.load(name);
        const master = { name, doc, root: doc.documentElement, kind: 'master', layouts: [] };
        const mr = await pkg.rels(name); master.rels = mr.map;
        for (const l of kids(child(doc.documentElement, NS_P, 'sldLayoutIdLst'))) {
          const lname = mr.map[l.getAttributeNS(NS_R, 'id')].abs; const ldoc = await pkg.load(lname);
          const lay = { name: lname, doc: ldoc, root: ldoc.documentElement, kind: 'layout', master, rels: (await pkg.rels(lname)).map };
          master.layouts.push(lay); this.layouts.push(lay);
        }
        this.masters.push(master);
      }
      const byName = {}; for (const l of this.layouts) byName[l.name] = l;
      const slist = child(this.pres.documentElement, NS_P, 'sldIdLst');
      for (const s of kids(slist)) {
        const name = pr.map[s.getAttributeNS(NS_R, 'id')].abs; const doc = await pkg.load(name);
        const sr = await pkg.rels(name);
        const lrel = Object.values(sr.map).find((r) => r.type === RT + 'slideLayout');
        this.slides.push({ name, doc, root: doc.documentElement, kind: 'slide', layout: byName[lrel.abs], rels: sr.map });
      }
      for (const p of [...this.masters, ...this.layouts, ...this.slides]) pkg.dirty.add(p.name);
    }
    tree(part) { return path(part.root, [[NS_P, 'cSld'], [NS_P, 'spTree']]); }
    shapes(part) { return shapesOf(this.tree(part)).map((el) => ({ el, part })); }
    placeholders(part) { return this.shapes(part).filter((s) => phEl(s.el)); }
    layoutName(lay) { const c = child(lay.root, NS_P, 'cSld'); return c ? (c.getAttribute('name') || '') : ''; }
    base(sh) { // python-pptx _base_placeholder
      const { el, part } = sh;
      if (part.kind === 'slide') { const idx = phIdx(el); return this.placeholders(part.layout).find((p) => phIdx(p.el) === idx) || null; }
      if (part.kind === 'layout') {
        const bt = BASE_TYPE[phType(el)]; if (bt === undefined) return null;
        return this.placeholders(part.master).find((p) => phType(p.el) === bt) || null;
      }
      return null;
    }
    chain(sh) { const out = [sh]; let cur = sh; for (let i = 0; i < 3; i++) { cur = this.base(cur); if (!cur) break; out.push(cur); } return out; }
    geom(sh) { // [left, top, width, height] with placeholder inheritance; null parts when missing
      const g = [null, null, null, null];
      const xf = ownXfrm(sh.el);
      if (xf) {
        const off = child(xf, NS_A, 'off'), ext = child(xf, NS_A, 'ext');
        if (off) { g[0] = parseInt(off.getAttribute('x'), 10); g[1] = parseInt(off.getAttribute('y'), 10); }
        if (ext) { g[2] = parseInt(ext.getAttribute('cx'), 10); g[3] = parseInt(ext.getAttribute('cy'), 10); }
      }
      if (g.includes(null) && phEl(sh.el)) {
        const b = this.base(sh);
        if (b) { const bg = this.geom(b); for (let i = 0; i < 4; i++) if (g[i] === null) g[i] = bg[i]; }
      }
      return g;
    }
    setGeom(sh, l, t, w, h) {
      const doc = sh.el.ownerDocument; let sp = child(sh.el, NS_P, 'spPr');
      if (!sp) { sp = Pn(doc, 'spPr'); const n = nv(sh.el); after(n, sp); }
      let xf = child(sp, NS_A, 'xfrm');
      if (!xf) { xf = A(doc, 'xfrm'); sp.insertBefore(xf, sp.firstChild); }
      let off = child(xf, NS_A, 'off'), ext = child(xf, NS_A, 'ext');
      if (!off) { off = A(doc, 'off'); xf.insertBefore(off, xf.firstChild); }
      if (!ext) { ext = A(doc, 'ext'); after(off, ext); }
      off.setAttribute('x', String(l)); off.setAttribute('y', String(t)); ext.setAttribute('cx', String(w)); ext.setAttribute('cy', String(h));
    }
  }

  // ---------------------------------------------------------------- 1. stacking / overlays
  function box(m, sh) { const [l, t, w, h] = m.geom(sh); if (l === null || w === null || t === null || h === null) return null; return [l, t, l + w, t + h]; }
  function covered(s, b) {
    const w = Math.min(s[2], b[2]) - Math.max(s[0], b[0]), h = Math.min(s[3], b[3]) - Math.max(s[1], b[1]);
    const area = (s[2] - s[0]) * (s[3] - s[1]); return area > 0 && Math.max(0, w) * Math.max(0, h) >= 0.5 * area;
  }
  function wordy(el) { return isSp(el) && (!!phEl(el) || !!textOf(el).trim()); }
  function artKind(el) { // python-pptx shape types 1 autoshape, 5 freeform, 9 line, 13 picture
    if (phEl(el)) return 0;
    if (isA(el, NS_P, 'cxnSp')) return 9;
    if (isA(el, NS_P, 'pic')) return 13;
    if (!isSp(el)) return 0;
    const c = child(nv(el), NS_P, 'cNvSpPr'); if (c && attr(c, 'txBox') === '1') return 17;
    const sp = child(el, NS_P, 'spPr'); if (sp && child(sp, NS_A, 'custGeom')) return 5;
    return isAutoShape(el) ? 1 : 0;
  }
  function textOverArt(m, part) {
    let moved = 0;
    for (let it = 0; it < 50; it++) {
      const items = m.shapes(part).map((sh) => [sh, box(m, sh)]); let change = false;
      for (let j = 0; j < items.length && !change; j++) {
        const [txt, tb] = items[j]; if (!tb || !wordy(txt.el)) continue;
        for (const [art, ab] of items.slice(j + 1)) {
          if (!ab || phEl(art.el) || wordy(art.el) || ![1, 5, 9, 13].includes(artKind(art.el))) continue;
          if (covered(tb, ab)) { before(txt.el, art.el); moved++; change = true; break; }
        }
      }
      if (!change) break;
    }
    return moved;
  }
  function fixTree(m, part) {
    let moved = 0;
    for (let it = 0; it < 20; it++) {
      const items = m.shapes(part).map((sh) => [sh, box(m, sh)]); let change = false;
      for (let j = 0; j < items.length && !change; j++) {
        const [low, lb] = items[j]; if (!lb || !wordy(low.el)) continue;
        const la = (lb[2] - lb[0]) * (lb[3] - lb[1]);
        for (const [high, hb] of items.slice(j + 1)) {
          if (!hb || !isSp(high.el)) continue;
          const ha = (hb[2] - hb[0]) * (hb[3] - hb[1]);
          if (ha > la && covered(lb, hb)) { after(high.el, low.el); moved++; change = true; break; }
        }
      }
      if (!change) break;
    }
    return moved;
  }
  function unclickOverlays(m, part) {
    let n = 0;
    for (const sh of m.shapes(part)) {
      if (shapeName(sh.el) !== 'Photo overlay') continue;
      const c = desc(sh.el, NS_P, 'cNvSpPr'); if (!c) continue;
      let lk = child(c, NS_A, 'spLocks'); if (!lk) { lk = A(sh.el.ownerDocument, 'spLocks'); c.appendChild(lk); }
      if (lk.getAttribute('noSelect') !== '1') { lk.setAttribute('noSelect', '1'); n++; }
    }
    return n;
  }

  // Brand artwork stays fixed on slides: the footer brand line ("Growth. Orchestrated." | 2X logo) and the baton
  // watermark behind the title. The page number stays editable.
  const FOOTER_FIXED = ['Tagline', 'Footer divider', '2X logo', 'Baton watermark'];
  const LOCKS = { sp: ['cNvSpPr', 'spLocks', ['noGrp', 'noSelect', 'noRot', 'noMove', 'noResize', 'noTextEdit']],
    pic: ['cNvPicPr', 'picLocks', ['noGrp', 'noSelect', 'noRot', 'noChangeAspect', 'noMove', 'noResize']],
    cxnSp: ['cNvCxnSpPr', 'cxnSpLocks', ['noGrp', 'noSelect', 'noRot', 'noMove', 'noResize']] };
  function lockFooter(m, part) {
    let n = 0;
    for (const sh of m.shapes(part)) {
      const L = LOCKS[ln(sh.el)]; if (!L || !FOOTER_FIXED.includes(shapeName(sh.el))) continue;
      const cp = cNvPr(sh.el); if (!cp || cp.getAttribute('descr') !== ART_TAG) continue;
      const c = desc(sh.el, NS_P, L[0]); if (!c) continue;
      let lk = child(c, NS_A, L[1]); if (!lk) { lk = A(sh.el.ownerDocument, L[1]); c.insertBefore(lk, c.firstChild); }
      if (L[2].some((a) => lk.getAttribute(a) !== '1')) { for (const a of L[2]) lk.setAttribute(a, '1'); n++; }
    }
    return n;
  }

  // ---------------------------------------------------------------- 3. detach layout artwork onto slides
  async function copyArt(m, src, slide, ref) {
    let n = 0; const doc = slide.doc;
    for (const el of kids(m.tree(src))) {
      if (el.namespaceURI !== NS_P || !['sp', 'pic', 'grpSp', 'cxnSp', 'graphicFrame'].includes(ln(el))) continue;
      if (desc(el, NS_P, 'ph')) continue;
      const neu = clone(doc, el);
      for (const node of iterAll(neu)) {
        for (const a of ['embed', 'link', 'id']) {
          if (!node.hasAttributeNS(NS_R, a)) continue;
          const rid = node.getAttributeNS(NS_R, a), rel = src.rels[rid]; if (!rel) continue;
          const nid = await m.pkg.relateTo(slide.name, rel.external ? rel.target : rel.abs, rel.type, rel.external);
          node.setAttributeNS(NS_R, 'r:' + a, nid);
        }
      }
      const c = desc(neu, NS_P, 'cNvPr') || (isA(neu, NS_P, 'cNvPr') ? neu : null);
      if (c) c.setAttribute('descr', ART_TAG);
      before(ref, neu); n++;
    }
    return n;
  }
  async function detach(m) {
    let total = 0;
    for (const s of m.slides) {
      if (s.root.getAttribute('showMasterSp') === '0') continue;
      const tree = m.tree(s); let first = kids(tree).find((e) => e.namespaceURI === NS_P && SHAPE_TAGS.includes(ln(e))); let temp = null;
      if (!first) { first = temp = Pn(s.doc, 'sp'); tree.appendChild(first); }
      const lay = s.layout;
      if (lay.root.getAttribute('showMasterSp') !== '0') total += await copyArt(m, lay.master, s, first);
      total += await copyArt(m, lay, s, first);
      if (temp) remove(temp);
      s.root.setAttribute('showMasterSp', '0');
    }
    return total;
  }

  // ---------------------------------------------------------------- 4. titles over eyebrows
  function anchorOf(m, sh) {
    for (const s of m.chain(sh)) { const bp = desc(s.el, NS_A, 'bodyPr'); if (bp && attr(bp, 'anchor')) return bp.getAttribute('anchor'); }
    return 't';
  }
  function untangleTitles(m, part) {
    let n = 0;
    const items = m.shapes(part).filter((sh) => isSp(sh.el) && m.geom(sh)[2]);
    for (const t of items) {
      if (!(phEl(t.el) && [1, 3].includes(phType(t.el)))) continue;
      if (anchorOf(m, t) !== 'b') continue;
      let [tl, tt, tw, th] = m.geom(t); const bottom = tt + th;
      for (const e of items) {
        const [el_, et, ew, eh] = m.geom(e); [tl, tt, tw, th] = m.geom(t);
        if (e === t || eh >= th) continue;
        const horiz = Math.min(el_ + ew, tl + tw) - Math.max(el_, tl);
        if (horiz <= 0 || !(tt <= et + eh && et + eh <= bottom - 300000)) continue;
        if (et + eh + GAP > tt) {
          const nt = et + eh + GAP;
          if (bottom - nt >= 400000) {
            m.setGeom(t, tl, nt, tw, bottom - nt);
            const tb = child(t.el, NS_P, 'txBody'); const bp = child(tb, NS_A, 'bodyPr');
            for (const c of kids(bp)) if (['normAutofit', 'spAutoFit', 'noAutofit'].includes(ln(c)) && c.namespaceURI === NS_A) remove(c);
            bp.appendChild(A(t.el.ownerDocument, 'noAutofit')); bp.setAttribute('anchor', 'b'); n++;
          }
        }
      }
    }
    return n;
  }

  // ---------------------------------------------------------------- 5. short placeholders -> text boxes
  const PPR_ORDER = ['lnSpc', 'spcBef', 'spcAft', 'buClrTx', 'buClr', 'buSzTx', 'buSzPct', 'buSzPts', 'buFontTx', 'buFont', 'buNone', 'buAutoNum', 'buChar', 'buBlip', 'tabLst', 'defRPr', 'extLst'];
  function sortKids(el, keyFn) {
    const ks = kids(el).map((c, i) => [keyFn(c), i, c]); ks.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
    for (const [, , c] of ks) el.appendChild(c);
  }
  function mergePpr(doc, base, over) {
    let out = base ? clone(doc, base) : null;
    if (!over) return out;
    if (!out) return clone(doc, over);
    for (let i = 0; i < over.attributes.length; i++) { const a = over.attributes[i]; if (!a.name.startsWith('xmlns')) out.setAttribute(a.name, a.value); }
    const groups = [['buNone', 'buAutoNum', 'buChar', 'buBlip'], ['buClrTx', 'buClr'], ['buSzTx', 'buSzPct', 'buSzPts'], ['buFontTx', 'buFont']];
    for (const g of groups) if (g.some((x) => child(over, NS_A, x))) for (const x of g) for (const e of kids(out)) if (isA(e, NS_A, x)) remove(e);
    for (const ch of kids(over)) {
      const name = ln(ch), old = child(out, NS_A, name);
      if (name === 'defRPr' && old) replaceEl(old, mergePpr(doc, old, ch));
      else if (old) replaceEl(old, clone(doc, ch));
      else out.appendChild(clone(doc, ch));
    }
    sortKids(out, (e) => { const i = PPR_ORDER.indexOf(ln(e)); return i < 0 ? 99 : i; });
    return out;
  }
  function effectiveText(m, sh, master) {
    const doc = sh.el.ownerDocument;
    const kind = [1, 3].includes(phType(sh.el)) ? 'titleStyle' : 'bodyStyle';
    const txs = desc(master.root, NS_P, 'txStyles'); const tx = txs ? child(txs, NS_P, kind) : null;
    const levels = {}; for (const lv of LVLS) levels[lv] = tx ? child(tx, NS_A, lv) : null;
    let body = null;
    for (const s of m.chain(sh).reverse()) {
      const tb = child(s.el, NS_P, 'txBody'); if (!tb) continue;
      const bp = child(tb, NS_A, 'bodyPr');
      if (bp) {
        if (!body) body = clone(doc, bp);
        else {
          for (let i = 0; i < bp.attributes.length; i++) { const a = bp.attributes[i]; if (!a.name.startsWith('xmlns')) body.setAttribute(a.name, a.value); }
          const fits = kids(bp).filter((c) => ['noAutofit', 'normAutofit', 'spAutoFit'].includes(ln(c)));
          if (fits.length) { for (const c of kids(body)) if (['noAutofit', 'normAutofit', 'spAutoFit'].includes(ln(c))) remove(c); body.appendChild(clone(doc, fits[0])); }
        }
      }
      const ls = child(tb, NS_A, 'lstStyle');
      if (ls) for (const lv of LVLS) { const o = child(ls, NS_A, lv); if (o) levels[lv] = mergePpr(doc, levels[lv], o); }
    }
    const lst = A(doc, 'lstStyle');
    for (const lv of LVLS) if (levels[lv]) { let e = clone(doc, levels[lv]); if (ln(e) !== lv) { const r = A(doc, lv); for (let i = 0; i < e.attributes.length; i++) r.setAttribute(e.attributes[i].name, e.attributes[i].value); for (const c of kids(e)) r.appendChild(c); e = r; } lst.appendChild(e); }
    if (!body) body = A(doc, 'bodyPr');
    for (const c of kids(body)) if (['normAutofit', 'spAutoFit'].includes(ln(c))) remove(c);
    if (!child(body, NS_A, 'noAutofit')) body.appendChild(A(doc, 'noAutofit'));
    return [body, lst];
  }
  function toTextboxes(m, slide) {
    let n = 0; const master = slide.layout.master;
    for (const ph of m.placeholders(slide)) {
      const t = phType(ph.el);
      const [l, tp, w, h] = m.geom(ph);
      if ([1, 3, 18].includes(t) || !isSp(ph.el) || h === null || h >= SHORT) continue;
      if (!textOf(ph.el).trim()) continue;
      const doc = slide.doc;
      const [body, lst] = effectiveText(m, ph, master);
      const groups = [['prstGeom', 'custGeom'], ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'], ['ln'], ['effectLst', 'effectDag'], ['scene3d'], ['sp3d']];
      const chain = m.chain(ph); const chosen = [];
      for (const g of groups) {
        let pick = null;
        for (const s of chain) {
          const ssp = child(s.el, NS_P, 'spPr'); if (!ssp) continue;
          const hit = kids(ssp).filter((k) => g.includes(ln(k)));
          if (hit.length) { pick = clone(doc, hit[0]); break; }
        }
        if (!pick && g[0] === 'prstGeom') { pick = A(doc, 'prstGeom', { prst: 'rect' }); pick.appendChild(A(doc, 'avLst')); }
        if (!pick && g[0] === 'noFill') pick = A(doc, 'noFill');
        if (pick) chosen.push(pick);
      }
      let style = null;
      if (!child(ph.el, NS_P, 'style')) for (const s of chain.slice(1)) { const st = child(s.el, NS_P, 'style'); if (st) { style = clone(doc, st); break; } }
      const el = ph.el, nvp = child(el, NS_P, 'nvSpPr'), nvpr = child(nvp, NS_P, 'nvPr'), pe = child(nvpr, NS_P, 'ph');
      if (pe) remove(pe);
      const c = child(nvp, NS_P, 'cNvSpPr'); for (const k of kids(c)) remove(k); c.setAttribute('txBox', '1');
      child(nvp, NS_P, 'cNvPr').setAttribute('descr', TEXTBOX_TAG);
      const sp = child(el, NS_P, 'spPr'); for (const k of kids(sp)) remove(k);
      const xf = A(doc, 'xfrm'); xf.appendChild(A(doc, 'off', { x: l, y: tp })); xf.appendChild(A(doc, 'ext', { cx: w, cy: h })); sp.appendChild(xf);
      for (const k of chosen) sp.appendChild(k);
      if (style) before(child(el, NS_P, 'txBody'), style);
      const tb = child(el, NS_P, 'txBody'); const oldBp = child(tb, NS_A, 'bodyPr'), oldLs = child(tb, NS_A, 'lstStyle');
      if (oldBp) replaceEl(oldBp, body); else tb.insertBefore(body, tb.firstChild);
      if (oldLs) replaceEl(oldLs, lst); else after(body, lst);
      n++;
    }
    return n;
  }

  // ---------------------------------------------------------------- 6. grow short layout placeholders
  function insets(m, sh) {
    let t = 45720, b = 45720;
    for (const s of m.chain(sh).reverse()) { const bp = path(s.el, [[NS_P, 'txBody'], [NS_A, 'bodyPr']]); if (bp) { if (attr(bp, 'tIns') !== null) t = parseInt(bp.getAttribute('tIns'), 10); if (attr(bp, 'bIns') !== null) b = parseInt(bp.getAttribute('bIns'), 10); } }
    return [t, b];
  }
  function painted(m, sh) {
    for (const s of m.chain(sh)) {
      const sp = child(s.el, NS_P, 'spPr'); if (!sp) continue;
      for (const k of kids(sp)) { const n = ln(k); if (['solidFill', 'gradFill', 'blipFill', 'pattFill'].includes(n)) return true; if (n === 'noFill') break; }
      const l = child(sp, NS_A, 'ln'); if (l && !child(l, NS_A, 'noFill') && kids(l).length) return true;
    }
    return !!child(sh.el, NS_P, 'style');
  }
  function growShort(m, lay) {
    const phs = m.placeholders(lay); const boxes = new Map();
    for (const ph of phs) boxes.set(phIdx(ph.el), m.geom(ph));
    const out = {};
    for (const ph of phs) {
      const idx = phIdx(ph.el); if (NO_GROW.includes(phType(ph.el)) || !boxes.has(idx)) continue;
      const [l, t, w, h] = boxes.get(idx);
      if ([l, t, w, h].includes(null) || h >= TALL || w <= 0 || painted(m, ph)) continue;
      const need = TALL - h; let above = t, below = m.slideH - (t + h);
      for (const [i2, [l2, t2, w2, h2]] of boxes) {
        if (i2 === idx || [l2, t2, w2, h2].includes(null) || l2 >= l + w || l2 + w2 <= l) continue;
        if (t2 + h2 <= t) above = Math.min(above, t - (t2 + h2) - GAP); else if (t2 >= t + h) below = Math.min(below, t2 - (t + h) - GAP);
      }
      above = Math.max(0, above); below = Math.max(0, below);
      let up = Math.min(Math.floor(need / 2), above); let down = Math.min(need - up, below); up = Math.min(need - down, above);
      if (up + down < need) {
        up += need - up - down;
        if (t - up < 0) { down += up - t; up = t; } // never past the top of the page
        const tree = ph.el.parentNode, ext = child(tree, NS_P, 'extLst');
        if (ext) before(ext, ph.el); else tree.appendChild(ph.el);
      }
      const [tins, bins] = insets(m, ph);
      m.setGeom(ph, l, t - up, w, h + up + down);
      const tb = child(ph.el, NS_P, 'txBody'); if (!tb) continue;
      let bp = child(tb, NS_A, 'bodyPr'); if (!bp) { bp = A(lay.doc, 'bodyPr'); tb.insertBefore(bp, tb.firstChild); }
      bp.setAttribute('tIns', String(tins + up)); bp.setAttribute('bIns', String(bins + down));
      out[idx] = [up, down, tins, bins]; boxes.set(idx, [l, t - up, w, h + up + down]);
    }
    return out;
  }
  function followLayout(m, slide, grown) {
    for (const ph of m.placeholders(slide)) {
      const idx = phIdx(ph.el); if (!(idx in grown)) continue;
      const [up, down, t0, b0] = grown[idx];
      const tb = child(ph.el, NS_P, 'txBody'); if (!tb) continue;
      let bp = child(tb, NS_A, 'bodyPr'); if (!bp) { bp = A(slide.doc, 'bodyPr'); tb.insertBefore(bp, tb.firstChild); }
      if (path(ph.el, [[NS_P, 'spPr'], [NS_A, 'xfrm']])) {
        if (attr(bp, 'tIns') === null) bp.setAttribute('tIns', String(t0));
        if (attr(bp, 'bIns') === null) bp.setAttribute('bIns', String(b0));
        continue;
      }
      if (attr(bp, 'tIns') !== null) bp.setAttribute('tIns', String(parseInt(bp.getAttribute('tIns'), 10) + up));
      if (attr(bp, 'bIns') !== null) bp.setAttribute('bIns', String(parseInt(bp.getAttribute('bIns'), 10) + down));
    }
  }

  // ---------------------------------------------------------------- 7. panels become their body box's shape
  function allInsets(m, sh) {
    const ins = { lIns: 91440, tIns: 45720, rIns: 91440, bIns: 45720 };
    for (const s of m.chain(sh).reverse()) { const bp = path(s.el, [[NS_P, 'txBody'], [NS_A, 'bodyPr']]); if (bp) for (const k in ins) if (attr(bp, k) !== null) ins[k] = parseInt(bp.getAttribute(k), 10); }
    return ins;
  }
  function pin(sph, bx, ins) {
    const el = sph.el, doc = el.ownerDocument; let sp = child(el, NS_P, 'spPr');
    if (!child(sp, NS_A, 'xfrm')) { const [l, t, w, h] = bx; const xf = A(doc, 'xfrm'); xf.appendChild(A(doc, 'off', { x: l, y: t })); xf.appendChild(A(doc, 'ext', { cx: w, cy: h })); sp.insertBefore(xf, sp.firstChild); }
    for (const k of kids(sp)) if (['prstGeom', 'custGeom', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'ln'].includes(ln(k))) remove(k);
    const g = A(doc, 'prstGeom', { prst: 'rect' }); g.appendChild(A(doc, 'avLst')); sp.appendChild(g);
    sp.appendChild(A(doc, 'noFill')); const l2 = A(doc, 'ln'); l2.appendChild(A(doc, 'noFill')); sp.appendChild(l2);
    const tb = child(el, NS_P, 'txBody');
    if (tb) { let bp = child(tb, NS_A, 'bodyPr'); if (!bp) { bp = A(doc, 'bodyPr'); tb.insertBefore(bp, tb.firstChild); } for (const k in ins) if (attr(bp, k) === null) bp.setAttribute(k, String(ins[k])); }
    const st = child(el, NS_P, 'style'); if (st) remove(st);
  }
  function panelize(m) {
    let done = 0;
    for (const lay of m.layouts) {
      const jobs = [], taken = new Set();
      for (const sh of m.shapes(lay)) {
        if (phEl(sh.el) || !isAutoShape(sh.el) || !shapeName(sh.el).endsWith('_Bg')) continue;
        const [pl, pt, pw, ph_] = m.geom(sh); let best = null;
        for (const cand of m.placeholders(lay)) {
          const idx = phIdx(cand.el);
          if (taken.has(idx) || NO_GROW.includes(phType(cand.el)) || !isSp(cand.el) || painted(m, cand)) continue;
          const [cl, ct, cw, chh] = m.geom(cand);
          if ([cl, ct, cw, chh].includes(null) || !(pl <= cl && pt <= ct && cl + cw <= pl + pw && ct + chh <= pt + ph_)) continue;
          const nm = shapeName(cand.el);
          const rank = [/(Body|Text|Bio|Note)$/.test(nm) ? 0 : /Value$/.test(nm) ? 1 : 2, -(cw * chh)];
          if (!best || rank[0] < best[0][0] || (rank[0] === best[0][0] && rank[1] < best[0][1])) best = [rank, cand];
        }
        if (!best) continue;
        const host = best[1]; taken.add(phIdx(host.el));
        jobs.push([sh, host, [pl, pt, pw, ph_], m.geom(host)]);
      }
      if (!jobs.length) continue;
      const users = m.slides.filter((s) => s.layout === lay);
      for (const [art, host, pb, hb] of jobs) {
        const idx = phIdx(host.el), ins = allInsets(m, host);
        const inside = new Set(m.placeholders(lay).filter((p2) => {
          const [x, y, w, h] = m.geom(p2); return phIdx(p2.el) !== idx && x !== null && pb[0] <= x && pb[1] <= y && x + w <= pb[0] + pb[2] && y + h <= pb[1] + pb[3];
        }).map((p2) => phIdx(p2.el)));
        for (const s of users) {
          const mine = m.placeholders(s).filter((p) => phIdx(p.el) === idx); if (!mine.length) continue;
          if (s.root.getAttribute('showMasterSp') === '0') { pin(mine[0], hb, ins); continue; }
          const el = mine[0].el; const sibs = m.placeholders(s).filter((p) => inside.has(phIdx(p.el))).map((p) => p.el);
          const first = [...sibs, el].sort((a, b) => index(a) - index(b))[0];
          if (first !== el) before(first, el);
        }
        let neu = { lIns: ins.lIns + hb[0] - pb[0], tIns: ins.tIns + hb[1] - pb[1], rIns: ins.rIns + (pb[0] + pb[2]) - (hb[0] + hb[2]), bIns: ins.bIns + (pb[1] + pb[3]) - (hb[1] + hb[3]) };
        const g = path(art.el, [[NS_P, 'spPr'], [NS_A, 'prstGeom']]);
        if (g && g.getAttribute('prst') === 'roundRect') {
          const gd = path(g, [[NS_A, 'avLst'], [NS_A, 'gd']]);
          const adj = gd ? parseInt(gd.getAttribute('fmla').trim().split(/\s+/).pop(), 10) : 16667;
          const cut = Math.trunc(adj / 100000 * Math.min(pb[2], pb[3]) * 0.29289);
          for (const k in neu) neu[k] = Math.max(0, neu[k] - cut);
        }
        const hel = host.el, doc = lay.doc; const sp = child(hel, NS_P, 'spPr'); for (const k of kids(sp)) remove(k);
        const xf = A(doc, 'xfrm'); xf.appendChild(A(doc, 'off', { x: pb[0], y: pb[1] })); xf.appendChild(A(doc, 'ext', { cx: pb[2], cy: pb[3] })); sp.appendChild(xf);
        for (const k of kids(child(art.el, NS_P, 'spPr'))) if (ln(k) !== 'xfrm') sp.appendChild(clone(doc, k));
        const ast = child(art.el, NS_P, 'style'); if (ast && !child(hel, NS_P, 'style')) before(child(hel, NS_P, 'txBody'), clone(doc, ast));
        const tb = child(hel, NS_P, 'txBody'); let bp = child(tb, NS_A, 'bodyPr'); if (!bp) { bp = A(doc, 'bodyPr'); tb.insertBefore(bp, tb.firstChild); }
        for (const k in neu) bp.setAttribute(k, String(neu[k]));
        before(art.el, hel); remove(art.el); done++;
      }
    }
    return done;
  }

  // ---------------------------------------------------------------- 8. panel numbers become placeholders
  const NUMBER_RE = /^((?:Card|Step|Item|Rung)\d+)_Number$/;
  const RPR_ORDER = ['ln', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill', 'effectLst', 'effectDag', 'highlight', 'uLnTx', 'uLn', 'uFillTx', 'uFill', 'latin', 'ea', 'cs', 'sym', 'hlinkClick', 'hlinkMouseOver', 'rtl', 'extLst'];
  function mergedPpr(m, art, doc) {
    const base = path(m.pres.documentElement, [[NS_P, 'defaultTextStyle'], [NS_A, 'lvl1pPr']]);
    const tb = child(art.el, NS_P, 'txBody'); const own = path(tb, [[NS_A, 'lstStyle'], [NS_A, 'lvl1pPr']]);
    const p0 = child(tb, NS_A, 'p'); const ppr = p0 ? child(p0, NS_A, 'pPr') : null;
    let r0 = null; if (p0) for (const r of kids(p0)) if (isA(r, NS_A, 'r')) { const rp = child(r, NS_A, 'rPr'); if (rp) { r0 = rp; break; } }
    let out = base ? clone(doc, base) : A(doc, 'lvl1pPr');
    for (const o of [own, ppr]) if (o) out = mergePpr(doc, out, o);
    if (ln(out) !== 'lvl1pPr') { const r = A(doc, 'lvl1pPr'); for (let i = 0; i < out.attributes.length; i++) if (!out.attributes[i].name.startsWith('xmlns')) r.setAttribute(out.attributes[i].name, out.attributes[i].value); for (const c of kids(out)) r.appendChild(c); out = r; }
    if (r0) {
      let d = child(out, NS_A, 'defRPr'); if (!d) { d = A(doc, 'defRPr'); out.appendChild(d); }
      for (let i = 0; i < r0.attributes.length; i++) { const a = r0.attributes[i]; if (a.name !== 'lang' && !a.name.startsWith('xmlns')) d.setAttribute(a.name, a.value); }
      for (const c of kids(r0)) { for (const old of kids(d)) if (ln(old) === ln(c)) remove(old); d.appendChild(clone(doc, c)); }
      sortKids(d, (c) => { const i = RPR_ORDER.indexOf(ln(c)); return i < 0 ? 99 : i; });
    }
    out.setAttribute('marL', attr(out, 'marL') ?? '0'); out.setAttribute('indent', attr(out, 'indent') ?? '0');
    for (const k of ['lnSpc', 'spcBef', 'spcAft']) {
      if (child(out, NS_A, k)) continue;
      const e = A(doc, k); e.appendChild(A(doc, k === 'lnSpc' ? 'spcPct' : 'spcPts', { val: k === 'lnSpc' ? '100000' : '0' }));
      const fr = child(out, NS_A, 'defRPr'); if (fr) before(fr, e); else out.appendChild(e);
    }
    if (!child(out, NS_A, 'buNone')) {
      for (const b of kids(out)) if (ln(b).startsWith('bu')) remove(b);
      const bn = A(doc, 'buNone'); const fr = child(out, NS_A, 'defRPr'); if (fr) before(fr, bn); else out.appendChild(bn);
    }
    const rank = { lnSpc: 0, spcBef: 1, spcAft: 2, tabLst: 4, defRPr: 5, extLst: 6 };
    sortKids(out, (c) => (ln(c) in rank ? rank[ln(c)] : ln(c).startsWith('bu') ? 3 : 7));
    return out;
  }
  function numberize(m) {
    const made = {};
    for (const lay of m.layouts) {
      if (m.layoutName(lay).startsWith('Agenda')) continue;
      const arts = new Map(); for (const sh of m.shapes(lay)) if (!phEl(sh.el)) arts.set(shapeName(sh.el), sh);
      const used = m.placeholders(lay).map((p) => phIdx(p.el)); let nxt = Math.max(...used, 10) + 1;
      for (const [name, art] of [...arts]) {
        const mm = NUMBER_RE.exec(name || ''); if (!mm || !isSp(art.el) || !textOf(art.el).trim()) continue;
        const text = textOf(art.el).trim(), doc = lay.doc, el = art.el, ph = clone(doc, el);
        const nvp = child(ph, NS_P, 'nvSpPr'), cnv = child(nvp, NS_P, 'cNvSpPr');
        while (cnv.attributes.length) cnv.removeAttribute(cnv.attributes[0].name);
        if (!child(cnv, NS_A, 'spLocks')) cnv.appendChild(A(doc, 'spLocks', { noGrp: '1' }));
        child(nvp, NS_P, 'cNvPr').setAttribute('descr', NUMBER_TAG);
        const nvpr = child(nvp, NS_P, 'nvPr');
        for (const a of [...Array.from(nvpr.attributes)]) if (!a.name.startsWith('xmlns')) nvpr.removeAttribute(a.name);
        for (const c of kids(nvpr)) remove(c);
        nvpr.appendChild(Pn(doc, 'ph', { type: 'body', sz: 'quarter', idx: String(nxt), hasCustomPrompt: '1' }));
        const tb = child(ph, NS_P, 'txBody'); const ls = child(tb, NS_A, 'lstStyle'); for (const c of kids(ls)) remove(c);
        ls.appendChild(mergedPpr(m, art, doc));
        const badge = arts.get(mm[1] + '_Badge'); let anchorEl = el;
        if (badge && isAutoShape(badge.el)) {
          const sp = child(ph, NS_P, 'spPr'); for (const c of kids(sp)) remove(c);
          for (const c of kids(child(badge.el, NS_P, 'spPr'))) sp.appendChild(clone(doc, c));
          child(tb, NS_A, 'bodyPr').setAttribute('anchor', 'ctr'); anchorEl = badge.el;
        }
        before(anchorEl, ph); remove(el); if (anchorEl !== el) remove(anchorEl);
        (made[lay.name] = made[lay.name] || {})[nxt] = text; nxt++;
      }
    }
    return made;
  }
  function setText(el, text) {
    const doc = el.ownerDocument, tb = child(el, NS_P, 'txBody');
    for (const p of kids(tb)) if (isA(p, NS_A, 'p')) remove(p);
    const p = A(doc, 'p'), r = A(doc, 'r'), t = A(doc, 't'); t.textContent = text; r.appendChild(t); p.appendChild(r); tb.appendChild(p);
  }
  function fillNumbers(m, made) {
    let n = 0;
    if (!made) {
      made = {};
      for (const lay of m.layouts) for (const ph of m.placeholders(lay)) {
        const c = path(ph.el, [[NS_P, 'nvSpPr'], [NS_P, 'cNvPr']]);
        if (c && c.getAttribute('descr') === NUMBER_TAG && isSp(ph.el) && textOf(ph.el).trim()) (made[lay.name] = made[lay.name] || {})[phIdx(ph.el)] = textOf(ph.el).trim();
      }
    }
    for (const s of m.slides) {
      const nums = made[s.layout.name]; if (!nums) continue;
      const have = new Map(); for (const ph of m.placeholders(s)) have.set(phIdx(ph.el), ph);
      for (const idx in nums) { const ph = have.get(parseInt(idx, 10)); if (!ph) continue; if (isSp(ph.el) && !textOf(ph.el).trim()) { setText(ph.el, nums[idx]); n++; } }
    }
    return n;
  }

  // ---------------------------------------------------------------- ids
  function uniqueIds(rootEl) {
    const seen = new Set(); let fixed = 0;
    const nodes = allDesc(rootEl, NS_P, 'cNvPr');
    let top = Math.max(1, ...nodes.map((n) => n.getAttribute('id') || '').filter((i) => /^\d+$/.test(i)).map((i) => parseInt(i, 10)));
    const ordered = nodes.map((n, i) => [n.getAttribute('descr') === ART_TAG ? 1 : 0, i, n]).sort((a, b) => a[0] - b[0] || a[1] - b[1]).map((x) => x[2]);
    for (const n of ordered) { const i = n.getAttribute('id'); if (seen.has(i) || !/^\d+$/.test(i || '')) { top++; n.setAttribute('id', String(top)); fixed++; } seen.add(n.getAttribute('id')); }
    return fixed;
  }

  // ---------------------------------------------------------------- main
  // ---------------------------------------------------------------- header tidy; mirrors tidy.py
  // Titles start 0.608 in from the top (0.29 in under the eyebrow's text), cover eyebrows sit the same distance
  // above the cover title, agenda numbers sit centred over their red stroke on the item's text line. Title boxes keep
  // a 0.62 in minimum (equal top/bottom insets) so PowerPoint for Mac still draws a selection frame.
  const T_TOP = 555955, T_LINE = 491491, T_PAD = 37718, T_GAP_EB = 265176, T_MIN_H = 566928 - 2000, T_NEAR = 91440;
  const T_E = 914400, T_UPM = 2048;
  const T_W = {" ": 485, "!": 692, "\"": 1130, "#": 1329, "$": 1341, "%": 2080, "&": 1376, "'": 694, "(": 772, ")": 772, "*": 1145, "+": 1390, ",": 684, "-": 958, ".": 684, "/": 795, "0": 1381, "1": 883, "2": 1290, "3": 1322, "4": 1385, "5": 1309, "6": 1330, "7": 1191, "8": 1333, "9": 1330, ":": 684, ";": 702, "<": 1390, "=": 1390, ">": 1390, "?": 1146, "@": 2081, "A": 1529, "B": 1355, "C": 1515, "D": 1479, "E": 1244, "F": 1202, "G": 1537, "H": 1530, "I": 575, "J": 1197, "K": 1473, "L": 1158, "M": 1908, "N": 1561, "O": 1578, "P": 1327, "Q": 1591, "R": 1345, "S": 1341, "T": 1367, "U": 1499, "V": 1529, "W": 2125, "X": 1512, "Y": 1497, "Z": 1360, "[": 772, "\\": 795, "]": 772, "^": 997, "_": 975, "`": 748, "a": 1189, "b": 1291, "c": 1205, "d": 1291, "e": 1220, "f": 815, "g": 1294, "h": 1275, "i": 555, "j": 555, "k": 1188, "l": 555, "m": 1869, "n": 1275, "o": 1256, "p": 1291, "q": 1291, "r": 834, "s": 1147, "t": 750, "u": 1275, "v": 1228, "w": 1741, "x": 1188, "y": 1233, "z": 1173, "{": 960, "|": 761, "}": 960, "~": 1390, "’": 636, "‘": 636, "“": 1106, "”": 1089, "—": 2048, "–": 1024, "·": 684, "×": 1390, "…": 2052, "€": 1402, "£": 1308, "•": 971, "→": 1954};
  const tW = (s, sz) => { let u = 0; for (const ch of s) u += (ch in T_W ? T_W[ch] : T_W['n']); return u * sz / T_UPM; };
  function tLines(text, widthPt, sz) {
    let n = 0;
    for (const para of text.replace(/\v/g, '\n').split('\n')) {
      n++; let cur = '';
      for (const word of para.split(' ')) {
        const t = cur ? cur + ' ' + word : word;
        if (cur && tW(t, sz) > widthPt) { n++; cur = word; } else cur = t;
      }
    }
    return n;
  }
  const fdiv = (a, b) => Math.floor(a / b);
  const pyStrip2 = (t) => t.replace(/^\s+|\s+$/g, '');
  const tChain = (m, sh) => (phEl(sh.el) ? m.chain(sh) : [sh]);   // python-pptx: only placeholders inherit
  function tBp(m, sh, key, dflt) {
    for (const c of tChain(m, sh)) { const b = desc(c.el, NS_A, 'bodyPr'); if (b && b.hasAttribute(key)) return b.getAttribute(key); }
    return dflt;
  }
  function tSizeH(m, sh) { // size in hundredths of a point
    const s = allDesc(sh.el, NS_A, 'rPr').filter((r) => r.hasAttribute('sz')).map((r) => parseInt(r.getAttribute('sz'), 10));
    if (s.length) return Math.max(...s);
    for (const c of tChain(m, sh)) {
      const lv = allDesc(c.el, NS_A, 'lvl1pPr').find((x) => child(x, NS_A, 'defRPr'));
      const d = lv ? child(lv, NS_A, 'defRPr') : null;
      if (d && d.getAttribute('sz')) return parseInt(d.getAttribute('sz'), 10);
    }
    return 3200;
  }
  const tL = (szh) => fdiv(T_LINE * szh, 3200);
  function tSetBox(m, sh, top, height, ins) {
    const g = m.geom(sh); m.setGeom(sh, g[0], top, g[2], height);
    const bp = child(child(sh.el, NS_P, 'txBody'), NS_A, 'bodyPr');
    bp.setAttribute('tIns', String(ins)); bp.setAttribute('bIns', String(ins)); bp.setAttribute('anchor', 'b');
  }
  function tPin(m, sh) { const g = m.geom(sh); m.setGeom(sh, g[0], g[1], g[2], g[3]); }
  const tTitle = (list) => list.find((x) => phEl(x.el) && [1, 3].includes(phType(x.el))) || null;
  const tFirst = (m, t, n, L) => { const g = m.geom(t); return g[1] + g[3] - parseInt(tBp(m, t, 'bIns', 45720), 10) - n * L; };
  const tPlace = (m, t, n, szh) => tSetBox(m, t, T_TOP - T_PAD, n * tL(szh) + 2 * T_PAD, T_PAD);

  function tidyLayoutTitle(m, lay) {
    const t = tTitle(m.placeholders(lay)); if (!t) return 0;
    const g = m.geom(t); if (g[1] === null || tBp(m, t, 'anchor', 't') !== 'b') return 0;
    const first = tFirst(m, t, 1, T_LINE);   // layout titles hold one line; the box may be taller
    if (!(T_TOP - 228600 < first && first < T_TOP + 228600) || Math.abs(first - T_TOP) <= 2000) return 0;
    tPlace(m, t, 1, 3200); return 1;
  }
  function tidySlideTitle(m, s) {
    const shapes = m.shapes(s);
    if (!shapes.some((x) => shapeName(x.el) === 'Baton watermark')) return 0;
    const t = tTitle(shapes); if (!t || !child(t.el, NS_P, 'txBody')) return 0;
    const g = m.geom(t); if (g[2] === null || g[1] === null) return 0;
    const szh = tSizeH(m, t), L = tL(szh);
    const width = (g[2] - parseInt(tBp(m, t, 'lIns', 91440), 10) - parseInt(tBp(m, t, 'rIns', 91440), 10)) / 12700;
    const n = tLines(textOf(t.el), width, szh / 100);
    if (tBp(m, t, 'anchor', 't') === 'b' && Math.abs(tFirst(m, t, n, L) - T_TOP) <= 2000 && g[3] >= T_MIN_H) return 0;
    const below = shapes.filter((x) => x !== t && shapeName(x.el) !== 'Baton watermark').map((x) => m.geom(x)[1])
      .filter((y) => y !== null && y > T_TOP + fdiv(T_LINE, 2));
    if (below.length && T_TOP + n * L > Math.min(...below) - T_NEAR) { tPin(m, t); return 0; }
    tPlace(m, t, n, szh); return 1;
  }
  function tidyCover(m, part, isLayout) {
    const t = tTitle(isLayout ? m.placeholders(part) : m.shapes(part)); if (!t) return 0;
    let tg = m.geom(t); if (tg[1] === null) return 0;
    let eb = null;
    if (isLayout) eb = m.shapes(part).find((x) => shapeName(x.el) === 'Eyebrow') || null;
    else {
      const cands = m.shapes(part).filter((x) => {
        if (x === t || !pyStrip2(textOf(x.el))) return false;
        const g = m.geom(x); if (g[1] === null || g[3] === null) return false;
        return g[1] + g[3] <= tg[1] + 182880 && tg[1] - (g[1] + g[3]) < 548640 && Math.abs(g[0] - tg[0]) < 182880;
      });
      for (const c of cands) if (!eb || m.geom(c)[1] > m.geom(eb)[1]) eb = c;
    }
    if (!eb) return 0;
    const szh = tSizeH(m, t), L = tL(szh);
    const tt = parseInt(tBp(m, t, 'tIns', 45720), 10), tb = parseInt(tBp(m, t, 'bIns', 45720), 10);
    const n = isLayout ? Math.max(1, fdiv(tg[3] - tt - tb + fdiv(L, 2), L))
      : tLines(textOf(t.el), (tg[2] - parseInt(tBp(m, t, 'lIns', 91440), 10) - parseInt(tBp(m, t, 'rIns', 91440), 10)) / 12700, szh / 100);
    const bottom = tg[1] + tg[3]; let first;
    if (tBp(m, t, 'anchor', 't') === 'b') {
      first = bottom - tb - n * L;
      if (isLayout && tg[1] < first - tt) { m.setGeom(t, tg[0], first - tt, tg[2], bottom - (first - tt)); tg = m.geom(t); }
    } else first = tg[1] + tt;
    const want = first - T_GAP_EB;
    const ti = parseInt(tBp(m, eb, 'tIns', 45720), 10), bi = parseInt(tBp(m, eb, 'bIns', 45720), 10);
    const eg = m.geom(eb);
    if (!isLayout) {
      if (Math.abs(eg[1] + ti - want) <= 1000) return 0;
      m.setGeom(eb, eg[0], want - ti, eg[2], eg[3]); return 1;
    }
    if (Math.abs(eg[1] + ti - want) <= 1000 && eg[1] + eg[3] <= tg[1]) return 0;
    const textH = tL(tSizeH(m, eb));
    const h = Math.max(eg[3], 566928);
    const top = Math.min(tg[1] - h, want);
    const nti = want - top, nbi = Math.max(0, h - nti - textH);
    m.setGeom(eb, eg[0], top, eg[2], h);
    const b = desc(eb.el, NS_A, 'bodyPr'); b.setAttribute('tIns', String(nti)); b.setAttribute('bIns', String(nbi));
    return 1;
  }
  function tidyAgenda(m, part) {
    const shapes = m.shapes(part); let changed = 0;
    for (const h of shapes) {
      if (shapeName(h.el) === 'Agenda heading' && Math.abs(tFirst(m, h, 1, T_LINE) - T_TOP) > 2000) {
        tSetBox(m, h, T_TOP - T_PAD, T_LINE + 2 * T_PAD, T_PAD); changed++;
      }
    }
    const itemRe = /^Item\d+_Title$/;
    let its = shapes.filter((s) => phEl(s.el) && itemRe.test(shapeName(s.el)));
    if (!its.length && part.kind === 'slide') {
      const idx = {}; for (const x of m.placeholders(part.layout)) idx[phIdx(x.el)] = shapeName(x.el);
      its = shapes.filter((s) => phEl(s.el) && itemRe.test(idx[phIdx(s.el)] || ''));
    }
    const nums = shapes.filter((s) => isSp(s.el) && !phEl(s.el) && /^\d{1,2}$/.test(pyStrip2(textOf(s.el)) || 'x'));
    const dashes = shapes.filter((s) => { if (!isA(s.el, NS_P, 'cxnSp')) return false; const g = m.geom(s); return g[3] === 0 && g[2] < 548640; });
    const G = (s) => m.geom(s);
    const rows = [];
    for (const n of nums) {
      const ng = G(n), nc = ng[1] + fdiv(ng[3], 2);
      let it = null, itk = 0;
      for (const i of its) { const ig = G(i); if (ig[0] > ng[0] && Math.abs(ig[1] + fdiv(ig[3], 2) - nc) < 320040) { const k = ig[0] - ng[0]; if (!it || k < itk) { it = i; itk = k; } } }
      let d = null, dk = 0;
      for (const x of dashes) { const dg = G(x); const dy = dg[1] - ng[1]; if (dy > 0 && dy < 822960 && Math.abs(dg[0] - ng[0]) < 411480) { if (!d || dy < dk) { d = x; dk = dy; } } }
      if (it && d) rows.push([n, it, d]);
    }
    const cols = {};
    for (const [n, it] of rows) { const k = fdiv(G(it)[0], 91440); (cols[k] = cols[k] || []).push([G(it)[0], Math.trunc(tW(pyStrip2(textOf(n.el)), 32) * 12700)]); }
    for (const [n, it, d] of rows) {
      const ig = G(it), col = cols[fdiv(ig[0], 91440)];
      const cx = Math.min(...col.map((x) => x[0])) - 182880 - fdiv(Math.max(...col.map((x) => x[1])), 2);
      const c = ig[1] + fdiv(ig[3], 2);
      const ng = G(n), dg = G(d);
      const neu = [cx - fdiv(ng[2], 2), c - fdiv(ng[3], 2), cx - fdiv(dg[2], 2), c + 274320];
      const bp = desc(n.el, NS_A, 'bodyPr');
      if (ng[0] === neu[0] && ng[1] === neu[1] && dg[0] === neu[2] && dg[1] === neu[3] && bp.getAttribute('anchor') === 'ctr' && bp.getAttribute('lIns') === '0') continue;
      m.setGeom(n, neu[0], neu[1], ng[2], ng[3]);
      for (const k of ['lIns', 'tIns', 'rIns', 'bIns']) bp.setAttribute(k, '0');
      bp.setAttribute('anchor', 'ctr');
      for (const p of allDesc(n.el, NS_A, 'p')) {
        let pp = child(p, NS_A, 'pPr'); if (!pp) { pp = A(p.ownerDocument, 'pPr'); p.insertBefore(pp, p.firstChild); }
        pp.setAttribute('algn', 'ctr');
      }
      m.setGeom(d, neu[2], neu[3], dg[2], dg[3]);
      changed++;
    }
    return changed;
  }
  function tidy(m) {
    let n = 0;
    for (const lay of m.layouts) {
      const nm = m.layoutName(lay);
      if (nm.startsWith('Agenda')) n += tidyAgenda(m, lay); else if (nm.startsWith('Cover')) n += tidyCover(m, lay, true); else n += tidyLayoutTitle(m, lay);
    }
    for (const s of m.slides) {
      const nm = m.layoutName(s.layout);
      if (nm.startsWith('Agenda')) n += tidyAgenda(m, s); else if (nm.startsWith('Cover')) n += tidyCover(m, s, false); else n += tidySlideTitle(m, s);
    }
    return n;
  }

  async function fixPackage(data, env, opts) {
    opts = opts || {};
    const zip = await env.JSZip.loadAsync(data); const pkg = new Pkg(zip, env); const m = new Model(pkg); await m.init();
    const isInterim = m.layouts.some((l) => /^(Cards_Numbered|IconRow|Cover_Hero|Divider_Section)/.test(m.layoutName(l)));
    const report = { slides: m.slides.length, layouts: m.layouts.length, interim: isInterim };
    // every existing slide first gets its own copy of the layout artwork, so converting layout panels and numbers
    // into placeholders can't take anything away from it
    report.artCopied = await detach(m);
    report.panels = panelize(m);
    report.numbersMade = Object.values(numberize(m)).reduce((a, o) => a + Object.keys(o).length, 0);
    report.numbersFilled = fillNumbers(m);
    report.tidied = tidy(m);
    let moved = 0, locked = 0, converted = 0, footer = 0;
    for (const lay of m.layouts) { moved += untangleTitles(m, lay); moved += fixTree(m, lay); locked += unclickOverlays(m, lay); }
    for (const s of m.slides) { converted += toTextboxes(m, s); moved += untangleTitles(m, s); moved += fixTree(m, s); locked += unclickOverlays(m, s); footer += lockFooter(m, s); }
    const grown = {}; let grew = 0;
    for (const lay of m.layouts) { grown[lay.name] = growShort(m, lay); grew += Object.keys(grown[lay.name]).length; }
    for (const s of m.slides) followLayout(m, s, grown[s.layout.name] || {});
    for (const lay of m.layouts) moved += textOverArt(m, lay);
    report.tidied += tidy(m);   // again, after the box fixes have settled what slides inherit
    for (const mm of m.masters) uniqueIds(mm.root);
    for (const lay of m.layouts) uniqueIds(lay.root);
    for (const s of m.slides) uniqueIds(s.root);
    Object.assign(report, { moved, overlays: locked, textBoxes: converted, grown: grew, footerLocked: footer });
    report.cleaned = opts.clean ? cleanUp(m) : [];
    return { data: await pkg.save(), report };
  }

  // ---------------------------------------------------------------- clean-up (opt-in); mirrors clean_deck.py
  const FOOTER_BAND = 0.84, PASTED_FOOTER = /^GROWTH\. ORCHESTRATED\.\s*\d+\s*\/\s*\d+$/i, MEDIA = ['videoFile', 'audioFile', 'quickTimeFile'];
  function isMedia(el) {
    if (!isA(el, NS_P, 'pic')) return false;
    const p = path(el, [[NS_P, 'nvPicPr'], [NS_P, 'nvPr']]); return !!p && kids(p).some((c) => MEDIA.includes(ln(c)));
  }
  function descrOf(el) { const c = cNvPr(el); return c ? (c.getAttribute('descr') || '') : ''; }
  const pyStrip = (t) => t.replace(/^[\s\v]+|[\s\v]+$/g, '');
  function cleanUp(m) {
    const W = m.slideW, H = m.slideH, band = FOOTER_BAND * H, log = [];
    m.slides.forEach((s, k) => {
      const shapes = m.shapes(s);
      const own = shapes.filter((sh) => !phEl(sh.el) && descrOf(sh.el) !== ART_TAG);
      const phs = shapes.filter((sh) => phEl(sh.el));
      const blank = !own.length && phs.length && !phs.some((sh) => pyStrip(textOf(sh.el)) || desc(sh.el, NS_A, 'blip'));
      const fullMedia = shapes.some((sh) => { if (!isMedia(sh.el)) return false; const g = m.geom(sh); return (g[2] || 0) >= 0.98 * W && (g[3] || 0) >= 0.98 * H; });
      for (const sh of shapes) {
        const el = sh.el, d = descrOf(el), ph = phEl(el), top = m.geom(sh)[1] || 0, txt = pyStrip(textOf(el));
        let why = null;
        if (fullMedia && d === ART_TAG) why = 'template footer hidden behind full-screen video';
        else if (top > band && !ph && d !== ART_TAG && PASTED_FOOTER.test(txt)) why = 'pasted duplicate footer';
        else if (top > band && isA(el, NS_P, 'pic') && !isMedia(el) && d === '2X logo') why = 'second 2X logo in footer';
        else if (!blank && ph && attr(ph, 'type') !== 'pic' && d !== NUMBER_TAG && child(el, NS_P, 'txBody') && !txt) why = 'empty text box';
        if (why) { log.push({ slide: k + 1, name: shapeName(el), why }); el.parentNode.removeChild(el); }
      }
    });
    return log;
  }

  const api = { fixPackage };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.TwoXFix = api;
})(typeof self !== 'undefined' ? self : this);
