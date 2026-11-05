// Ikon garis Neuron: subset Lucide (ISC; sebagian turunan Feather, MIT — lihat THIRD_PARTY_NOTICES.txt) digambar
// sebagai BENTUK NATIVE PowerPoint (custGeom tanpa isian, garis berwarna) — bisa diwarnai ulang/diubah ukurannya di
// PowerPoint, tidak pernah gambar. Data: assets/icons/lucide.json (elemen SVG asli, grid 24, stroke-width 2).
//
// Aturan desain (brief): warna 1F2937 (putih di atas latar gelap), tidak pernah merah, maks. 1 ikon per kartu,
// selalu berlabel teks di sebelahnya (label dibuat pemanggil).
//
//   const { icon, ICON_NAMES } = require('./neuron_icons');
//   icon(slide, 'truck', { x: 1, y: 1, size: 0.4 });
//
// Pengurai SVG di bawah menangani elemen yang dipakai subset: path (semua perintah, angka rapat seperti "1.5.5" dan
// "-1-2", flag busur rapat), circle, ellipse, rect (rx/ry), line, polyline, polygon. Busur → kubik (≤ 90° per potong);
// lingkaran/elips → 4 kubik; Q/T → kubik. Hanya modul bawaan Node.
const fs = require('fs');
const path = require('path');

const DATA = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'assets', 'icons', 'lucide.json'), 'utf8'));
const META = DATA._meta;
const ICON_NAMES = Object.keys(DATA).filter((k) => k !== '_meta');
// Nama lama yang diganti Lucide (berkas aliasnya masih ada di lucide-static) → nama baru.
const ALIASES = { 'building-2': 'building-complex' };
const GRID = 24;          // viewBox Lucide 0 0 24 24
const STROKE = 2;         // stroke-width Lucide pada grid 24
const DOT = 0.1;          // subpath terbuka sepanjang < DOT unit grid = "titik" (Lucide: "h.01" + round cap)
const DOT_R = 0.15;       // titik digambar sebagai lingkaran kecil; tebal garis menjadikannya cakram ≈ round cap
const K = 0.5522847498307936;   // 4/3·(√2−1): kubik untuk seperempat lingkaran

// ---------- pengurai path SVG ----------
// Hasil: [{ start: [x, y], segs: [{ t: 'L', x, y } | { t: 'C', x1, y1, x2, y2, x, y }], closed }]
const NUM = /[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/y;
const ARGS = { M: 2, L: 2, H: 1, V: 1, C: 6, S: 4, Q: 4, T: 2, A: 7, Z: 0 };

function parsePath(d) {
  const s = String(d);
  let i = 0;
  const err = (m) => { throw new Error(`[neuron_icons] path "${s}": ${m} (posisi ${i})`); };
  const skip = () => { while (i < s.length && /[\s,]/.test(s[i])) i++; };
  const num = () => {
    skip(); NUM.lastIndex = i;
    const m = NUM.exec(s);
    if (!m) err('angka diharapkan');
    i = NUM.lastIndex;
    return Number(m[0]);
  };
  const flag = () => {
    skip();
    if (s[i] !== '0' && s[i] !== '1') err('flag busur 0/1 diharapkan');
    return s[i++] === '1';
  };

  const subs = [];
  let sub = null, cx = 0, cy = 0, sx = 0, sy = 0;
  let prevC = null, prevQ = null;   // titik kendali terakhir untuk S / T (dipantulkan)
  const moveTo = (x, y) => { sub = { start: [x, y], segs: [], closed: false }; subs.push(sub); cx = sx = x; cy = sy = y; };
  const ensure = () => { if (!sub) err('path harus diawali M'); if (sub.closed) moveTo(sx, sy); };
  const line = (x, y) => { ensure(); sub.segs.push({ t: 'L', x, y }); cx = x; cy = y; };
  const cubic = (x1, y1, x2, y2, x, y) => { ensure(); sub.segs.push({ t: 'C', x1, y1, x2, y2, x, y }); cx = x; cy = y; };

  let cmd = null;
  skip();
  while (i < s.length) {
    if (/[A-Za-z]/.test(s[i])) {
      cmd = s[i++];
      if (!(cmd.toUpperCase() in ARGS)) err(`perintah "${cmd}" tidak dikenal`);
    } else if (!cmd) err('perintah diharapkan');
    else if (cmd === 'Z' || cmd === 'z') err('angka setelah Z');
    const C = cmd.toUpperCase(), rel = cmd !== C;
    const ox = rel ? cx : 0, oy = rel ? cy : 0;
    let q = null, c2 = null;
    switch (C) {
      case 'M': { const x = num() + ox, y = num() + oy; moveTo(x, y); cmd = rel ? 'l' : 'L'; break; }
      case 'L': { const x = num() + ox, y = num() + oy; line(x, y); break; }
      case 'H': line(num() + ox, cy); break;
      case 'V': line(cx, num() + oy); break;
      case 'C': {
        const x1 = num() + ox, y1 = num() + oy, x2 = num() + ox, y2 = num() + oy, x = num() + ox, y = num() + oy;
        cubic(x1, y1, x2, y2, x, y); c2 = [x2, y2]; break;
      }
      case 'S': {
        const [x1, y1] = prevC ? [2 * cx - prevC[0], 2 * cy - prevC[1]] : [cx, cy];
        const x2 = num() + ox, y2 = num() + oy, x = num() + ox, y = num() + oy;
        cubic(x1, y1, x2, y2, x, y); c2 = [x2, y2]; break;
      }
      case 'Q': case 'T': {
        let qx, qy;
        if (C === 'Q') { qx = num() + ox; qy = num() + oy; } else [qx, qy] = prevQ ? [2 * cx - prevQ[0], 2 * cy - prevQ[1]] : [cx, cy];
        const x = num() + ox, y = num() + oy, x0 = cx, y0 = cy;
        cubic(x0 + (2 / 3) * (qx - x0), y0 + (2 / 3) * (qy - y0), x + (2 / 3) * (qx - x), y + (2 / 3) * (qy - y), x, y);
        q = [qx, qy]; break;
      }
      case 'A': {
        const rx = num(), ry = num(), rot = num(), large = flag(), sweep = flag(), x = num() + ox, y = num() + oy;
        ensure();
        for (const sg of arcToCubic(cx, cy, rx, ry, rot, large, sweep, x, y)) {
          if (sg.t === 'L') line(sg.x, sg.y); else cubic(sg.x1, sg.y1, sg.x2, sg.y2, sg.x, sg.y);
        }
        cx = x; cy = y; break;
      }
      case 'Z': if (sub && !sub.closed) { sub.closed = true; cx = sx; cy = sy; } break;
    }
    prevC = c2; prevQ = q;
    skip();
  }
  return subs;
}

// Busur SVG (endpoint) → kubik. Algoritma SVG 1.1 Implementation Notes F.6.5–F.6.6 (endpoint → center), dipotong
// per ≤ 90°. Mengembalikan segmen absolut; rx/ry 0 → garis lurus; titik akhir = titik awal → tanpa segmen.
function arcToCubic(x1, y1, rx, ry, rotDeg, large, sweep, x2, y2) {
  if (x1 === x2 && y1 === y2) return [];
  rx = Math.abs(rx); ry = Math.abs(ry);
  if (rx === 0 || ry === 0) return [{ t: 'L', x: x2, y: y2 }];
  const phi = (rotDeg * Math.PI) / 180, cos = Math.cos(phi), sin = Math.sin(phi);
  const dx = (x1 - x2) / 2, dy = (y1 - y2) / 2;
  const x1p = cos * dx + sin * dy, y1p = -sin * dx + cos * dy;
  const lambda = (x1p * x1p) / (rx * rx) + (y1p * y1p) / (ry * ry);
  if (lambda > 1) { const k = Math.sqrt(lambda); rx *= k; ry *= k; }
  const num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p;
  const den = rx * rx * y1p * y1p + ry * ry * x1p * x1p;
  let coef = Math.sqrt(Math.max(0, num / den));
  if (large === sweep) coef = -coef;
  const cxp = (coef * rx * y1p) / ry, cyp = (-coef * ry * x1p) / rx;
  const ccx = cos * cxp - sin * cyp + (x1 + x2) / 2, ccy = sin * cxp + cos * cyp + (y1 + y2) / 2;
  const ang = (ux, uy, vx, vy) => {
    const a = Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy);
    return a;
  };
  const ux = (x1p - cxp) / rx, uy = (y1p - cyp) / ry, vx = (-x1p - cxp) / rx, vy = (-y1p - cyp) / ry;
  const th1 = ang(1, 0, ux, uy);
  let dth = ang(ux, uy, vx, vy);
  if (!sweep && dth > 0) dth -= 2 * Math.PI;
  else if (sweep && dth < 0) dth += 2 * Math.PI;
  const n = Math.max(1, Math.ceil(Math.abs(dth) / (Math.PI / 2) - 1e-9));
  const step = dth / n, t = (4 / 3) * Math.tan(step / 4);
  const pt = (a) => { const ex = rx * Math.cos(a), ey = ry * Math.sin(a); return [cos * ex - sin * ey + ccx, sin * ex + cos * ey + ccy]; };
  const dv = (a) => { const ex = -rx * Math.sin(a), ey = ry * Math.cos(a); return [cos * ex - sin * ey, sin * ex + cos * ey]; };
  const out = [];
  for (let k = 0; k < n; k++) {
    const a1 = th1 + k * step, a2 = a1 + step;
    const p1 = pt(a1), p2 = k === n - 1 ? [x2, y2] : pt(a2), d1 = dv(a1), d2 = dv(a2);
    out.push({ t: 'C', x1: p1[0] + t * d1[0], y1: p1[1] + t * d1[1], x2: p2[0] - t * d2[0], y2: p2[1] - t * d2[1], x: p2[0], y: p2[1] });
  }
  return out;
}

// ---------- elemen SVG → subpath ----------
function ellipse(cx, cy, rx, ry) {
  const kx = K * rx, ky = K * ry;
  return { start: [cx + rx, cy], closed: true, segs: [
    { t: 'C', x1: cx + rx, y1: cy + ky, x2: cx + kx, y2: cy + ry, x: cx, y: cy + ry },
    { t: 'C', x1: cx - kx, y1: cy + ry, x2: cx - rx, y2: cy + ky, x: cx - rx, y: cy },
    { t: 'C', x1: cx - rx, y1: cy - ky, x2: cx - kx, y2: cy - ry, x: cx, y: cy - ry },
    { t: 'C', x1: cx + kx, y1: cy - ry, x2: cx + rx, y2: cy - ky, x: cx + rx, y: cy },
  ] };
}
function rect({ x = 0, y = 0, width: w, height: h, rx, ry }) {
  // SVG: rx/ry yang tidak diisi mengikuti yang lain; dibatasi setengah lebar/tinggi.
  if (rx === undefined && ry !== undefined) rx = ry;
  if (ry === undefined && rx !== undefined) ry = rx;
  rx = Math.min(Math.abs(rx || 0), w / 2); ry = Math.min(Math.abs(ry || 0), h / 2);
  if (!rx || !ry) return { start: [x, y], closed: true, segs: [{ t: 'L', x: x + w, y }, { t: 'L', x: x + w, y: y + h }, { t: 'L', x, y: y + h }] };
  const kx = K * rx, ky = K * ry, r = x + w, b = y + h;
  return { start: [x + rx, y], closed: true, segs: [
    { t: 'L', x: r - rx, y },
    { t: 'C', x1: r - rx + kx, y1: y, x2: r, y2: y + ry - ky, x: r, y: y + ry },
    { t: 'L', x: r, y: b - ry },
    { t: 'C', x1: r, y1: b - ry + ky, x2: r - rx + kx, y2: b, x: r - rx, y: b },
    { t: 'L', x: x + rx, y: b },
    { t: 'C', x1: x + rx - kx, y1: b, x2: x, y2: b - ry + ky, x, y: b - ry },
    { t: 'L', x, y: y + ry },
    { t: 'C', x1: x, y1: y + ry - ky, x2: x + rx - kx, y2: y, x: x + rx, y },
  ] };
}
function poly(points, closed) {
  const n = (String(points).match(/[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/g) || []).map(Number);
  if (n.length < 4 || n.length % 2) throw new Error(`[neuron_icons] points "${points}" tidak sah`);
  const segs = [];
  for (let k = 2; k < n.length; k += 2) segs.push({ t: 'L', x: n[k], y: n[k + 1] });
  return { start: [n[0], n[1]], segs, closed };
}
function elementToSubpaths([tag, a]) {
  switch (tag) {
    case 'path': return parsePath(a.d);
    case 'circle': return [ellipse(+a.cx || 0, +a.cy || 0, +a.r, +a.r)];
    case 'ellipse': return [ellipse(+a.cx || 0, +a.cy || 0, +a.rx, +a.ry)];
    case 'rect': return [rect({ x: +a.x || 0, y: +a.y || 0, width: +a.width, height: +a.height,
      rx: a.rx === undefined ? undefined : +a.rx, ry: a.ry === undefined ? undefined : +a.ry })];
    case 'line': return [{ start: [+a.x1 || 0, +a.y1 || 0], segs: [{ t: 'L', x: +a.x2 || 0, y: +a.y2 || 0 }], closed: false }];
    case 'polyline': return [poly(a.points, false)];
    case 'polygon': return [poly(a.points, true)];
    default: throw new Error(`[neuron_icons] elemen SVG "${tag}" tidak didukung`);
  }
}

// Subpath terbuka yang sangat pendek (titik Lucide, mis. "M12 17h.01") tak terlihat tanpa round cap →
// ganti dengan lingkaran kecil tertutup di tengahnya; tebal garis membuatnya cakram selebar garis.
function dotsToCircles(subs) {
  return subs.map((p) => {
    if (p.closed) return p;
    const xs = [p.start[0]], ys = [p.start[1]];
    for (const g of p.segs) { xs.push(g.x); ys.push(g.y); if (g.t === 'C') xs.push(g.x1, g.x2), ys.push(g.y1, g.y2); }
    const w = Math.max(...xs) - Math.min(...xs), h = Math.max(...ys) - Math.min(...ys);
    if (Math.hypot(w, h) >= DOT) return p;
    return ellipse((Math.max(...xs) + Math.min(...xs)) / 2, (Math.max(...ys) + Math.min(...ys)) / 2, DOT_R, DOT_R);
  });
}

function resolve(name) {
  const key = ALIASES[name] || name;
  if (!Object.prototype.hasOwnProperty.call(DATA, key) || key === '_meta') {
    throw new Error(`[neuron_icons] ikon "${name}" tidak ada; tersedia: ${ICON_NAMES.join(', ')}`);
  }
  return key;
}
// Subpath ikon dalam koordinat grid 24 (untuk uji & pemakaian lanjutan).
function iconSubpaths(name) { return dotsToCircles(DATA[resolve(name)].flatMap(elementToSubpaths)); }

// Subpath grid → titik custGeom pptxgenjs (inci, relatif kotak shape).
const r5 = (v) => Math.round(v * 1e5) / 1e5;
function toPoints(subs, scale) {
  const P = (x, y) => ({ x: r5(x * scale), y: r5(y * scale) });
  const pts = [];
  for (const p of subs) {
    pts.push({ ...P(p.start[0], p.start[1]), moveTo: true });
    for (const g of p.segs) {
      if (g.t === 'L') pts.push(P(g.x, g.y));
      else {
        const c1 = P(g.x1, g.y1), c2 = P(g.x2, g.y2);
        pts.push({ ...P(g.x, g.y), curve: { type: 'cubic', x1: c1.x, y1: c1.y, x2: c2.x, y2: c2.y } });
      }
    }
    if (p.closed) pts.push({ close: true });
  }
  return pts;
}

// icon(slide, name, { x, y, size, color, stroke, objectName, altText }) → slide
//   size: sisi kotak (inci, bawaan 0,4); stroke: tebal garis pt (bawaan Lucide: size·2/24 inci = size·6 pt).
//   Satu shape custGeom per ikon, tanpa isian.
function icon(slide, name, { x, y, size = 0.4, color = '1F2937', stroke, objectName, altText } = {}) {
  const key = resolve(name);
  if (!slide || typeof slide.addShape !== 'function') throw new Error('[neuron_icons] slide pptxgenjs wajib diisi');
  if (!Number.isFinite(x) || !Number.isFinite(y)) throw new Error(`[neuron_icons] ikon "${name}": x dan y (inci) wajib berupa angka`);
  if (!(size > 0)) throw new Error(`[neuron_icons] ikon "${name}": size harus > 0`);
  const hex = String(color).replace(/^#/, '').toUpperCase();
  if (!/^[0-9A-F]{6}$/.test(hex)) throw new Error(`[neuron_icons] ikon "${name}": color "${color}" bukan heks 6 digit`);
  const width = stroke !== undefined ? stroke : (size * STROKE / GRID) * 72;
  if (!(width > 0)) throw new Error(`[neuron_icons] ikon "${name}": stroke harus > 0`);
  const opts = {
    x, y, w: size, h: size,
    points: toPoints(iconSubpaths(key), size / GRID),
    line: { color: hex, width: Math.round(width * 100) / 100 },
    objectName: objectName || `ikon-${key}`,
  };
  // pptxgenjs 4.0.1 tidak menulis descr untuk shape; tetap diteruskan agar berlaku bila versi berikut mendukung.
  if (altText) opts.altText = altText;
  slide.addShape('custGeom', opts);
  return slide;
}

// Pasca-proses opsional untuk XML slide (dipanggil dari finalize/penyimpan): pptxgenjs 4.0.1 tidak punya opsi ujung &
// sambungan garis untuk shape, jadi garis ikon memakai bawaan (ujung rata). Fungsi ini menambahkan cap="rnd" dan
// <a:round/> pada <a:ln> setiap shape bernama "ikon-…" — sama dengan stroke-linecap/linejoin="round" Lucide.
function roundIconLines(xml) {
  return xml.replace(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?<\/p:sp>/g, (sp) => {
    if (!/<p:cNvPr\b[^>]*\bname="ikon-/.test(sp)) return sp;
    return sp.replace(/<a:ln\b([^>]*)>([\s\S]*?)<\/a:ln>/, (m, attrs, inner) => {
      const a = /\bcap="/.test(attrs) ? attrs : `${attrs} cap="rnd"`;
      if (/<a:(round|bevel|miter)\b/.test(inner)) return `<a:ln${a}>${inner}</a:ln>`;
      // urutan CT_LineProperties: isian, prstDash, sambungan, headEnd, tailEnd
      const at = inner.search(/<a:(headEnd|tailEnd)\b/);
      const body = at < 0 ? `${inner}<a:round/>` : `${inner.slice(0, at)}<a:round/>${inner.slice(at)}`;
      return `<a:ln${a}>${body}</a:ln>`;
    });
  });
}

module.exports = { ICON_NAMES, ALIASES, META, icon, iconSubpaths, parsePath, arcToCubic, elementToSubpaths, roundIconLines };
