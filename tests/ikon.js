// Uji ikon garis (scripts/neuron_icons.js):  node tests/ikon.js <folder-keluaran>
// (a) unit test pengurai SVG, (b) deck contoh semua ikon → <folder>/ikon.pptx (ukuran 0,4" dan 0,8", putih di atas
// persegi 1F2937, ukuran besar untuk memeriksa busur/titik) + <folder>/ikon-bulat.pptx (pasca-proses roundIconLines).
// Keluar ≠ 0 bila gagal.
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const T = require(path.join(__dirname, '..', 'scripts', 'neuron_theme.js'));
const I = require(path.join(__dirname, '..', 'scripts', 'neuron_icons.js'));
const JSZip = require(path.join(__dirname, '..', 'scripts', 'vendor', 'pptxgenjs.bundle.cjs')).JSZip;

const out = process.argv[2] || fs.mkdtempSync(path.join(require('os').tmpdir(), 'ikon-'));
fs.mkdirSync(out, { recursive: true });
let n = 0;
const ok = (m) => { n++; console.log(`  [LULUS] ${m}`); };
const near = (a, b, what, eps = 1e-6) => assert(Math.abs(a - b) <= eps, `${what}: ${a} ≠ ${b}`);
const nearPt = (p, x, y, what, eps) => { near(p[0], x, `${what} x`, eps); near(p[1], y, `${what} y`, eps); };
const end = (sub) => { const g = sub.segs[sub.segs.length - 1]; return [g.x, g.y]; };
// titik pada kubik di t
const bez = (p0, g, t) => {
  const u = 1 - t;
  return [u ** 3 * p0[0] + 3 * u * u * t * g.x1 + 3 * u * t * t * g.x2 + t ** 3 * g.x,
    u ** 3 * p0[1] + 3 * u * u * t * g.y1 + 3 * u * t * t * g.y2 + t ** 3 * g.y];
};

(async () => {
  // ---------- (a) pengurai ----------
  {
    const [p] = I.parsePath('M0 0h10v10H0z');
    assert.deepStrictEqual(p, { start: [0, 0], closed: true,
      segs: [{ t: 'L', x: 10, y: 0 }, { t: 'L', x: 10, y: 10 }, { t: 'L', x: 0, y: 10 }] });
    ok('M0 0h10v10H0z → persegi tertutup');
  }
  {
    const subs = I.parsePath('M1.5.5l-1-2 .25.25');
    assert.deepStrictEqual(subs[0].start, [1.5, 0.5]);
    assert.deepStrictEqual(subs[0].segs.map((g) => [g.x, g.y]), [[0.5, -1.5], [0.75, -1.25]]);
    const [q] = I.parsePath('M1e1,2E-1L-.5-.5');
    assert.deepStrictEqual([q.start, end(q)], [[10, 0.2], [-0.5, -0.5]]);
    ok('angka rapat (1.5.5, -1-2, .25.25, eksponen) dan l implisit setelah m');
  }
  {
    // setengah lingkaran r=5 searah jarum jam (sweep=1, y ke bawah) dari (0,0) ke (10,0): lewat (5,-5)
    const [p] = I.parsePath('M0 0A5 5 0 0 1 10 0');
    assert.strictEqual(p.segs.length, 2);
    nearPt(end(p), 10, 0, 'ujung busur');
    nearPt([p.segs[0].x, p.segs[0].y], 5, -5, 'tengah busur');
    // semua titik pada kurva berjarak ~5 dari pusat (5,0) — galat aproksimasi kubik 90° < 0,03%
    let prev = p.start;
    for (const g of p.segs) { for (let t = 0; t <= 1; t += 0.125) { const q = bez(prev, g, t); near(Math.hypot(q[0] - 5, q[1]), 5, 'jari-jari', 5 * 3e-4); } prev = [g.x, g.y]; }
    // flag rapat + sweep=0 → lewat (5,5)
    const [r] = I.parsePath('M0 0a5 5 0 0010 0');
    nearPt(end(r), 10, 0, 'ujung busur rapat'); nearPt([r.segs[0].x, r.segs[0].y], 5, 5, 'tengah busur sweep 0');
    // jari-jari terlalu kecil diperbesar (F.6.6) → tetap berujung di titik akhir
    const [s] = I.parsePath('M0 0A1 1 0 0 1 10 0');
    nearPt(end(s), 10, 0, 'ujung busur jari-jari diperbesar'); nearPt([s.segs[0].x, s.segs[0].y], 5, -5, 'tengah busur diperbesar');
    // busur besar 270°: 3 potong, berujung tepat
    const [b] = I.parsePath('M5 0A5 5 0 1 1 0 5');
    assert.strictEqual(b.segs.length, 3); nearPt(end(b), 0, 5, 'ujung busur besar');
    // busur elips berotasi: ujung tepat
    const [e] = I.parsePath('M2 3A4 2 30 1 0 9 7');
    nearPt(end(e), 9, 7, 'ujung busur elips berotasi');
    // rx/ry 0 → garis; titik akhir = awal → tidak ada segmen
    assert.deepStrictEqual(I.parsePath('M0 0A0 5 0 0 1 4 4')[0].segs, [{ t: 'L', x: 4, y: 4 }]);
    assert.strictEqual(I.parsePath('M3 3A5 5 0 0 1 3 3')[0].segs.length, 0);
    ok('busur → kubik (setengah lingkaran, flag rapat, jari-jari diperbesar, 270°, elips berotasi, kasus tepi)');
  }
  {
    const [c] = I.parsePath('M0 0C0 1 1 1 1 0S2-1 2 0');
    assert.deepStrictEqual([c.segs[1].x1, c.segs[1].y1], [1, -1]);
    const [q] = I.parsePath('M0 0Q1 1 2 0T4 0');
    nearPt([q.segs[0].x1, q.segs[0].y1], 2 / 3, 2 / 3, 'Q c1'); nearPt([q.segs[0].x2, q.segs[0].y2], 4 / 3, 2 / 3, 'Q c2');
    nearPt([q.segs[1].x1, q.segs[1].y1], 2 + 2 / 3, -2 / 3, 'T dipantulkan');
    const subs = I.parsePath('M2 2h2z m1 1h1 Z l1 1');
    assert.strictEqual(subs.length, 3);
    assert.deepStrictEqual(subs[1].start, [3, 3], 'm relatif setelah Z dihitung dari awal subpath');
    assert.deepStrictEqual([subs[2].start, end(subs[2])], [[3, 3], [4, 4]], 'perintah setelah Z membuka subpath baru');
    ok('S/Q/T (kendali dipantulkan), Z lalu m relatif / perintah tanpa M');
  }
  {
    const [c] = I.elementToSubpaths(['circle', { cx: 12, cy: 12, r: 10 }]);
    assert(c.closed && c.segs.length === 4);
    assert.deepStrictEqual(c.segs.map((g) => [g.x, g.y]), [[12, 22], [2, 12], [12, 2], [22, 12]]);
    let prev = c.start;
    for (const g of c.segs) { near(Math.hypot(...bez(prev, g, 0.5).map((v) => v - 12)), 10, 'lingkaran', 10 * 3e-4); prev = [g.x, g.y]; }
    const [e] = I.elementToSubpaths(['ellipse', { cx: 12, cy: 5, rx: 9, ry: 3 }]);
    assert.deepStrictEqual(e.segs.map((g) => [g.x, g.y]), [[12, 8], [3, 5], [12, 2], [21, 5]]);
    const [r] = I.elementToSubpaths(['rect', { x: 2, y: 4, width: 20, height: 16, rx: 2 }]);
    assert(r.closed); assert.deepStrictEqual(r.start, [4, 4]);
    assert.deepStrictEqual(r.segs.map((g) => [g.x, g.y]), [[20, 4], [22, 6], [22, 18], [20, 20], [4, 20], [2, 18], [2, 6], [4, 4]]);
    const [r2] = I.elementToSubpaths(['rect', { x: 0, y: 0, width: 4, height: 2, rx: 9 }]);
    assert.deepStrictEqual(r2.start, [2, 0], 'rx dibatasi setengah lebar');
    const [r3] = I.elementToSubpaths(['rect', { x: 1, y: 1, width: 2, height: 2 }]);
    assert.strictEqual(r3.segs.length, 3);
    const [l] = I.elementToSubpaths(['line', { x1: 1, y1: 2, x2: 3, y2: 4 }]);
    assert.deepStrictEqual(l, { start: [1, 2], segs: [{ t: 'L', x: 3, y: 4 }], closed: false });
    const [pl] = I.elementToSubpaths(['polyline', { points: '3.29 7 12 12 20.71 7' }]);
    assert.deepStrictEqual([pl.start, pl.segs.length, pl.closed], [[3.29, 7], 2, false]);
    const [pg] = I.elementToSubpaths(['polygon', { points: '0,0 4,0 2,3' }]);
    assert(pg.closed && pg.segs.length === 2);
    ok('circle/ellipse (4 kubik), rect (rx, rx dibatasi, tanpa rx), line, polyline, polygon');
  }
  {
    // titik Lucide ("h.01") → lingkaran kecil tertutup di tempatnya
    const dots = I.iconSubpaths('triangle-alert').filter((p) => p.closed && p.segs.length === 4);
    assert.strictEqual(dots.length, 1);
    nearPt([dots[0].segs[2].x, dots[0].segs[2].y], 12.005, 17 - 0.15, 'titik tanda seru', 1e-9);
    assert.strictEqual(I.iconSubpaths('server').filter((p) => p.closed && p.segs.length === 4).length, 2);
    ok('titik (subpath < 0,1 unit) menjadi lingkaran kecil');
  }
  {
    assert.throws(() => I.parsePath('M0 0 L x'), /angka diharapkan/);
    assert.throws(() => I.parsePath('L1 1'), /diawali M/);
    assert.throws(() => I.parsePath('M0 0 A1 1 0 2 0 1 1'), /flag busur/);
    assert.throws(() => I.parsePath('M0 0 X1'), /tidak dikenal/);
    assert.throws(() => I.elementToSubpaths(['text', {}]), /tidak didukung/);
    const s = { addShape() {} };
    assert.throws(() => I.icon(s, 'kucing', { x: 0, y: 0 }), /ikon "kucing" tidak ada; tersedia: users, /);
    assert.throws(() => I.icon(s, 'users', { y: 0 }), /x dan y/);
    assert.throws(() => I.icon(s, 'users', { x: 0, y: 0, color: 'merah' }), /heks/);
    assert.throws(() => I.icon(s, 'users', { x: 0, y: 0, size: 0 }), /size/);
    assert.throws(() => I.icon(s, '_meta', { x: 0, y: 0 }), /tidak ada/);
    ok('galat jelas: path rusak, elemen/ikon tak dikenal, x/y/warna/ukuran salah');
  }
  {
    assert(I.ICON_NAMES.length >= 24 && I.ICON_NAMES.length <= 30, `jumlah ikon ${I.ICON_NAMES.length}`);
    assert(/^lucide-static@\d+\.\d+\.\d+$/.test(I.META.source) && I.META.license === 'ISC');
    for (const name of I.ICON_NAMES) {
      const subs = I.iconSubpaths(name);
      assert(subs.length > 0, name);
      for (const p of subs) for (const v of [p.start, ...p.segs.flatMap((g) => (g.t === 'C' ? [[g.x1, g.y1], [g.x2, g.y2], [g.x, g.y]] : [[g.x, g.y]]))]) {
        assert(v.every(Number.isFinite), `${name}: koordinat bukan angka`);
        assert(v.every((c) => c >= -0.5 && c <= 24.5), `${name}: koordinat ${v} di luar grid 24`);
      }
    }
    assert.deepStrictEqual(I.iconSubpaths('building-2'), I.iconSubpaths('building-complex'), 'alias building-2');
    // opsi bawaan: tebal Lucide size·2/24 inci → pt; objectName; titik custGeom dalam inci
    const calls = [];
    const s = { addShape: (t, o) => calls.push([t, o]) };
    assert.strictEqual(I.icon(s, 'user', { x: 1, y: 2 }), s);
    I.icon(s, 'building-2', { x: 0, y: 0, size: 0.8, color: '#ffffff', stroke: 1.5, objectName: 'x', altText: 'Gedung' });
    const [t0, o0] = calls[0], o1 = calls[1][1];
    assert.strictEqual(t0, 'custGeom');
    assert.deepStrictEqual([o0.x, o0.y, o0.w, o0.h, o0.line, o0.objectName, o0.fill], [1, 2, 0.4, 0.4, { color: '1F2937', width: 2.4 }, 'ikon-user', undefined]);
    assert.deepStrictEqual([o1.line, o1.objectName, o1.altText], [{ color: 'FFFFFF', width: 1.5 }, 'x', 'Gedung']);
    assert(o0.points[0].moveTo && o0.points.every((p) => p.close || (p.x >= -0.01 && p.x <= 0.41 && p.y >= -0.01 && p.y <= 0.41)));
    ok(`${I.ICON_NAMES.length} ikon terurai dalam grid; opsi icon() (tebal ${o0.line.width} pt di 0,4", warna, objectName, alias)`);
  }
  {
    const xml = '<p:sp><p:nvSpPr><p:cNvPr id="2" name="ikon-user"/></p:nvSpPr><p:spPr><a:ln w="30480"><a:solidFill><a:srgbClr val="1F2937"/></a:solidFill><a:prstDash val="solid"/></a:ln></p:spPr></p:sp>'
      + '<p:sp><p:nvSpPr><p:cNvPr id="3" name="kotak"/></p:nvSpPr><p:spPr><a:ln w="12700"><a:prstDash val="solid"/></a:ln></p:spPr></p:sp>';
    const r = I.roundIconLines(xml);
    assert(r.includes('<a:ln w="30480" cap="rnd"><a:solidFill><a:srgbClr val="1F2937"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>'));
    assert(r.includes('<a:ln w="12700"><a:prstDash val="solid"/></a:ln>'), 'shape lain tidak disentuh');
    assert.strictEqual(I.roundIconLines(r), r, 'idempoten');
    ok('roundIconLines: cap="rnd" + <a:round/> hanya untuk shape ikon-…, idempoten');
  }

  // ---------- (b) deck contoh ----------
  const deck = T.createDeck({ title: 'Uji ikon Neuron' });
  const names = I.ICON_NAMES;
  const label = (s, text, x, y, w, color = T.TEXT) => s.addText(text, { x, y, w, h: 0.2, fontFace: T.FONT, fontSize: 9, color,
    align: 'center', valign: 'top', margin: 0, objectName: `label-${text}` });
  const title = (s, text) => s.addText(text, { x: 0.5, y: 0.2, w: 9, h: 0.3, fontFace: T.FONT, fontSize: 12, bold: true, color: T.INK, margin: 0 });
  const grid = (text, size, cellW, cellH, y0, { dark = false } = {}) => {
    const s = deck.pres.addSlide();
    title(s, text);
    const cols = 10;
    names.forEach((name, k) => {
      const cx = 0.5 + (k % cols) * cellW + cellW / 2, y = y0 + Math.floor(k / cols) * cellH;
      if (dark) s.addShape('rect', { x: cx - size / 2 - 0.1, y: y - 0.1, w: size + 0.2, h: size + 0.2, fill: { color: T.INK }, objectName: 'latar-gelap' });
      I.icon(s, name, { x: cx - size / 2, y, size, color: dark ? 'FFFFFF' : T.INK });
      label(s, name, cx - cellW / 2, y + size + (dark ? 0.16 : 0.06), cellW);
    });
    return s;
  };
  grid('Ikon 0,4" · 1F2937', 0.4, 0.9, 1.4, 1.0);
  grid('Ikon 0,8" · 1F2937', 0.8, 0.9, 1.6, 0.75);
  grid('Ikon 0,4" putih di atas persegi 1F2937', 0.4, 0.9, 1.4, 1.0, { dark: true });
  // detail: busur, titik, lingkaran
  const detail = ['triangle-alert', 'server', 'smartphone', 'database', 'globe', 'settings', 'cloud', 'handshake', 'rocket', 'users'];
  const s4 = deck.pres.addSlide();
  title(s4, 'Detail 1,4" (busur, titik, lingkaran)');
  detail.forEach((name, k) => {
    const x = 0.6 + (k % 5) * 1.8, y = 0.8 + Math.floor(k / 5) * 2.3;
    s4.addShape('rect', { x, y, w: 1.4, h: 1.4, line: { color: T.RULE, width: 0.5, dashType: 'dash' }, objectName: 'kotak-ikon' });
    I.icon(s4, name, { x, y, size: 1.4 });
    label(s4, name, x - 0.2, y + 1.55, 1.8);
  });
  const file = await deck.save(path.join(out, 'ikon.pptx'));
  // varian dengan ujung & sambungan bulat (pasca-proses yang disarankan untuk finalize)
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  for (const name of Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))) {
    zip.file(name, I.roundIconLines(await zip.file(name).async('string')));
  }
  const bulat = path.join(out, 'ikon-bulat.pptx');
  fs.writeFileSync(bulat, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));

  // periksa XML: satu custGeom per ikon, tanpa isian, garis berwarna
  const z = await JSZip.loadAsync(fs.readFileSync(file));
  const s1 = await z.file('ppt/slides/slide1.xml').async('string');
  const sps = s1.match(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?<\/p:sp>/g).filter((x) => /name="ikon-/.test(x));
  assert.strictEqual(sps.length, names.length);
  for (const sp of sps) {
    assert(/<a:custGeom>/.test(sp) && /<\/a:custGeom><a:noFill\/><a:ln w="30480"[^>]*><a:solidFill><a:srgbClr val="1F2937"\/>/.test(sp), 'custGeom tanpa isian, garis 2,4 pt 1F2937');
    assert.strictEqual((sp.match(/<a:path /g) || []).length, 1);
  }
  const s3 = await z.file('ppt/slides/slide3.xml').async('string');
  assert.strictEqual((s3.match(/name="ikon-[^"]*"[\s\S]*?<a:ln w="30480"[^>]*><a:solidFill><a:srgbClr val="FFFFFF"\/>/g) || []).length, names.length, 'ikon putih di latar gelap');
  const b1 = await (await JSZip.loadAsync(fs.readFileSync(bulat))).file('ppt/slides/slide1.xml').async('string');
  assert.strictEqual((b1.match(/cap="rnd"/g) || []).length, names.length);
  ok(`deck ${path.basename(file)} (4 slide: ${names.length} ikon × 3 + ${detail.length} detail) + ikon-bulat.pptx; XML: 1 custGeom/ikon, noFill, garis 1F2937`);
  console.log(`Ikon: ${n} uji lulus → ${file}`);
})().catch((e) => { console.error(e && e.stack ? e.stack : e); process.exit(1); });
