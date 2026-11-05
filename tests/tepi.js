// Uji input tepi & regresi temuan tinjauan. Dipanggil tests/uji.sh:  node tests/tepi.js <folder-keluaran>
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const T = require(path.join(__dirname, '..', 'scripts', 'neuron_theme.js'));
const JSZip = require(path.join(__dirname, '..', 'scripts', 'vendor', 'pptxgenjs.bundle.cjs')).JSZip;

const out = process.argv[2] || fs.mkdtempSync(path.join(require('os').tmpdir(), 'tepi-'));
fs.mkdirSync(out, { recursive: true });
const warnings = [];
const origWarn = console.warn;
console.warn = (m) => warnings.push(String(m));
const throws = (fn, re, what) => { assert.throws(fn, re, what); };
let n = 0;
const ok = (m) => { n++; origWarn.call(console, `  [LULUS] ${m}`); };

(async () => {
  // --- galat jelas untuk input wajib/salah ---
  const d0 = T.createDeck({});
  throws(() => d0.bullets({ bullets: ['a'] }), /title wajib/, 'judul kosong');
  throws(() => d0.chart({ title: 'x', series: [] }), /series wajib/, 'seri kosong');
  throws(() => d0.chart({ title: 'x', labels: ['a'], series: [{ name: 's', values: [1] }, { name: 't', values: [2] }], focus: 9 }), /focus 9/, 'focus di luar rentang');
  throws(() => d0.chart({ title: 'x', labels: ['a'], series: [{ name: 's', values: [1] }], focus: null }), /focus null/, 'focus null');
  throws(() => d0.table({ title: 'x', head: ['A', 'B', 'C'], rows: [['1']] }), /baris 1 berisi 1 sel/, 'baris tabel kurang sel');
  throws(() => d0.table({ title: 'x', head: [], rows: [] }), /head wajib/, 'header kosong');
  throws(() => d0.stats({ title: 'x', stats: [{ label: 'a' }] }), /value wajib/, 'angka kosong');
  throws(() => d0.stats({ title: 'x', stats: [{ value: '1' }], highlight: 5 }), /highlight 5/, 'highlight angka di luar rentang');
  throws(() => d0.compare({ title: 'x', rows: [{ name: 'a', after: '1' }] }), /before wajib/, 'compare tanpa before');
  throws(() => d0.funnel({ title: 'x', stages: [{ name: 'a', value: 0 }] }), /minimal satu tahap/, 'corong semua nol');
  throws(() => d0.architecture({ title: 'x', nodes: [{ id: 'a', title: 'A', col: 0, row: 0 }], links: [{ from: 'a', to: 'a' }] }), /dirinya sendiri/, 'link ke diri sendiri');
  throws(() => d0.architecture({ title: 'x', boundary: { cols: [2, 1] }, nodes: [{ id: 'a', title: 'A', col: 0, row: 0 }] }), /boundary\.cols/, 'batas terbalik');
  throws(() => d0.timeline({ title: 'x', periods: ['a', 'b'], phases: [{ name: 'p', start: 0, end: 1 }], today: 9 }), /today 9/, 'today di luar kisi');
  ok('input wajib kosong/salah memberi galat yang jelas');
  // format angka label di luar grafik
  assert.strictEqual(T.formatNum(0.048, '0.0%'), '4,8%');
  assert.strictEqual(T.formatNum(4.8, '"Rp "#,##0.0" jt"'), 'Rp 4,8 jt');
  assert.strictEqual(T.formatNum(12400, '#,##0'), '12.400');
  ok('formatNum: persen ×100, awalan/akhiran literal, ribuan titik');

  // --- peringatan meluap ---
  warnings.length = 0;
  const d1 = T.createDeck({});
  d1.closing({ steps: Array.from({ length: 4 }, () => ({ text: 'Langkah yang sangat panjang sekali '.repeat(4), owner: 'Pemilik yang panjang · tanggal' })) });
  assert(warnings.some((w) => /closing: langkah diperkirakan/.test(w)), 'closing meluap harus diperingatkan');
  d1.table({ title: 'Tabel padat', caption: 'Tabel 9 · Padat', head: ['A', 'B'], rows: Array.from({ length: 7 }, () => ['teks sel yang panjang sekali '.repeat(4), 'x']) });
  assert(warnings.some((w) => /table: tabel diperkirakan/.test(w)), 'tabel meluap harus diperingatkan');
  ok('closing & tabel yang meluap memicu peringatan');

  // --- deck tepi: section, doughnut, company, save tanpa .pptx ---
  warnings.length = 0;
  const d = T.createDeck({ company: 'PT A&B <c> "d"', footerLabel: 'Uji & <tepi>' });
  d.cover({ title: 'Sampul & <uji>' });
  d.section({ number: '01', title: 'Bab satu' });
  d.bullets({ chapter: '01 · Bab satu', title: 'Poin', bullets: ['satu'] });
  d.section({ number: '02', title: 'Bab satu' });              // judul kembar → section unik
  d.chart({ title: 'Donat', type: 'doughnut', labels: ['A', 'B', 'C'], series: [{ name: 's', values: [5, 3, 2] }] });
  d.chart({ title: 'Garis', type: 'line', labels: ['1', '2', '3'], series: [{ name: 'a', values: [1, 2, 3] }, { name: 'b', values: [3, 2, 1] }], numFmt: '0.0' });
  d.chart({ title: 'Batang', type: 'bar', labels: ['x', 'y'], series: [{ name: 'a', values: [1000.5, 2000] }], highlight: 1 });
  d.chart({ title: 'Donat fokus 1', type: 'doughnut', labels: ['A', 'B', 'C'], series: [{ name: 's', values: [5, 3, 2] }], focus: 1 });
  d.chart({ title: 'Donat tanpa sorotan', type: 'doughnut', labels: ['A', 'B'], series: [{ name: 's', values: [6, 4] }], focus: -1 });
  throws(() => d0.architecture({ title: 'x', caption: 'c', columns: ['a', 'b', 'c', 'd'], boundary: { label: 'b', cols: [1, 2] }, legend: true, source: 's', nodes: [0, 1, 2].map((r) => ({ id: `n${r}`, title: 'N', col: 1, row: r })) }), /tidak muat/, 'arsitektur 3 baris terlalu padat');
  // C4 kontainer 3 baris yang patuh brief (caption + batas + kunci) harus muat
  d.architecture({ title: 'Arsitektur tiga baris padat', caption: 'Diagram kontainer · Uji', boundary: { label: 'Batas', cols: [1, 2] },
    nodes: [{ id: 'u', title: 'Pengguna', kind: 'person', col: 0, row: 0 }, { id: 'w', title: 'Web', col: 1, row: 0 }, { id: 'a', title: 'API', kind: 'main', col: 1, row: 1 },
      { id: 'd', title: 'DB', kind: 'db', col: 1, row: 2 }, { id: 'x', title: 'Luar', kind: 'external', col: 3, row: 1 }],
    links: [{ from: 'u', to: 'w', label: 'Pakai' }, { from: 'w', to: 'a', label: 'JSON' }, { from: 'a', to: 'd', label: 'SQL' }, { from: 'a', to: 'x', label: 'REST' }] });
  // konteks: sistem kami bentang 3 baris dengan orang di 3 baris → tinggi tidak diringkas, garis tetap lurus
  d.architecture({ title: 'Konteks tiga baris', nodes: [0, 1, 2].map((r) => ({ id: `o${r}`, title: `Orang ${r}`, kind: 'person', col: 0, row: r }))
      .concat([{ id: 'sys', title: 'Sistem', kind: 'main', col: 1, colSpan: 2, row: 0, rowSpan: 3 }, { id: 'e0', title: 'Luar', kind: 'external', col: 3, row: 0 }]),
    links: [0, 1, 2].map((r) => ({ from: `o${r}`, to: 'sys', label: 'Memakai' })).concat([{ from: 'sys', to: 'e0', label: 'Kirim' }]) });
  d.pyramid({ title: 'Piramida', layers: ['A', 'B', 'C'] });
  d.funnel({ title: 'Corong dengan nol', stages: [{ name: 'a', value: 10 }, { name: 'b', value: 0 }, { name: 'c', value: 0 }] });
  d.closing({ steps: ['Satu'] });
  d.closing({ steps: ['Dua'] });                               // closing kedua → "Penutup (2)"
  const file = await d.save(path.join(out, 'tepi'));            // tanpa .pptx
  assert.strictEqual(file, path.join(out, 'tepi.pptx'));
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const read = (n) => zip.file(n).async('string');

  const pres = await read('ppt/presentation.xml');
  const sections = [...pres.matchAll(/<p14:section name="([^"]*)"[^>]*>([\s\S]*?)<\/p14:section>/g)].map((m) => [m[1], (m[2].match(/<p14:sldId /g) || []).length]);
  assert(sections.length, 'section harus ada');
  assert(!sections.some(([s]) => /^Default/.test(s)), `tidak boleh ada section Default: ${JSON.stringify(sections)}`);
  assert(sections.every(([, c]) => c > 0), `tiap section berisi slide: ${JSON.stringify(sections)}`);
  assert.deepStrictEqual(sections.map(([s]) => s), ['Pembuka', 'Bab satu', 'Bab satu (2)', 'Penutup', 'Penutup (2)']);
  const all = await Promise.all(Object.keys(zip.files).filter((k) => /^ppt\/slides\/slide\d+\.xml$/.test(k)).sort((x, y) => parseInt(x.match(/\d+/)[0]) - parseInt(y.match(/\d+/)[0])).map(read));
  const bySlideTitle = (t) => all.find((x) => x.replace(/\u00A0/g, ' ').includes(`<a:t>${t}</a:t>`));   // judul memakai spasi tak-putus anti-yatim
  assert(!/NaN|undefined/.test(all.join('')), 'tidak ada teks NaN/undefined');
  const dz = bySlideTitle('Donat tanpa sorotan');
  assert(dz && !/name="grafik-donat-tengah"[\s\S]*?B3121B/.test(dz.split('grafik-donat-tengah')[1]?.slice(0, 900) || ''), 'donat focus -1 tanpa angka merah');
  const py = bySlideTitle('Piramida');
  assert(py && (py.match(/<a:custGeom>/g) || []).length >= 3 && !/prst="trapezoid"/.test(py), 'piramida dari custGeom, bukan trapesium preset');
  const ar = bySlideTitle('Arsitektur tiga baris padat');
  for (const m of ar.matchAll(/name="diagram-aliran-\d+(?:-\d+)?"[\s\S]*?<a:xfrm( flipV="1")?( flipH="1")?[^>]*>/g)) assert(!m[1], 'panah vertikal arsitektur tidak terbalik');
  const cx3 = bySlideTitle('Konteks tiga baris');
  assert(cx3 && !/diagram-aliran-\d+-2"/.test(cx3), 'konteks 3 baris: semua garis lurus (tanpa siku)');
  ok('donat focus 1 / -1, piramida custGeom, arsitektur 3 baris (C4 + konteks) rapi, tanpa NaN/undefined');
  ok(`section PowerPoint: ${sections.map(([s, c]) => `${s}=${c}`).join(', ')}`);

  const app = await read('docProps/app.xml');
  assert(app.includes('<Company>PT A&amp;B &lt;c&gt; &quot;d&quot;</Company>'), 'company ter-escape');
  ok('company ter-escape; save() tanpa .pptx mengembalikan path .pptx');

  const charts = Object.keys(zip.files).filter((k) => /^ppt\/charts\/chart\d+\.xml$/.test(k)).sort();
  const donut = await read(charts[0]);
  assert(/<c:doughnutChart>/.test(donut) && !/<c:showPercent val="1"\/>/.test(donut), 'donat tanpa label di irisan');
  const slides = await Promise.all(Object.keys(zip.files).filter((k) => /^ppt\/slides\/slide\d+\.xml$/.test(k)).map(read));
  assert(slides.some((x) => /name="grafik-donat-tengah"[\s\S]*?<a:t>50%<\/a:t>/.test(x)), 'persen fokus donat (50%) di tengah');
  for (const c of charts) {
    const x = await read(c);
    for (const m of x.matchAll(/formatCode="([^"]*)"/g)) assert(m[1] === 'General' || m[1].startsWith('[$-421]'), `${c}: ${m[1]}`);
  }
  ok('donat: persen fokus di tengah, tanpa label irisan; format angka grafik bertanda id-ID');

  for (const k of Object.keys(zip.files).filter((k) => /^ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(k))) {
    const ids = [...(await read(k)).matchAll(/<p:cNvPr\b[^>]*?\bid="(\d+)"/g)].map((m) => m[1]);
    assert.strictEqual(new Set(ids).size, ids.length, `${k}: id shape ganda ${ids}`);
  }
  ok('id shape unik di setiap slide/layout');

  const theme = await read('ppt/theme/theme1.xml');
  assert(/<a:clrScheme name="Neuron"><a:dk1><a:srgbClr val="111111"\/>/.test(theme) && /<a:accent1><a:srgbClr val="B3121B"\/>/.test(theme));
  ok('tema warna Neuron tertulis tanpa apply_theme.js');

  console.warn = origWarn;
  fs.writeFileSync(path.join(out, 'tepi-warnings.log'), warnings.join('\n'));
  console.log(`  ${n} kelompok uji tepi lulus -> ${file}`);
})().catch((e) => { console.warn = origWarn; console.error('  [GAGAL]', e.message); process.exit(1); });
