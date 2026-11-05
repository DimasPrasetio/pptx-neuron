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
  ok('input wajib kosong/salah memberi galat yang jelas');

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
  ok(`section PowerPoint: ${sections.map(([s, c]) => `${s}=${c}`).join(', ')}`);

  const app = await read('docProps/app.xml');
  assert(app.includes('<Company>PT A&amp;B &lt;c&gt; &quot;d&quot;</Company>'), 'company ter-escape');
  ok('company ter-escape; save() tanpa .pptx mengembalikan path .pptx');

  const charts = Object.keys(zip.files).filter((k) => /^ppt\/charts\/chart\d+\.xml$/.test(k)).sort();
  const donut = await read(charts[0]);
  assert(/<c:doughnutChart>/.test(donut) && /formatCode="\[\$-421\]0%"/.test(donut), 'doughnut memakai 0% (id-ID)');
  for (const c of charts) {
    const x = await read(c);
    for (const m of x.matchAll(/formatCode="([^"]*)"/g)) assert(m[1] === 'General' || m[1].startsWith('[$-421]'), `${c}: ${m[1]}`);
  }
  ok('label persen doughnut 0%; format angka grafik bertanda id-ID');

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
