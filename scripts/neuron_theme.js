// Desain presentasi Neuron (PowerPoint 16:9).
// Hanya DESAIN: warna, huruf, tata letak, logo, motif Tiga Goresan, footer. Isi slide selalu dari pemanggil.
// Token sama persis dengan docx-neuron/scripts/neuron_theme.js; tata letak dari brief desain (riset 2026-11-05,
// revisi skala huruf 2026-11-05: judul aksi 22 pt, isi 14 pt).
//
// Sepuluh jenis slide: cover, agenda, section, bullets, twoColumn, table, chart, image, stats, closing.
const fs = require('fs');
const path = require('path');
// pptxgenjs dibundel di scripts/vendor (tanpa npm install). Bila bundel hilang, jatuh ke paket terpasang.
const loadPptx = () => { try { return require('./vendor/pptxgenjs.bundle.cjs'); } catch (e) { return require('pptxgenjs'); } };
const PptxGenJS = loadPptx();
// jszip ikut di bundel (PptxGenJS.JSZip); pada paket npm biasa, ambil jszip yang terpasang di samping pptxgenjs.
const JSZip = PptxGenJS.JSZip || require(require.resolve('jszip', { paths: [require.resolve('pptxgenjs')] }));

const LOGO_PATH = path.join(__dirname, '..', 'assets', 'logo_neuron.png');
const LOGO_RATIO = 612 / 176;   // 3,477 : 1 — jangan diregangkan; latar logo putih opak → hanya di area putih
// Ruang putih di dalam berkas logo (diukur dari piksel): tinta mulai 4,41% dari kiri, berakhir 8,66% dari kanan,
// 22,73% dari atas, 19,89% dari bawah. Dipakai agar TINTA logo sejajar margin, bukan kotak gambarnya.
const LOGO_PAD = { l: 0.0441, r: 0.0866, t: 0.2273, b: 0.1989 };

// ---------- token desain (sama dengan docx-neuron) ----------
const ACC = 'B3121B';       // merah logo: label bab, Tiga Goresan, sorotan data (maks. 2 elemen merah per slide isi)
const ACC_DARK = '7F0D14';  // turunan B3121B x 0,71: hanya goresan ke-2 motif
const INK = '1F2937';       // judul; latar pembatas bab
const MUTED = '5B6470';     // sekunder: sumber, footer, keterangan
const TEXT = '111111';      // teks isi
const RULE = 'C9CDD3';      // garis tabel/grafik; teks hanya di latar gelap
const HEAD_FILL = '2B3440'; // header tabel
const ZEBRA = 'F4F5F7';     // baris selang-seling, panel
const GREY = '8C939D';      // turunan: seri pembanding grafik (bukan teks)
const WHITE = 'FFFFFF';
const FONT = 'Arial';
const LANG = 'id-ID';
// Urutan warna seri grafik: seri fokus merah, sisanya gelap → abu (maks. 4 seri; lebih → gabung "Lainnya").
// (riset visualisasi V2, model "sorot + netral"): fokus merah → abu → gelap → sedang. B3121B dan 5B6470 hampir sama
// terangnya (rasio 1,16) → tidak boleh bersebelahan. Maks. 4 warna data per grafik.
const SERIES = [ACC, GREY, INK, MUTED];
const NEUTRAL = [INK, GREY, MUTED];   // multi-seri tanpa fokus

// Tema PowerPoint (pemilih warna di PowerPoint menampilkan palet Neuron).
const THEME = {
  name: 'Neuron', headFontFace: FONT, bodyFontFace: FONT,
  colors: { dk1: TEXT, lt1: WHITE, dk2: INK, lt2: ZEBRA, accent1: ACC, accent2: INK, accent3: MUTED, accent4: GREY,
    accent5: HEAD_FILL, accent6: ACC_DARK, hlink: ACC, folHlink: MUTED },
};

// ---------- grid (inci, kanvas 10 x 5,625) ----------
const M = 0.5, W = 9.0, GAP = 0.3;
const Z = { chapterY: 0.45, titleY: 0.72, titleH: 0.75, bodyY: 1.65, bodyEnd: 4.75, sourceY: 4.72, footerY: 5.12 };
// Skala huruf (pt) pada kanvas 10". Setara x1,33 di kanvas lebar 13,33" — judul 22 pt ~ 29 pt, isi 14 pt ~ 19 pt.
// Revisi 2026-11-05: skala v1 (judul 36, isi 16) dinilai pengguna terlalu besar / kurang profesional.
const TYPE = { coverTitle: 30, sectionNum: 40, sectionTitle: 28, title: 22, stat: 40, statOne: 54, keyValue: 32,
  subhead: 15, coverSub: 14, body: 14, key: 14, table: 12, tableDense: 11, small: 11, chart: 10, label: 10,
  meta: 11, source: 9, footer: 9,
  // visualisasi (V3): nama simpul/langkah 12, deskripsi/label 10, tag/label garis 9, nomor lingkaran 11, pita fase 11
  v: { node: 12, desc: 10, tag: 9, num: 11, ribbon: 11, kpi: 24, compare: 24 } };
const col = (n) => n * 0.475 + (n - 1) * 0.3;          // lebar n kolom dari grid 12 kolom
const colX = (start) => M + start * (0.475 + 0.3);      // x awal kolom ke-start (0..11)

// ---------- helper teks ----------
// Angka tidak dipisah dari satuannya di akhir baris ("Rp 101 jt", "14 bln"): spasi diganti spasi tak-putus.
const nb = (s) => String(s).replace(/\bRp (?=\d)/g, 'Rp\u00A0')
  .replace(/(\d) (?=(jt|juta|M|miliar|T|triliun|rb|ribu|bln|bulan|hari|minggu|menit|jam|tahun|thn|unit|orang|%)(?![\p{L}]))/gu, '$1\u00A0')
  .replace(/\b(tahap|langkah|bab|tabel|gambar|slide|fase|Q|kuartal|triwulan|semester) (?=\d)/giu, '$1\u00A0');
// Cegah satu kata yatim di baris terakhir judul: dua kata terakhir disatukan dengan spasi tak-putus.
const noOrphan = (s) => { const str = String(s); const k = str.lastIndexOf(' '); return k > 0 && str.split(' ').length >= 4 && str.length - k <= 14 ? `${str.slice(0, k)}\u00A0${str.slice(k + 1)}` : str; };
// rich('teks **tebal** biasa', o): '**...**' menjadi tebal (sama seperti docx-neuron).
const rich = (s, o = {}) => nb(s).split('**').map((seg, k) => [seg, k % 2 === 1]).filter(([seg]) => seg !== '')
  .map(([seg, bold]) => ({ text: seg, options: { ...o, bold: o.bold || bold } }));
const plain = (s) => String(s).replace(/\*\*/g, '');
const base = (o = {}) => ({ fontFace: FONT, lang: LANG, isTextBox: true, margin: 0, ...o });

// Perkiraan kasar jumlah baris Arial (lebar rata-rata huruf 0,5 em; tebal 0,55 em). Hanya untuk PERINGATAN
// teks meluap — verifikasi akhir tetap lewat render.
const lines = (s, wIn, pt, bold = false) => Math.max(1, Math.ceil((plain(s).length * pt * (bold ? 0.55 : 0.5)) / 72 / wIn));
const warn = (msg) => console.warn(`[neuron_theme] ${msg}`);
function checkFit(kind, texts, wIn, pt, hIn, { spacing = 1.15, paraAfterPt = 8, bold = false } = {}) {
  const n = texts.reduce((a, t) => a + lines(t, wIn, pt, bold), 0);
  const need = (n * pt * 1.2 * spacing + Math.max(0, texts.length - 1) * paraAfterPt) / 72;
  if (need > hIn) warn(`${kind}: teks diperkirakan ${need.toFixed(2)}" tinggi, ruang ${hIn.toFixed(2)}" — persingkat atau pecah slide.`);
}
const textH = (t, wIn, pt, { spacing = 1.15, bold = false } = {}) => (lines(t, wIn, pt, bold) * pt * 1.2 * spacing) / 72;
// Pecah judul 2 baris secara seimbang (hindari satu kata yatim di baris kedua).
function balance(t, wIn, pt) {
  const str = plain(t);
  if (str.includes('\n') || lines(str, wIn, pt, true) !== 2) return t;
  const words = String(t).split(' ');
  let best = null;
  for (let k = 1; k < words.length; k++) {
    const a = words.slice(0, k).join(' '), b = words.slice(k).join(' ');
    if (lines(a, wIn, pt, true) > 1 || lines(b, wIn, pt, true) > 1) continue;
    const d = Math.abs(plain(a).length - plain(b).length);
    if (!best || d < best.d) best = { d, t: `${a}\n${b}` };
  }
  return best ? best.t : t;
}
function checkTitle(kind, t) {
  if (lines(t, W, TYPE.title, true) > 2) warn(`${kind}: judul "${plain(t).slice(0, 40)}…" lebih dari 2 baris — persingkat (± 90 karakter).`);
  if (/\.\s*$/.test(plain(t))) warn(`${kind}: judul diakhiri titik — hapus titiknya.`);
}

// Bullet: satu paragraf per poin. pptxgenjs memulai paragraf baru pada run ber-bullet, jadi hanya run pertama
// poin yang membawa bullet; run terakhir menutup baris.
const bulletRuns = (items, o = {}) => items.flatMap((b, i) => rich(b, o).map((r, k, arr) => ({ text: r.text,
  options: { ...r.options, ...(k === 0 ? { bullet: { indent: 16 } } : {}), ...(k === arr.length - 1 && i < items.length - 1 ? { breakLine: true } : {}) } })));

// Ukuran piksel gambar PNG atau JPEG (dibaca dari header berkas).
function imageInfo(buf, file) {
  if (buf.length > 24 && buf.readUInt32BE(0) === 0x89504e47) return { type: 'png', w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
  if (buf.length > 4 && buf[0] === 0xff && buf[1] === 0xd8) {
    for (let i = 2; i + 9 < buf.length;) {
      if (buf[i] !== 0xff) { i++; continue; }
      const m = buf[i + 1];
      if (m === 0xff) { i++; continue; }   // byte pengisi sebelum penanda
      if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) return { type: 'jpg', h: buf.readUInt16BE(i + 5), w: buf.readUInt16BE(i + 7) };
      i += 2 + buf.readUInt16BE(i + 2);
    }
  }
  throw new Error(`[neuron_theme] image: ${file} bukan PNG/JPEG yang valid.`);
}

// ---------- motif Tiga Goresan (dari simbol logo) ----------
// Tiga jajaran genjang ramping condong ±20° (bawah-kiri -> atas-kanan). Dekoratif; dibuat terakhir (urutan baca).
function strokes(slide, { x, y, h, sw, colors, transparency = [0, 0, 0], name = 'motif-neuron' }) {
  const slant = h * Math.tan((20 * Math.PI) / 180);
  const gap = sw * 0.6;
  colors.forEach((c, i) => {
    slide.addShape('custGeom', {
      x: x + i * (sw + gap), y, w: slant + sw, h, objectName: `${name}-${i + 1}`, altText: '',
      fill: { color: c, transparency: transparency[i] }, line: { type: 'none' },
      points: [{ x: 0, y: h }, { x: slant, y: 0 }, { x: slant + sw, y: 0 }, { x: sw, y: h }, { close: true }],
    });
  });
  return slant + 3 * sw + 2 * gap;   // lebar total
}
const MINI = { h: 0.16, sw: 0.035 };
const miniStrokes = (slide, x, y, name = 'motif-mini') => strokes(slide, { x, y, ...MINI, colors: [ACC, ACC_DARK, INK], name });

// Penyelesaian berkas setelah pptxgenjs menulis (satu kali buka-tutup zip):
// 1. pptxgenjs 4.0.1 menulis <a:pPr> untuk SETIAP run, sehingga paragraf ber-run banyak (mis. '**tebal**' di tengah
//    kalimat) berisi <a:pPr> di tengah paragraf: tidak sah menurut skema dan run berikutnya membawa <a:buNone/> yang
//    menghapus bullet. pPr hanya boleh menjadi anak pertama <a:p>: buang setiap pPr yang mengikuti run/br/fld.
// 2. ID shape (<p:cNvPr id>) dinomori ulang per slide/layout: pptxgenjs bisa memberi id ganda (nomor slide vs objek ke-24).
// 3. Format angka grafik diberi penanda bahasa Indonesia [$-421] agar desimal/ribuan tampil "4,8" / "1.410".
// 3b. Ikon (shape bernama "ikon-…") diberi ujung & sambungan garis bulat seperti Lucide.
// 4. Tema: palet warna, nama tema, dan nama skema huruf Neuron ditulis ke ppt/theme/theme*.xml
//    (pptxgenjs hanya menulis huruf; tanpa ini warna skema = palet Office).
const SLOTS = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
async function finalize(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const pPr = /(<\/a:r>|<a:br\/>|<\/a:fld>)<a:pPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:pPr>)/g;
  const edit = async (name, fn) => { const xml = await zip.file(name).async('string'); const out = fn(xml); if (out !== xml) zip.file(name, out); };
  for (const name of Object.keys(zip.files)) {
    if (/^ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(name)) {
      await edit(name, (xml) => {
        let x = xml, prev;
        do { prev = x; x = x.replace(pPr, '$1'); } while (x !== prev);
        let id = 0;
        x = x.replace(/(<p:cNvPr\b[^>]*?\bid=")\d+(")/g, (m, a, b) => `${a}${++id}${b}`);
        // ikon Lucide: ujung & sambungan garis bulat (pptxgenjs tidak punya opsinya)
        return x.includes('name="ikon-') ? loadIcons().roundIconLines(x) : x;
      });
    } else if (/^ppt\/charts\/chart\d+\.xml$/.test(name)) {
      await edit(name, (xml) => {
        let x = xml.replace(/formatCode="(?!General")(?!\[\$-)([^"]*)"/g, 'formatCode="[$-421]$1"');
        // donat: pptxgenjs memaksa nama kategori + persen di tingkat seri (ditimpa per titik) → matikan langsung
        if (/<c:doughnutChart>/.test(x)) x = x.replace(/<c:showCatName val="1"\/>/g, '<c:showCatName val="0"/>').replace(/<c:showPercent val="1"\/>/g, '<c:showPercent val="0"/>');
        return x;
      });
    } else if (/^ppt\/theme\/theme\d+\.xml$/.test(name)) {
      await edit(name, (xml) => {
        const clr = SLOTS.map((k) => `<a:${k}><a:srgbClr val="${THEME.colors[k]}"/></a:${k}>`).join('');
        if (!/<a:clrScheme\b/.test(xml)) throw new Error(`[neuron_theme] ${name} tanpa <a:clrScheme>; tema tidak bisa ditulis.`);
        return xml.replace(/<a:clrScheme\b[^>]*>[\s\S]*?<\/a:clrScheme>/, `<a:clrScheme name="${THEME.name}">${clr}</a:clrScheme>`)
          .replace(/(<a:theme\b[^>]*?\bname=")[^"]*(")/, `$1${THEME.name}$2`)
          .replace(/(<a:fontScheme\b[^>]*?\bname=")[^"]*(")/, `$1${THEME.name}$2`);
      });
    }
  }
  fs.writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
}
const xmlEsc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const need = (v, what) => { if (v === undefined || v === null || String(v).trim() === '') throw new Error(`[neuron_theme] ${what} wajib diisi.`); return v; };

// Skala sumbu "bagus" (1/2/2,5/5 × 10^n), ± 5 garis bantu.
function niceScale(lo, hi) {
  if (!(hi > lo)) hi = lo + 1;
  const raw = (hi - lo) / 5, p = 10 ** Math.floor(Math.log10(raw));
  const step = [1, 2, 2.5, 5, 10].map((m) => m * p).find((v) => v >= raw);
  return { min: Math.floor(lo / step) * step, max: Math.ceil(hi / step) * step, step };
}
// Angka untuk teks di luar grafik (label ujung, panel) gaya Indonesia, mengikuti jumlah desimal & akhiran "…" numFmt.
// Mendukung bentuk umum: '#,##0', '0.0', '0.0%' (×100), '"Rp "#,##0.0" jt"' (awalan/akhiran literal), '0.0"%"'.
function formatNum(v, numFmt = '#,##0') {
  const f = String(numFmt).replace(/^\[\$-[0-9A-Fa-f]+\]/, '').split(';')[0];
  const prefix = (/^"([^"]*)"/.exec(f) || [, ''])[1];
  const suffix = (/"([^"]*)"\s*$/.exec(f) || [, ''])[1];
  const core = f.replace(/"[^"]*"/g, '');
  const pct = core.includes('%');
  const dec = (/0\.(0+)/.exec(core) || [, ''])[1].length;
  if (!/[0#]/.test(core)) warn(`format angka "${numFmt}" tidak dikenali — label di luar grafik memakai #,##0.`);
  const n = new Intl.NumberFormat('id-ID', { minimumFractionDigits: dec, maximumFractionDigits: dec, useGrouping: core.includes(',') || !/^0+(\.0+)?%?$/.test(core) })
    .format(pct ? v * 100 : v);
  return `${prefix}${n.replace(/^-/, '−')}${pct ? '%' : ''}${suffix}`;   // tanda minus tipografis
}

// Ikon Lucide sebagai bentuk native (scripts/neuron_icons.js). Dimuat saat dipakai; bila modul tidak ada, ikon dilewati.
let iconsMod;
function loadIcons() {
  if (iconsMod === undefined) {
    try { iconsMod = require('./neuron_icons.js'); } catch (e) { iconsMod = null; warn('ikon: scripts/neuron_icons.js tidak ditemukan — ikon dilewati.'); }
  }
  return iconsMod;
}

// ---------- deck ----------
function createDeck({ title = 'Presentasi Neuron', author = 'Neuronworks Indonesia', company = 'Neuronworks Indonesia',
  footerLabel = '' } = {}) {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_16x9';
  pres.title = title; pres.author = author;
  pres.company = xmlEsc(company);   // ditulis tanpa escape oleh pptxgenjs
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };

  const titlePh = (o) => ({ placeholder: { options: { name: 'title', type: 'title', fontFace: FONT, bold: true, color: INK,
    valign: 'top', align: 'left', margin: 0, lang: LANG, lineSpacingMultiple: 0.9, ...o }, text: '' } });
  // logoAt({ left | right, top | middle | bottom, w }): posisi berdasarkan tepi TINTA logo.
  const logoAt = ({ left, right, top, middle, bottom, w }) => {
    const h = w / LOGO_RATIO;
    const x = left !== undefined ? left - LOGO_PAD.l * w : right - w + LOGO_PAD.r * w;
    const y = top !== undefined ? top - LOGO_PAD.t * h
      : bottom !== undefined ? bottom - h + LOGO_PAD.b * h : middle - ((LOGO_PAD.t + 1 - LOGO_PAD.b) / 2) * h;
    return { image: { path: LOGO_PATH, x, y, w, h, altText: 'Logo Neuron' } };
  };
  const footerText = footerLabel ? [{ text: { text: footerLabel, options: base({ x: M, y: Z.footerY, w: 5.5, h: 0.22, fontSize: TYPE.footer, color: MUTED, valign: 'middle' }) } }] : [];
  const slideNo = (color) => ({ x: 7.75, y: Z.footerY, w: 0.6, h: 0.22, fontFace: FONT, fontSize: TYPE.footer, color, align: 'right' });

  // Layout (pptxgenjs menyebutnya "slide master"; di PowerPoint menjadi layout di bawah satu master).
  pres.defineSlideMaster({ title: 'NEURON_SAMPUL', background: { color: WHITE },
    objects: [logoAt({ left: M, top: M, w: 1.9 }), titlePh({ x: M, y: 2.1, w: 6.4, h: 1.0, fontSize: TYPE.coverTitle })] });
  pres.defineSlideMaster({ title: 'NEURON_ISI', background: { color: WHITE }, margin: [0.5, 0.5, 0.75, 0.5],
    slideNumber: slideNo(MUTED),
    objects: [titlePh({ x: M, y: Z.titleY, w: W, h: Z.titleH, fontSize: TYPE.title }), ...footerText,
      logoAt({ right: M + W, middle: Z.footerY + 0.11, w: 0.95 })] });
  // Pembatas bab: latar gelap, tanpa logo (logo berlatar putih opak), tanpa label dokumen.
  pres.defineSlideMaster({ title: 'NEURON_BAB', background: { color: INK }, slideNumber: { ...slideNo(RULE), x: M, align: 'left', margin: [0, 0, 0, 0] },
    objects: [titlePh({ x: M, y: 2.35, w: 6.2, h: 1.0, fontSize: TYPE.sectionTitle, color: WHITE })] });
  // Penutup: latar putih, judul di posisi slide isi, logo besar kanan bawah, tanpa footer.
  pres.defineSlideMaster({ title: 'NEURON_PENUTUP', background: { color: WHITE },
    objects: [titlePh({ x: M, y: Z.titleY, w: W, h: Z.titleH, fontSize: TYPE.title }),
      logoAt({ left: colX(8), bottom: 5.0, w: 2.0 })] });

  // Bagian PowerPoint (sections): slide pertama masuk "Pembuka"; section() membuka bagian baru; closing() "Penutup".
  // pptxgenjs menaruh slide tanpa sectionTitle di section baru "Default-N", jadi judul section aktif selalu diteruskan.
  // Judul section harus unik (pptxgenjs mencari menurut judul): judul kembar diberi akhiran " (2)", " (3)", …
  const sectionNames = new Set();
  let current = null;
  const openSection = (t) => {
    let name = plain(t).trim() || 'Bagian', k = 1;
    while (sectionNames.has(name)) name = `${plain(t).trim() || 'Bagian'} (${++k})`;
    sectionNames.add(name); pres.addSection({ title: name }); current = name;
  };
  const add = (masterName, notes) => {
    if (!current) openSection('Pembuka');
    const s = pres.addSlide({ masterName, sectionTitle: current });
    if (notes) s.addNotes(notes);
    return s;
  };
  // Kepala slide isi: label bab (merah + goresan mini) dan judul aksi di placeholder judul.
  const head = (s, kind, chapter, t) => {
    need(t, `${kind}: title`);
    checkTitle(kind, t);
    if (chapter) s.addText(chapter, base({ x: M + 0.24, y: Z.chapterY, w: 6, h: 0.22, fontSize: TYPE.label, bold: true, color: ACC,
      valign: 'middle', objectName: 'label-bab' }));
    s.addText(noOrphan(nb(t)), { placeholder: 'title' });
    s._mini = !!chapter;
  };
  const finish = (s) => { if (s._mini) miniStrokes(s, M, Z.chapterY + 0.03); return s; };
  const source = (s, text, y = Z.sourceY) => {
    if (text) s.addText(text, base({ x: M, y, w: W, h: 0.2, fontSize: TYPE.source, color: MUTED, valign: 'top', objectName: 'sumber' }));
  };
  const bottomOf = (src) => (src ? Z.sourceY - 0.2 : Z.bodyEnd);
  const bulletBox = (s, kind, items, x, y, w, h, o = {}) => {
    if (items.length > 5) warn(`${kind}: ${items.length} poin; maksimal 5 per blok — pecah slide.`);
    checkFit(kind, items, w - 0.25, o.fontSize || TYPE.body, h, { paraAfterPt: o.paraSpaceAfter ?? 8 });
    s.addText(bulletRuns(items), base({ x, y, w, h, fontSize: TYPE.body, color: TEXT, valign: 'top', paraSpaceAfter: 8,
      lineSpacingMultiple: 1.15, ...o }));
  };

  const visuals = require('./neuron_visuals.js').build({ add, head, finish, source, bottomOf, base, rich, nb, plain, lines, textH, warn, need, bulletBox,
    T: { ACC, INK, MUTED, TEXT, RULE, ZEBRA, GREY, WHITE }, TYPE, M, W, Z, GAP, col, colX, formatNum });
  const api = {
    pres,
    // 1. Sampul — "Apa ini, untuk siapa, dari siapa?"
    cover({ type = '', title: t, subtitle = '', meta = '', classification = '', notes } = {}) {
      need(t, 'cover: title');
      const s = add('NEURON_SAMPUL', notes);
      if (lines(t, 6.4, TYPE.coverTitle, true) > 2) warn(`cover: judul lebih dari 2 baris — persingkat.`);
      if (type) s.addText(String(type).toUpperCase(), base({ x: M, y: 1.8, w: 6.4, h: 0.22, fontSize: TYPE.label, bold: true, color: ACC, charSpacing: 1.5, valign: 'middle', objectName: 'jenis-dokumen' }));
      s.addText(nb(balance(t, 6.4, TYPE.coverTitle)), { placeholder: 'title' });
      if (subtitle) s.addText(subtitle, base({ x: M, y: 3.25, w: 6.0, h: 0.5, fontSize: TYPE.coverSub, color: MUTED, valign: 'top', objectName: 'subjudul' }));
      if (meta) s.addText(meta, base({ x: M, y: 4.45, w: 5.9, h: 0.22, fontSize: TYPE.meta, color: INK, valign: 'top', objectName: 'penyusun' }));
      if (classification) s.addText(classification, base({ x: M, y: 4.75, w: 5.9, h: 0.2, fontSize: TYPE.source, color: MUTED, valign: 'top', objectName: 'klasifikasi' }));
      strokes(s, { x: 7.35, y: 1.0, h: 3.6, sw: 0.16, colors: [ACC, ACC_DARK, INK], name: 'motif-sampul' });
      return s;
    },

    // 2. Agenda — "Apa saja yang dibahas dan sekarang di mana?" items: string | { text, note }; active: indeks butir aktif.
    agenda({ chapter = 'Agenda', title: t = 'Yang akan dibahas', items = [], active, notes } = {}) {
      if (items.length > 6) warn(`agenda: ${items.length} butir; maksimal 6.`);
      const s = add('NEURON_ISI', notes);
      head(s, 'agenda', chapter, t);
      const rowH = 0.46;
      items.forEach((it, i) => {
        const { text, note } = typeof it === 'string' ? { text: it } : it || {};
        need(text, `agenda: items[${i}].text`);
        if (lines(text, 6.2, TYPE.body, true) > 1) warn(`agenda: butir ${i + 1} lebih dari 1 baris — persingkat.`);
        const y = Z.bodyY + i * rowH, on = active === i, dim = active !== undefined && active !== null && !on;
        s.addText(String(i + 1).padStart(2, '0'), base({ x: M, y, w: 0.6, h: rowH, fontSize: TYPE.subhead, bold: true, color: dim ? MUTED : INK, valign: 'middle', objectName: `agenda-no-${i + 1}` }));
        s.addText(plain(text), base({ x: M + 0.95, y, w: 6.2, h: rowH, fontSize: TYPE.body, bold: on, color: dim ? MUTED : TEXT, valign: 'middle', objectName: `agenda-${i + 1}` }));
        if (note) s.addText(note, base({ x: M + 7.3, y, w: 1.7, h: rowH, fontSize: TYPE.small, color: MUTED, align: 'right', valign: 'middle', objectName: `agenda-ket-${i + 1}` }));
      });
      if (active !== undefined && items[active] !== undefined) miniStrokes(s, M + 0.5, Z.bodyY + active * rowH + (rowH - MINI.h) / 2, 'motif-agenda');
      return finish(s);
    },

    // 3. Pembatas bab — "Bagian baru dimulai: apa?" Latar gelap, nomor besar, goresan merah memudar.
    section({ number = '', title: t, description = '', notes } = {}) {
      need(t, 'section: title');
      openSection(t);
      const s = add('NEURON_BAB', notes);
      if (number !== '') s.addText(String(number), base({ x: M, y: 1.35, w: 3, h: 0.8, fontSize: TYPE.sectionNum, bold: true, color: WHITE, valign: 'bottom', objectName: 'nomor-bab' }));
      s.addText(nb(balance(t, 6.2, TYPE.sectionTitle)), { placeholder: 'title' });
      // deskripsi menempel 0,25" di bawah judul (1 atau 2 baris), bukan di posisi tetap
      const dy = 2.35 + Math.min(2, lines(t, 6.2, TYPE.sectionTitle, true)) * TYPE.sectionTitle * 0.9 * 1.2 / 72 + 0.25;
      if (description) s.addText(description, base({ x: M, y: dy, w: 6.0, h: 0.7, fontSize: TYPE.body, color: RULE, valign: 'top', lineSpacingMultiple: 1.15, objectName: 'deskripsi-bab' }));
      strokes(s, { x: 7.35, y: 1.0, h: 3.6, sw: 0.16, colors: [ACC, ACC, ACC], transparency: [0, 35, 60], name: 'motif-bab' });
      return s;
    },

    // 4. Isi poin — "Apa kesimpulannya dan apa 3–5 alasannya?" Panel Inti kanan opsional.
    bullets({ chapter, title: t, bullets = [], keyPoint, source: src, notes } = {}) {
      const s = add('NEURON_ISI', notes);
      head(s, 'bullets', chapter, t);
      const bottom = bottomOf(src);
      bulletBox(s, 'bullets', bullets, M, Z.bodyY, keyPoint ? col(8) : W, bottom - Z.bodyY, { objectName: 'poin' });
      if (keyPoint) {
        const px = colX(8), pw = col(4);
        checkFit('bullets.keyPoint', [keyPoint], pw - 0.44, TYPE.key, bottom - Z.bodyY - 0.75, { bold: true });
        // tinggi panel mengikuti isi terpanjang (poin atau Inti), minimal 1,6", agar tidak menyisakan rongga
        const estB = bullets.reduce((a, b) => a + textH(b, col(8) - 0.25, TYPE.body) + 8 / 72, 0);
        const estK = 0.44 + 0.3 + textH(keyPoint, pw - 0.44, TYPE.key, { bold: true });
        const ph = Math.min(bottom - Z.bodyY, Math.max(1.6, estB, estK));
        s.addShape('rect', { x: px, y: Z.bodyY, w: pw, h: ph, fill: { color: ZEBRA }, line: { type: 'none' }, objectName: 'panel-inti' });
        s.addText([{ text: 'Inti', options: { fontSize: TYPE.label, bold: true, color: MUTED, breakLine: true } },
          ...rich(keyPoint, { fontSize: TYPE.key, bold: true, color: INK })],
          base({ x: px + 0.22, y: Z.bodyY + 0.22, w: pw - 0.44, h: ph - 0.44, valign: 'top', paraSpaceAfter: 6, lineSpacingMultiple: 1.15, objectName: 'inti' }));
      }
      source(s, src);
      return finish(s);
    },

    // 5. Dua kolom — "Apa bedanya A dan B?" left/right: { heading, bullets }; recommend: 'left' | 'right'.
    twoColumn({ chapter, title: t, left = {}, right = {}, recommend, source: src, notes } = {}) {
      const s = add('NEURON_ISI', notes);
      head(s, 'twoColumn', chapter, t);
      const bottom = bottomOf(src), cw = col(6);
      [['left', left, M], ['right', right, colX(6)]].forEach(([side, c, x]) => {
        const rec = recommend === side, pad = rec ? 0.22 : 0;
        if (rec) s.addShape('rect', { x, y: Z.bodyY, w: cw, h: bottom - Z.bodyY, fill: { color: ZEBRA }, line: { type: 'none' }, objectName: `panel-${side}` });
        if (rec) s.addText('Rekomendasi', base({ x: x + pad, y: Z.bodyY + 0.22, w: cw - 2 * pad, h: 0.2, fontSize: TYPE.label, bold: true, color: ACC, valign: 'middle', objectName: `rekomendasi-${side}` }));
        // subjudul kedua kolom sejajar: kolom tanpa panel ikut turun setinggi label "Rekomendasi"
        let y = Z.bodyY + (recommend ? 0.48 : 0);
        if (c.heading) { s.addText(c.heading, base({ x: x + pad, y, w: cw - 2 * pad, h: 0.3, fontSize: TYPE.subhead, bold: true, color: INK, valign: 'top', objectName: `subjudul-${side}` })); y += 0.45; }
        bulletBox(s, `twoColumn.${side}`, c.bullets || [], x + pad, y, cw - 2 * pad, bottom - y - pad, { objectName: `poin-${side}` });
      });
      source(s, src);
      return finish(s);
    },

    // 6. Tabel — "Bagaimana perbandingan angka/atribut ini?" Header 2B3440, zebra F4F5F7, garis horizontal saja.
    table({ chapter, title: t, caption, head: hd, rows = [], align = [], widths, source: src, notes } = {}) {
      const maxRows = caption ? 7 : 8;
      if (rows.length > maxRows) warn(`table: ${rows.length} baris; maksimal ${maxRows} per slide — pecah ke slide "(lanjutan)" dengan header yang sama.`);
      if (widths && Math.abs(widths.reduce((a, b) => a + b, 0) - W) > 0.01) warn(`table: jumlah lebar kolom ${widths.reduce((a, b) => a + b, 0)}", seharusnya ${W}".`);
      if (!Array.isArray(hd) || !hd.length) throw new Error('[neuron_theme] table: head wajib diisi.');
      if (widths && widths.length !== hd.length) throw new Error(`[neuron_theme] table: ${hd.length} kolom header tetapi ${widths.length} lebar.`);
      rows.forEach((r, i) => { if (!Array.isArray(r) || r.length !== hd.length) throw new Error(`[neuron_theme] table: baris ${i + 1} berisi ${Array.isArray(r) ? r.length : 0} sel, header ${hd.length} kolom.`); });
      const s = add('NEURON_ISI', notes);
      head(s, 'table', chapter, t);
      let y = Z.bodyY;
      if (caption) {
        const [label, ...rest] = String(caption).split(' · ');
        s.addText([{ text: label + (rest.length ? ' · ' : ''), options: { bold: true } }, { text: rest.join(' · ') }],
          base({ x: M, y, w: W, h: 0.22, fontSize: TYPE.label, color: INK, valign: 'middle', objectName: 'judul-tabel' }));
        y += 0.3;
      }
      const fsz = hd.length > 5 || rows.length > 6 ? TYPE.tableDense : TYPE.table;
      const hair = { type: 'solid', pt: 0.5, color: RULE }, none = { type: 'none' };
      const al = (i) => (align[i] === 'r' ? 'right' : align[i] === 'c' ? 'center' : 'left');
      const cell = (txt, i, isHead, zebra) => ({
        text: isHead ? plain(txt) : rich(txt, {}),
        options: { fontFace: FONT, lang: LANG, fontSize: fsz, bold: isHead, color: isHead ? WHITE : TEXT, align: al(i),
          valign: 'middle', fill: { color: isHead ? HEAD_FILL : zebra ? ZEBRA : WHITE }, margin: [0, 0.1, 0, 0.1],
          border: [isHead ? none : hair, none, hair, none] },
      });
      const rowH = 0.32, colW = widths || hd.map(() => W / hd.length);
      // tinggi perkiraan tiap baris (sel yang membungkus menambah tinggi) → peringatan meluap + letak sumber
      const estRow = (r, bold) => Math.max(rowH, ...r.map((c, i) => textH(String(c ?? ''), colW[i] - 0.2, fsz, { spacing: 1, bold }) + 0.1));
      const tableH = estRow(hd, true) + rows.reduce((a, r) => a + estRow(r, false), 0);
      if (y + tableH > bottomOf(src) + 0.01) warn(`table: tabel diperkirakan ${tableH.toFixed(2)}" tinggi, ruang ${(bottomOf(src) - y).toFixed(2)}" — kurangi baris/teks sel atau pecah slide.`);
      s.addTable([hd.map((h, i) => cell(h, i, true)), ...rows.map((r, ri) => r.map((c, i) => cell(c ?? '', i, false, ri % 2 === 1)))],
        { x: M, y, w: W, colW, rowH, autoPage: false, objectName: 'tabel' });
      source(s, src, Math.min(Z.sourceY, y + tableH + 0.12));   // sumber menempel di bawah tabel
      return finish(s);
    },

    // 7. Grafik — "Tren/perbandingan apa yang harus dilihat?" Grafik native (bisa diedit di PowerPoint).
    // type: 'column' | 'bar' | 'line' | 'doughnut' | 'stacked100'. Model "sorot + netral": satu merah per slide.
    // series: [{ name, values, dashed?, target? }]; focus: indeks seri merah (bawaan seri 0 bila > 1 seri; -1 = tanpa sorotan);
    // highlight (satu seri batang/kolom): indeks KATEGORI merah (yang disebut di judul), batang lain abu.
    // column + seri { target: true } = kolom + garis target putus-putus berlabel (satu sumbu, satuan sama).
    // keyValue/keyText: panel angka kanan (column/bar/line). numFmt: format angka (tampil gaya Indonesia).
    chart({ chapter, title: t, type = 'column', labels = [], series = [], focus, highlight, keyValue, keyText, numFmt = '#,##0',
      showValue = true, source: src, notes } = {}) {
      const TYPES = ['column', 'bar', 'line', 'doughnut', 'stacked100'];
      if (!TYPES.includes(type)) throw new Error(`[neuron_theme] chart: type '${type}' tidak dikenal (${TYPES.join(', ')}).`);
      if (!Array.isArray(series) || !series.length) throw new Error('[neuron_theme] chart: series wajib berisi minimal 1 seri.');
      series.forEach((sr, i) => { if (!sr || !Array.isArray(sr.values) || sr.values.length !== labels.length) throw new Error(`[neuron_theme] chart: series[${i}].values harus berisi ${labels.length} nilai (sejumlah labels).`); });
      const targets = type === 'column' ? series.filter((sr) => sr.target) : [];
      const main = series.filter((sr) => !targets.includes(sr));
      // focus = indeks seri; untuk donat = indeks irisan (kategori)
      const fMax = type === 'doughnut' ? labels.length : main.length;
      if (focus !== undefined && focus !== -1 && !(Number.isInteger(focus) && focus >= 0 && focus < fMax))
        throw new Error(`[neuron_theme] chart: focus ${focus} di luar rentang 0..${fMax - 1} (atau -1 tanpa sorotan).`);
      if (main.length > 4) warn(`chart: ${main.length} seri; maksimal 4 — gabung ke "Lainnya" atau pecah grafik.`);
      if (type === 'column' && main.length > 3) warn(`chart: kolom berkelompok ${main.length} seri; maksimal 3.`);
      const s = add('NEURON_ISI', notes);
      head(s, 'chart', chapter, t);
      const bottom = bottomOf(src);
      const fmt = (v) => formatNum(v, numFmt);
      const axis = { color: MUTED, fontFace: FONT, fontSize: TYPE.chart };
      const common = {
        objectName: 'grafik', altText: plain(t),
        catAxisLabelColor: axis.color, catAxisLabelFontFace: axis.fontFace, catAxisLabelFontSize: axis.fontSize,
        valAxisLabelColor: axis.color, valAxisLabelFontFace: axis.fontFace, valAxisLabelFontSize: axis.fontSize,
        catAxisLineShow: true, catAxisLineColor: RULE, catAxisLineSize: 0.75, valAxisLineShow: false, valAxisLabelFormatCode: numFmt,
        valGridLine: { color: RULE, size: 0.5 }, catGridLine: { style: 'none' },
        catAxisMajorTickMark: 'none', valAxisMajorTickMark: 'none', catAxisMinorTickMark: 'none', valAxisMinorTickMark: 'none',
        showLegend: false, legendFontFace: FONT, legendFontSize: TYPE.chart, legendColor: MUTED, legendPos: 't',
        dataLabelFormatCode: numFmt, dataLabelColor: INK, dataLabelFontFace: FONT, dataLabelFontSize: TYPE.chart,
      };
      // legenda manual (kotak 0,12" + nama) — urutannya pasti sama dengan urutan seri
      const legend = (items, y) => {
        let x = M;
        items.forEach(([name, color], i) => {
          s.addShape('rect', { x, y: y + 0.05, w: 0.12, h: 0.12, fill: { color }, line: { type: 'none' }, objectName: `grafik-legenda-${i + 1}` });
          s.addText(plain(name), base({ x: x + 0.18, y, w: 2.5, h: 0.22, fontSize: TYPE.chart, color: MUTED, valign: 'middle', objectName: `grafik-legenda-${i + 1}-teks` }));
          x += 0.18 + plain(name).length * 0.07 + 0.3;
        });
      };

      // --- donat: 2–3 bagian, irisan fokus merah mulai jam 12, persen fokus di tengah, daftar label di kanan ---
      if (type === 'doughnut') {
        const vals = main[0].values, total = vals.reduce((a, b) => a + b, 0);
        if (labels.length > 3) warn(`chart: donat ${labels.length} bagian; maksimal 3 — gabung "Lainnya" atau pakai batang.`);
        if (!(total > 0)) throw new Error('[neuron_theme] chart: donat butuh nilai positif.');
        const fi = focus === -1 ? -1 : focus !== undefined ? focus : 0;
        const order = labels.map((_, i) => i).sort((a, b) => (a === fi ? -1 : b === fi ? 1 : vals[b] - vals[a]));   // fokus pertama, lalu besar→kecil
        const col3 = fi >= 0 ? { [order[0]]: ACC, [order[1]]: GREY, [order[2]]: INK } : { [order[0]]: INK, [order[1]]: GREY, [order[2]]: MUTED };
        const ci = fi >= 0 ? fi : order[0];   // angka tengah: irisan fokus, atau irisan terbesar bila tanpa sorotan
        const d = Math.min(2.6, bottom - Z.bodyY - 0.1), cx = M + 0.3, cy = Z.bodyY + 0.05;
        s.addChart('doughnut', [{ name: main[0].name, labels: order.map((i) => labels[i]), values: order.map((i) => vals[i]) }], {
          ...common, x: cx, y: cy, w: d, h: d, holeSize: 65, firstSliceAng: 0, showValue: false, showPercent: false, showLabel: false,
          chartColors: order.map((i) => col3[i] || MUTED), dataBorder: { pt: 1, color: 'FFFFFF' } });
        const pct = (v) => `${formatNum((v / total) * 100, '0')}%`;
        s.addText(pct(vals[ci]), base({ x: cx + d * 0.2, y: cy + d / 2 - 0.25, w: d * 0.6, h: 0.5, fontSize: TYPE.v.kpi, bold: true, color: fi >= 0 ? ACC : INK, align: 'center', valign: 'middle', objectName: 'grafik-donat-tengah' }));
        const lx = cx + d + 0.5;
        order.forEach((i, k) => {
          const y = cy + 0.35 + k * 0.4;
          s.addShape('rect', { x: lx, y: y + 0.06, w: 0.12, h: 0.12, fill: { color: col3[i] || MUTED }, line: { type: 'none' }, objectName: `grafik-donat-kunci-${k + 1}` });
          s.addText([{ text: plain(labels[i]), options: { bold: i === ci } }], base({ x: lx + 0.22, y, w: 2.4, h: 0.24, fontSize: TYPE.v.node, color: TEXT, valign: 'middle', objectName: `grafik-donat-label-${k + 1}` }));
          s.addText(pct(vals[i]), base({ x: lx + 2.6, y, w: 0.7, h: 0.24, fontSize: TYPE.v.node, bold: i === ci, color: TEXT, align: 'right', valign: 'middle', objectName: `grafik-donat-persen-${k + 1}` }));
        });
        if (keyText) {
          const ky = cy + 0.35 + order.length * 0.4 + 0.25, kw = M + W - lx;
          checkFit('chart.keyText', [keyText], kw, TYPE.body, bottom - ky);
          s.addText(rich(keyText), base({ x: lx, y: ky, w: kw, h: bottom - ky, fontSize: TYPE.body, color: TEXT, valign: 'top', lineSpacingMultiple: 1.15, objectName: 'penjelasan' }));
        }
        source(s, src);
        return finish(s);
      }

      // --- tumpuk 100% horizontal: segmen fokus menempel kiri, label % di dalam segmen, legenda atas ---
      if (type === 'stacked100') {
        if (labels.length > 5) warn(`chart: tumpuk 100% ${labels.length} batang; maksimal 5.`);
        const fi = focus === -1 ? -1 : focus !== undefined ? focus : 0;
        if (fi > 0) warn('chart: tumpuk 100% — letakkan segmen fokus sebagai seri pertama (menempel sumbu kiri).');
        if (fi < 0 && main.length > 3) warn('chart: tumpuk 100% tanpa fokus maks. 3 segmen berwarna berbeda — segmen ke-4 memakai C9CDD3.');
        const segCol = main.map((_, i) => (i === fi ? ACC : fi >= 0 ? [GREY, INK, MUTED][(i > fi ? i - 1 : i) % 3] : [INK, GREY, MUTED, RULE][i % 4]));
        legend(main.map((sr, i) => [sr.name, segCol[i]]), Z.bodyY);
        const box = { x: M, y: Z.bodyY + 0.35, w: W, h: bottom - Z.bodyY - 0.35 };
        const layout = { x: 1.5 / W, y: 0.02, w: 1 - 1.5 / W - 0.01, h: 0.96 };
        s.addChart('bar', main.map((sr) => ({ name: sr.name, labels: [...labels].reverse(), values: [...sr.values].reverse() })), {
          ...common, ...box, layout, barDir: 'bar', barGrouping: 'percentStacked', barGapWidthPct: 50, chartColors: segCol,
          valAxisHidden: true, valGridLine: { style: 'none' }, catAxisLineShow: false, showValue: false, dataBorder: { pt: 1, color: 'FFFFFF' } });
        // label % manual (warna per segmen: putih di merah/gelap, 1F2937 di abu — kontras ≥ 4,5)
        const px = box.x + layout.x * box.w, pw = layout.w * box.w, band = (layout.h * box.h) / labels.length, py = box.y + layout.y * box.h;
        labels.forEach((_, r) => {
          const sum = main.reduce((a, sr) => a + (sr.values[r] || 0), 0);
          let acc = 0;
          main.forEach((sr, i) => {
            const p = sum ? (sr.values[r] || 0) / sum : 0, w = p * pw;
            // di atas abu 8C939D teks gelap minimal 12 pt tebal (kontras teks besar); segmen lain 10 pt putih
            if (p >= 0.08 && w >= 0.45) s.addText(`${formatNum(p * 100, '0')}%`, base({ x: px + acc * pw, y: py + r * band, w, h: band, fontSize: segCol[i] === GREY || segCol[i] === RULE ? 12 : TYPE.chart, bold: true,
              color: segCol[i] === GREY || segCol[i] === RULE ? INK : WHITE, align: 'center', valign: 'middle', objectName: `grafik-segmen-${r + 1}-${i + 1}` }));
            acc += p;
          });
        });
        source(s, src);
        return finish(s);
      }

      // --- kolom / batang / garis ---
      const panel = keyValue !== undefined || keyText;
      const cw = panel ? col(8) : W;
      const f = focus !== undefined ? focus : main.length > 1 ? 0 : -1;
      const pointHi = highlight !== undefined && highlight !== -1 && main.length === 1 && (type === 'column' || type === 'bar');
      if (highlight !== undefined && highlight !== -1 && !pointHi) warn('chart: highlight hanya untuk grafik batang/kolom satu seri; pakai focus untuk menyorot seri.');
      if (highlight === undefined && main.length === 1 && (type === 'column' || type === 'bar')) warn('chart: batang/kolom satu seri tanpa highlight — sorot kategori yang disebut judul (highlight: i) atau tulis highlight: -1.');
      if (pointHi && !(Number.isInteger(highlight) && highlight >= 0 && highlight < labels.length)) throw new Error(`[neuron_theme] chart: highlight ${highlight} di luar rentang kategori.`);
      // warna: seri fokus merah, lainnya netral berurutan; satu seri tanpa fokus = 5B6470; satu seri + highlight = abu + 1 merah
      const colors = f >= 0 ? main.map((_, i) => (i === f ? ACC : SERIES.slice(1)[(i > f ? i - 1 : i) % 3]))
        : main.length === 1 ? [MUTED] : main.map((_, i) => NEUTRAL[i % 3]);
      const redInChart = pointHi || f >= 0;   // satu sorotan merah per slide → angka panel gelap bila grafik sudah merah
      // batang horizontal: kategori pertama digambar di bawah → dibalik agar urutan baca atas→bawah sama dengan data
      const rev = type === 'bar';
      const lab = rev ? [...labels].reverse() : labels;
      const data = main.map((sr) => ({ name: sr.name, labels: lab, values: rev ? [...sr.values].reverse() : sr.values }));
      const all = series.flatMap((sr) => sr.values.filter((v) => typeof v === 'number'));
      const scale = niceScale(Math.min(0, ...all), Math.max(0, ...all));
      const lockLabels = type === 'line' || targets.length > 0;   // label ujung (garis / target) butuh area plot terkunci
      const multiLegend = type !== 'line' && main.length > 1;
      if (multiLegend) legend(main.map((sr, i) => [sr.name, colors[i]]), Z.bodyY);
      const box = { x: M, y: Z.bodyY + (multiLegend ? 0.35 : 0), w: cw, h: bottom - Z.bodyY - (multiLegend ? 0.35 : 0) };
      const layout = lockLabels ? { x: 0.08, y: 0.06, w: Math.max(0.5, (cw - 1.3) / cw - 0.08), h: 0.82 } : undefined;
      const manyBars = labels.length * main.length > 12;
      const opts = {
        ...common, ...box, ...(layout ? { layout } : {}),
        valAxisMinVal: scale.min, valAxisMaxVal: scale.max, valAxisMajorUnit: scale.step, lineDataSymbol: 'none',
      };
      const barOpts = {
        barDir: type === 'bar' ? 'bar' : 'col', barGapWidthPct: 60, catAxisLabelPos: 'low',   // label kategori tidak tertimpa batang negatif
        chartColors: pointHi ? lab.map((_, i) => (i === (rev ? labels.length - 1 - highlight : highlight) ? ACC : GREY)) : colors,
        // label nilai di ujung batang; sumbu nilai disembunyikan agar tidak dobel (kecuali batang terlalu banyak)
        showValue: showValue && !manyBars, dataLabelPosition: 'outEnd',
      };
      const hideAxis = barOpts.showValue;
      if (type === 'line') {
        // garis: satu tipe per seri (combo) agar tebal/putus-putus per seri — fokus 2,25 pt, lainnya 1,5 pt, dashed (target) 1 pt
        s.addChart(main.map((sr, i) => ({ type: 'line', data: [data[i]], options: { chartColors: [sr.dashed ? MUTED : colors[i]],
          lineSize: i === f ? 2.25 : sr.dashed ? 1 : 1.5, lineDash: sr.dashed ? 'dash' : 'solid', lineDataSymbol: 'none' } })), opts);
        visuals.lineEndLabels(s, { ...box, layout, min: scale.min, max: scale.max, series: main, focus: f, fmt });
      } else if (targets.length) {
        // garis target digambar lebih dulu (di belakang batang); sumbu nilai disembunyikan; nilai batang ditulis manual
        // di atas batang, kecuali yang akan bertabrakan dengan garis target (batang sorotan tetap berlabel, di dalam batang bila perlu)
        s.addChart([...targets.map((tg) => ({ type: 'line', data: [{ name: tg.name, labels: lab, values: tg.values }], options: { chartColors: [MUTED], lineSize: 1, lineDash: 'dash', lineDataSymbol: 'none' } })),
          { type: 'bar', data, options: { ...barOpts, showValue: false } }],
          { ...opts, valAxisHidden: true, valGridLine: { style: 'none' } });
        visuals.lineEndLabels(s, { ...box, layout, min: scale.min, max: scale.max, series: targets, focus: -1, fmt });
        const ppx = box.x + layout.x * box.w, ppw = layout.w * box.w, ppy = box.y + layout.y * box.h, pph = layout.h * box.h, band = ppw / labels.length;
        const yOf = (v) => ppy + (1 - (v - scale.min) / (scale.max - scale.min)) * pph;
        main[0].values.forEach((v, k) => {
          if (typeof v !== 'number') return;
          const top = yOf(v), hit = targets.some((tg) => Math.abs(yOf(tg.values[k]) - (top - 0.12)) < 0.16), hi = pointHi && k === highlight;
          if (hit && !hi) return;
          const inside = hit && hi;
          s.addText(fmt(v), base({ x: ppx + k * band, y: inside ? top + 0.04 : top - 0.24, w: band, h: 0.2, fontSize: TYPE.chart, bold: hi,
            color: inside ? WHITE : INK, align: 'center', valign: inside ? 'top' : 'bottom', objectName: `grafik-nilai-${k + 1}` }));
        });
      } else {
        s.addChart('bar', data, { ...opts, ...barOpts, valAxisHidden: hideAxis, ...(hideAxis ? { valGridLine: { style: 'none' } } : {}) });
      }
      if (panel) {
        const px = colX(8) + 0.1, pw = col(4) - 0.1;
        if (keyValue !== undefined) s.addText(nb(keyValue), base({ x: px, y: Z.bodyY + 0.1, w: pw, h: 0.6, fontSize: TYPE.keyValue, bold: true, color: redInChart ? INK : ACC, valign: 'top', objectName: 'angka-kunci' }));
        if (keyText) {
          checkFit('chart.keyText', [keyText], pw, TYPE.body, bottom - Z.bodyY - 0.85);
          s.addText(rich(keyText), base({ x: px, y: Z.bodyY + 0.85, w: pw, h: bottom - Z.bodyY - 0.85, fontSize: TYPE.body, color: TEXT, valign: 'top', lineSpacingMultiple: 1.15, objectName: 'penjelasan' }));
        }
      }
      source(s, src);
      return finish(s);
    },

    // 8. Gambar + keterangan — "Seperti apa wujudnya?" Rasio gambar dijaga; keterangan "Gambar N · …" di bawah.
    // wide: true → gambar 8 kolom + teks 4 kolom (tangkapan layar lebar); bawaan 7 + 5.
    image({ chapter, title: t, image, alt, caption, bullets = [], wide = false, source: src, notes } = {}) {
      if (!alt) warn('image: alt (teks alternatif) wajib diisi untuk aksesibilitas.');
      const buf = fs.readFileSync(image);
      const info = imageInfo(buf, image);
      const s = add('NEURON_ISI', notes);
      head(s, 'image', chapter, t);
      const bottom = bottomOf(src);
      const boxW = col(wide ? 8 : 7), boxH = bottom - Z.bodyY - (caption ? 0.35 : 0);
      const scale = Math.min(boxW / info.w, boxH / info.h);
      const iw = info.w * scale, ih = info.h * scale;
      s.addImage({ data: `image/${info.type === 'jpg' ? 'jpeg' : 'png'};base64,${buf.toString('base64')}`, x: M, y: Z.bodyY, w: iw, h: ih,
        altText: alt || '', objectName: 'gambar' });
      if (caption) {
        const [label, ...rest] = String(caption).split(' · ');
        s.addText([{ text: label + (rest.length ? ' · ' : ''), options: { bold: true } }, { text: rest.join(' · ') }],
          base({ x: M, y: Z.bodyY + ih + 0.1, w: Math.max(iw, 3), h: 0.22, fontSize: TYPE.label, color: MUTED, valign: 'top', objectName: 'keterangan-gambar' }));
      }
      if (bullets.length) {
        if (bullets.length > 3) warn(`image: ${bullets.length} poin; maksimal 3 di samping gambar.`);
        const tx = M + boxW + GAP;
        bulletBox(s, 'image', bullets, tx, Z.bodyY, M + W - tx, bottom - Z.bodyY, { objectName: 'poin' });
      }
      source(s, src);
      return finish(s);
    },

    // 9. Angka penting — "Berapa angkanya dan apa artinya?" stats: [{ value, label, note }]; highlight: indeks merah.
    stats({ chapter, title: t, stats = [], highlight = 0, explanation, source: src, notes } = {}) {
      if (stats.length > 4) warn(`stats: ${stats.length} angka; maksimal 4 per slide (3 bila tanpa ikon).`);
      const n = stats.length, one = n === 1, four = n >= 4;
      if (highlight !== -1 && n && !(Number.isInteger(highlight) && highlight >= 0 && highlight < n)) throw new Error(`[neuron_theme] stats: highlight ${highlight} di luar rentang 0..${n - 1} (atau -1 tanpa sorotan).`);
      stats.forEach((st, i) => need(st && st.value, `stats: stats[${i}].value`));
      const s = add('NEURON_ISI', notes);
      head(s, 'stats', chapter, t);
      // hierarki angka (V3): 1 angka 54 pt · 2–3 angka 40 pt · 4 angka 32 pt; satu ukuran & satu angka merah per slide
      const cw = one ? col(6) : four ? col(3) : col(4);
      const big = one ? TYPE.statOne : four ? TYPE.keyValue : TYPE.stat, vh = big / 72 * 1.25;
      const hasIcon = stats.some((st) => st && st.icon), iconS = 0.4, ih = hasIcon ? iconS + 0.15 : 0;
      const icons = hasIcon ? loadIcons() : null;
      // blok (ikon + angka + label + keterangan) di tengah vertikal zona isi
      const blockH = ih + vh + 0.12 + 0.36 + Math.max(0, ...stats.map((st) => (st && st.note ? textH(st.note, cw, TYPE.small) : 0)));
      const top = Z.bodyY + Math.max(0, (bottomOf(src) - Z.bodyY - blockH) / 2 - 0.1);
      stats.forEach((st, i) => {
        need(st && st.value, `stats: stats[${i}].value`);
        const x = one ? M : colX(i * (four ? 3 : 4));
        if (st.label && lines(st.label, cw, TYPE.body, true) > 1) warn(`stats: label angka ${i + 1} lebih dari 1 baris — persingkat.`);
        if (lines(st.value, cw, big, true) > 1) warn(`stats: angka ${i + 1} ("${st.value}") terlalu lebar untuk kolomnya — persingkat (mis. "Rp 1,2 M").`);
        if (st.icon && icons) icons.icon(s, st.icon, { x, y: top, size: iconS, color: INK, objectName: `ikon-angka-${i + 1}` });
        const vy = top + ih;
        s.addText(nb(st.value), base({ x, y: vy, w: cw, h: vh, fontSize: big, bold: true, color: i === highlight ? ACC : INK, valign: 'bottom', objectName: `angka-${i + 1}` }));
        let y = vy + vh + 0.12;
        if (st.label) { s.addText(rich(st.label), base({ x, y, w: cw, h: 0.3, fontSize: TYPE.body, bold: true, color: INK, valign: 'top', objectName: `label-angka-${i + 1}` })); y += 0.36; }
        if (st.note) s.addText(rich(st.note), base({ x, y, w: cw, h: 0.6, fontSize: TYPE.small, color: MUTED, valign: 'top', lineSpacingMultiple: 1.15, objectName: `ket-angka-${i + 1}` }));
      });
      if (one && explanation) {
        const x = colX(6);
        checkFit('stats.explanation', [explanation], col(6), TYPE.body, bottomOf(src) - Z.bodyY - 0.2);
        s.addText(rich(explanation), base({ x, y: Z.bodyY + 0.2, w: col(6), h: bottomOf(src) - Z.bodyY - 0.2, fontSize: TYPE.body, color: TEXT, valign: 'top', lineSpacingMultiple: 1.2, objectName: 'penjelasan' }));
      }
      source(s, src);
      return finish(s);
    },

    // 10. Penutup — "Apa yang harus terjadi setelah ini, dan siapa yang dihubungi?"
    // steps: string | { text, owner }; contact: { name, role, email, phone }.
    closing({ chapter = 'Penutup', title: t = 'Langkah berikutnya', steps = [], contact, notes } = {}) {
      if (steps.length > 4) warn(`closing: ${steps.length} langkah; maksimal 4.`);
      openSection('Penutup');
      const s = add('NEURON_PENUTUP', notes);
      head(s, 'closing', chapter, t);
      const tw = col(7) - 0.45;
      let y = Z.bodyY;
      steps.forEach((st, i) => {
        const { text, owner } = typeof st === 'string' ? { text: st } : st || {};
        need(text, `closing: steps[${i}].text`);
        const h = textH(text, tw, TYPE.body, { spacing: 1.1 }) + (owner ? textH(owner, tw, TYPE.small, { spacing: 1.1 }) : 0) + 0.05;
        s.addText(String(i + 1), base({ x: M, y, w: 0.4, h: 0.3, fontSize: TYPE.subhead, bold: true, color: INK, valign: 'top', objectName: `langkah-no-${i + 1}` }));
        s.addText([...rich(text, { fontSize: TYPE.body, color: TEXT }), ...(owner ? [{ text: '', options: { breakLine: true } }, { text: owner, options: { fontSize: TYPE.small, color: MUTED } }] : [])],
          base({ x: M + 0.45, y, w: tw, h, valign: 'top', lineSpacingMultiple: 1.1, objectName: `langkah-${i + 1}` }));
        y += h + 0.18;
      });
      if (y - 0.18 > Z.bodyEnd) warn(`closing: langkah diperkirakan sampai y ${(y - 0.18).toFixed(2)}", batas ${Z.bodyEnd}" — persingkat atau kurangi langkah.`);
      if (contact) {
        const x = colX(8), w = col(4);
        const runs = [{ text: 'Kontak', options: { fontSize: TYPE.label, bold: true, color: MUTED, breakLine: true } }];
        if (contact.name) runs.push({ text: contact.name, options: { fontSize: TYPE.body, bold: true, color: INK, breakLine: true } });
        if (contact.role) runs.push({ text: contact.role, options: { fontSize: TYPE.small, color: MUTED, breakLine: true } });
        [contact.email, contact.phone].filter(Boolean).forEach((v) => runs.push({ text: v, options: { fontSize: TYPE.small, color: TEXT, breakLine: true } }));
        runs[runs.length - 1].options.breakLine = false;
        s.addText(runs, base({ x, y: Z.bodyY, w, h: 1.6, valign: 'top', paraSpaceAfter: 4, objectName: 'kontak' }));
      }
      return finish(s);
    },

    // Visualisasi (neuron_visuals.js): diagram proses, matriks 2×2, arsitektur C4 — lihat SKILL.md.
    process: visuals.process, cycle: visuals.cycle, timeline: visuals.timeline, orgChart: visuals.orgChart,
    compare: visuals.compare, pyramid: visuals.pyramid, funnel: visuals.funnel, matrix: visuals.matrix, architecture: visuals.architecture,
    // Ikon garis (Lucide, bentuk native) pada slide hasil fungsi tema: deck.icon(slide, 'users', { x, y, size, color }).
    icon(slide, name, opts) { const m = loadIcons(); if (!m) throw new Error('[neuron_theme] ikon tidak tersedia (scripts/neuron_icons.js hilang).'); return m.icon(slide, name, opts); },

    // Menulis berkas, merapikan XML, menerapkan tema Neuron. Mengembalikan path berkas (pptxgenjs menambah .pptx bila tidak ada).
    async save(out) {
      const file = await pres.writeFile({ fileName: out });
      await finalize(file);
      return file;
    },
  };
  return api;
}

module.exports = { createDeck, THEME, TYPE, SERIES, NEUTRAL, niceScale, formatNum, ACC, ACC_DARK, INK, MUTED, TEXT, RULE, HEAD_FILL, ZEBRA, GREY, FONT,
  LOGO_PATH, LOGO_RATIO, imageInfo };
