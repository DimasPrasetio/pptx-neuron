// Desain presentasi Neuron (PowerPoint 16:9).
// Hanya DESAIN: warna, huruf, tata letak, logo, motif Tiga Goresan, footer. Isi slide selalu dari pemanggil.
// Token sama persis dengan docx-neuron/scripts/neuron_theme.js; tata letak dari brief desain (riset 2026-11-05).
//
// TAHAP RANCANGAN: baru 3 jenis slide (sampul, isi poin, tabel). Tujuh jenis lain menyusul di Konstruksi.
const fs = require('fs');
const path = require('path');
// pptxgenjs dibundel di scripts/vendor (tanpa npm install). Bila bundel hilang, jatuh ke paket terpasang.
const loadPptx = () => { try { return require('./vendor/pptxgenjs.bundle.cjs'); } catch (e) { return require('pptxgenjs'); } };
const PptxGenJS = loadPptx();
const { applyTheme } = require('./apply_theme.js');

const LOGO_PATH = path.join(__dirname, '..', 'assets', 'logo_neuron.png');
const LOGO_RATIO = 612 / 176;   // 3,477 : 1 — jangan diregangkan; latar logo putih opak → hanya di area putih
// Ruang putih di dalam berkas logo (diukur dari piksel): tinta mulai 4,41% dari kiri, berakhir 8,66% dari kanan,
// 22,73% dari atas, 19,89% dari bawah. Dipakai agar TINTA logo sejajar margin, bukan kotak gambarnya.
const LOGO_PAD = { l: 0.0441, r: 0.0866, t: 0.2273, b: 0.1989 };

// ---------- token desain (sama dengan docx-neuron) ----------
const ACC = 'B3121B';       // merah logo: label bab, Tiga Goresan, sorotan data (maks. 2 elemen per slide isi)
const ACC_DARK = '7F0D14';  // turunan B3121B x 0,71: hanya goresan ke-2 motif
const INK = '1F2937';       // judul; latar pembatas bab
const MUTED = '5B6470';     // sekunder: sumber, footer, keterangan
const TEXT = '111111';      // teks isi
const RULE = 'C9CDD3';      // garis tabel/grafik (bukan teks di latar terang)
const HEAD_FILL = '2B3440'; // header tabel
const ZEBRA = 'F4F5F7';     // baris selang-seling, panel "Inti"
const WHITE = 'FFFFFF';
const FONT = 'Arial';
const LANG = 'id-ID';

// Tema PowerPoint (pemilih warna di PowerPoint menampilkan palet Neuron).
const THEME = {
  name: 'Neuron', headFontFace: FONT, bodyFontFace: FONT,
  colors: { dk1: TEXT, lt1: WHITE, dk2: INK, lt2: ZEBRA, accent1: ACC, accent2: INK, accent3: MUTED, accent4: '8C939D',
    accent5: HEAD_FILL, accent6: ACC_DARK, hlink: ACC, folHlink: MUTED },
};

// ---------- grid (inci, kanvas 10 x 5,625) ----------
const M = 0.5, W = 9.0, GAP = 0.3;
const Z = { chapterY: 0.45, titleY: 0.72, titleH: 0.75, bodyY: 1.65, bodyEnd: 4.75, sourceY: 4.72, footerY: 5.12 };
// Skala huruf (pt) pada kanvas 10". Setara x1,33 di kanvas lebar 13,33" — judul 22 pt ~ 29 pt, isi 14 pt ~ 19 pt.
// Revisi 2026-11-05: skala v1 (judul 36, isi 16) dinilai pengguna terlalu besar / kurang profesional.
const TYPE = { coverTitle: 30, title: 22, coverSub: 14, body: 14, key: 14, table: 12, tableDense: 11, label: 10,
  meta: 11, source: 9, footer: 9 };
const col = (n) => n * 0.475 + (n - 1) * 0.3;   // lebar n kolom dari grid 12 kolom

// ---------- helper teks ----------
// rich('teks **tebal** biasa', o): '**...**' menjadi tebal (sama seperti docx-neuron).
const rich = (s, o = {}) => String(s).split('**').map((seg, k) => [seg, k % 2 === 1]).filter(([seg]) => seg !== '')
  .map(([seg, bold]) => ({ text: seg, options: { ...o, bold: o.bold || bold } }));
const base = (o = {}) => ({ fontFace: FONT, lang: LANG, isTextBox: true, margin: 0, ...o });

// ---------- motif Tiga Goresan (dari simbol logo) ----------
// Tiga jajaran genjang ramping condong ±20° (bawah-kiri -> atas-kanan). Dekoratif; dibuat terakhir (urutan baca).
function strokes(slide, { x, y, h, sw, colors, transparency = [0, 0, 0], name = 'motif-neuron' }) {
  const slant = h * Math.tan((20 * Math.PI) / 180);
  const gap = sw * 0.6;
  colors.forEach((c, i) => {
    const x0 = x + i * (sw + gap);
    slide.addShape('custGeom', {
      x: x0, y, w: slant + sw, h, objectName: `${name}-${i + 1}`, altText: '',
      fill: { color: c, transparency: transparency[i] }, line: { type: 'none' },
      points: [{ x: 0, y: h }, { x: slant, y: 0 }, { x: slant + sw, y: 0 }, { x: sw, y: h }, { close: true }],
    });
  });
  return slant + 3 * sw + 2 * gap;   // lebar total
}

// Label bab merah 12 pt + Tiga Goresan mini: penanda posisi di deck (pengganti garis aksen Word).
function chapterLabel(slide, text) {
  if (!text) return;
  slide.addText(text, base({ x: M + 0.24, y: Z.chapterY, w: 6, h: 0.22, fontSize: TYPE.label, bold: true, color: ACC, valign: 'middle',
    objectName: 'label-bab' }));
  slide._neuronMini = true;
}
const finishSlide = (slide) => {
  if (slide._neuronMini) strokes(slide, { x: M, y: Z.chapterY + 0.03, h: 0.16, sw: 0.035, colors: [ACC, ACC_DARK, INK], name: 'motif-mini' });
};

function source(slide, text, y = Z.sourceY) {
  if (!text) return;
  slide.addText(text, base({ x: M, y, w: W, h: 0.2, fontSize: TYPE.source, color: MUTED, valign: 'top', objectName: 'sumber' }));
}

// pptxgenjs 4.0.1 menulis <a:pPr> untuk SETIAP run, sehingga paragraf ber-run banyak (mis. '**tebal**' di tengah
// kalimat) berisi <a:pPr> di tengah paragraf: tidak sah menurut skema dan run berikutnya membawa <a:buNone/> yang
// menghapus bullet. pPr hanya boleh menjadi anak pertama <a:p>: buang setiap pPr yang mengikuti run/br.
async function fixParagraphs(file) {
  const zip = await PptxGenJS.JSZip.loadAsync(fs.readFileSync(file));
  const re = /(<\/a:r>|<a:br\/>|<\/a:fld>)<a:pPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:pPr>)/g;
  for (const name of Object.keys(zip.files).filter((n) => /^ppt\/(slides|slideLayouts)\/[^/]+\.xml$/.test(n))) {
    const xml = await zip.file(name).async('string');
    let fixed = xml, prev;
    do { prev = fixed; fixed = fixed.replace(re, '$1'); } while (fixed !== prev);
    if (fixed !== xml) zip.file(name, fixed);
  }
  fs.writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
}

// ---------- deck ----------
function createDeck({ title = 'Presentasi Neuron', author = 'Neuronworks Indonesia', company = 'Neuronworks Indonesia',
  footerLabel = '' } = {}) {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_16x9';
  pres.title = title; pres.author = author;
  pres.company = String(company).replace(/&/g, '&amp;');   // ditulis tanpa escape oleh pptxgenjs
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };

  const titlePh = (o) => ({ placeholder: { options: { name: 'title', type: 'title', fontFace: FONT, bold: true, color: INK,
    valign: 'top', align: 'left', margin: 0, lang: LANG, lineSpacingMultiple: 0.9, ...o }, text: '' } });
  // logoAt({ left | right, top | middle, w }): posisi berdasarkan tepi tinta logo.
  const logoAt = ({ left, right, top, middle, w }) => {
    const h = w / LOGO_RATIO;
    const x = left !== undefined ? left - LOGO_PAD.l * w : right - w + LOGO_PAD.r * w;
    const y = top !== undefined ? top - LOGO_PAD.t * h : middle - ((LOGO_PAD.t + 1 - LOGO_PAD.b) / 2) * h;
    return { image: { path: LOGO_PATH, x, y, w, h, altText: 'Logo Neuron' } };
  };

  // Layout sampul: logo kiri atas, judul 40 pt.
  pres.defineSlideMaster({ title: 'NEURON_SAMPUL', background: { color: WHITE },
    objects: [logoAt({ left: M, top: M, w: 1.9 }), titlePh({ x: M, y: 2.1, w: 6.4, h: 1.0, fontSize: TYPE.coverTitle })] });
  // Layout isi: judul aksi 36 pt di posisi tetap, footer label dokumen + nomor slide + logo kanan bawah.
  pres.defineSlideMaster({ title: 'NEURON_ISI', background: { color: WHITE },
    margin: [0.5, 0.5, 0.75, 0.5],
    slideNumber: { x: 7.75, y: Z.footerY, w: 0.6, h: 0.22, fontFace: FONT, fontSize: TYPE.footer, color: MUTED, align: 'right' },
    objects: [
      titlePh({ x: M, y: Z.titleY, w: W, h: Z.titleH, fontSize: TYPE.title }),
      ...(footerLabel ? [{ text: { text: footerLabel, options: base({ x: M, y: Z.footerY, w: 5.5, h: 0.22, fontSize: TYPE.footer, color: MUTED, valign: 'middle' }) } }] : []),
      logoAt({ right: M + W, middle: Z.footerY + 0.11, w: 0.95 }),
    ] });

  const slides = [];
  const add = (masterName, notes) => { const s = pres.addSlide({ masterName }); if (notes) s.addNotes(notes); slides.push(s); return s; };

  const api = {
    pres,
    // 1. Sampul — "Apa ini, untuk siapa, dari siapa?"
    cover({ type = '', title: t, subtitle = '', meta = '', classification = '', notes } = {}) {
      const s = add('NEURON_SAMPUL', notes);
      if (type) s.addText(String(type).toUpperCase(), base({ x: M, y: 1.8, w: 6.4, h: 0.22, fontSize: TYPE.label, bold: true, color: ACC, charSpacing: 1.5, valign: 'middle', objectName: 'jenis-dokumen' }));
      s.addText(t, { placeholder: 'title' });
      if (subtitle) s.addText(subtitle, base({ x: M, y: 3.25, w: 6.0, h: 0.5, fontSize: TYPE.coverSub, color: MUTED, valign: 'top', objectName: 'subjudul' }));
      if (meta) s.addText(meta, base({ x: M, y: 4.45, w: 5.9, h: 0.22, fontSize: TYPE.meta, color: INK, valign: 'top', objectName: 'penyusun' }));
      if (classification) s.addText(classification, base({ x: M, y: 4.75, w: 5.9, h: 0.2, fontSize: TYPE.source, color: MUTED, valign: 'top', objectName: 'klasifikasi' }));
      strokes(s, { x: 7.35, y: 1.0, h: 3.6, sw: 0.16, colors: [ACC, ACC_DARK, INK], name: 'motif-sampul' });
      return s;
    },
    // 4. Isi poin — "Apa kesimpulannya dan apa 3–5 alasannya?" Panel Inti kanan opsional.
    bullets({ chapter, title: t, bullets = [], keyPoint, source: src, notes } = {}) {
      if (bullets.length > 5) console.warn(`[neuron_theme] bullets: ${bullets.length} poin; maksimal 5 per slide — pecah slide.`);
      const s = add('NEURON_ISI', notes);
      chapterLabel(s, chapter);
      s.addText(t, { placeholder: 'title' });
      const textW = keyPoint ? col(8) : W;   // grid 8 + 4: baris poin cukup panjang agar jarang terbungkus
      const bottom = src ? Z.sourceY - GAP : Z.bodyEnd;
      // Satu paragraf per poin: semua run poin membawa bullet yang sama; run terakhir menutup baris.
      // pptxgenjs memulai paragraf baru pada run ber-bullet: hanya run pertama poin yang membawa bullet.
      const items = bullets.flatMap((b, i) => rich(b).map((r, k, arr) => ({ text: r.text,
        options: { ...r.options, ...(k === 0 ? { bullet: { indent: 18 } } : {}), ...(k === arr.length - 1 && i < bullets.length - 1 ? { breakLine: true } : {}) } })));
      s.addText(items,
        base({ x: M, y: Z.bodyY, w: textW, h: bottom - Z.bodyY, fontSize: TYPE.body, color: TEXT, valign: 'top', paraSpaceAfter: 8,
          lineSpacingMultiple: 1.15, objectName: 'poin' }));
      if (keyPoint) {
        const px = M + col(8) + GAP;
        s.addShape('rect', { x: px, y: Z.bodyY, w: col(4), h: bottom - Z.bodyY, fill: { color: ZEBRA }, line: { type: 'none' }, objectName: 'panel-inti' });
        s.addText([{ text: 'Inti', options: { fontSize: TYPE.label, bold: true, color: MUTED, breakLine: true } },
          ...rich(keyPoint, { fontSize: TYPE.key, bold: true, color: INK })],
          base({ x: px + 0.22, y: Z.bodyY + 0.22, w: col(4) - 0.44, h: bottom - Z.bodyY - 0.44, valign: 'top', paraSpaceAfter: 6, lineSpacingMultiple: 1.15, objectName: 'inti' }));
      }
      source(s, src);
      finishSlide(s);
      return s;
    },
    // 6. Tabel — "Bagaimana perbandingan angka/atribut ini?" Header 2B3440, zebra F4F5F7, garis horizontal saja.
    table({ chapter, title: t, caption, head, rows, align = [], widths, source: src, notes } = {}) {
      const maxRows = caption ? 7 : 8;
      if (rows.length > maxRows) console.warn(`[neuron_theme] table: ${rows.length} baris; maksimal ${maxRows} per slide — pecah ke slide "(lanjutan)".`);
      const s = add('NEURON_ISI', notes);
      chapterLabel(s, chapter);
      s.addText(t, { placeholder: 'title' });
      let y = Z.bodyY;
      if (caption) {
        const [label, ...rest] = String(caption).split(' · ');
        s.addText([{ text: label + (rest.length ? ' · ' : ''), options: { bold: true } }, { text: rest.join(' · ') }],
          base({ x: M, y, w: W, h: 0.22, fontSize: TYPE.label, color: INK, valign: 'middle', objectName: 'judul-tabel' }));
        y += 0.3;
      }
      const fs_ = head.length > 5 || rows.length > 6 ? TYPE.tableDense : TYPE.table;
      const hair = { type: 'solid', pt: 0.5, color: RULE }, none = { type: 'none' };
      const al = (i) => (align[i] === 'r' ? 'right' : align[i] === 'c' ? 'center' : 'left');
      const cell = (txt, i, isHead, zebra) => ({
        text: isHead ? String(txt) : rich(txt, {}),
        options: { fontFace: FONT, lang: LANG, fontSize: fs_, bold: isHead, color: isHead ? WHITE : TEXT, align: isHead && align[i] !== 'r' ? 'left' : al(i),
          valign: 'middle', fill: { color: isHead ? HEAD_FILL : zebra ? ZEBRA : WHITE }, margin: [0, 0.1, 0, 0.1],
          border: [isHead ? none : hair, none, hair, none] },
      });
      const colW = widths || head.map(() => W / head.length);
      s.addTable([head.map((h, i) => cell(h, i, true)), ...rows.map((r, ri) => r.map((c, i) => cell(c, i, false, ri % 2 === 1)))],
        { x: M, y, w: W, colW, rowH: 0.32, autoPage: false, objectName: 'tabel' });
      source(s, src, Math.min(Z.sourceY, y + (rows.length + 1) * 0.32 + 0.12));   // sumber menempel di bawah tabel
      finishSlide(s);
      return s;
    },
    async save(out) {
      await pres.writeFile({ fileName: out });
      await fixParagraphs(out);
      await applyTheme(out, THEME);
      return out;
    },
  };
  return api;
}

module.exports = { createDeck, THEME, TYPE, ACC, ACC_DARK, INK, MUTED, TEXT, RULE, HEAD_FILL, ZEBRA, FONT, LOGO_PATH, LOGO_RATIO };
