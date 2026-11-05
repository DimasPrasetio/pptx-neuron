// Visualisasi Neuron: diagram proses, infografis, dan diagram teknis dari BENTUK NATIVE PowerPoint (bisa diedit),
// serta penanda label ujung grafik garis. Aturan dari brief desain › "Visualisasi (riset 2026-11-05)":
// model "sorot + netral" (satu merah per slide), simpul isian F4F5F7 tanpa garis tepi, simpul utama 1F2937 teks putih,
// pihak luar putih + garis 1 pt 8C939D, batas sistem putus-putus 8C939D, panah 1 pt 5B6470 satu arah lurus/siku.
//
// Modul ini dipanggil dari neuron_theme.js (createDeck) dengan konteks bersama; jangan di-require langsung.

function build(ctx) {
  const { add, head, finish, source, bottomOf, base, rich, nb, plain, lines, textH, warn, need, T, TYPE, M, W, Z, GAP, col, colX } = ctx;
  const { ACC, INK, MUTED, TEXT, RULE, ZEBRA, GREY, WHITE } = T;
  const V = TYPE.v;   // skala huruf visualisasi (V3)

  // ---------- primitif ----------
  // Garis lurus dari (x1,y1) ke (x2,y2); arrow = ujung segitiga di (x2,y2). Horizontal/vertikal/diagonal lewat flip.
  function seg(s, x1, y1, x2, y2, { color = MUTED, width = 1, arrow = false, dash = false, name = 'diagram-garis' } = {}) {
    const flipH = x2 < x1, flipV = y2 < y1;
    s.addShape('line', {
      x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1), flipH, flipV, objectName: name,
      line: { color, width, dashType: dash ? 'dash' : 'solid', ...(arrow ? { endArrowType: 'triangle' } : {}) },
    });
  }
  // Polyline siku: titik [[x,y],…]; panah hanya di segmen terakhir.
  function path(s, pts, o = {}) {
    for (let i = 1; i < pts.length; i++) seg(s, pts[i - 1][0], pts[i - 1][1], pts[i][0], pts[i][1], { ...o, arrow: !!o.arrow && i === pts.length - 1, name: `${o.name || 'diagram-garis'}-${i}` });
  }
  // Simpul: kotak dengan judul + (tag) + deskripsi. kind: 'node' | 'main' | 'external' | 'person' | 'db'
  function node(s, { x, y, w, h, title, tag, desc, kind = 'node', focus = false, name = 'diagram-simpul', alt }) {
    const dark = kind === 'main';
    const fill = dark ? INK : kind === 'external' ? WHITE : ZEBRA;
    const line = kind === 'external' ? { color: GREY, width: 1 } : focus ? { color: ACC, width: 1.5 } : { type: 'none' };
    const shape = kind === 'db' ? 'can' : 'rect';
    s.addShape(shape, { x, y, w, h, fill: { color: fill }, line, objectName: name, ...(alt ? { altText: alt } : {}),
      ...(kind === 'db' ? { adjustPoint: 0.18 } : {}) });
    const tc = dark ? WHITE : INK, sc = dark ? RULE : MUTED;
    const runs = [{ text: plain(title), options: { fontSize: V.node, bold: true, color: tc, breakLine: !!(tag || desc) } }];
    if (tag) runs.push({ text: `[${tag}]`, options: { fontSize: V.tag, color: sc, breakLine: !!desc } });
    if (desc) runs.push({ text: plain(desc), options: { fontSize: V.tag, color: sc } });
    const top = kind === 'db' ? 0.14 : 0.06;
    s.addText(runs, base({ x: x + 0.08, y: y + top, w: w - 0.16, h: h - top - 0.04, align: 'center', valign: 'middle', lineSpacingMultiple: 1.0, objectName: `${name}-teks` }));
  }
  const fmtID = (v, opts) => new Intl.NumberFormat('id-ID', opts).format(v);

  return {
    // ---------- Diagram proses: alur langkah (V4 #7) ----------
    // steps: [{ title, text }] 3–6; active: indeks langkah yang dibahas (lingkaran merah + nama tebal); variant 'ribbon' = pita fase.
    process({ chapter, title: t, steps = [], active, variant = 'circles', source: src, notes } = {}) {
      if (steps.length < 3 || steps.length > 6) warn(`process: ${steps.length} langkah; idealnya 3–6 (pecah slide bila lebih).`);
      const s = add('NEURON_ISI', notes);
      head(s, 'process', chapter, t);
      const n = steps.length, w = (W - (n - 1) * GAP) / n;
      steps.forEach((st, i) => {
        need(st && st.title, `process: steps[${i}].title`);
        const x = M + i * (w + GAP), on = active === i;
        if (variant === 'ribbon') {
          if (st.title.split(/\s+/).length > 2) warn(`process(ribbon): nama fase "${st.title}" > 2 kata — pakai variant 'circles'.`);
          s.addShape(i === 0 ? 'homePlate' : 'chevron', { x: x - (i ? 0.12 : 0), y: Z.bodyY, w: w + (i ? 0.12 : 0) + 0.12 - 0.06, h: 0.5,
            fill: { color: on ? INK : ZEBRA }, line: { type: 'none' }, objectName: `diagram-fase-${i + 1}`, ...(i === 0 ? { altText: `Alur ${n} fase: ${steps.map((x) => plain(x.title)).join(', ')}` } : {}) });
          s.addText(plain(st.title), base({ x: x + 0.12, y: Z.bodyY, w: w - 0.24, h: 0.5, fontSize: V.ribbon, bold: true, color: on ? WHITE : INK, align: 'center', valign: 'middle', objectName: `diagram-fase-${i + 1}-teks` }));
          if (st.text) s.addText(rich(st.text), base({ x, y: Z.bodyY + 0.7, w, h: bottomOf(src) - Z.bodyY - 0.7, fontSize: V.desc + 1, color: TEXT, valign: 'top', lineSpacingMultiple: 1.15, objectName: `diagram-fase-${i + 1}-isi` }));
          return;
        }
        // blok diagram (lingkaran + nama + deskripsi) ditengahkan vertikal di zona isi, maks. turun 0,5"
        const maxDesc = Math.max(0, ...steps.map((x) => (x && x.text ? textH(x.text, w, V.desc) : 0)));
        const d = 0.36, cy = Z.bodyY + Math.min(0.5, Math.max(0.1, (bottomOf(src) - Z.bodyY - (d + 0.62 + maxDesc)) / 2));
        s.addShape('ellipse', { x, y: cy, w: d, h: d, fill: { color: on ? ACC : INK }, line: { type: 'none' }, objectName: `diagram-langkah-${i + 1}`,
          ...(i === 0 ? { altText: `Alur ${n} langkah: ${steps.map((x) => plain(x.title)).join(' → ')}` } : {}) });
        s.addText(String(i + 1), base({ x, y: cy, w: d, h: d, fontSize: V.num, bold: true, color: WHITE, align: 'center', valign: 'middle', objectName: `diagram-langkah-${i + 1}-no` }));
        if (i < n - 1) seg(s, x + d + 0.08, cy + d / 2, x + w + GAP - 0.08, cy + d / 2, { arrow: true, name: `diagram-panah-${i + 1}` });
        const ty = cy + d + 0.14;
        if (lines(st.title, w, V.node, true) > 2) warn(`process: nama langkah ${i + 1} lebih dari 2 baris.`);
        s.addText(plain(st.title), base({ x, y: ty, w, h: 0.42, fontSize: V.node, bold: true, color: INK, valign: 'top', lineSpacingMultiple: 1.0, objectName: `diagram-langkah-${i + 1}-nama` }));
        if (st.text) {
          if (lines(st.text, w, V.desc) > 4) warn(`process: deskripsi langkah ${i + 1} lebih dari 4 baris — persingkat (± 12 kata).`);
          s.addText(rich(st.text), base({ x, y: ty + 0.48, w, h: bottomOf(src) - ty - 0.48, fontSize: V.desc, color: MUTED, valign: 'top', lineSpacingMultiple: 1.15, objectName: `diagram-langkah-${i + 1}-isi` }));
        }
      });
      source(s, src);
      return finish(s);
    },

    // ---------- Infografis: matriks 2×2 (V4 #13) ----------
    // axes: { x: { label, low, high }, y: { label, low, high } }; quadrants: { tl, tr, bl, br } (nama); best: kuadran berisian.
    // points: [{ label, x, y }] dengan x,y 0..1; focus: indeks titik merah; explanation: teks kanan (14 pt).
    matrix({ chapter, title: t, axes = {}, quadrants = {}, best = 'tr', points = [], focus, explanation, source: src, notes } = {}) {
      if (points.length > 10) warn(`matrix: ${points.length} titik; maksimal 10.`);
      const s = add('NEURON_ISI', notes);
      head(s, 'matrix', chapter, t);
      const px = M + 0.45, py = Z.bodyY + 0.3, pw = explanation ? 4.5 : 8.1, ph = bottomOf(src) - py - 0.35, hw = pw / 2, hh = ph / 2;
      const q = { tl: [px, py], tr: [px + hw, py], bl: [px, py + hh], br: [px + hw, py + hh] };
      Object.entries(q).forEach(([k, [x, y]]) => {
        s.addShape('rect', { x, y, w: hw, h: hh, fill: { color: k === best ? ZEBRA : WHITE }, line: { color: RULE, width: 1 }, objectName: `matriks-${k}`,
          ...(k === 'tl' ? { altText: `Matriks 2×2: ${axes.y?.label || ''} terhadap ${axes.x?.label || ''}; ${points.length} butir` } : {}) });
        if (quadrants[k]) s.addText(quadrants[k], base({ x: x + 0.08, y: y + 0.06, w: hw - 0.16, h: 0.2, fontSize: V.desc, bold: true, color: MUTED, align: k.endsWith('r') ? 'right' : 'left', valign: 'top', objectName: `matriks-${k}-nama` }));
      });
      // label sumbu: y di kiri-atas (tanpa teks diputar), x di bawah
      if (axes.y) s.addText(`${axes.y.label} ↑`, base({ x: M, y: Z.bodyY, w: 4, h: 0.22, fontSize: V.desc, bold: true, color: MUTED, valign: 'middle', objectName: 'matriks-sumbu-y' }));
      if (axes.y?.high) s.addText(axes.y.high, base({ x: M - 0.05, y: py, w: 0.45, h: 0.2, fontSize: V.tag, color: MUTED, align: 'right', objectName: 'matriks-y-tinggi' }));
      if (axes.y?.low) s.addText(axes.y.low, base({ x: M - 0.05, y: py + ph - 0.2, w: 0.45, h: 0.2, fontSize: V.tag, color: MUTED, align: 'right', objectName: 'matriks-y-rendah' }));
      if (axes.x) s.addText([{ text: `${axes.x.low || ''}  ←  `, options: { bold: false } }, { text: axes.x.label, options: { bold: true } }, { text: `  →  ${axes.x.high || ''}`, options: { bold: false } }],
        base({ x: px, y: py + ph + 0.08, w: pw, h: 0.22, fontSize: V.desc, color: MUTED, align: 'center', valign: 'middle', objectName: 'matriks-sumbu-x' }));
      points.forEach((p, i) => {
        if (!(p.x >= 0 && p.x <= 1 && p.y >= 0 && p.y <= 1)) throw new Error(`[neuron_theme] matrix: titik ${i + 1} di luar 0..1.`);
        const d = 0.14, cx = px + p.x * pw, cy = py + (1 - p.y) * ph, on = focus === i;
        s.addShape('ellipse', { x: cx - d / 2, y: cy - d / 2, w: d, h: d, fill: { color: on ? ACC : INK }, line: { type: 'none' }, objectName: `matriks-titik-${i + 1}` });
        const right = p.x < 0.8;
        s.addText(plain(p.label), base({ x: right ? cx + 0.1 : cx - 1.6, y: cy - 0.11, w: 1.5, h: 0.22, fontSize: V.desc, bold: on, color: TEXT, align: right ? 'left' : 'right', valign: 'middle', objectName: `matriks-titik-${i + 1}-label` }));
      });
      if (explanation) {
        const ex = px + pw + 0.5, ew = M + W - ex;
        if (Array.isArray(explanation)) ctx.bulletBox(s, 'matrix', explanation, ex, Z.bodyY + 0.3, ew, bottomOf(src) - Z.bodyY - 0.3, { objectName: 'matriks-penjelasan' });
        else s.addText(rich(explanation), base({ x: ex, y: Z.bodyY + 0.3, w: ew, h: bottomOf(src) - Z.bodyY - 0.3, fontSize: TYPE.body, color: TEXT, valign: 'top', lineSpacingMultiple: 1.2, objectName: 'matriks-penjelasan' }));
      }
      source(s, src);
      return finish(s);
    },

    // ---------- Diagram teknis: arsitektur C4 (V4 #15–16) ----------
    // nodes: [{ id, title, tag, desc, kind: 'person'|'external'|'container'|'db'|'main', col: 0..3, row: 0..2 }]
    // links: [{ from, to, label, protocol, focus, dashed }]; boundary: { label, cols: [a,b] } (batas sistem putus-putus);
    // caption: "Diagram kontainer · Sistem Inventori"; legend: true menampilkan kunci notasi.
    architecture({ chapter, title: t, caption, nodes = [], links = [], boundary, legend = true, source: src, notes } = {}) {
      if (nodes.length > 9) warn(`architecture: ${nodes.length} elemen; maksimal 9 (6 kontainer + 3 luar) — pecah diagram.`);
      const s = add('NEURON_ISI', notes);
      head(s, 'architecture', chapter, t);
      // grid 4 kolom × 2 baris; celah antarkolom 0,8 / 1,0 / 1,2" untuk label garis 9 pt
      const COLX = [0.5, 2.8, 5.3, 8.0], NW = 1.5, NH = 0.8;
      const top = Z.bodyY + (caption ? 0.62 : 0.35);
      const ROWY = [top, top + 1.15];
      const pos = {};
      nodes.forEach((nd, i) => {
        need(nd && nd.id, `architecture: nodes[${i}].id`); need(nd.title, `architecture: nodes[${i}].title`);
        if (!(nd.col >= 0 && nd.col <= 3 && nd.row >= 0 && nd.row <= 1)) throw new Error(`[neuron_theme] architecture: ${nd.id} col 0..3, row 0..1.`);
        pos[nd.id] = { x: COLX[nd.col], y: ROWY[nd.row], w: NW, h: NH, ...nd };
      });
      if (caption) s.addText(caption, base({ x: M, y: Z.bodyY, w: W, h: 0.22, fontSize: V.desc, bold: true, color: MUTED, valign: 'middle', objectName: 'diagram-judul' }));
      if (boundary) {
        const [a, b] = boundary.cols || [1, 2];
        const bx = COLX[a] - 0.22, bw = COLX[b] + NW + 0.22 - bx, by = top - 0.32, bh = ROWY[1] + NH + 0.15 - by;
        s.addShape('rect', { x: bx, y: by, w: bw, h: bh, fill: { type: 'none' }, line: { color: GREY, width: 1, dashType: 'dash' }, objectName: 'diagram-batas' });
        s.addText(boundary.label || '', base({ x: bx + 0.08, y: by + 0.04, w: bw - 0.16, h: 0.18, fontSize: V.tag, color: MUTED, valign: 'top', objectName: 'diagram-batas-label' }));
      }
      const alt = `Diagram arsitektur: ${nodes.map((n) => n.title).join(', ')}. ${links.map((l) => `${pos[l.from]?.title} → ${pos[l.to]?.title}`).join('; ')}`;
      nodes.forEach((nd, i) => {
        const p = pos[nd.id];
        const kind = nd.kind === 'container' ? 'node' : nd.kind === 'person' ? 'node' : nd.kind;
        node(s, { ...p, title: nd.title, tag: nd.tag, desc: nd.desc, kind, name: `diagram-${nd.id}`, alt: i === 0 ? alt : undefined });
      });
      links.forEach((l, i) => {
        const a = pos[l.from], b = pos[l.to];
        if (!a || !b) throw new Error(`[neuron_theme] architecture: link ${i + 1} merujuk simpul yang tidak ada (${l.from} → ${l.to}).`);
        const o = { color: l.focus ? ACC : MUTED, width: l.focus ? 1.5 : 1, arrow: true, dash: !!l.dashed, name: `diagram-aliran-${i + 1}` };
        const lab = (x, y, w, align = 'center') => {
          const need_ = Math.max(plain(l.label || '').length, (l.protocol || '').length + 2) * 0.064;
          if (need_ > w + 0.05) warn(`architecture: label garis "${l.label}" terlalu panjang untuk celah ${w.toFixed(2)}" — persingkat (± ${Math.floor(w / 0.064)} karakter).`);
          if (l.label) s.addText(l.label, base({ x, y: y - 0.21, w, h: 0.18, fontSize: V.tag, color: MUTED, align, valign: 'bottom', objectName: `diagram-aliran-${i + 1}-label` }));
          if (l.protocol) s.addText(`[${l.protocol}]`, base({ x, y: y + 0.03, w, h: 0.18, fontSize: V.tag, color: MUTED, align, valign: 'top', objectName: `diagram-aliran-${i + 1}-protokol` }));
        };
        if (a.row === b.row) {                                   // horizontal
          const ltr = b.x > a.x, y = a.y + NH / 2;
          const x1 = ltr ? a.x + NW + 0.05 : a.x - 0.05, x2 = ltr ? b.x - 0.05 : b.x + NW + 0.05;
          seg(s, x1, y, x2, y, o); lab(Math.min(x1, x2), y, Math.abs(x2 - x1));
        } else if (a.col === b.col) {                            // vertikal
          const down = b.y > a.y, x = a.x + NW / 2;
          const y1 = down ? a.y + NH + 0.05 : a.y - 0.05, y2 = down ? b.y - 0.05 : b.y + NH + 0.05;
          seg(s, x, y1, x, y2, o);
          if (l.label) s.addText(l.label, base({ x: x + 0.08, y: (y1 + y2) / 2 - 0.2, w: 1.4, h: 0.18, fontSize: V.tag, color: MUTED, valign: 'bottom', objectName: `diagram-aliran-${i + 1}-label` }));
          if (l.protocol) s.addText(`[${l.protocol}]`, base({ x: x + 0.08, y: (y1 + y2) / 2 + 0.02, w: 1.4, h: 0.18, fontSize: V.tag, color: MUTED, valign: 'top', objectName: `diagram-aliran-${i + 1}-protokol` }));
        } else {                                                 // siku Z: keluar sisi samping (seperempat bawah/atas), turun/naik di tengah celah, masuk sisi samping tujuan
          const ltr = b.x > a.x, down = b.y > a.y;
          const y1 = a.y + NH * (down ? 0.75 : 0.25), y2 = b.y + NH / 2;
          const x1 = ltr ? a.x + NW + 0.05 : a.x - 0.05, x2 = ltr ? b.x - 0.05 : b.x + NW + 0.05;
          const gapL = ltr ? a.x + NW : b.x + NW, gapR = ltr ? b.x : a.x, xm = (gapL + gapR) / 2;
          path(s, [[x1, y1], [xm, y1], [xm, y2], [x2, y2]], o);
          // label di samping segmen vertikal (sisi yang lebih lapang = sisi tujuan), boleh 2 baris
          const lw = Math.max(0.4, gapR - xm - 0.08);
          const txt = [l.label && { text: l.label, options: { breakLine: !!l.protocol } }, l.protocol && { text: `[${l.protocol}]` }].filter(Boolean);
          if (txt.length) s.addText(txt, base({ x: xm + 0.06, y: Math.min(y1, y2) + 0.08, w: lw, h: Math.abs(y2 - y1) - 0.16, fontSize: V.tag, color: MUTED, valign: 'middle', objectName: `diagram-aliran-${i + 1}-label` }));
        }
      });
      if (legend) {
        const ly = bottomOf(src) - 0.05;
        const items = [['main', 'sistem kami'], ['node', 'kontainer / orang'], ['external', 'sistem luar']];
        let lx = M;
        items.forEach(([k, txt]) => {
          s.addShape('rect', { x: lx, y: ly + 0.04, w: 0.14, h: 0.12, fill: { color: k === 'main' ? INK : k === 'external' ? WHITE : ZEBRA }, line: k === 'external' ? { color: GREY, width: 0.75 } : { type: 'none' }, objectName: `diagram-kunci-${k}` });
          s.addText(txt, base({ x: lx + 0.2, y: ly, w: 1.5, h: 0.2, fontSize: V.tag, color: MUTED, valign: 'middle', objectName: `diagram-kunci-${k}-teks` }));
          lx += 0.2 + txt.length * 0.055 + 0.3;
        });
        seg(s, lx, ly + 0.1, lx + 0.3, ly + 0.1, { arrow: true, name: 'diagram-kunci-panah' });
        s.addText('arah data · label: apa [protokol]', base({ x: lx + 0.38, y: ly, w: 3, h: 0.2, fontSize: V.tag, color: MUTED, valign: 'middle', objectName: 'diagram-kunci-panah-teks' }));
      }
      source(s, src);
      return finish(s);
    },

    // ---------- Label ujung untuk grafik garis (dipakai chart({ type: 'line' })) ----------
    lineEndLabels(s, { x, y, w, h, layout, min, max, series, colors, focus, fmt }) {
      const px = x + (layout.x + layout.w) * w + 0.08, py = y + layout.y * h, ph = layout.h * h;
      const items = series.map((sr, i) => {
        const v = sr.values[sr.values.length - 1];
        return { i, v, y: py + (1 - (v - min) / (max - min)) * ph - 0.1, text: `${sr.name} ${fmt(v)}` };
      }).sort((a, b) => a.y - b.y);
      for (let k = 1; k < items.length; k++) if (items[k].y - items[k - 1].y < 0.2) items[k].y = items[k - 1].y + 0.2;   // cegah tumpuk
      items.forEach((it) => s.addText(it.text, base({ x: px, y: it.y, w: x + w - px, h: 0.2, fontSize: V.desc, bold: it.i === focus, color: TEXT, valign: 'middle', objectName: `grafik-label-ujung-${it.i + 1}` })));
    },
    fmtID,
  };
}

module.exports = { build };
