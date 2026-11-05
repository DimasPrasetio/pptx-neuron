// Visualisasi Neuron: diagram proses, infografis, dan diagram teknis dari BENTUK NATIVE PowerPoint (bisa diedit),
// serta penanda label ujung grafik garis. Aturan dari brief desain › "Visualisasi (riset 2026-11-05)":
// model "sorot + netral" (satu merah per slide), simpul isian F4F5F7 tanpa garis tepi, simpul utama 1F2937 teks putih,
// pihak luar putih + garis 1 pt 8C939D, batas sistem putus-putus 8C939D, panah 1 pt 5B6470 satu arah lurus/siku.
//
// Modul ini dipanggil dari neuron_theme.js (createDeck) dengan konteks bersama; jangan di-require langsung.

function build(ctx) {
  const { add, head, finish, source, bottomOf, base, rich, nb, plain, lines, textH, warn, need, T, TYPE, M, W, Z, GAP, col, colX, formatNum } = ctx;
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

    // ---------- Diagram proses: siklus (V4 #8) ----------
    // steps: [{ title, text? }] 3–5; active: tahap yang dibahas (isian gelap, nomor merah); center: label tengah.
    cycle({ chapter, title: t, steps = [], active, center, source: src, notes } = {}) {
      const n = steps.length;
      if (n < 3 || n > 5) warn(`cycle: ${n} tahap; siklus 3–5 tahap (proses linear → pakai process).`);
      const s = add('NEURON_ISI', notes);
      head(s, 'cycle', chapter, t);
      const zoneH = bottomOf(src) - Z.bodyY, cx = M + W / 2, cy = Z.bodyY + zoneH / 2;
      const bw = n === 5 ? 1.8 : 2.2, bh = n === 5 ? 0.62 : 0.72;
      // pusat simpul: 4 = sudut persegi (searah jarum jam dari kiri-atas), 3 = segitiga, 5 = segi lima
      const rx = n === 4 ? 2.2 : 2.6, ry = (zoneH - bh) / 2 - 0.05;
      const centers = n === 4 ? [[cx - rx, cy - ry], [cx + rx, cy - ry], [cx + rx, cy + ry], [cx - rx, cy + ry]]
        : steps.map((_, i) => { const a = -Math.PI / 2 + (2 * Math.PI * i) / n; return [cx + rx * Math.cos(a), cy + ry * Math.sin(a)]; });
      const boxes = centers.map(([x, y]) => ({ x: x - bw / 2, y: y - bh / 2, w: bw, h: bh, cx: x, cy: y }));
      // titik potong garis antarpusat dengan tepi kotak (agar panah berhenti 0,05" sebelum simpul)
      const clip = (b, dx, dy) => { const k = Math.min(dx ? (b.w / 2) / Math.abs(dx) : Infinity, dy ? (b.h / 2) / Math.abs(dy) : Infinity); return [b.cx + dx * k, b.cy + dy * k]; };
      boxes.forEach((b, i) => {
        const st = steps[i] || {}; need(st.title, `cycle: steps[${i}].title`);
        const on = active === i;
        s.addShape('rect', { x: b.x, y: b.y, w: b.w, h: b.h, fill: { color: on ? INK : ZEBRA }, line: { type: 'none' }, objectName: `diagram-tahap-${i + 1}`,
          ...(i === 0 ? { altText: `Siklus ${n} tahap: ${steps.map((x) => plain(x.title)).join(' → ')} → kembali ke awal` } : {}) });
        // tanpa merah di siklus: nomor gelap/abu, tahap aktif cukup dibedakan isian gelap (merah tidak terbaca di 1F2937)
        s.addText([{ text: `${i + 1}  `, options: { bold: true, color: on ? WHITE : MUTED } }, { text: plain(st.title), options: { bold: true, color: on ? WHITE : INK, breakLine: !!st.text } },
          ...(st.text ? [{ text: plain(st.text), options: { fontSize: V.tag, color: on ? RULE : MUTED } }] : [])],
          base({ x: b.x + 0.1, y: b.y, w: b.w - 0.2, h: b.h, fontSize: V.node, align: 'center', valign: 'middle', objectName: `diagram-tahap-${i + 1}-teks` }));
        const nb2 = boxes[(i + 1) % n], dx = nb2.cx - b.cx, dy = nb2.cy - b.cy, len = Math.hypot(dx, dy);
        const [x1, y1] = clip(b, dx, dy), [x2, y2] = clip(nb2, -dx, -dy);
        seg(s, x1 + (dx / len) * 0.08, y1 + (dy / len) * 0.08, x2 - (dx / len) * 0.08, y2 - (dy / len) * 0.08, { arrow: true, name: `diagram-siklus-panah-${i + 1}` });
      });
      if (center) s.addText(center, base({ x: cx - 1.1, y: cy - 0.25, w: 2.2, h: 0.5, fontSize: V.node, bold: true, color: MUTED, align: 'center', valign: 'middle', objectName: 'diagram-siklus-tengah' }));
      source(s, src);
      return finish(s);
    },

    // ---------- Diagram proses: linimasa (V4 #9) ----------
    // milestones: [{ date, label, status: 'done'|'now'|'plan' }] ≤ 6, atau phases: [{ name, start, end, status }] (indeks kolom
    // periode, end eksklusif) dengan periods: ['Jan', …] dan today (indeks pecahan, opsional).
    timeline({ chapter, title: t, milestones, phases, periods = [], today, source: src, notes } = {}) {
      const s = add('NEURON_ISI', notes);
      head(s, 'timeline', chapter, t);
      if (milestones) {
        const n = milestones.length;
        if (n > 6) warn(`timeline: ${n} tonggak; maksimal 6 — pecah per semester.`);
        const ly = Z.bodyY + (bottomOf(src) - Z.bodyY) / 2 - 0.1, slot = W / n;
        seg(s, M, ly, M + W, ly, { color: GREY, width: 1, name: 'diagram-linimasa-garis' });
        milestones.forEach((m, i) => {
          need(m && m.label, `timeline: milestones[${i}].label`);
          const x = M + slot * (i + 0.5), st = m.status || 'plan', d = st === 'now' ? 0.2 : 0.16;
          s.addShape('ellipse', { x: x - d / 2, y: ly - d / 2, w: d, h: d, fill: { color: st === 'now' ? ACC : st === 'done' ? MUTED : WHITE },
            line: st === 'plan' ? { color: GREY, width: 1 } : { type: 'none' }, objectName: `diagram-tonggak-${i + 1}`,
            ...(i === 0 ? { altText: `Linimasa: ${milestones.map((x) => `${x.date} ${plain(x.label)}`).join('; ')}` } : {}) });
          s.addText(m.date || '', base({ x: x - slot / 2 + 0.05, y: ly - 0.5, w: slot - 0.1, h: 0.3, fontSize: V.desc, bold: true, color: INK, align: 'center', valign: 'bottom', objectName: `diagram-tonggak-${i + 1}-tanggal` }));
          if (lines(m.label, slot - 0.1, V.desc) > 3) warn(`timeline: label tonggak ${i + 1} lebih dari 3 baris.`);
          s.addText([{ text: plain(m.label), options: { breakLine: st === 'now' } }, ...(st === 'now' ? [{ text: 'Sekarang', options: { fontSize: V.tag, bold: true, color: ACC } }] : [])],
            base({ x: x - slot / 2 + 0.05, y: ly + 0.2, w: slot - 0.1, h: 0.75, fontSize: V.desc, color: TEXT, align: 'center', valign: 'top', lineSpacingMultiple: 1.1, objectName: `diagram-tonggak-${i + 1}-label` }));
        });
      } else if (phases) {
        if (phases.length > 6) warn(`timeline: ${phases.length} fase; maksimal 6 baris.`);
        const gx = M + 2.1, gw = M + W - gx, cwid = gw / Math.max(1, periods.length), rh = 0.28, gap = 0.14, y0 = Z.bodyY + 0.3;
        periods.forEach((p, i) => s.addText(p, base({ x: gx + i * cwid, y: Z.bodyY, w: cwid, h: 0.22, fontSize: V.tag, color: MUTED, align: 'center', valign: 'middle', objectName: `diagram-periode-${i + 1}` })));
        phases.forEach((ph, i) => {
          need(ph && ph.name, `timeline: phases[${i}].name`);
          if (!(ph.start >= 0 && ph.end > ph.start && ph.end <= periods.length)) throw new Error(`[neuron_theme] timeline: fase ${i + 1} start/end di luar 0..${periods.length}.`);
          const y = y0 + i * (rh + gap), st = ph.status || 'plan';
          s.addText(plain(ph.name), base({ x: M, y, w: 2.0, h: rh, fontSize: V.desc, bold: st === 'active', color: TEXT, valign: 'middle', objectName: `diagram-fase-${i + 1}-nama` }));
          s.addShape('rect', { x: gx + ph.start * cwid + 0.02, y, w: (ph.end - ph.start) * cwid - 0.04, h: rh, fill: { color: st === 'active' ? ACC : st === 'done' ? MUTED : GREY }, line: { type: 'none' },
            objectName: `diagram-fase-${i + 1}`, ...(i === 0 ? { altText: `Jadwal: ${phases.map((x) => `${plain(x.name)} ${periods[x.start]}–${periods[x.end - 1]}`).join('; ')}` } : {}) });
        });
        if (today !== undefined) {
          const x = gx + today * cwid, yb = y0 + phases.length * (rh + gap) - gap + 0.1;
          seg(s, x, Z.bodyY + 0.22, x, yb, { color: MUTED, width: 1, dash: true, name: 'diagram-hari-ini' });
          s.addText('Hari ini', base({ x: x - 0.5, y: yb + 0.02, w: 1.0, h: 0.18, fontSize: V.tag, color: MUTED, align: 'center', objectName: 'diagram-hari-ini-label' }));
        }
        const ky = Math.min(bottomOf(src) - 0.2, y0 + phases.length * (rh + gap) + 0.35);
        [['done', MUTED, 'selesai'], ['active', ACC, 'berjalan'], ['plan', GREY, 'rencana']].forEach(([k, c, txt], i) => {
          s.addShape('rect', { x: M + i * 1.3, y: ky + 0.04, w: 0.12, h: 0.12, fill: { color: c }, line: { type: 'none' }, objectName: `diagram-kunci-${k}` });
          s.addText(txt, base({ x: M + i * 1.3 + 0.18, y: ky, w: 1.0, h: 0.2, fontSize: V.tag, color: MUTED, valign: 'middle', objectName: `diagram-kunci-${k}-teks` }));
        });
      } else throw new Error('[neuron_theme] timeline: isi milestones atau phases.');
      source(s, src);
      return finish(s);
    },

    // ---------- Diagram proses: struktur organisasi (V4 #10) ----------
    // root: { title, role }; units: [{ title, role?, items?: [teks tingkat 3] }] ≤ 6; focus: indeks unit yang dibahas.
    orgChart({ chapter, title: t, root = {}, units = [], focus, source: src, notes } = {}) {
      need(root.title, 'orgChart: root.title');
      if (units.length > 6) warn(`orgChart: ${units.length} unit; maksimal 6 per tingkat — pecah per divisi.`);
      const s = add('NEURON_ISI', notes);
      head(s, 'orgChart', chapter, t);
      const n = Math.max(1, units.length), uw = (W - (n - 1) * GAP) / n, rw = 2.2, rh = 0.55;
      const rx = M + (W - rw) / 2, ry = Z.bodyY;
      node(s, { x: rx, y: ry, w: rw, h: rh, title: root.title, desc: root.role, kind: 'main', name: 'diagram-org-puncak',
        alt: `Struktur organisasi: ${root.title} membawahi ${units.map((u) => u.title).join(', ')}` });
      const busY = ry + rh + 0.2, uy = busY + 0.18;
      seg(s, M + W / 2, ry + rh, M + W / 2, busY, { color: GREY, name: 'diagram-org-garis-puncak' });
      if (n > 1) seg(s, M + uw / 2, busY, M + W - uw / 2, busY, { color: GREY, name: 'diagram-org-bus' });
      units.forEach((u, i) => {
        need(u && u.title, `orgChart: units[${i}].title`);
        const x = M + i * (uw + GAP);
        seg(s, x + uw / 2, busY, x + uw / 2, uy, { color: GREY, name: `diagram-org-garis-${i + 1}` });
        node(s, { x, y: uy, w: uw, h: rh, title: u.title, desc: u.role, kind: focus === i ? 'main' : 'node', name: `diagram-org-unit-${i + 1}` });
        if (u.items && u.items.length) ctx.bulletBox(s, `orgChart.${i}`, u.items, x, uy + rh + 0.15, uw, bottomOf(src) - uy - rh - 0.15, { fontSize: V.desc, paraSpaceAfter: 2, objectName: `diagram-org-unit-${i + 1}-isi` });
      });
      source(s, src);
      return finish(s);
    },

    // ---------- Infografis: sebelum/sesudah (V4 #12) ----------
    // rows: [{ name, before, after, delta? }] ≤ 4; focus: indeks baris yang deltanya merah (yang disebut judul).
    compare({ chapter, title: t, labels = ['Sebelum', 'Sesudah'], rows = [], focus, source: src, notes } = {}) {
      if (rows.length > 4) warn(`compare: ${rows.length} atribut; maksimal 4.`);
      const s = add('NEURON_ISI', notes);
      head(s, 'compare', chapter, t);
      const pw = 4.2, ph = bottomOf(src) - Z.bodyY, xs = [M, M + W - pw], rowH = Math.min(0.85, (ph - 0.55) / Math.max(1, rows.length));
      xs.forEach((x, k) => {
        s.addShape('rect', { x, y: Z.bodyY, w: pw, h: ph, fill: { color: ZEBRA }, line: { type: 'none' }, objectName: `perbandingan-panel-${k + 1}`,
          ...(k === 0 ? { altText: `Perbandingan ${labels[0]} dan ${labels[1]}: ${rows.map((r) => `${plain(r.name)} ${r.before} → ${r.after}`).join('; ')}` } : {}) });
        s.addText(labels[k], base({ x: x + 0.2, y: Z.bodyY + 0.15, w: pw - 0.4, h: 0.25, fontSize: V.node, bold: true, color: k ? INK : MUTED, valign: 'middle', objectName: `perbandingan-kepala-${k + 1}` }));
        rows.forEach((r, i) => {
          need(r && r.name, `compare: rows[${i}].name`);
          const y = Z.bodyY + 0.5 + i * rowH;
          s.addText(plain(r.name), base({ x: x + 0.2, y, w: pw - 0.4, h: 0.2, fontSize: V.desc, color: MUTED, valign: 'top', objectName: `perbandingan-${k + 1}-${i + 1}-nama` }));
          s.addText(nb(k ? r.after : r.before), base({ x: x + 0.2, y: y + 0.18, w: pw - 1.4, h: 0.4, fontSize: V.compare, bold: true, color: INK, valign: 'middle', objectName: `perbandingan-${k + 1}-${i + 1}-nilai` }));
          if (k && r.delta) s.addText(nb(r.delta), base({ x: x + pw - 1.2, y: y + 0.18, w: 1.0, h: 0.4, fontSize: V.node, bold: true, color: focus === i ? ACC : INK, align: 'right', valign: 'middle', objectName: `perbandingan-delta-${i + 1}` }));
        });
      });
      s.addShape('chevron', { x: M + W / 2 - 0.15, y: Z.bodyY + ph / 2 - 0.2, w: 0.3, h: 0.4, fill: { color: INK }, line: { type: 'none' }, objectName: 'perbandingan-panah' });
      source(s, src);
      return finish(s);
    },

    // ---------- Infografis: piramida konsep (tanpa angka) & corong data (batang) (V4 #14) ----------
    // pyramid: layers atas→bawah (3–5), focus = lapis yang dibahas, notes kanan opsional per lapis (descs).
    pyramid({ chapter, title: t, layers = [], descs = [], focus, source: src, notes } = {}) {
      const n = layers.length;
      if (n < 3 || n > 5) warn(`pyramid: ${n} lapis; 3–5 lapis.`);
      const s = add('NEURON_ISI', notes);
      head(s, 'pyramid', chapter, t);
      const ph = bottomOf(src) - Z.bodyY, lh = (ph - (n - 1) * 0.06) / n, baseW = 4.35, cx = M + baseW / 2;
      layers.forEach((ly, i) => {
        need(ly, `pyramid: layers[${i}]`);
        const top = baseW * (i / n), bot = baseW * ((i + 1) / n), y = Z.bodyY + i * (lh + 0.06), on = focus === i;
        if (i === 0) s.addShape('triangle', { x: cx - bot / 2, y, w: bot, h: lh, fill: { color: on ? INK : ZEBRA }, line: { type: 'none' }, objectName: 'piramida-1', altText: `Piramida (konsep, bukan jumlah): ${layers.join(' → ')}` });
        else s.addShape('trapezoid', { x: cx - bot / 2, y, w: bot, h: lh, fill: { color: on ? INK : ZEBRA }, line: { type: 'none' }, objectName: `piramida-${i + 1}`, adjustPoint: (bot - top) / 2 / Math.min(bot, lh) });
        const tw = Math.max(1.0, (top + bot) / 2 - 0.3);
        s.addText(plain(ly), base({ x: cx - tw / 2, y: y + (i === 0 ? lh * 0.35 : 0), w: tw, h: lh - (i === 0 ? lh * 0.35 : 0), fontSize: V.ribbon, bold: true, color: on ? WHITE : INK, align: 'center', valign: 'middle', objectName: `piramida-${i + 1}-teks` }));
        if (descs[i]) s.addText(rich(descs[i]), base({ x: M + baseW + 0.4, y, w: W - baseW - 0.4, h: lh, fontSize: V.desc + 2, color: TEXT, valign: 'middle', objectName: `piramida-${i + 1}-keterangan` }));
      });
      source(s, src);
      return finish(s);
    },
    // funnel: stages [{ name, value }] ≤ 6; batang rata kiri sebanding nilai; focus bawaan = penurunan terbesar.
    funnel({ chapter, title: t, stages = [], focus, numFmt = '#,##0', source: src, notes } = {}) {
      if (stages.length > 6) warn(`funnel: ${stages.length} tahap; maksimal 6.`);
      const s = add('NEURON_ISI', notes);
      head(s, 'funnel', chapter, t);
      const n = Math.max(1, stages.length), avail = bottomOf(src) - Z.bodyY - 0.1;
      const max = Math.max(...stages.map((x) => x.value)), bx = M + 1.9, bwMax = 5.4, gap = 0.16, bh = Math.min(0.45, (avail - (n - 1) * gap) / n);
      const conv = stages.map((x, i) => (i ? x.value / stages[i - 1].value : null));
      const f = focus !== undefined ? focus : conv.reduce((best, c, i) => (c !== null && (best < 0 || c < conv[best]) ? i : best), -1);
      stages.forEach((st, i) => {
        need(st && st.name, `funnel: stages[${i}].name`);
        const y = Z.bodyY + 0.1 + i * (bh + gap), w = Math.max(0.05, (st.value / max) * bwMax);
        s.addText(plain(st.name), base({ x: M, y, w: 1.8, h: bh, fontSize: V.desc, color: TEXT, valign: 'middle', objectName: `corong-${i + 1}-nama` }));
        s.addShape('rect', { x: bx, y, w, h: bh, fill: { color: i === f ? ACC : GREY }, line: { type: 'none' }, objectName: `corong-${i + 1}`,
          ...(i === 0 ? { altText: `Corong: ${stages.map((x) => `${plain(x.name)} ${formatNum(x.value, numFmt)}`).join(', ')}` } : {}) });
        s.addText(formatNum(st.value, numFmt), base({ x: bx + w + 0.08, y, w: 1.0, h: bh, fontSize: V.desc, bold: i === f, color: INK, valign: 'middle', objectName: `corong-${i + 1}-nilai` }));
        if (conv[i] !== null) s.addText(`→ ${formatNum(conv[i] * 100, '0')}%`, base({ x: M + W - 1.0, y, w: 1.0, h: bh, fontSize: V.desc, bold: i === f, color: MUTED, align: 'right', valign: 'middle', objectName: `corong-${i + 1}-konversi` }));
      });
      s.addText('konversi dari tahap sebelumnya', base({ x: M + W - 2.2, y: Z.bodyY - 0.2, w: 2.2, h: 0.18, fontSize: V.tag, color: MUTED, align: 'right', objectName: 'corong-keterangan' }));
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

    // ---------- Diagram teknis: arsitektur C4 konteks/kontainer & alur data (V4 #15–17) ----------
    // nodes: [{ id, title, tag?, desc?, kind: 'person'|'container'|'external'|'db'|'main', col: 0..3, row: 0..2, colSpan?, rowSpan?, focus? }]
    // links: [{ from, to, label, protocol?, focus?, dashed? }] — satu arah; garis lurus bila simpul sejajar, selain itu siku Z.
    // boundary: { label, cols: [a, b] } batas sistem putus-putus; caption: "Diagram kontainer · …"; columns: judul kolom
    // (alur data: ['Sumber', 'Proses', 'Simpan', 'Keluaran']); legend: kunci notasi C4 (bawaan: true bila tanpa columns).
    architecture({ chapter, title: t, caption, nodes = [], links = [], boundary, columns, legend, source: src, notes } = {}) {
      if (nodes.length > 9) warn(`architecture: ${nodes.length} elemen; maksimal 9 (6 kontainer + 3 luar) — pecah diagram.`);
      const showLegend = legend !== undefined ? legend : !columns;
      const s = add('NEURON_ISI', notes);
      head(s, 'architecture', chapter, t);
      // grid 4 kolom; celah antarkolom 0,8 / 1,0 / 1,2" untuk label garis 9 pt; baris 1–3
      const COLX = [0.5, 2.8, 5.3, 8.0], NW = 1.5;
      const R = Math.max(1, ...nodes.map((n) => (n.row || 0) + (n.rowSpan || 1)));
      if (R > 3) throw new Error('[neuron_theme] architecture: maksimal 3 baris (row 0..2).');
      const NH = R === 3 ? 0.58 : 0.8;
      const top = Z.bodyY + (caption ? 0.3 : 0) + (columns ? 0.3 : 0) + (boundary ? 0.32 : 0.05);
      const bottomLimit = bottomOf(src) - (showLegend ? 0.3 : 0) - (boundary ? 0.15 : 0);
      const step = R > 1 ? Math.min(1.25, (bottomLimit - top - NH) / (R - 1)) : 0;
      const pos = {};
      nodes.forEach((nd, i) => {
        need(nd && nd.id, `architecture: nodes[${i}].id`); need(nd.title, `architecture: nodes[${i}].title`);
        const cs = nd.colSpan || 1, rs = nd.rowSpan || 1;
        if (!(nd.col >= 0 && nd.col + cs <= 4 && nd.row >= 0 && nd.row + rs <= 3)) throw new Error(`[neuron_theme] architecture: ${nd.id} di luar grid (col 0..3, row 0..2).`);
        if (pos[nd.id]) throw new Error(`[neuron_theme] architecture: id "${nd.id}" ganda.`);
        const x = COLX[nd.col], y = top + nd.row * step;
        pos[nd.id] = { x, y, w: COLX[nd.col + cs - 1] + NW - x, h: (nd.row + rs - 1) * step + NH - nd.row * step, ...nd };
      });
      let hy = Z.bodyY;
      if (caption) { s.addText(caption, base({ x: M, y: hy, w: W, h: 0.22, fontSize: V.desc, bold: true, color: MUTED, valign: 'middle', objectName: 'diagram-judul' })); hy += 0.3; }
      if (columns) columns.forEach((c, i) => s.addText(c, base({ x: COLX[i], y: hy, w: NW, h: 0.22, fontSize: V.tag, bold: true, color: MUTED, valign: 'middle', objectName: `diagram-kolom-${i + 1}` })));
      if (boundary) {
        const [a, b] = boundary.cols || [1, 2];
        const bx = COLX[a] - 0.22, bw = COLX[b] + NW + 0.22 - bx, by = top - 0.32, bh = top + (R - 1) * step + NH + 0.15 - by;
        s.addShape('rect', { x: bx, y: by, w: bw, h: bh, fill: { type: 'none' }, line: { color: GREY, width: 1, dashType: 'dash' }, objectName: 'diagram-batas' });
        s.addText(boundary.label || '', base({ x: bx + 0.08, y: by + 0.04, w: bw - 0.16, h: 0.18, fontSize: V.tag, color: MUTED, valign: 'top', objectName: 'diagram-batas-label' }));
      }
      const alt = `Diagram: ${nodes.map((n) => n.title).join(', ')}. Aliran: ${links.map((l) => `${pos[l.from]?.title} → ${pos[l.to]?.title}${l.label ? ` (${l.label})` : ''}`).join('; ')}`;
      nodes.forEach((nd, i) => {
        const p = pos[nd.id];
        const kind = nd.kind === 'container' || nd.kind === 'person' ? 'node' : nd.kind || 'node';
        node(s, { ...p, title: nd.title, tag: nd.tag, desc: nd.desc, kind, focus: !!nd.focus, name: `diagram-${nd.id}`, alt: i === 0 ? alt : undefined });
      });
      const ov = (a1, a2, b1, b2) => [Math.max(a1, b1), Math.min(a2, b2)];
      links.forEach((l, i) => {
        const a = pos[l.from], b = pos[l.to];
        if (!a || !b) throw new Error(`[neuron_theme] architecture: link ${i + 1} merujuk simpul yang tidak ada (${l.from} → ${l.to}).`);
        const o = { color: l.focus ? ACC : MUTED, width: l.focus ? 1.5 : 1, arrow: true, dash: !!l.dashed, name: `diagram-aliran-${i + 1}` };
        const [vy1, vy2] = ov(a.y, a.y + a.h, b.y, b.y + b.h), [hx1, hx2] = ov(a.x, a.x + a.w, b.x, b.x + b.w);
        if (vy2 - vy1 > 0.2 && hx2 <= hx1) {                    // sejajar mendatar → garis lurus horizontal
          const ltr = b.x > a.x, y = (vy1 + vy2) / 2;
          const x1 = ltr ? a.x + a.w + 0.05 : a.x - 0.05, x2 = ltr ? b.x - 0.05 : b.x + b.w + 0.05;
          seg(s, x1, y, x2, y, o);
          const w = Math.abs(x2 - x1), x = Math.min(x1, x2);
          const needW = Math.max(plain(l.label || '').length, (l.protocol || '').length + 2) * 0.064;
          if (needW > w + 0.05) warn(`architecture: label garis "${l.label}" terlalu panjang untuk celah ${w.toFixed(2)}" — persingkat (± ${Math.floor(w / 0.064)} karakter).`);
          if (l.label) s.addText(l.label, base({ x, y: y - 0.21, w, h: 0.18, fontSize: V.tag, color: MUTED, align: 'center', valign: 'bottom', objectName: `diagram-aliran-${i + 1}-label` }));
          if (l.protocol) s.addText(`[${l.protocol}]`, base({ x, y: y + 0.03, w, h: 0.18, fontSize: V.tag, color: MUTED, align: 'center', valign: 'top', objectName: `diagram-aliran-${i + 1}-protokol` }));
        } else if (hx2 - hx1 > 0.2 && vy2 <= vy1) {             // sejajar menurun → garis lurus vertikal
          const down = b.y > a.y, x = (hx1 + hx2) / 2;
          const y1 = down ? a.y + a.h + 0.05 : a.y - 0.05, y2 = down ? b.y - 0.05 : b.y + b.h + 0.05;
          seg(s, x, y1, x, y2, o);
          const txt = [l.label && { text: l.label, options: { breakLine: !!l.protocol } }, l.protocol && { text: `[${l.protocol}]` }].filter(Boolean);
          if (txt.length) s.addText(txt, base({ x: x + 0.08, y: Math.min(y1, y2), w: 1.5, h: Math.abs(y2 - y1), fontSize: V.tag, color: MUTED, valign: 'middle', objectName: `diagram-aliran-${i + 1}-label` }));
        } else {                                                 // siku Z: keluar sisi samping, belok di tengah celah, masuk sisi samping tujuan
          const ltr = b.x > a.x, down = b.y > a.y;
          const y1 = a.y + a.h * (down ? 0.75 : 0.25), y2 = b.y + b.h / 2;
          const x1 = ltr ? a.x + a.w + 0.05 : a.x - 0.05, x2 = ltr ? b.x - 0.05 : b.x + b.w + 0.05;
          const gapL = ltr ? a.x + a.w : b.x + b.w, gapR = ltr ? b.x : a.x, xm = (gapL + gapR) / 2;
          path(s, [[x1, y1], [xm, y1], [xm, y2], [x2, y2]], o);
          const lw = Math.max(0.4, gapR - xm - 0.08);
          const txt = [l.label && { text: l.label, options: { breakLine: !!l.protocol } }, l.protocol && { text: `[${l.protocol}]` }].filter(Boolean);
          if (txt.length) s.addText(txt, base({ x: xm + 0.06, y: Math.min(y1, y2) + 0.08, w: lw, h: Math.abs(y2 - y1) - 0.16, fontSize: V.tag, color: MUTED, valign: 'middle', objectName: `diagram-aliran-${i + 1}-label` }));
        }
      });
      if (showLegend) {
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
