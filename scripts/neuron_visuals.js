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
    s.addShape(shape, { x, y, w, h, fill: { color: fill }, line, objectName: name, ...(alt ? { altText: alt } : {}) });
    const tc = dark ? WHITE : INK, sc = dark ? RULE : MUTED;
    const runs = [{ text: plain(title), options: { fontSize: V.node, bold: true, color: tc, breakLine: !!(tag || desc) } }];
    if (tag) runs.push({ text: `[${tag}]`, options: { fontSize: V.tag, color: sc, breakLine: !!desc } });
    if (desc) runs.push({ text: plain(desc), options: { fontSize: V.tag, color: sc } });
    const top = kind === 'db' ? 0.14 : 0.06;
    s.addText(runs, base({ x: x + 0.08, y: y + top, w: w - 0.16, h: h - top - 0.04, align: 'center', valign: 'middle', lineSpacingMultiple: 1.0, objectName: `${name}-teks` }));
  }
  const fmtID = (v, opts) => new Intl.NumberFormat('id-ID', opts).format(v);

  // Panel "Inti" satu baris di bawah diagram (opsional): implikasi diagram dalam satu kalimat, 14 pt tebal, isian F4F5F7.
  const KEY_H = 0.5;
  const keyPanel = (s, src, text) => {
    const y = bottomOf(src) - KEY_H;
    if (lines(text, W - 1.2, TYPE.key, true) > 1) warn('keyPoint: kalimat Inti lebih dari 1 baris — persingkat (± 75 karakter).');
    s.addShape('rect', { x: M, y, w: W, h: KEY_H, fill: { color: ZEBRA }, line: { type: 'none' }, objectName: 'panel-inti' });
    s.addText([{ text: 'Inti   ', options: { fontSize: TYPE.label, bold: true, color: MUTED } }, ...rich(text, { fontSize: TYPE.key, bold: true, color: INK })],
      base({ x: M + 0.22, y, w: W - 0.44, h: KEY_H, valign: 'middle', objectName: 'inti' }));
    return y - GAP;   // batas bawah baru untuk diagram
  };

  return {
    // ---------- Diagram proses: alur langkah (V4 #7) ----------
    // steps: [{ title, text }] 3–6; active: langkah yang dibahas (lingkaran merah + nama tebal); variant 'ribbon' = pita fase;
    // keyPoint: panel Inti satu baris di bawah (diagram lalu di posisi brief, y 1,75).
    process({ chapter, title: t, steps = [], active, variant = 'circles', keyPoint, source: src, notes } = {}) {
      if (steps.length < 3 || steps.length > 6) warn(`process: ${steps.length} langkah; idealnya 3–6 (pecah slide bila lebih).`);
      const s = add('NEURON_ISI', notes);
      head(s, 'process', chapter, t);
      const bottom = keyPoint ? keyPanel(s, src, keyPoint) : bottomOf(src);
      const n = Math.max(1, steps.length), w = (W - (n - 1) * GAP) / n;
      const maxDesc = Math.max(0, ...steps.map((x) => (x && x.text ? textH(x.text, w, V.desc) : 0)));
      const d = 0.36;
      // tanpa panel: blok (lingkaran + nama + deskripsi) di tengah vertikal; dengan panel: di posisi brief
      const cy = keyPoint ? Z.bodyY + 0.1 : Z.bodyY + Math.max(0.1, (bottom - Z.bodyY - (d + 0.62 + maxDesc)) / 2);
      steps.forEach((st, i) => {
        need(st && st.title, `process: steps[${i}].title`);
        const x = M + i * (w + GAP), on = active === i;
        if (variant === 'ribbon') {
          if (st.title.split(/\s+/).length > 2) warn(`process(ribbon): nama fase "${st.title}" > 2 kata — pakai variant 'circles'.`);
          s.addShape(i === 0 ? 'homePlate' : 'chevron', { x: x - (i ? 0.12 : 0), y: Z.bodyY, w: w + (i ? 0.12 : 0) + 0.12 - 0.06, h: 0.5,
            fill: { color: on ? INK : ZEBRA }, line: { type: 'none' }, objectName: `diagram-fase-${i + 1}`, ...(i === 0 ? { altText: `Alur ${n} fase: ${steps.map((x) => plain(x.title)).join(', ')}` } : {}) });
          s.addText(plain(st.title), base({ x: x + 0.12, y: Z.bodyY, w: w - 0.24, h: 0.5, fontSize: V.ribbon, bold: true, color: on ? WHITE : INK, align: 'center', valign: 'middle', objectName: `diagram-fase-${i + 1}-teks` }));
          if (st.text) s.addText(rich(st.text), base({ x, y: Z.bodyY + 0.7, w, h: bottom - Z.bodyY - 0.7, fontSize: V.desc + 1, color: TEXT, valign: 'top', lineSpacingMultiple: 1.15, objectName: `diagram-fase-${i + 1}-isi` }));
          return;
        }
        s.addShape('ellipse', { x, y: cy, w: d, h: d, fill: { color: on ? ACC : INK }, line: { type: 'none' }, objectName: `diagram-langkah-${i + 1}`,
          ...(i === 0 ? { altText: `Alur ${n} langkah: ${steps.map((x) => plain(x.title)).join(', lalu ')}` } : {}) });
        s.addText(String(i + 1), base({ x, y: cy, w: d, h: d, fontSize: V.num, bold: true, color: WHITE, align: 'center', valign: 'middle', objectName: `diagram-langkah-${i + 1}-no` }));
        if (i < n - 1) seg(s, x + d + 0.08, cy + d / 2, x + w + GAP - 0.08, cy + d / 2, { arrow: true, name: `diagram-panah-${i + 1}` });
        const ty = cy + d + 0.14;
        if (lines(st.title, w, V.node, true) > 2) warn(`process: nama langkah ${i + 1} lebih dari 2 baris.`);
        s.addText(plain(st.title), base({ x, y: ty, w, h: 0.42, fontSize: V.node, bold: true, color: INK, valign: 'top', lineSpacingMultiple: 1.0, objectName: `diagram-langkah-${i + 1}-nama` }));
        if (st.text) {
          if (ty + 0.48 + textH(st.text, w, V.desc) > bottom + 0.02) warn(`process: deskripsi langkah ${i + 1} terlalu panjang — persingkat (± 12 kata).`);
          s.addText(rich(st.text), base({ x, y: ty + 0.48, w, h: Math.max(0.3, bottom - ty - 0.48), fontSize: V.desc, color: MUTED, valign: 'top', lineSpacingMultiple: 1.15, objectName: `diagram-langkah-${i + 1}-isi` }));
        }
      });
      source(s, src);
      return finish(s);
    },

    // ---------- Diagram proses: siklus (V4 #8) ----------
    // steps: [{ title, text? }] 3–5; active: tahap yang dibahas (isian gelap); center: label tengah. Tanpa merah.
    cycle({ chapter, title: t, steps = [], active, center, source: src, notes } = {}) {
      const n = steps.length;
      if (n < 3 || n > 5) warn(`cycle: ${n} tahap; siklus 3–5 tahap (proses linear → pakai process).`);
      if (n < 2) throw new Error('[neuron_theme] cycle: minimal 2 tahap.');
      const s = add('NEURON_ISI', notes);
      head(s, 'cycle', chapter, t);
      const zoneH = bottomOf(src) - Z.bodyY, cx = M + W / 2, cy = Z.bodyY + zoneH / 2;
      const bw = n === 5 ? 1.8 : 2.2, bh = n === 5 ? 0.62 : 0.72;
      // pusat simpul: 4 = sudut persegi (searah jarum jam dari kiri-atas), 3 = segitiga, 5 = segi lima
      const rx = n === 4 ? 2.2 : 2.6, ry = (zoneH - bh) / 2 - 0.05;
      const centers = n === 4 ? [[cx - rx, cy - ry], [cx + rx, cy - ry], [cx + rx, cy + ry], [cx - rx, cy + ry]]
        : steps.map((_, i) => { const a = -Math.PI / 2 + (2 * Math.PI * i) / n; return [cx + rx * Math.cos(a), cy + ry * Math.sin(a)]; });
      const boxes = centers.map(([x, y]) => ({ x: x - bw / 2, y: y - bh / 2, w: bw, h: bh, cx: x, cy: y }));
      // titik potong garis antarpusat dengan tepi kotak (agar panah berhenti sebelum simpul)
      const clip = (b, dx, dy) => { const k = Math.min(dx ? (b.w / 2) / Math.abs(dx) : Infinity, dy ? (b.h / 2) / Math.abs(dy) : Infinity); return [b.cx + dx * k, b.cy + dy * k]; };
      boxes.forEach((b, i) => {
        const st = steps[i] || {}; need(st.title, `cycle: steps[${i}].title`);
        const on = active === i;
        s.addShape('rect', { x: b.x, y: b.y, w: b.w, h: b.h, fill: { color: on ? INK : ZEBRA }, line: { type: 'none' }, objectName: `diagram-tahap-${i + 1}`,
          ...(i === 0 ? { altText: `Siklus ${n} tahap: ${steps.map((x) => plain(x.title)).join(', lalu ')}, lalu kembali ke awal` } : {}) });
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
    // milestones: [{ date, label, status: 'done'|'now'|'plan' }] ≤ 6, atau phases: [{ name, start, end, status }] (indeks
    // periode, end eksklusif, boleh pecahan) dengan periods: ['Jan', …] dan today (indeks pecahan 0..periods.length, opsional).
    // keyPoint: panel Inti satu baris di bawah.
    timeline({ chapter, title: t, milestones, phases, periods = [], today, keyPoint, source: src, notes } = {}) {
      if (!milestones && !phases) throw new Error('[neuron_theme] timeline: isi milestones atau phases.');
      if (phases && !periods.length) throw new Error('[neuron_theme] timeline: phases butuh periods.');
      if (phases && today !== undefined && !(today >= 0 && today <= periods.length)) throw new Error(`[neuron_theme] timeline: today ${today} di luar 0..${periods.length}.`);
      const s = add('NEURON_ISI', notes);
      head(s, 'timeline', chapter, t);
      const bottom = keyPoint ? keyPanel(s, src, keyPoint) : bottomOf(src);
      if (milestones) {
        const n = Math.max(1, milestones.length);
        if (milestones.length > 6) warn(`timeline: ${milestones.length} tonggak; maksimal 6 — pecah per semester.`);
        // pita tanggal-garis-label (± 1,4") di tengah vertikal zona yang tersisa
        const ly = Z.bodyY + Math.max(0.55, (bottom - Z.bodyY - 1.4) / 2 + 0.55), slot = W / n;
        seg(s, M, ly, M + W, ly, { color: GREY, width: 1, name: 'diagram-linimasa-garis' });
        milestones.forEach((m, i) => {
          need(m && m.label, `timeline: milestones[${i}].label`);
          const x = M + slot * (i + 0.5), st = m.status || 'plan', d = st === 'now' ? 0.2 : 0.16;
          s.addShape('ellipse', { x: x - d / 2, y: ly - d / 2, w: d, h: d, fill: { color: st === 'now' ? ACC : st === 'done' ? MUTED : WHITE },
            line: st === 'plan' ? { color: GREY, width: 1 } : { type: 'none' }, objectName: `diagram-tonggak-${i + 1}`,
            ...(i === 0 ? { altText: `Linimasa: ${milestones.map((x) => `${x.date || ''} ${plain(x.label)}`).join('; ')}` } : {}) });
          s.addText(m.date || '', base({ x: x - slot / 2 + 0.05, y: ly - 0.5, w: slot - 0.1, h: 0.3, fontSize: V.desc, bold: true, color: INK, align: 'center', valign: 'bottom', objectName: `diagram-tonggak-${i + 1}-tanggal` }));
          if (lines(m.label, slot - 0.1, V.desc) > 3) warn(`timeline: label tonggak ${i + 1} lebih dari 3 baris.`);
          // "Sekarang" gelap tebal: titik merah sudah menjadi satu-satunya sorotan
          s.addText([{ text: plain(m.label), options: { breakLine: st === 'now' } }, ...(st === 'now' ? [{ text: 'Sekarang', options: { fontSize: V.tag, bold: true, color: INK } }] : [])],
            base({ x: x - slot / 2 + 0.05, y: ly + 0.2, w: slot - 0.1, h: 0.75, fontSize: V.desc, color: TEXT, align: 'center', valign: 'top', lineSpacingMultiple: 1.1, objectName: `diagram-tonggak-${i + 1}-label` }));
        });
      } else {
        if (phases.length > 6) warn(`timeline: ${phases.length} fase; maksimal 6 baris.`);
        if (!periods.length) throw new Error('[neuron_theme] timeline: phases butuh periods.');
        if (today !== undefined && !(today >= 0 && today <= periods.length)) throw new Error(`[neuron_theme] timeline: today ${today} di luar 0..${periods.length}.`);
        const gx = M + 2.1, gw = M + W - gx, cwid = gw / periods.length, rh = 0.28, gap = 0.14;
        // kunci status (hanya status yang dipakai) di baris kepala kiri, di atas nama fase (legenda di atas plot)
        const used = ['done', 'active', 'plan'].filter((k) => phases.some((p) => (p.status || 'plan') === k));
        const KEY = { done: [MUTED, 'selesai'], active: [ACC, 'berjalan'], plan: [GREY, 'rencana'] };
        used.forEach((k, i) => {
          s.addShape('rect', { x: M + i * 0.66, y: Z.bodyY + Math.max(0, (bottom - Z.bodyY - (0.3 + phases.length * (rh + gap) + 0.25)) / 2 - 0.1) + 0.05, w: 0.12, h: 0.12, fill: { color: KEY[k][0] }, line: { type: 'none' }, objectName: `diagram-kunci-${k}` });
          s.addText(KEY[k][1], base({ x: M + i * 0.66 + 0.16, y: Z.bodyY + Math.max(0, (bottom - Z.bodyY - (0.3 + phases.length * (rh + gap) + 0.25)) / 2 - 0.1), w: 0.5, h: 0.22, fontSize: V.tag, color: MUTED, valign: 'middle', objectName: `diagram-kunci-${k}-teks` }));
        });
        // kepala periode + baris fase sebagai satu blok di tengah vertikal zona isi
        const blockH = 0.3 + phases.length * (rh + gap) + 0.25;
        const hy = Z.bodyY + Math.max(0, (bottom - Z.bodyY - blockH) / 2 - 0.1), y0 = hy + 0.3;
        periods.forEach((p, i) => s.addText(p, base({ x: gx + i * cwid, y: hy, w: cwid, h: 0.22, fontSize: V.tag, color: MUTED, align: 'center', valign: 'middle', objectName: `diagram-periode-${i + 1}` })));
        phases.forEach((ph, i) => {
          need(ph && ph.name, `timeline: phases[${i}].name`);
          if (!(ph.start >= 0 && ph.end > ph.start && ph.end <= periods.length)) throw new Error(`[neuron_theme] timeline: fase ${i + 1} start/end di luar 0..${periods.length}.`);
          const y = y0 + i * (rh + gap), st = ph.status || 'plan';
          s.addText(plain(ph.name), base({ x: M, y, w: 2.0, h: rh, fontSize: V.desc, bold: st === 'active', color: TEXT, valign: 'middle', objectName: `diagram-fase-${i + 1}-nama` }));
          s.addShape('rect', { x: gx + ph.start * cwid + 0.02, y, w: (ph.end - ph.start) * cwid - 0.04, h: rh, fill: { color: st === 'active' ? ACC : st === 'done' ? MUTED : GREY }, line: { type: 'none' },
            objectName: `diagram-fase-${i + 1}`, ...(i === 0 ? { altText: `Jadwal: ${phases.map((x) => `${plain(x.name)} ${periods[Math.floor(x.start)]}–${periods[Math.ceil(x.end) - 1]}`).join('; ')}` } : {}) });
        });
        if (today !== undefined) {
          const x = gx + today * cwid, yb = y0 + phases.length * (rh + gap) - gap + 0.08;
          seg(s, x, y0 - 0.06, x, yb, { color: MUTED, width: 1, dash: true, name: 'diagram-hari-ini' });
          s.addText('Hari ini', base({ x: x - 0.5, y: yb + 0.02, w: 1.0, h: 0.18, fontSize: V.tag, color: MUTED, align: 'center', objectName: 'diagram-hari-ini-label' }));
        }
      }
      source(s, src);
      return finish(s);
    },

    // ---------- Diagram proses: struktur organisasi (V4 #10) ----------
    // root: { title, role }; units: [{ title, role?, items?: [teks tingkat 3] }] ≤ 6; focus: unit yang dibahas (isian gelap,
    // pakai hanya bila judul menyebut unit itu); keyPoint: panel Inti satu baris di bawah.
    orgChart({ chapter, title: t, root = {}, units = [], focus, keyPoint, source: src, notes } = {}) {
      need(root.title, 'orgChart: root.title');
      if (units.length > 6) warn(`orgChart: ${units.length} unit; maksimal 6 per tingkat — pecah per divisi.`);
      const s = add('NEURON_ISI', notes);
      head(s, 'orgChart', chapter, t);
      const bottom = keyPoint ? keyPanel(s, src, keyPoint) : bottomOf(src);
      const n = Math.max(1, units.length), uw = (W - (n - 1) * GAP) / n, rw = 2.2, rh = 0.55;
      const maxItems = Math.max(0, ...units.map((u) => (u && u.items ? u.items.length : 0)));
      const blockH = rh + 0.38 + rh + (maxItems ? 0.15 + maxItems * 0.26 : 0);
      const ry = Z.bodyY + Math.max(0, (bottom - Z.bodyY - blockH) / 2), rx = M + (W - rw) / 2;
      node(s, { x: rx, y: ry, w: rw, h: rh, title: root.title, desc: root.role, kind: 'main', name: 'diagram-org-puncak',
        alt: `Struktur organisasi: ${root.title} membawahi ${units.map((u) => u && u.title).join(', ')}` });
      if (!units.length) { source(s, src); return finish(s); }
      const busY = ry + rh + 0.2, uy = busY + 0.18;
      seg(s, M + W / 2, ry + rh, M + W / 2, busY, { color: GREY, name: 'diagram-org-garis-puncak' });
      if (n > 1) seg(s, M + uw / 2, busY, M + W - uw / 2, busY, { color: GREY, name: 'diagram-org-bus' });
      units.forEach((u, i) => {
        need(u && u.title, `orgChart: units[${i}].title`);
        const x = M + i * (uw + GAP);
        seg(s, x + uw / 2, busY, x + uw / 2, uy, { color: GREY, name: `diagram-org-garis-${i + 1}` });
        node(s, { x, y: uy, w: uw, h: rh, title: u.title, desc: u.role, kind: focus === i ? 'main' : 'node', name: `diagram-org-unit-${i + 1}` });
        if (u.items && u.items.length) ctx.bulletBox(s, `orgChart.${i}`, u.items, x, uy + rh + 0.15, uw, Math.max(0.3, bottom - uy - rh - 0.15), { fontSize: V.desc, paraSpaceAfter: 2, objectName: `diagram-org-unit-${i + 1}-isi` });
      });
      source(s, src);
      return finish(s);
    },

    // ---------- Infografis: sebelum/sesudah (V4 #12) ----------
    // rows: [{ name, before, after, delta? }] ≤ 4 (atribut sama di kedua sisi); focus: baris yang deltanya merah (yang disebut judul).
    compare({ chapter, title: t, labels = ['Sebelum', 'Sesudah'], rows = [], focus, source: src, notes } = {}) {
      if (rows.length > 4) warn(`compare: ${rows.length} atribut; maksimal 4.`);
      rows.forEach((r, i) => { need(r && r.name, `compare: rows[${i}].name`); need(r.before, `compare: rows[${i}].before`); need(r.after, `compare: rows[${i}].after`); });
      const s = add('NEURON_ISI', notes);
      head(s, 'compare', chapter, t);
      const pw = 4.2, ph = bottomOf(src) - Z.bodyY, xs = [M, M + W - pw], rowH = Math.min(0.85, (ph - 0.55) / Math.max(1, rows.length));
      xs.forEach((x, k) => {
        s.addShape('rect', { x, y: Z.bodyY, w: pw, h: ph, fill: { color: ZEBRA }, line: { type: 'none' }, objectName: `perbandingan-panel-${k + 1}`,
          ...(k === 0 ? { altText: `Perbandingan ${labels[0]} dan ${labels[1]}: ${rows.map((r) => `${plain(r.name)} dari ${r.before} menjadi ${r.after}`).join('; ')}` } : {}) });
        s.addText(labels[k], base({ x: x + 0.2, y: Z.bodyY + 0.15, w: pw - 0.4, h: 0.25, fontSize: V.node, bold: true, color: k ? INK : MUTED, valign: 'middle', objectName: `perbandingan-kepala-${k + 1}` }));
        rows.forEach((r, i) => {
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
    // pyramid: layers atas→bawah (3–5), focus = lapis yang dibahas, descs = keterangan kanan per lapis.
    // Lapis digambar sebagai bentuk bebas (custGeom) dengan SATU kemiringan agar tepi piramida segaris (preset trapesium
    // pptxgenjs tidak bisa diatur kemiringannya).
    pyramid({ chapter, title: t, layers = [], descs = [], focus, source: src, notes } = {}) {
      const n = layers.length;
      if (n < 3 || n > 5) warn(`pyramid: ${n} lapis; 3–5 lapis.`);
      if (!n) throw new Error('[neuron_theme] pyramid: layers wajib diisi.');
      const s = add('NEURON_ISI', notes);
      head(s, 'pyramid', chapter, t);
      const G = 0.06, ph = bottomOf(src) - Z.bodyY, lh = (ph - (n - 1) * G) / n, baseW = 4.35, cx = M + baseW / 2;
      const widthAt = (y) => (baseW * (y - Z.bodyY)) / ph;   // lebar pada ketinggian y (puncak di Z.bodyY)
      layers.forEach((ly, i) => {
        need(ly, `pyramid: layers[${i}]`);
        const y = Z.bodyY + i * (lh + G), on = focus === i, top = widthAt(y), bot = widthAt(y + lh), inset = (bot - top) / 2;
        s.addShape('custGeom', { x: cx - bot / 2, y, w: bot, h: lh, fill: { color: on ? INK : ZEBRA }, line: { type: 'none' }, objectName: `piramida-${i + 1}`,
          points: top < 0.01 ? [{ x: bot / 2, y: 0 }, { x: bot, y: lh }, { x: 0, y: lh }, { close: true }] : [{ x: inset, y: 0 }, { x: bot - inset, y: 0 }, { x: bot, y: lh }, { x: 0, y: lh }, { close: true }],
          ...(i === 0 ? { altText: `Piramida (konsep, bukan jumlah): ${layers.map(plain).join(', ')}` } : {}) });
        const tw = Math.max(1.0, (top + bot) / 2 - 0.3), tOff = i === 0 ? lh * 0.35 : 0;
        s.addText(plain(ly), base({ x: cx - tw / 2, y: y + tOff, w: tw, h: lh - tOff, fontSize: V.ribbon, bold: true, color: on ? WHITE : INK, align: 'center', valign: 'middle', objectName: `piramida-${i + 1}-teks` }));
        if (descs[i]) s.addText(rich(descs[i]), base({ x: M + baseW + 0.4, y, w: W - baseW - 0.4, h: lh, fontSize: V.desc + 2, color: TEXT, valign: 'middle', objectName: `piramida-${i + 1}-keterangan` }));
      });
      source(s, src);
      return finish(s);
    },
    // funnel: stages [{ name, value }] ≤ 6; batang rata kiri sebanding nilai; focus bawaan = konversi terendah.
    funnel({ chapter, title: t, stages = [], focus, numFmt = '#,##0', source: src, notes } = {}) {
      if (stages.length > 6) warn(`funnel: ${stages.length} tahap; maksimal 6.`);
      stages.forEach((st, i) => { need(st && st.name, `funnel: stages[${i}].name`); if (!(typeof st.value === 'number' && st.value >= 0)) throw new Error(`[neuron_theme] funnel: stages[${i}].value harus angka ≥ 0.`); });
      const max = Math.max(0, ...stages.map((x) => x.value));
      if (!(max > 0)) throw new Error('[neuron_theme] funnel: minimal satu tahap bernilai > 0.');
      if (focus !== undefined && !(Number.isInteger(focus) && focus >= -1 && focus < stages.length)) throw new Error(`[neuron_theme] funnel: focus ${focus} di luar rentang.`);
      const s = add('NEURON_ISI', notes);
      head(s, 'funnel', chapter, t);
      const n = Math.max(1, stages.length), avail = bottomOf(src) - Z.bodyY - 0.3;
      const bx = M + 1.9, bwMax = 5.4, gap = 0.16, bh = Math.min(0.45, (avail - (n - 1) * gap) / n);
      const y0 = Z.bodyY + 0.3 + Math.max(0, (avail - (n * bh + (n - 1) * gap)) / 2);
      const conv = stages.map((x, i) => (i && stages[i - 1].value > 0 ? x.value / stages[i - 1].value : null));
      const f = focus !== undefined ? focus : conv.reduce((best, c, i) => (c !== null && (best < 0 || c < conv[best]) ? i : best), -1);
      s.addText('konversi dari tahap sebelumnya', base({ x: M + W - 2.4, y: Z.bodyY, w: 2.4, h: 0.2, fontSize: V.tag, color: MUTED, align: 'right', valign: 'middle', objectName: 'corong-keterangan' }));
      stages.forEach((st, i) => {
        const y = y0 + i * (bh + gap), w = Math.max(0.03, (st.value / max) * bwMax);
        s.addText(plain(st.name), base({ x: M, y, w: 1.8, h: bh, fontSize: V.desc, color: TEXT, valign: 'middle', objectName: `corong-${i + 1}-nama` }));
        s.addShape('rect', { x: bx, y, w, h: bh, fill: { color: i === f ? ACC : GREY }, line: { type: 'none' }, objectName: `corong-${i + 1}`,
          ...(i === 0 ? { altText: `Corong: ${stages.map((x) => `${plain(x.name)} ${formatNum(x.value, numFmt)}`).join(', ')}` } : {}) });
        s.addText(formatNum(st.value, numFmt), base({ x: bx + w + 0.08, y, w: 1.0, h: bh, fontSize: V.desc, bold: i === f, color: INK, valign: 'middle', objectName: `corong-${i + 1}-nilai` }));
        if (i) s.addText(conv[i] === null ? '–' : `${formatNum(conv[i] * 100, '0')}%`, base({ x: M + W - 0.8, y, w: 0.8, h: bh, fontSize: V.desc, bold: i === f, color: MUTED, align: 'right', valign: 'middle', objectName: `corong-${i + 1}-konversi` }));
      });
      source(s, src);
      return finish(s);
    },

    // ---------- Infografis: matriks 2×2 (V4 #13) ----------
    // axes: { x: { label, low, high }, y: { label, low, high } }; quadrants: { tl, tr, bl, br } (nama); best: kuadran berisian.
    // points: [{ label, x, y }] dengan x,y 0..1; focus: titik merah; explanation: teks atau daftar poin kanan (14 pt).
    matrix({ chapter, title: t, axes = {}, quadrants = {}, best = 'tr', points = [], focus, explanation, source: src, notes } = {}) {
      if (points.length > 10) warn(`matrix: ${points.length} titik; maksimal 10.`);
      const s = add('NEURON_ISI', notes);
      head(s, 'matrix', chapter, t);
      const ax = axes.x || {}, ay = axes.y || {};
      const px = M + 0.55, py = Z.bodyY + 0.3, pw = explanation ? 4.6 : 8.0, ph = bottomOf(src) - py - 0.35, hw = pw / 2, hh = ph / 2;
      const q = { tl: [px, py], tr: [px + hw, py], bl: [px, py + hh], br: [px + hw, py + hh] };
      Object.entries(q).forEach(([k, [x, y]]) => {
        s.addShape('rect', { x, y, w: hw, h: hh, fill: { color: k === best ? ZEBRA : WHITE }, line: { color: RULE, width: 1 }, objectName: `matriks-${k}`,
          ...(k === 'tl' ? { altText: `Matriks 2×2: ${ay.label || 'sumbu tegak'} terhadap ${ax.label || 'sumbu datar'}; ${points.map((p) => plain(p.label)).join(', ')}` } : {}) });
        if (quadrants[k]) s.addText(quadrants[k], base({ x: x + 0.08, y: y + 0.06, w: hw - 0.16, h: 0.2, fontSize: V.desc, bold: true, color: MUTED, align: k.endsWith('r') ? 'right' : 'left', valign: 'top', objectName: `matriks-${k}-nama` }));
      });
      // label sumbu tanpa panah (glyph panah tidak selalu ada di Arial pengganti): tegak di kiri-atas, ujung rendah/tinggi di tepi plot
      if (ay.label) s.addText(ay.label, base({ x: M, y: Z.bodyY, w: 4, h: 0.22, fontSize: V.desc, bold: true, color: MUTED, valign: 'middle', objectName: 'matriks-sumbu-y' }));
      if (ay.high) s.addText(ay.high, base({ x: M, y: py, w: 0.5, h: 0.2, fontSize: V.tag, color: MUTED, align: 'right', objectName: 'matriks-y-tinggi' }));
      if (ay.low) s.addText(ay.low, base({ x: M, y: py + ph - 0.2, w: 0.5, h: 0.2, fontSize: V.tag, color: MUTED, align: 'right', objectName: 'matriks-y-rendah' }));
      const xy = py + ph + 0.08;
      if (ax.low) s.addText(ax.low, base({ x: px, y: xy, w: 1.2, h: 0.22, fontSize: V.tag, color: MUTED, valign: 'middle', objectName: 'matriks-x-rendah' }));
      if (ax.label) s.addText(ax.label, base({ x: px + hw - 1.5, y: xy, w: 3.0, h: 0.22, fontSize: V.desc, bold: true, color: MUTED, align: 'center', valign: 'middle', objectName: 'matriks-sumbu-x' }));
      if (ax.high) s.addText(ax.high, base({ x: px + pw - 1.2, y: xy, w: 1.2, h: 0.22, fontSize: V.tag, color: MUTED, align: 'right', valign: 'middle', objectName: 'matriks-x-tinggi' }));
      points.forEach((p, i) => {
        if (!(p.x >= 0 && p.x <= 1 && p.y >= 0 && p.y <= 1)) throw new Error(`[neuron_theme] matrix: titik ${i + 1} di luar 0..1.`);
        const d = 0.14, cx = px + p.x * pw, cy = py + (1 - p.y) * ph, on = focus === i;
        s.addShape('ellipse', { x: cx - d / 2, y: cy - d / 2, w: d, h: d, fill: { color: on ? ACC : INK }, line: { type: 'none' }, objectName: `matriks-titik-${i + 1}` });
        // label tetap di dalam kuadran titiknya: kanan bila muat, kiri bila muat, selain itu di bawah titik (boleh 2 baris)
        const lw = Math.min(1.8, (plain(p.label).length * V.desc * (on ? 0.55 : 0.5)) / 72 + 0.05);
        const qx0 = p.x < 0.5 ? px : px + hw, qx1 = p.x < 0.5 ? px + hw : px + pw;
        let bx, al, by = cy - 0.11, bh = 0.22;
        if (cx + 0.1 + lw <= qx1 - 0.04) { bx = cx + 0.1; al = 'left'; }
        else if (cx - 0.1 - lw >= qx0 + 0.04) { bx = cx - 0.1 - lw; al = 'right'; }
        else {
          // di bawah atau di atas titik, tetap di dalam kuadran tegaknya
          const bw = Math.min(qx1 - qx0 - 0.1, Math.max(1.0, lw)), qy0 = p.y >= 0.5 ? py : py + hh, qy1 = p.y >= 0.5 ? py + hh : py + ph;
          bx = Math.max(qx0 + 0.05, Math.min(cx - bw / 2, qx1 - 0.05 - bw)); al = 'center'; bh = 0.38;
          by = cy + 0.1 + bh <= qy1 - 0.03 ? cy + 0.1 : cy - 0.1 - bh;
          if (by < qy0 + 0.02) warn(`matrix: label "${plain(p.label)}" tidak muat di kuadrannya — persingkat atau geser titik.`);
        }
        s.addText(plain(p.label), base({ x: bx, y: by, w: al === 'center' ? Math.min(qx1 - qx0 - 0.1, Math.max(1.0, lw)) : lw, h: bh, fontSize: V.desc, bold: on, color: TEXT, align: al, valign: al === 'center' ? 'top' : 'middle', objectName: `matriks-titik-${i + 1}-label` }));
      });
      if (explanation) {
        const ex = px + pw + 0.45, ew = M + W - ex;
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
      // grid 4 kolom; celah antarkolom 0,8 / 1,0 / 1,2" untuk label garis 9 pt; baris 1–3
      const COLX = [0.5, 2.8, 5.3, 8.0], NW = 1.5;
      const R = Math.max(1, ...nodes.map((n) => (n.row || 0) + (n.rowSpan || 1)));
      if (R > 3) throw new Error('[neuron_theme] architecture: maksimal 3 baris (row 0..2).');
      const top = Z.bodyY + (caption ? 0.26 : 0) + (columns ? 0.28 : 0) + (boundary ? 0.28 : 0.05);
      const bottomLimit = bottomOf(src) - (showLegend ? 0.25 : 0) - (boundary ? 0.1 : 0);
      // tinggi simpul & jarak baris: celah antarbaris minimal 0,35" (panah + label); simpul diperkecil bila perlu
      const GAPMIN = 0.35;
      let NH = R === 3 ? 0.58 : 0.8;
      if (R > 1 && (bottomLimit - top - R * NH) / (R - 1) < GAPMIN) NH = (bottomLimit - top - (R - 1) * GAPMIN) / R;
      if (NH < 0.4) throw new Error(`[neuron_theme] architecture: ${R} baris tidak muat (simpul ${NH.toFixed(2)}") — kurangi baris, atau hilangkan columns/source (C4 kontainer 3 baris muat dengan caption + batas + kunci).`);
      if (boundary) {
        const [a, b] = boundary.cols || [1, 2];
        if (!(Number.isInteger(a) && Number.isInteger(b) && a >= 0 && b <= 3 && a <= b)) throw new Error(`[neuron_theme] architecture: boundary.cols [${a}, ${b}] harus 0 ≤ a ≤ b ≤ 3.`);
      }
      const step = R > 1 ? Math.min(1.25, (bottomLimit - top - NH) / (R - 1)) : 0;
      const pos = {};
      nodes.forEach((nd, i) => {
        need(nd && nd.id, `architecture: nodes[${i}].id`); need(nd.title, `architecture: nodes[${i}].title`);
        const cs = nd.colSpan || 1, rs = nd.rowSpan || 1;
        if (!(Number.isInteger(nd.col) && Number.isInteger(nd.row) && nd.col >= 0 && nd.col + cs <= 4 && nd.row >= 0 && nd.row + rs <= 3)) throw new Error(`[neuron_theme] architecture: ${nd.id} di luar grid (col 0..3, row 0..2).`);
        if (pos[nd.id]) throw new Error(`[neuron_theme] architecture: id "${nd.id}" ganda.`);
        // "sistem kami" yang membentang beberapa sel (diagram konteks) dibuat ringkas (maks. 2,4" × 1,0") di tengah bentangnya —
        // tingginya hanya diringkas bila setiap simpul tetangga di kiri/kanan masih sejajar ≥ 0,2" (garis tetap lurus);
        // simpul lain mengisi bentangnya penuh agar garis dari tiap baris tetap lurus (alur data)
        const sx = COLX[nd.col], sy = top + nd.row * step, sw = COLX[nd.col + cs - 1] + NW - sx, sh = (rs - 1) * step + NH;
        const isMain = nd.kind === 'main' && (cs > 1 || rs > 1);
        const w = isMain ? Math.min(sw, 2.4) : sw;
        let h = sh;
        if (isMain && rs > 1) {
          const ch = Math.min(sh, Math.max(NH, 1.0)), cy0 = sy + (sh - ch) / 2;
          const peers = links.filter((l) => l.from === nd.id || l.to === nd.id).map((l) => nodes.find((m) => m.id === (l.from === nd.id ? l.to : l.from))).filter(Boolean);
          const ok = peers.every((m) => { const my = top + m.row * step, mh = ((m.rowSpan || 1) - 1) * step + NH; return Math.min(cy0 + ch, my + mh) - Math.max(cy0, my) >= 0.2; });
          if (ok) h = ch;
        }
        pos[nd.id] = { ...nd, x: sx + (sw - w) / 2, y: sy + (sh - h) / 2, w, h };
      });
      links.forEach((l, i) => {
        if (!pos[l.from] || !pos[l.to]) throw new Error(`[neuron_theme] architecture: link ${i + 1} merujuk simpul yang tidak ada (${l.from} → ${l.to}).`);
        if (l.from === l.to) throw new Error(`[neuron_theme] architecture: link ${i + 1} dari "${l.from}" ke dirinya sendiri.`);
      });
      const s = add('NEURON_ISI', notes);
      head(s, 'architecture', chapter, t);
      const ids = Object.keys(pos);
      ids.forEach((a, i) => ids.slice(i + 1).forEach((b) => {
        const A = pos[a], B = pos[b];
        if (A.x < B.x + B.w && B.x < A.x + A.w && A.y < B.y + B.h && B.y < A.y + A.h) warn(`architecture: simpul "${a}" dan "${b}" bertumpuk — pindahkan col/row.`);
      }));
      let hy = Z.bodyY;
      if (caption) { s.addText(caption, base({ x: M, y: hy, w: W, h: 0.22, fontSize: V.desc, bold: true, color: MUTED, valign: 'middle', objectName: 'diagram-judul' })); hy += 0.26; }
      if (columns) columns.forEach((c, i) => s.addText(c, base({ x: COLX[i], y: hy, w: NW, h: 0.22, fontSize: V.tag, bold: true, color: MUTED, valign: 'middle', objectName: `diagram-kolom-${i + 1}` })));
      let bL = null, bR = null;
      if (boundary) {
        const [a, b] = boundary.cols || [1, 2];
        const bx = COLX[a] - 0.12, bw = COLX[b] + NW + 0.12 - bx, by = top - 0.28, bh = top + (R - 1) * step + NH + 0.1 - by;
        bL = bx; bR = bx + bw;
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
        if (l.from === l.to) throw new Error(`[neuron_theme] architecture: link ${i + 1} dari "${l.from}" ke dirinya sendiri.`);
        const o = { color: l.focus ? ACC : MUTED, width: l.focus ? 1.5 : 1, arrow: true, dash: !!l.dashed, name: `diagram-aliran-${i + 1}` };
        const [vy1, vy2] = ov(a.y, a.y + a.h, b.y, b.y + b.h), [hx1, hx2] = ov(a.x, a.x + a.w, b.x, b.x + b.w);
        if (vy2 - vy1 > 0.2 && hx2 <= hx1) {                    // sejajar mendatar → garis lurus horizontal
          const ltr = b.x > a.x, y = (vy1 + vy2) / 2;
          const x1 = ltr ? a.x + a.w + 0.05 : a.x - 0.05, x2 = ltr ? b.x - 0.05 : b.x + b.w + 0.05;
          seg(s, x1, y, x2, y, o);
          const lo = Math.min(x1, x2), hi = Math.max(x1, x2);
          ids.forEach((k) => { const N = pos[k]; if (N !== a && N !== b && N.x < hi && N.x + N.w > lo && N.y < y && N.y + N.h > y) warn(`architecture: garis ${l.from} → ${l.to} melintasi simpul "${k}" — atur ulang col/row.`); });
          // label tidak boleh tertimpa garis batas sistem: pakai bagian garis di sisi batas yang lebih panjang
          let x = lo, w = hi - lo;
          for (const bx of [bL, bR]) if (bx !== null && bx > x + 0.05 && bx < x + w - 0.05) { if (bx - x >= x + w - bx) w = bx - x - 0.05; else { w = x + w - bx - 0.05; x = bx + 0.05; } }
          const needW = Math.max(plain(l.label || '').length, (l.protocol || '').length + 2) * 0.064;
          if (needW > w + 0.05) warn(`architecture: label garis "${l.label}" terlalu panjang untuk celah ${w.toFixed(2)}" — persingkat (± ${Math.floor(w / 0.064)} karakter).`);
          if (l.label) s.addText(l.label, base({ x, y: y - 0.21, w, h: 0.18, fontSize: V.tag, color: MUTED, align: 'center', valign: 'bottom', objectName: `diagram-aliran-${i + 1}-label` }));
          if (l.protocol) s.addText(`[${l.protocol}]`, base({ x, y: y + 0.03, w, h: 0.18, fontSize: V.tag, color: MUTED, align: 'center', valign: 'top', objectName: `diagram-aliran-${i + 1}-protokol` }));
        } else if (hx2 - hx1 > 0.2 && vy2 <= vy1) {             // sejajar menurun → garis lurus vertikal
          const down = b.y > a.y, x = (hx1 + hx2) / 2;
          const y1 = down ? a.y + a.h + 0.05 : a.y - 0.05, y2 = down ? b.y - 0.05 : b.y + b.h + 0.05;
          if (Math.abs(y2 - y1) < 0.15 || (down ? y2 < y1 : y2 > y1)) warn(`architecture: garis ${l.from} → ${l.to} terlalu pendek — beri jarak baris lebih besar.`);
          seg(s, x, y1, x, y2, o);
          const txt = [l.label && { text: l.label, options: { breakLine: !!l.protocol } }, l.protocol && { text: `[${l.protocol}]` }].filter(Boolean);
          // label di kanan garis, dibatasi garis batas sistem; bila sempit (< 0,9") pindah ke kiri garis
          const need_ = Math.max(plain(l.label || '').length, (l.protocol || '').length + 2) * 0.064 + 0.05;
          let lx = x + 0.08, lw = 1.5, al = 'left';
          if (bR !== null && x < bR && x + 0.08 + need_ > bR - 0.05) { lw = Math.max(0, bR - 0.05 - lx); if (lw < need_) { lw = Math.min(1.5, x - 0.08 - ((bL ?? M) + 0.05)); lx = x - 0.08 - lw; al = 'right'; } }
          if (lw < need_) warn(`architecture: label garis "${l.label}" tidak muat di samping garis tegak — persingkat.`);
          if (txt.length) s.addText(txt, base({ x: lx, y: Math.min(y1, y2), w: Math.max(0.3, lw), h: Math.abs(y2 - y1), fontSize: V.tag, color: MUTED, align: al, valign: 'middle', objectName: `diagram-aliran-${i + 1}-label` }));
        } else {                                                 // siku Z: keluar sisi samping, belok di tengah celah, masuk sisi samping tujuan
          const ltr = b.x > a.x, down = b.y > a.y;
          const y1 = a.y + a.h * (down ? 0.75 : 0.25), y2 = b.y + b.h / 2;
          const x1 = ltr ? a.x + a.w + 0.05 : a.x - 0.05, x2 = ltr ? b.x - 0.05 : b.x + b.w + 0.05;
          const gapL = ltr ? a.x + a.w : b.x + b.w, gapR = ltr ? b.x : a.x;
          // belokan dekat simpul asal (0,25") agar segmen mendatar terakhir cukup panjang untuk label di atasnya
          const xm = gapR - gapL > 0.9 ? (ltr ? gapL + 0.25 : gapR - 0.25) : (gapL + gapR) / 2;
          path(s, [[x1, y1], [xm, y1], [xm, y2], [x2, y2]], o);
          // label: di atas segmen mendatar terakhir bila cukup panjang (≥ 0,7"), selain itu di samping segmen tegak
          const tail = Math.abs(x2 - xm), txt = [l.label && { text: l.label, options: { breakLine: !!l.protocol } }, l.protocol && { text: `[${l.protocol}]` }].filter(Boolean);
          if (txt.length && tail >= 0.7) s.addText(txt, base({ x: Math.min(xm, x2) + 0.04, y: y2 - (l.protocol ? 0.4 : 0.22), w: tail - 0.08, h: l.protocol ? 0.38 : 0.2, fontSize: V.tag, color: MUTED, align: 'center', valign: 'bottom', objectName: `diagram-aliran-${i + 1}-label` }));
          else if (txt.length) s.addText(txt, base({ x: xm + 0.06, y: Math.min(y1, y2) + 0.08, w: Math.max(0.4, gapR - xm - 0.08), h: Math.abs(y2 - y1) - 0.16, fontSize: V.tag, color: MUTED, valign: 'middle', objectName: `diagram-aliran-${i + 1}-label` }));
        }
      });
      if (showLegend) {
        const ly = bottomOf(src) - 0.05;
        const items = [['main', 'sistem kami'], ['node', 'kontainer / orang'], ['external', 'sistem luar']];
        // kunci notasi rata kanan bawah (brief V4 #15)
        const keyW = items.reduce((a, [, t2]) => a + 0.2 + t2.length * 0.055 + 0.3, 0) + 0.38 + 'arah data · label: apa [protokol]'.length * 0.055;
        let lx = M + W - keyW;
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
        const vals = sr.values.filter((v) => typeof v === 'number' && Number.isFinite(v));
        if (!vals.length) return null;
        const v = vals[vals.length - 1];   // nilai terakhir yang ada (seri boleh berakhir dengan null)
        return { i, v, y: py + (1 - (v - min) / (max - min)) * ph - 0.1, text: `${sr.name} ${fmt(v)}` };
      }).filter(Boolean).sort((a, b) => a.y - b.y);
      for (let k = 1; k < items.length; k++) if (items[k].y - items[k - 1].y < 0.2) items[k].y = items[k - 1].y + 0.2;   // cegah tumpuk
      items.forEach((it) => s.addText(it.text, base({ x: px, y: it.y, w: x + w - px, h: 0.2, fontSize: V.desc, bold: it.i === focus, color: TEXT, valign: 'middle', objectName: `grafik-label-ujung-${it.i + 1}` })));
    },
    fmtID,
  };
}

module.exports = { build };
