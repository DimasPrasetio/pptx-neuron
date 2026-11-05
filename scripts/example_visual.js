// Contoh visualisasi Neuron (tahap Rancangan visualisasi): grafik garis, alur langkah, matriks 2×2, arsitektur C4.
//   node scripts/example_visual.js [keluaran.pptx]
const T = require(require('path').join(__dirname, 'neuron_theme.js'));
(async () => {
  const d = T.createDeck({ footerLabel: 'Proposal Modernisasi Inventori  ·  PT Contoh Logistik' });
  d.chart({
    chapter: '01 · Latar belakang', title: 'Pesanan online melampaui pesanan toko sejak Mei dan terus naik',
    type: 'line', labels: ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt'],
    series: [
      { name: 'Online', values: [310, 335, 360, 395, 430, 470, 505, 540, 575, 612] },
      { name: 'Toko', values: [455, 450, 446, 440, 432, 428, 421, 415, 410, 405] },
      { name: 'Target', values: [400, 410, 420, 430, 440, 450, 460, 470, 480, 490], dashed: true },
    ],
    focus: 0, numFmt: '#,##0',
    source: 'Sumber: data pesanan PT Contoh Logistik, Januari–Oktober 2026 (ribu pesanan)',
  });
  d.process({
    chapter: '03 · Rencana tiga tahap', title: 'Barang masuk tercatat dalam lima langkah, langkah ketiga kini otomatis',
    steps: [
      { title: 'Terima barang', text: 'Petugas mencocokkan surat jalan dengan pesanan pembelian' },
      { title: 'Pindai kode', text: 'Pemindai genggam membaca kode barang dan jumlah' },
      { title: 'Catat ke ERP', text: 'Stok bertambah otomatis tanpa salin ulang' },
      { title: 'Simpan di rak', text: 'Sistem menyarankan lokasi rak yang kosong' },
      { title: 'Laporan harian', text: 'Selisih muncul di dasbor pukul 07.00' },
    ],
    active: 2,
  });
  d.matrix({
    chapter: '03 · Rencana tiga tahap', title: 'Pemindai di penerimaan memberi dampak terbesar dengan usaha terkecil',
    axes: { y: { label: 'Dampak terhadap selisih', high: 'Tinggi', low: 'Rendah' }, x: { label: 'Kemudahan', low: 'Sulit', high: 'Mudah' } },
    quadrants: { tl: 'Rencanakan', tr: 'Prioritas', bl: 'Tunda', br: 'Kerjakan cepat' },
    points: [
      { label: 'Pemindai penerimaan', x: 0.82, y: 0.86 }, { label: 'Kode barang tunggal', x: 0.62, y: 0.7 },
      { label: 'Integrasi kurir', x: 0.25, y: 0.62 }, { label: 'Opname bergilir', x: 0.58, y: 0.3 }, { label: 'Ganti ERP', x: 0.12, y: 0.2 },
    ],
    focus: 0,
    explanation: ['Tiga inisiatif kanan atas masuk tahap 1', 'Integrasi kurir menunggu API kurir stabil', 'Ganti ERP tidak diperlukan'],
    source: 'Sumber: penilaian Tim Engineering dan tim gudang, Oktober 2026',
  });
  d.architecture({
    chapter: '03 · Rencana tiga tahap', title: 'Aplikasi baru menambah satu API di samping ERP yang sudah ada',
    caption: 'Diagram kontainer · Sistem Inventori',
    boundary: { label: 'Sistem Inventori [batas sistem]', cols: [1, 2] },
    nodes: [
      { id: 'staf', title: 'Staf gudang', tag: 'Orang', kind: 'person', col: 0, row: 0 },
      { id: 'web', title: 'Aplikasi pemindai', tag: 'Kontainer · React', desc: 'Pindai & cari barang', kind: 'container', col: 1, row: 0 },
      { id: 'api', title: 'API inventori', tag: 'Kontainer · Go', desc: 'Aturan stok', kind: 'main', col: 2, row: 0 },
      { id: 'db', title: 'Basis data', tag: 'Kontainer · MySQL', kind: 'db', col: 2, row: 1 },
      { id: 'erp', title: 'ERP perusahaan', tag: 'Sistem luar', kind: 'external', col: 3, row: 0 },
      { id: 'kurir', title: 'API kurir', tag: 'Sistem luar', kind: 'external', col: 3, row: 1 },
    ],
    links: [
      { from: 'staf', to: 'web', label: 'Memindai', protocol: 'HTTPS' },
      { from: 'web', to: 'api', label: 'Barang masuk', protocol: 'JSON/HTTPS', focus: true },
      { from: 'api', to: 'db', label: 'Simpan stok', protocol: 'SQL' },
      { from: 'api', to: 'erp', label: 'Sinkron stok', protocol: 'REST' },
      { from: 'api', to: 'kurir', label: 'Pesan kirim', protocol: 'REST', dashed: true },
    ],
  });
  console.log(await d.save(process.argv[2] || 'contoh-visual-neuron.pptx'));
})().catch((e) => { console.error(e); process.exit(1); });
