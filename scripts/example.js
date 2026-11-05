// Contoh presentasi Neuron lengkap: ke-10 jenis slide dasar + semua visualisasi (grafik, diagram proses, infografis,
// diagram teknis). Titik awal tercepat untuk presentasi baru — salin, ambil slide yang perlu, ganti isinya.
//   node scripts/example.js [keluaran.pptx]
const fs = require('fs');
const os = require('os');
const path = require('path');
const zlib = require('zlib');
const { createDeck } = require(path.join(__dirname, 'neuron_theme.js'));

// Gambar contoh (tiruan dasbor) dibuat di sini agar contoh berjalan tanpa berkas luar. Di presentasi nyata,
// pakai tangkapan layar atau foto sungguhan: deck.image({ image: 'layar.png', ... }).
function mockDashboard(file) {
  const w = 1600, h = 900, px = Buffer.alloc(w * h * 3, 0xf4);   // latar F4F5F7 agar tepi gambar terlihat di slide putih
  const hex = (c) => [parseInt(c.slice(0, 2), 16), parseInt(c.slice(2, 4), 16), parseInt(c.slice(4, 6), 16)];
  const rect = (x, y, rw, rh, c) => { const [r, g, b] = hex(c); for (let j = y; j < y + rh; j++) for (let i = x; i < x + rw; i++) { const k = (j * w + i) * 3; px[k] = r; px[k + 1] = g; px[k + 2] = b; } };
  rect(0, 0, 260, h, '2B3440');                                   // menu samping
  for (let i = 0; i < 6; i++) rect(40, 140 + i * 70, 180, 18, i === 1 ? 'C9CDD3' : '5B6470');
  rect(260, 0, w - 260, 90, 'FFFFFF');                             // bilah atas
  // satu-satunya merah: kartu gudang yang selisihnya di atas batas (disebut di teks slide)
  [0, 1, 2].forEach((i) => { rect(320 + i * 420, 140, 380, 160, 'FFFFFF'); rect(350 + i * 420, 175, 160, 22, '5B6470'); rect(350 + i * 420, 220, 220, 50, i === 0 ? 'B3121B' : '1F2937'); });
  rect(320, 350, 1220, 500, 'FFFFFF');                             // panel grafik
  [220, 300, 260, 340, 380, 330, 410, 300, 360, 440].forEach((v, i) => rect(380 + i * 112, 800 - v, 70, v, '8C939D'));
  const raw = Buffer.alloc((w * 3 + 1) * h);
  for (let j = 0; j < h; j++) px.copy(raw, j * (w * 3 + 1) + 1, j * w * 3, (j + 1) * w * 3);
  const table = []; for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; table[n] = c >>> 0; }
  const crc = (b) => { let r = 0xffffffff; for (const x of b) r = table[(r ^ x) & 255] ^ (r >>> 8); return (r ^ 0xffffffff) >>> 0; };
  const chunk = (type, data) => { const len = Buffer.alloc(4); len.writeUInt32BE(data.length); const td = Buffer.concat([Buffer.from(type), data]); const c = Buffer.alloc(4); c.writeUInt32BE(crc(td)); return Buffer.concat([len, td, c]); };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 2;
  fs.writeFileSync(file, Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]));
  return file;
}

(async () => {
  const out = process.argv[2] || 'contoh-neuron.pptx';
  const gambar = mockDashboard(path.join(os.tmpdir(), `contoh-dasbor-${process.pid}.png`));
  const deck = createDeck({
    title: 'Modernisasi Sistem Inventori Gudang', author: 'Tim Engineering Neuronworks',
    footerLabel: 'Proposal Modernisasi Inventori  ·  PT Contoh Logistik',
  });
  const BAB1 = '01 · Latar belakang', BAB2 = '02 · Temuan audit stok', BAB3 = '03 · Rencana tiga tahap', BAB4 = '04 · Biaya dan dampak';

  // ===== Pembuka =====
  deck.cover({                                                        // [1] sampul
    type: 'Proposal kepada klien',
    title: 'Modernisasi sistem inventori gudang',
    subtitle: 'Rencana tiga tahap untuk memangkas selisih stok dan mempercepat pengiriman',
    meta: 'Tim Engineering  ·  Neuronworks Indonesia  ·  November 2026',
    classification: 'Klasifikasi: Rahasia — hanya untuk PT Contoh Logistik',
  });
  deck.agenda({                                                       // [2] agenda (active: butir yang sedang dibahas)
    items: [
      { text: 'Latar belakang', note: '5 menit' }, { text: 'Temuan audit stok', note: '10 menit' },
      { text: 'Rencana tiga tahap', note: '15 menit' }, { text: 'Biaya dan dampak', note: '10 menit' },
      { text: 'Langkah berikutnya', note: '5 menit' },
    ],
    active: 1,
  });

  // ===== Bab 1 =====
  deck.section({ number: '01', title: 'Latar belakang', description: 'Pertumbuhan pesanan online mengubah beban gudang.' });   // [3] pembatas bab
  deck.chart({                                                        // grafik garis: seri fokus merah + label di ujung garis
    chapter: BAB1, title: 'Pesanan online melampaui toko sejak Mei dan terus naik', type: 'line',
    labels: ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt'],
    series: [
      { name: 'Online', values: [310, 335, 360, 395, 430, 470, 505, 540, 575, 612] },
      { name: 'Toko', values: [455, 450, 446, 440, 432, 428, 421, 415, 410, 405] },
      { name: 'Target', values: [400, 410, 420, 430, 440, 450, 460, 470, 480, 490], dashed: true },
    ],
    focus: 0, source: 'Sumber: data pesanan PT Contoh Logistik, Januari–Oktober 2026 (ribu pesanan)',
  });
  deck.chart({                                                        // donat 2–3 bagian: persen fokus di tengah, daftar di kanan
    chapter: BAB1, title: 'Dua pertiga pesanan kini masuk lewat kanal digital', type: 'doughnut',
    labels: ['Kanal digital', 'Cabang', 'Lainnya'], series: [{ name: 'Pesanan', values: [62, 29, 9] }],
    keyText: 'Kanal digital naik dari **41%** pada Oktober 2025.', source: 'Sumber: data pesanan Oktober 2026',
  });

  // ===== Bab 2 =====
  deck.section({ number: '02', title: 'Temuan audit stok', description: 'Hasil audit tiga gudang selama Agustus sampai Oktober 2026.' });
  deck.bullets({                                                      // [4] isi poin + panel Inti
    chapter: BAB2, title: 'Selisih stok berasal dari pencatatan manual, bukan kehilangan barang',
    bullets: [
      '**71%** selisih muncul saat data kertas disalin ulang',
      'Barang masuk baru tercatat rata-rata **2 hari** setelah tiba',
      '**Tiga** gudang memakai kode berbeda untuk satu produk',
      'Opname bulanan menghentikan pengiriman **3 hari**',
    ],
    keyPoint: 'Memindai barang saat diterima menutup sebagian besar sumber selisih.',
    source: 'Sumber: audit stok tiga gudang PT Contoh Logistik, Agustus–Oktober 2026',
    notes: 'Tekankan bahwa masalahnya proses, bukan kejujuran staf gudang.',
  });
  deck.chart({                                                        // batang horizontal satu seri: satu batang disorot
    chapter: BAB2, title: 'Gudang Medan menyumbang selisih terbesar', type: 'bar',
    labels: ['Medan', 'Surabaya', 'Cikarang'], series: [{ name: 'Selisih (unit)', values: [128, 96, 74] }],
    highlight: 0, source: 'Sumber: audit stok Agustus–Oktober 2026 (unit selisih per 10.000 unit)',
  });
  deck.chart({                                                        // tumpuk 100% horizontal: segmen fokus menempel kiri
    chapter: BAB2, title: 'Staf Cikarang, gudang uji coba pemindai, paling puas dengan proses', type: 'stacked100',
    labels: ['Cikarang', 'Surabaya', 'Medan'],
    series: [{ name: 'Puas', values: [62, 48, 40] }, { name: 'Netral', values: [21, 30, 35] },
      { name: 'Tidak puas', values: [12, 15, 20] }, { name: 'Tidak menjawab', values: [5, 7, 5] }],
    source: 'Sumber: survei staf gudang, September 2026 (n = 214); sisa 5–7% tidak menjawab',
  });
  deck.twoColumn({                                                    // [5] dua kolom
    chapter: BAB2, title: 'Pemindaian di penerimaan memangkas langkah pencatatan dari lima menjadi dua',
    left: { heading: 'Proses sekarang', bullets: ['Petugas menulis barang masuk di formulir kertas', 'Admin menyalin formulir ke spreadsheet esok harinya', 'Kepala gudang mencocokkan spreadsheet dengan ERP'] },
    right: { heading: 'Proses usulan', bullets: ['Petugas memindai kode barang saat diterima', 'Stok di ERP bertambah saat itu juga', 'Selisih muncul otomatis di dasbor harian'] },
    recommend: 'right',
  });
  deck.compare({                                                      // sebelum/sesudah
    chapter: BAB2, title: 'Uji coba di Cikarang memangkas waktu pencatatan dari 2 hari menjadi 4 jam',
    rows: [
      { name: 'Waktu pencatatan barang masuk', before: '2 hari', after: '4 jam', delta: '−92%' },
      { name: 'Kesalahan input', before: '8%', after: '1,5%', delta: '−6,5 poin' },
      { name: 'Laporan stok', before: 'Manual', after: 'Otomatis' },
    ],
    focus: 0, source: 'Sumber: uji coba gudang Cikarang, Juni–Agustus 2026',
  });

  // ===== Bab 3 =====
  deck.section({ number: '03', title: 'Rencana tiga tahap', description: 'Lingkup, alur kerja baru, jadwal, dan arsitektur.' });
  deck.process({                                                      // keyPoint: panel Inti satu baris di bawah diagram                                                      // alur langkah
    chapter: BAB3, title: 'Barang masuk tercatat dalam lima langkah, langkah ketiga kini otomatis',
    steps: [
      { title: 'Terima barang', text: 'Petugas mencocokkan surat jalan dengan pesanan pembelian' },
      { title: 'Pindai kode', text: 'Pemindai genggam membaca kode barang dan jumlah' },
      { title: 'Catat ke ERP', text: 'Stok bertambah otomatis tanpa salin ulang' },
      { title: 'Simpan di rak', text: 'Sistem menyarankan lokasi rak yang kosong' },
      { title: 'Laporan harian', text: 'Selisih muncul di dasbor pukul 07.00' },
    ],
    active: 2,
    keyPoint: 'Langkah 3 otomatis: stok di ERP sama dengan isi rak sejak barang tiba.',
  });
  deck.cycle({                                                        // siklus
    chapter: BAB3, title: 'Opname bergilir berjalan setiap minggu, bukan sebulan sekali',
    steps: [{ title: 'Rencanakan', text: 'Pilih rak minggu ini' }, { title: 'Hitung', text: 'Pindai rak terpilih' },
      { title: 'Periksa', text: 'Bandingkan dengan ERP' }, { title: 'Perbaiki', text: 'Koreksi & cari penyebab' }],
    active: 2, center: 'Siklus mingguan',
  });
  deck.table({                                                        // [6] tabel
    chapter: BAB3, title: 'Tiga tahap selesai dalam 20 minggu dengan risiko terkendali',
    caption: 'Tabel 1 · Tahapan, risiko utama, durasi, dan biaya',
    head: ['Tahap', 'Lingkup', 'Risiko utama: mitigasi', 'Durasi', 'Biaya (Rp jt)'],
    rows: [
      ['1. Fondasi', 'Kode tunggal, pemindai', 'Kode ganda: bersihkan data', '6 minggu', '420'],
      ['2. Integrasi', 'Sinkron ERP dan kurir', 'API kurir berubah: adaptor', '8 minggu', '610'],
      ['3. Otomasi', 'Opname bergilir, dasbor', 'Staf belum terbiasa: pelatihan', '6 minggu', '380'],
      ['**Total**', '', '', '**20 minggu**', '**1.410**'],
    ],
    widths: [1.3, 2.2, 3.0, 1.2, 1.3], align: ['l', 'l', 'l', 'r', 'r'],
    source: 'Sumber: estimasi Tim Engineering Neuronworks, harga belum termasuk PPN',
  });
  deck.timeline({                                                     // linimasa tonggak
    chapter: BAB3, title: 'Rilis ke tiga gudang Mei 2027, enam bulan setelah persetujuan',
    milestones: [
      { date: 'Agu 2026', label: 'Uji coba pemindai di Cikarang', status: 'done' }, { date: 'Nov 2026', label: 'Proposal dan persetujuan', status: 'now' },
      { date: 'Feb 2027', label: 'Tahap 1: kode tunggal dan pemindai' }, { date: 'Apr 2027', label: 'Tahap 2: sinkron ERP dan kurir' },
      { date: 'Mei 2027', label: 'Tahap 3: rilis ke tiga gudang' },
    ],
    keyPoint: 'Persetujuan bulan ini menjaga rilis Mei, sebelum puncak pesanan Lebaran.',
  });
  deck.timeline({                                                     // linimasa fase (Gantt ringkas)
    chapter: BAB3, title: 'Tiga tahap selesai dalam 20 minggu, Januari sampai Mei 2027',
    periods: ['Jan 2027', 'Feb', 'Mar', 'Apr', 'Mei'],
    phases: [{ name: '1. Fondasi', start: 0, end: 1.5 }, { name: '2. Integrasi', start: 1.5, end: 3.5 }, { name: '3. Otomasi', start: 3.5, end: 5 }],
    source: 'Sumber: rencana proyek Tim Engineering, November 2026',
  });
  deck.matrix({                                                       // matriks 2×2
    chapter: BAB3, title: 'Pemindai di penerimaan memberi dampak terbesar dengan usaha terkecil',
    axes: { y: { label: 'Dampak terhadap selisih', high: 'Tinggi', low: 'Rendah' }, x: { label: 'Kemudahan', low: 'Sulit', high: 'Mudah' } },
    quadrants: { tl: 'Rencanakan', tr: 'Prioritas', bl: 'Tunda', br: 'Kerjakan cepat' },
    points: [
      { label: 'Pemindai penerimaan', x: 0.82, y: 0.86 }, { label: 'Kode barang tunggal', x: 0.62, y: 0.7 },
      { label: 'Integrasi kurir', x: 0.25, y: 0.62 }, { label: 'Opname bergilir', x: 0.58, y: 0.3 }, { label: 'Ganti ERP', x: 0.12, y: 0.2 },
    ],
    focus: 0,
    explanation: ['Tiga inisiatif kanan atas masuk tahap 1', 'Integrasi kurir menunggu API kurir stabil', 'Ganti ERP tidak diperlukan'],
    source: 'Sumber: penilaian Tim Engineering dan tim gudang, Oktober 2026 (posisi berdasarkan penilaian tim)',
  });
  deck.orgChart({                                                     // struktur organisasi
    chapter: BAB3, title: 'Neuronworks memimpin unit teknologi di bawah satu sponsor direksi',
    root: { title: 'Direktur Operasi', role: 'Sponsor proyek' },
    units: [
      { title: 'Gudang', role: 'Kepala gudang', items: ['Penerimaan', 'Opname'] }, { title: 'Keuangan', role: 'Controller', items: ['Nilai persediaan'] },
      { title: 'Teknologi', role: 'Neuronworks', items: ['Aplikasi', 'Integrasi'] }, { title: 'Pengadaan', role: 'Manajer', items: ['Pemasok'] },
    ],
    focus: 2,
    keyPoint: 'Satu sponsor direksi memutus keputusan lintas unit dalam seminggu.',
  });
  deck.architecture({                                                 // diagram konteks C4 (untuk klien/manajemen)
    chapter: BAB3, title: 'Sistem inventori melayani dua peran dan terhubung ke dua sistem luar', caption: 'Diagram konteks · Sistem Inventori',
    nodes: [
      { id: 'staf', title: 'Staf gudang', tag: 'Orang', kind: 'person', col: 0, row: 0 }, { id: 'mgr', title: 'Manajer', tag: 'Orang', kind: 'person', col: 0, row: 1 },
      { id: 'sys', title: 'Sistem Inventori', tag: 'Sistem kami', kind: 'main', col: 1, colSpan: 2, row: 0, rowSpan: 2 },
      { id: 'erp', title: 'ERP perusahaan', tag: 'Sistem luar', kind: 'external', col: 3, row: 0 }, { id: 'mail', title: 'Layanan email', tag: 'Sistem luar', kind: 'external', col: 3, row: 1 },
    ],
    links: [
      { from: 'staf', to: 'sys', label: 'Mencatat barang', protocol: 'HTTPS' }, { from: 'mgr', to: 'sys', label: 'Melihat laporan', protocol: 'HTTPS' },
      { from: 'sys', to: 'erp', label: 'Menyinkronkan stok', protocol: 'REST' }, { from: 'sys', to: 'mail', label: 'Mengirim notifikasi', protocol: 'SMTP' },
    ],
  });
  deck.architecture({                                                 // diagram kontainer C4 (untuk tim teknis)
    chapter: BAB3, title: 'Aplikasi baru menambah satu API di samping ERP yang sudah ada', caption: 'Diagram kontainer · Sistem Inventori',
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
      { from: 'staf', to: 'web', label: 'Memindai', protocol: 'HTTPS' }, { from: 'web', to: 'api', label: 'Mencatat stok', protocol: 'JSON/HTTPS', focus: true },
      { from: 'api', to: 'db', label: 'Menyimpan stok', protocol: 'SQL' }, { from: 'api', to: 'erp', label: 'Sinkron stok', protocol: 'REST' },
      { from: 'api', to: 'kurir', label: 'Pesan kurir', protocol: 'REST', dashed: true },
    ],
  });
  deck.architecture({                                                 // alur data (kiri → kanan)
    chapter: BAB3, title: 'Data transaksi mengalir dari kasir ke dasbor dalam empat langkah', columns: ['Sumber', 'Proses', 'Simpan', 'Keluaran'],
    nodes: [
      { id: 'pos', title: 'POS kasir', col: 0, row: 0 }, { id: 'mp', title: 'Marketplace', col: 0, row: 1 },
      { id: 'val', title: 'Validasi', col: 1, row: 0, rowSpan: 2 }, { id: 'dw', title: 'Gudang data', kind: 'db', col: 2, row: 0, rowSpan: 2 },
      { id: 'dash', title: 'Dasbor', col: 3, row: 0 }, { id: 'pdf', title: 'Laporan PDF', col: 3, row: 1 },
    ],
    links: [
      { from: 'pos', to: 'val', label: 'transaksi' }, { from: 'mp', to: 'val', label: 'pesanan' }, { from: 'val', to: 'dw', label: 'stok bersih', focus: true },
      { from: 'dw', to: 'dash', label: 'ringkasan' }, { from: 'dw', to: 'pdf', label: 'rekap' },
    ],
  });
  deck.image({                                                        // [8] gambar + keterangan
    chapter: BAB3, title: 'Dasbor selisih menampilkan gudang bermasalah setiap pagi',
    image: gambar, alt: 'Tiruan dasbor selisih stok: tiga kartu ringkasan dan grafik batang selisih harian',
    caption: 'Gambar 1 · Rancangan dasbor selisih harian',
    bullets: ['Kartu merah menandai gudang dengan selisih di atas batas', 'Grafik menunjukkan tren 10 hari terakhir', 'Kepala gudang menerima ringkasan lewat surel pukul 07.00'],
  });

  // ===== Bab 4 =====
  deck.section({ number: '04', title: 'Biaya dan dampak', description: 'Hasil yang diharapkan dan waktu balik modal.' });
  deck.chart({                                                        // [7] kolom satu seri, satu batang disorot + panel angka
    chapter: BAB4, title: 'Selisih stok turun bertahap hingga di bawah 1% pada akhir tahap 3', type: 'column',
    labels: ['Sekarang', 'Tahap 1', 'Tahap 2', 'Tahap 3'], series: [{ name: 'Selisih stok (%)', values: [4.8, 2.6, 1.5, 0.8] }],
    numFmt: '0.0"%"', highlight: 3,
    keyValue: '−83%', keyText: 'Selisih stok turun dari **4,8%** menjadi **0,8%** setelah ketiga tahap berjalan.',
    source: 'Sumber: proyeksi Tim Engineering berdasarkan audit Agustus–Oktober 2026',
  });
  deck.chart({                                                        // kolom + garis target (satu sumbu, satuan sama)
    chapter: BAB4, title: 'Pengiriman harian Oktober melampaui target untuk pertama kali', type: 'column',
    labels: ['Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt'],
    series: [{ name: 'Pengiriman', values: [410, 425, 432, 440, 446, 468] }, { name: 'Target', values: [450, 450, 450, 450, 450, 450], target: true }],
    highlight: 5, source: 'Sumber: data pengiriman gudang Cikarang, Mei–Oktober 2026 (rata-rata per hari)',
  });
  deck.chart({                                                        // kolom berkelompok (maks. 3 seri)
    chapter: BAB4, title: 'Cikarang paling cepat memproses barang masuk di tiga triwulan', type: 'column',
    labels: ['TW I', 'TW II', 'TW III'],
    series: [{ name: 'Cikarang', values: [42, 47, 55] }, { name: 'Surabaya', values: [38, 40, 41] }, { name: 'Medan', values: [30, 33, 35] }],
    focus: 0, source: 'Sumber: data operasional 2026 (palet per jam)',
  });
  deck.stats({                                                        // [9] angka penting (ikon opsional)
    chapter: BAB4, title: 'Investasi kembali dalam 14 bulan dari penghematan selisih dan lembur',
    stats: [
      { value: '14 bulan', label: 'Titik balik modal', note: 'Dihitung dari penghematan bulanan Rp 101 jt', icon: 'target' },
      { value: 'Rp 1,2 M', label: 'Penghematan per tahun', note: 'Selisih stok, lembur opname, dan klaim kurir', icon: 'wallet' },
      { value: '3 hari', label: 'Waktu opname dihapus', note: 'Diganti opname bergilir mingguan', icon: 'clock' },
    ],
    highlight: 0, source: 'Sumber: perhitungan Tim Engineering, asumsi volume 2026',
  });
  deck.stats({                                                        // angka besar + ikon, 4 angka (32 pt)
    chapter: BAB4, title: 'Tiga gudang, satu sistem: cakupan proyek dalam angka',
    stats: [
      { value: '3', label: 'Gudang', note: 'Cikarang, Surabaya, Medan', icon: 'warehouse' },
      { value: '214', label: 'Staf gudang', note: 'Dilatih 2 sesi per orang', icon: 'users' },
      { value: '18.500', label: 'Kode barang', note: 'Disatukan dari 3 katalog', icon: 'package' },
      { value: '42', label: 'Pemindai genggam', note: '14 unit per gudang', icon: 'scan-barcode' },
    ],
    highlight: 0, source: 'Sumber: data master PT Contoh Logistik, Oktober 2026',
  });
  deck.pyramid({                                                      // piramida konsep (tanpa angka)
    chapter: BAB4, title: 'Program inventori diturunkan langsung dari visi perusahaan',
    layers: ['Visi', 'Strategi', 'Program', 'Kegiatan'],
    descs: ['Pengiriman tepat waktu ke seluruh Indonesia', 'Stok akurat di semua gudang', 'Modernisasi inventori tiga tahap', 'Pemindai, opname bergilir, dasbor selisih'],
    focus: 2,
  });
  deck.funnel({                                                       // corong data (batang sebanding nilai)
    chapter: BAB4, title: 'Penurunan pesanan terbesar terjadi saat konfirmasi stok',
    stages: [{ name: 'Pesanan masuk', value: 12400 }, { name: 'Stok terkonfirmasi', value: 9050 }, { name: 'Dikemas', value: 8830 }, { name: 'Dikirim', value: 8710 }],
    focus: 1, source: 'Sumber: data pesanan Oktober 2026',
  });

  // ===== Penutup =====
  deck.closing({                                                      // [10] penutup
    chapter: '05 · Langkah berikutnya', title: 'Tiga keputusan yang kami perlukan bulan ini',
    steps: [
      { text: 'Setujui lingkup dan anggaran tahap 1', owner: 'Direksi PT Contoh Logistik  ·  20 November 2026' },
      { text: 'Tetapkan satu kode barang untuk ketiga gudang', owner: 'Tim gudang + Neuronworks  ·  4 Desember 2026' },
      { text: 'Mulai tahap 1 di ketiga gudang', owner: 'Neuronworks  ·  Januari 2027' },
    ],
    contact: { name: 'Nama Penanggung Jawab', role: 'Project Manager · Neuronworks', email: 'nama@neuronworks.co.id', phone: '+62 21 0000 0000' },
  });

  const file = await deck.save(out);
  fs.unlinkSync(gambar);
  console.log('Tersimpan:', file);
})().catch((e) => { console.error(e); process.exit(1); });
