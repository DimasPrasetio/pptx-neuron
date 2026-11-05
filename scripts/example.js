// Contoh presentasi Neuron lengkap: ke-10 jenis slide. Titik awal tercepat untuk presentasi baru — salin, ganti isinya.
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

  // 1. Sampul
  deck.cover({
    type: 'Proposal kepada klien',
    title: 'Modernisasi sistem inventori gudang',
    subtitle: 'Rencana tiga tahap untuk memangkas selisih stok dan mempercepat pengiriman',
    meta: 'Tim Engineering  ·  Neuronworks Indonesia  ·  November 2026',
    classification: 'Klasifikasi: Rahasia — hanya untuk PT Contoh Logistik',
  });

  // 2. Agenda (active: butir yang sedang dibahas; hilangkan untuk agenda biasa)
  deck.agenda({
    items: [
      { text: 'Latar belakang dan tujuan', note: '5 menit' },
      { text: 'Temuan audit stok', note: '10 menit' },
      { text: 'Rencana tiga tahap', note: '10 menit' },
      { text: 'Biaya dan dampak', note: '10 menit' },
      { text: 'Langkah berikutnya', note: '5 menit' },
    ],
    active: 1,
  });

  // 3. Pembatas bab
  deck.section({ number: '02', title: 'Temuan audit stok', description: 'Hasil audit tiga gudang selama Agustus sampai Oktober 2026.' });

  // 4. Isi poin (+ panel Inti)
  deck.bullets({
    chapter: '02 · Temuan audit stok',
    title: 'Selisih stok berasal dari pencatatan manual, bukan kehilangan barang',
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

  // 5. Dua kolom
  deck.twoColumn({
    chapter: '02 · Temuan audit stok',
    title: 'Pemindaian di penerimaan memangkas langkah pencatatan dari lima menjadi dua',
    left: { heading: 'Proses sekarang', bullets: ['Petugas menulis barang masuk di formulir kertas', 'Admin menyalin formulir ke spreadsheet esok harinya', 'Kepala gudang mencocokkan spreadsheet dengan ERP'] },
    right: { heading: 'Proses usulan', bullets: ['Petugas memindai kode barang saat diterima', 'Stok di ERP bertambah saat itu juga', 'Selisih muncul otomatis di dasbor harian'] },
    recommend: 'right',
  });

  // 6. Tabel
  deck.section({ number: '03', title: 'Rencana tiga tahap', description: 'Lingkup, durasi, dan biaya setiap tahap.' });
  deck.table({
    chapter: '03 · Rencana tiga tahap',
    title: 'Tiga tahap selesai dalam 20 minggu dengan risiko terkendali',
    caption: 'Tabel 1 · Tahapan, risiko utama, durasi, dan biaya',
    head: ['Tahap', 'Lingkup', 'Risiko utama → mitigasi', 'Durasi', 'Biaya (Rp jt)'],
    rows: [
      ['1. Fondasi', 'Kode tunggal, pemindai', 'Kode ganda → bersihkan data', '6 minggu', '420'],
      ['2. Integrasi', 'Sinkron ERP dan kurir', 'API kurir berubah → adaptor', '8 minggu', '610'],
      ['3. Otomasi', 'Opname bergilir, dasbor', 'Staf belum terbiasa → pelatihan', '6 minggu', '380'],
      ['**Total**', '', '', '**20 minggu**', '**1.410**'],
    ],
    widths: [1.3, 2.2, 3.0, 1.2, 1.3], align: ['l', 'l', 'l', 'r', 'r'],
    source: 'Sumber: estimasi Tim Engineering Neuronworks, harga belum termasuk PPN',
  });

  // 7. Grafik (native, bisa diedit di PowerPoint)
  deck.chart({
    chapter: '03 · Rencana tiga tahap',
    title: 'Selisih stok turun bertahap hingga di bawah 1% pada akhir tahap 3',
    type: 'column',
    labels: ['Sekarang', 'Tahap 1', 'Tahap 2', 'Tahap 3'],
    series: [{ name: 'Selisih stok (%)', values: [4.8, 2.6, 1.5, 0.8] }],
    numFmt: '0.0"%"',
    highlight: 3,   // batang Tahap 3 (disebut di judul) merah, lainnya abu
    keyValue: '−83%', keyText: 'Selisih stok turun dari **4,8%** menjadi **0,8%** setelah ketiga tahap berjalan.',
    source: 'Sumber: proyeksi Tim Engineering berdasarkan audit Agustus–Oktober 2026',
  });

  // 8. Gambar + keterangan
  deck.image({
    chapter: '03 · Rencana tiga tahap',
    title: 'Dasbor selisih menampilkan gudang bermasalah setiap pagi',
    image: gambar, alt: 'Tiruan dasbor selisih stok: tiga kartu ringkasan dan grafik batang selisih harian',
    caption: 'Gambar 1 · Rancangan dasbor selisih harian',
    bullets: ['Kartu merah menandai gudang dengan selisih di atas batas', 'Grafik menunjukkan tren 10 hari terakhir', 'Kepala gudang menerima ringkasan lewat surel pukul 07.00'],
  });

  // 9. Angka penting
  deck.section({ number: '04', title: 'Biaya dan dampak' });
  deck.stats({
    chapter: '04 · Biaya dan dampak',
    title: 'Investasi kembali dalam 14 bulan dari penghematan selisih dan lembur',
    stats: [
      { value: '14 bln', label: 'Titik balik modal', note: 'Dihitung dari penghematan bulanan Rp 101 jt' },
      { value: 'Rp 1,2 M', label: 'Penghematan per tahun', note: 'Selisih stok, lembur opname, dan klaim kurir' },
      { value: '3 hari', label: 'Waktu opname dihapus', note: 'Diganti opname bergilir harian' },
    ],
    highlight: 0,
    source: 'Sumber: perhitungan Tim Engineering, asumsi volume 2026',
  });

  // 10. Penutup
  deck.closing({
    chapter: '05 · Langkah berikutnya',
    title: 'Tiga keputusan yang kami perlukan bulan ini',
    steps: [
      { text: 'Setujui lingkup dan anggaran tahap 1', owner: 'Direksi PT Contoh Logistik  ·  20 November 2026' },
      { text: 'Tetapkan satu kode barang untuk ketiga gudang', owner: 'Tim gudang + Neuronworks  ·  4 Desember 2026' },
      { text: 'Mulai uji coba pemindai di gudang Cikarang', owner: 'Neuronworks  ·  Januari 2027' },
    ],
    contact: { name: 'Nama Penanggung Jawab', role: 'Project Manager · Neuronworks', email: 'nama@neuronworks.co.id', phone: '+62 21 0000 0000' },
  });

  await deck.save(out);
  fs.unlinkSync(gambar);
  console.log('Tersimpan:', out);
})().catch((e) => { console.error(e); process.exit(1); });
