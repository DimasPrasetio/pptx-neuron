// Contoh presentasi Neuron (tahap Rancangan: sampul, isi poin, tabel).
//   node scripts/example.js [keluaran.pptx]
const path = require('path');
const { createDeck } = require(path.join(__dirname, 'neuron_theme.js'));

(async () => {
  const out = process.argv[2] || 'contoh-neuron.pptx';
  const deck = createDeck({
    title: 'Modernisasi Sistem Inventori Gudang', author: 'Tim Engineering Neuronworks',
    footerLabel: 'Proposal Modernisasi Inventori  ·  PT Contoh Logistik',
  });

  deck.cover({
    type: 'Proposal kepada klien',
    title: 'Modernisasi sistem inventori gudang',
    subtitle: 'Rencana tiga tahap untuk memangkas selisih stok dan mempercepat pengiriman',
    meta: 'Tim Engineering  ·  Neuronworks Indonesia  ·  November 2026',
    classification: 'Klasifikasi: Rahasia — hanya untuk PT Contoh Logistik',
  });

  deck.bullets({
    chapter: '02 · Temuan audit',
    title: 'Selisih stok berasal dari pencatatan manual, bukan kehilangan barang',
    bullets: [
      '**71%** selisih muncul saat data kertas disalin ulang',
      'Barang masuk baru tercatat rata-rata **2 hari** setelah tiba',
      'Tiga gudang memakai kode berbeda untuk satu produk',
      'Opname bulanan menghentikan pengiriman 3 hari',
    ],
    keyPoint: 'Memindai barang saat diterima menutup sebagian besar sumber selisih.',
    source: 'Sumber: audit stok tiga gudang PT Contoh Logistik, Agustus–Oktober 2026',
    notes: 'Tekankan bahwa masalahnya proses, bukan kejujuran staf gudang.',
  });

  deck.table({
    chapter: '03 · Rencana',
    title: 'Tiga tahap selesai dalam 20 minggu dengan risiko terkendali',
    caption: 'Tabel 1 · Tahapan, durasi, dan hasil utama',
    head: ['Tahap', 'Lingkup', 'Durasi', 'Biaya (Rp jt)'],
    rows: [
      ['1. Fondasi', 'Kode barang tunggal, pemindai di penerimaan', '6 minggu', '420'],
      ['2. Integrasi', 'Sinkron dengan ERP dan aplikasi kurir', '8 minggu', '610'],
      ['3. Otomasi', 'Opname bergilir harian, dasbor selisih', '6 minggu', '380'],
      ['**Total**', '', '**20 minggu**', '**1.410**'],
    ],
    widths: [1.9, 4.3, 1.4, 1.4], align: ['l', 'l', 'r', 'r'],
    source: 'Sumber: estimasi Tim Engineering Neuronworks, harga belum termasuk PPN',
  });

  await deck.save(out);
  console.log('Tersimpan:', out);
})().catch((e) => { console.error(e); process.exit(1); });
