---
name: pptx-neuron
description: "Gunakan skill ini untuk membuat, membaca, atau mengedit presentasi PowerPoint (.pptx) bergaya Neuron / Neuronworks: slide laporan ke manajemen, proposal ke klien, materi pelatihan, presentasi internal, atau deck resmi apa pun yang harus memakai desain Neuron (logo Neuron, aksen merah B3121B, huruf Arial, motif tiga goresan, tabel header gelap, footer label dokumen + nomor slide). Pakai saat pengguna menyebut 'slide Neuron', 'presentasi Neuron', 'pptx-neuron', 'format Neuron' untuk PowerPoint, atau meminta slide/presentasi untuk Neuronworks. Semua presentasi BARU wajib lewat scripts/neuron_theme.js agar desain dan logo konsisten. Untuk slide/PowerPoint generik non-Neuron, pakai skill pptx biasa; untuk dokumen Word Neuron, pakai docx-neuron."
license: Proprietary. Lihat LICENSE.txt dan THIRD_PARTY_NOTICES.txt
---

# PPTX Neuron: membuat, mengedit, dan membaca presentasi

Skill ini turunan skill `pptx` dengan satu aturan tambahan: **setiap presentasi baru memakai desain Neuron**. Yang diambil dari template hanya desainnya; isinya selalu dari pengguna. Skill ini satu keluarga dengan `docx-neuron` (token warna, huruf, dan logo sama persis).

Semua path skrip di bawah relatif terhadap folder skill ini (`~/.claude/skills/pptx-neuron`).

## Prasyarat

| Kebutuhan | Untuk | Cek / pasang |
|---|---|---|
| Node.js 18+ (**wajib**) | membuat presentasi baru | `bash scripts/check_prereqs.sh` |
| pustaka `pptxgenjs` | sudah dibundel di `scripts/vendor/` | tidak perlu `npm install` |
| Python 3 | `read_pptx.py` (baca isi, tanpa modul tambahan) | biasanya sudah ada |
| Python `defusedxml`, `lxml` | `office/validate.py`, `clean.py`, `add_slide.py` | `bash scripts/setup.sh --full` |
| Python `python-pptx`, `Pillow` | `slide_text.py` (baca lengkap), `thumbnail.py` (grid slide) | `bash scripts/setup.sh --full` |
| LibreOffice + Poppler | pratinjau visual | `bash scripts/setup.sh --full` |

Jalankan `bash scripts/check_prereqs.sh` dulu. Bila yang wajib belum ada, minta pengguna menjalankan `bash scripts/setup.sh`; jangan memasang paket sendiri tanpa izin. Bila alat opsional tidak ada, lewati langkahnya dan sampaikan ke pengguna apa yang belum diperiksa (misalnya "tampilan belum diperiksa secara visual").

## Aturan desain Neuron (jangan diubah per presentasi)

| Elemen | Ketentuan |
|---|---|
| Kanvas | 16:9, 10" × 5,625" (`LAYOUT_16x9`), margin 0,5", grid 12 kolom (kolom 0,475", celah 0,3") |
| Huruf | Arial saja. Judul aksi 22 pt tebal `1F2937`; isi 14 pt `111111`; tabel 12 pt (11 bila padat); label bab 10 pt tebal merah; sumber & footer 9 pt `5B6470` |
| Warna | Aksen merah logo `B3121B` (hemat: label bab + satu sorotan per slide), judul `1F2937`, sekunder `5B6470`, garis `C9CDD3`, header tabel `2B3440`, zebra/panel `F4F5F7`; seri grafik `B3121B` → `1F2937` → `8C939D` → `5B6470` |
| Motif khas | **Tiga Goresan** dari simbol logo: besar di sampul & pembatas bab, mini di depan label bab setiap slide isi |
| Judul slide isi | Satu kalimat kesimpulan ("judul aksi"), sentence case, tanpa titik, maks. 2 baris, posisi tetap |
| Footer slide isi | Label dokumen kiri, nomor slide, logo kecil kanan bawah. Tanpa garis |
| Sampul & penutup | Latar putih (logo berlatar putih opak), logo besar |
| Pembatas bab | Latar `1F2937`, nomor + judul putih, goresan merah memudar, **tanpa logo** |
| Tabel | Header `2B3440` teks putih, zebra `F4F5F7`, hanya garis horizontal `C9CDD3`; "Tabel N · …" di atas; maks. 7 baris data per slide |
| Gambar | Rasio dijaga; "Gambar N · …" di bawah; alt text wajib |
| Sumber data | Wajib tampil di slide ("Sumber: …"), karena deck juga dibaca sebagai PDF |
| Dilarang | Garis di bawah judul, pita warna di tepi slide/kartu, latar merah, animasi/transisi, isi rata tengah |
| Logo | `assets/logo_neuron.png` (jangan diganti, jangan diregangkan, rasio 612:176, hanya di area putih) |

Perubahan desain dilakukan sekali di `scripts/neuron_theme.js`, bukan per presentasi. Alasan dan sumber setiap keputusan ada di brief desain proyek ini.

## Membuat presentasi baru (wajib lewat neuron_theme.js)

Jangan menulis gaya dari nol dan jangan `require('pptxgenjs')` langsung. Titik awal tercepat: salin `scripts/example.js` (contoh lengkap ke-10 jenis slide, siap dijalankan) lalu ganti isinya.

```js
const T = require(process.env.HOME + '/.claude/skills/pptx-neuron/scripts/neuron_theme.js');

const deck = T.createDeck({ title: 'Judul Presentasi', author: 'Nama', footerLabel: 'Laporan Triwulan III  ·  Divisi Operasional' });

deck.cover({ type: 'Laporan kepada manajemen', title: 'Judul presentasi', subtitle: '…',
             meta: 'Nama  ·  Unit  ·  November 2026', classification: 'Klasifikasi: Internal' });
deck.agenda({ items: [{ text: 'Latar belakang', note: '5 menit' }, 'Temuan', 'Rekomendasi'], active: 1 });
deck.section({ number: '02', title: 'Temuan', description: 'Satu kalimat konteks bab.' });
deck.bullets({ chapter: '02 · Temuan', title: 'Kesimpulan dalam satu kalimat', bullets: ['**71%** …', '…'],
               keyPoint: 'Satu kalimat inti.', source: 'Sumber: …', notes: 'Catatan pembicara' });
deck.twoColumn({ chapter, title, left: { heading: 'Sekarang', bullets: [...] }, right: { heading: 'Usulan', bullets: [...] }, recommend: 'right' });
deck.table({ chapter, title, caption: 'Tabel 1 · Judul tabel', head: ['A', 'B'], rows: [['a', '**b**']],
             widths: [4.5, 4.5], align: ['l', 'r'], source: 'Sumber: …' });          // jumlah widths = 9,0"
deck.chart({ chapter, title, type: 'column', labels: ['Q1', 'Q2'], series: [{ name: 'Pendapatan', values: [10, 12] }],
             keyValue: '+18%', keyText: 'Penjelasan singkat.', numFmt: '#,##0', source: 'Sumber: …' });
deck.image({ chapter, title, image: 'layar.png', alt: 'Deskripsi gambar', caption: 'Gambar 1 · Keterangan', bullets: [...] });
deck.stats({ chapter, title, stats: [{ value: '92%', label: 'Label', note: 'Keterangan' }], highlight: 0, source: 'Sumber: …' });
deck.closing({ steps: [{ text: 'Langkah', owner: 'Pemilik · tanggal' }], contact: { name, role, email, phone } });

const file = await deck.save('output.pptx');   // menulis, merapikan XML, menerapkan tema warna Neuron; mengembalikan path berkas
```

Referensi jenis slide:

| Fungsi | Jenis | Isi & batas |
|---|---|---|
| `cover` | Sampul | `type` (label merah kapital), `title` (maks. 2 baris), `subtitle`, `meta`, `classification` |
| `agenda` | Agenda | `items` (string atau `{ text, note }`, maks. 6), `active` (indeks butir yang sedang dibahas; untuk pelacak bab) |
| `section` | Pembatas bab | `number`, `title`, `description`; membuka *section* PowerPoint baru |
| `bullets` | Isi poin | `bullets` (maks. 5, idealnya 1 baris), `keyPoint` (panel "Inti" kanan, opsional) |
| `twoColumn` | Dua kolom | `left`/`right` = `{ heading, bullets }`; `recommend: 'left' | 'right'` memberi panel + label "Rekomendasi" |
| `table` | Tabel | `head`, `rows`, `widths` (inci, jumlah 9,0), `align` (`'l' | 'c' | 'r'`), `caption`; maks. 7 baris data (8 tanpa caption) |
| `chart` | Grafik | `type`: `column`, `bar` (horizontal, terbesar di atas), `line`, `doughnut`, `stacked100` — rincian di bagian Visualisasi |
| `image` | Gambar + keterangan | `image` (PNG/JPEG), `alt` (wajib), `caption`, `bullets` (maks. 3), `wide: true` untuk tangkapan layar lebar |
| `stats` | Angka penting / angka besar + ikon | `stats` maks. 4 `{ value, label, note, icon? }` (1 angka 54 pt, 2–3 angka 40 pt, 4 angka 32 pt); `highlight` = satu angka merah (yang disebut di judul); satu angka + `explanation` = varian besar |
| `closing` | Penutup | `steps` maks. 4 (string atau `{ text, owner }`), `contact`; membuka *section* "Penutup" |

Semua fungsi slide isi menerima `chapter` (label bab, mis. `'02 · Temuan'`), `title` (judul aksi), `source`, dan `notes` (catatan pembicara).

Aturan pemakaian:
- Teks **tebal** ditulis `**seperti ini**` di poin, sel tabel, panel, dan keterangan.
- Satu pesan per slide. Judul = kesimpulannya ("Pendapatan Q3 naik 18%"), bukan topiknya ("Pendapatan").
- Penjelasan panjang masuk `notes`, bukan ke slide.
- Tema mencetak peringatan `[neuron_theme] …` bila teks diperkirakan meluap, terlalu banyak poin/baris/seri, judul > 2 baris, judul bertitik, atau gambar tanpa alt. **Jangan abaikan**: persingkat, pecah slide, atau lanjutkan tabel di slide "(lanjutan)" dengan header yang sama. Jangan mengecilkan huruf di bawah skala tema.
- Spasi antara angka dan satuannya ("Rp 101 jt", "14 bln") otomatis dibuat tak-putus; judul sampul 2 baris dipecah seimbang.
- Bagian (*section*) PowerPoint dibuat otomatis: "Pembuka", satu per `section()`, dan "Penutup" — panel slide PowerPoint rapi per bab.
- Input wajib yang kosong (judul, `series`, `head`, nilai angka) atau tidak cocok (jumlah sel ≠ kolom header, `focus`/`highlight` di luar rentang) menghentikan skrip dengan pesan `[neuron_theme] …` — perbaiki isinya, jangan dibungkam.
- Grafik selalu native (bisa diedit di PowerPoint). Hanya jenis yang tidak ada di PowerPoint (Sankey, jaringan) yang boleh berupa gambar.
- Jenis presentasi lain (proposal, pelatihan, laporan) memakai modul yang sama; ubah hanya `cover.type`, `footerLabel`, dan isi.

## Visualisasi (grafik, diagram proses, infografis, diagram teknis)

Semua visual memakai **grafik dan bentuk asli PowerPoint** (bisa diedit, bukan gambar) dan model **"sorot + netral"**: hanya
SATU hal yang merah per slide — yang disebut judul aksi; sisanya abu `8C939D`, gelap `1F2937`, atau sedang `5B6470`.
Urutan warna seri: `B3121B` (fokus) → `8C939D` → `1F2937` → `5B6470` (maks. 4 warna; lebih → gabung "Lainnya" atau pecah).
Tanpa legenda bila label bisa ditulis langsung, tanpa tanda tick, garis bantu tipis, angka gaya Indonesia (4,8 · 1.410).

> **Pemisah desimal di PowerPoint:** angka di dalam grafik asli diberi format berpenanda bahasa Indonesia (`[$-421]`).
> LibreOffice/PDF menampilkannya "4,8"; PowerPoint dapat mengikuti setelan bahasa komputer penampil (bisa tampil "4.8" di
> komputer berbahasa Inggris) — belum diverifikasi di Microsoft PowerPoint. Angka di luar grafik (label ujung, panel, corong,
> donat) selalu gaya Indonesia. Bila wajib pasti, kirim deck sebagai PDF dari `render.sh`.
Alasan & sumber: brief desain proyek ini, bagian "Visualisasi".

| Kelompok | Fungsi | Kapan & batas | Opsi utama |
|---|---|---|---|
| Grafik | `chart({ type: 'column' })` | perbandingan besaran; ≤ 12 kolom; berkelompok ≤ 3 seri | satu seri: `highlight` = kategori merah (`-1` = sengaja tanpa sorotan; tanpa keduanya tema memperingatkan); multi-seri: `focus` = seri merah, legenda atas otomatis; seri `{ target: true }` = garis target putus-putus berlabel di belakang batang |
| | `chart({ type: 'bar' })` | label kategori panjang / > 6 kategori; ≤ 8 batang | urutkan data besar → kecil (digambar terbesar di atas); `highlight` |
| | `chart({ type: 'line' })` | tren waktu ≥ 5 titik; ≤ 4 garis | `focus` (2,25 pt merah), seri `{ dashed: true }` = rencana/target; nama + nilai terakhir ditulis di ujung garis |
| | `chart({ type: 'doughnut' })` | bagian-dari-keseluruhan 2–3 bagian (= 100%) | `focus` = indeks irisan merah (bawaan 0; `-1` tanpa merah); persen fokus di tengah, daftar label di kanan; `keyText` |
| | `chart({ type: 'stacked100' })` | komposisi beberapa kelompok; ≤ 5 batang × 4 segmen | seri pertama = segmen fokus (menempel kiri); label % di dalam segmen ≥ 8% |
| Diagram proses | `process({ steps: [{ title, text }], active, variant, keyPoint })` | urutan kerja 3–6 langkah | `active` = langkah merah; `variant: 'ribbon'` = pita fase (nama ≤ 2 kata); `keyPoint` = panel Inti satu baris di bawah (juga di `timeline` dan `orgChart`) |
| | `cycle({ steps, active, center })` | proses berulang 3–5 tahap | `active` = tahap berisian gelap; `center` = label tengah |
| | `timeline({ milestones: [{ date, label, status }] })` | ≤ 6 tonggak | `status`: `done` (abu gelap), `now` (merah + "Sekarang"), `plan` (lingkaran kosong) |
| | `timeline({ periods, phases: [{ name, start, end, status }], today })` | jadwal fase ≤ 6 baris | `start`/`end` = indeks periode (end eksklusif); `status`: `done` / `active` (merah) / `plan`; `today` = garis "Hari ini" |
| | `orgChart({ root: { title, role }, units: [{ title, role, items }], focus })` | ≤ 3 tingkat, ≤ 6 unit | tingkat 3 = `items` (daftar teks, bukan kotak); `focus` = unit berisian gelap |
| Infografis | `stats({ stats: [{ value, label, note, icon }] })` | 1–4 angka besar | `icon` = nama ikon (lihat Ikon), opsional, ikon tidak pernah merah |
| | `compare({ labels, rows: [{ name, before, after, delta }], focus })` | sebelum/sesudah ≤ 4 atribut, atribut sama di kedua sisi | `focus` = baris yang deltanya merah |
| | `matrix({ axes, quadrants, best, points: [{ label, x, y }], focus, explanation })` | prioritas 2×2, ≤ 10 titik (x, y = 0..1) | `best` = kuadran berisian (bawaan `tr`); `explanation` teks atau daftar poin |
| | `pyramid({ layers, descs, focus })` | KONSEP bertingkat 3–5 lapis, tanpa angka | lapis sama tinggi; jangan untuk data |
| | `funnel({ stages: [{ name, value }], focus, numFmt })` | corong DATA ≤ 6 tahap | batang sebanding nilai + konversi antar-tahap; `focus` bawaan = penurunan terbesar |
| Diagram teknis | `architecture({ caption, nodes, links, boundary })` | konteks C4 (klien) / kontainer C4 (tim teknis); ≤ 9 elemen | `nodes[{ id, title, tag, desc, kind, col 0–3, row 0–2, colSpan, rowSpan }]`, `kind`: `person`/`container` (abu muda), `main` (sistem kami, gelap), `external` (putih bergaris), `db` (silinder); `links[{ from, to, label, protocol, focus, dashed }]`; `boundary: { label, cols: [a, b] }` |
| | `architecture({ columns: ['Sumber', 'Proses', 'Simpan', 'Keluaran'], … })` | alur data kiri → kanan, ≤ 3 simpul per kolom | label garis = nama data yang mengalir; satu jalur `focus` |

Diagram arsitektur: label garis konteks/kontainer = kata kerja + objek ("Mencatat barang", ± 14 karakter per celah);
"sistem kami" yang membentang beberapa sel dibuat ringkas bila garis dari tetangganya tetap lurus. Pada 3 baris simpul
menjadi lebih pendek (± 0,4–0,5"): bila judul + tag + deskripsi tidak muat, tema berhenti dengan pesan — persingkat
deskripsi, hilangkan `source`/`columns`, atau pakai 2 baris. Label dua garis tegak yang berdampingan saling menghindar;
bila tetap tidak muat, tema memperingatkan agar label dipersingkat.

Aturan pemakaian visual:
- Satu visual dominan per slide; judul aksi menyebut apa yang disorot ("Medan menyumbang selisih terbesar" → `highlight` Medan).
- Garis diagram selalu satu arah, lurus atau siku (dibuat otomatis); label garis singkat (± 12 karakter per baris) — tema
  memperingatkan bila tidak muat.
- Jangan pakai donat untuk > 3 bagian, piramida untuk angka, sumbu ganda, atau grafik pelangi (warna berbeda per kategori).
- Alt text setiap visual dibuat otomatis dari datanya; ringkasan diagram juga baik ditulis di `notes`.

### Ikon

`stats({ … icon: 'users' })` atau langsung `deck.icon(slide, 'users', { x, y, size: 0.4, color })` pada slide hasil fungsi
tema. Ikon garis Lucide (lisensi ISC) digambar sebagai bentuk asli PowerPoint (bisa diwarnai ulang). Warna `1F2937` (putih di
latar gelap), tidak pernah merah, maks. satu per kartu, selalu berlabel. Daftar nama: `require('<skill>/scripts/neuron_icons.js').ICON_NAMES`.

## Verifikasi hasil (wajib)

```bash
python3 scripts/office/validate.py output.pptx          # struktur, relasi, grafik; tiap kegagalan menyebut perbaikannya
python3 scripts/read_pptx.py output.pptx                # isi per slide: cek salah ketik, urutan, teks tertinggal
bash scripts/render.sh output.pptx render/              # menghasilkan render/output.pdf dan render/slide-*.jpg
```

Baca setiap gambar slide dengan mata segar (subagent bagus untuk ini) dan periksa: teks meluap atau terpotong (cek ini dulu), elemen bertumpuk, sumber/footer bertabrakan dengan isi, judul slide sejenis tidak sejajar, angka terpisah dari satuannya, logo terlihat di sampul/penutup/slide isi dan tidak ada di pembatas bab, sisa teks contoh. Perbaiki di skrip pembuatnya, render ulang. Bila `render.sh` keluar dengan kode 2 (LibreOffice/Poppler tidak ada), sampaikan bahwa tampilan belum diperiksa secara visual. LibreOffice memakai Arial yang sama ukurannya, jadi pratinjau ukuran teks dapat dipercaya; tampilan akhir tetap di PowerPoint pengguna.

## Membaca presentasi yang sudah ada

- `python3 scripts/read_pptx.py deck.pptx` — judul, teks, tabel, data grafik, alt text gambar, dan catatan pembicara per slide (`<!-- Slide number: N -->`). Hanya butuh Python standar.
- `python3 scripts/slide_text.py deck.pptx` — versi lengkap dari skill pptx dasar (butuh `python-pptx`).
- `python3 scripts/thumbnail.py deck.pptx <prefix>` — grid bergambar semua slide untuk memilih layout (butuh `Pillow` + LibreOffice). Selalu beri `<prefix>` sesuai nama deck agar grid deck lain tidak tertimpa.
- Berkas `.ppt` lama dikonversi dulu: `python3 scripts/office/soffice.py --headless --convert-to pptx file.ppt`.

## Mengedit presentasi yang sudah ada

- **Presentasi Neuron**: pertahankan layout, logo, dan gaya apa adanya; jangan menyuntik gaya baru. Ikuti ukuran huruf yang sudah ada.
- **Presentasi non-Neuron yang diminta "dijadikan format Neuron"**: buat ulang lewat `neuron_theme.js`; ambil teks, angka, dan gambarnya (`read_pptx.py`), bukan gaya lamanya.

pptxgenjs tidak bisa membuka berkas yang sudah ada, jadi pengeditan dilakukan langsung pada XML:

```bash
python3 -c "import sys,zipfile; zipfile.ZipFile(sys.argv[1]).extractall('unpacked')" deck.pptx
python3 scripts/add_slide.py unpacked/ slide2.xml --after slide2.xml   # duplikasi slide (atau slideLayoutN.xml); mencetak path slide baru
# urutkan / hapus slide = edit <p:sldIdLst> di ppt/presentation.xml
python3 scripts/clean.py unpacked/                                     # setelah menghapus: buang slide, media, rels yatim
# edit isi slide di unpacked/ppt/slides/slideN.xml
(cd unpacked && rm -f ../out.pptx && zip -Xr ../out.pptx .)            # zip dari DALAM folder; rm dulu agar bagian terhapus tidak tertinggal
python3 scripts/office/validate.py out.pptx --original deck.pptx
```

- **Kerjakan semua perubahan struktur (tambah, hapus, urutkan) sebelum mengedit isi slide.** `add_slide.py` menyalin slide apa adanya, dan `clean.py` menghapus slide yang tidak ada di `<p:sldIdLst>`.
- Jangan menyalin berkas slide dengan tangan — `add_slide.py` mengurus semua pendaftarannya. Langsung ke berkas: `add_slide.py deck.pptx slide2.xml -o out.pptx` (**selalu beri `-o`**, tanpa itu deck asli ditimpa). Slide duplikat masih *merujuk* grafik slide asalnya; mengedit grafik satu slide mengubah yang lain.
- Parse XML dengan `defusedxml.minidom` bila menulis skrip transformasi — `xml.etree.ElementTree` menulis ulang prefiks namespace dan merusak deck.
- Satu `<a:p>` per butir daftar; salin `<a:pPr>` saudaranya untuk menjaga spasi; `<a:pPr>` selalu anak pertama `<a:p>`. Bullet diwarisi dari layout; jangan mengetik `•`.
- Teks dengan spasi di awal/akhir butuh `xml:space="preserve"` pada `<a:t>`.
- Slot template ≠ jumlah isi: bila template punya 4 kartu dan isinya 3, hapus seluruh grup kartu ke-4, bukan hanya teksnya.
- Cek sisa teks contoh: `python3 scripts/read_pptx.py out.pptx | grep -iE "\bx{3,}\b|lorem|ipsum|\bTODO|\[insert"`.
- Bila memakai `python-pptx`: ia tidak bisa menduplikasi slide, `text_frame.text = "…"` menghapus format (ubah `run.text`), dan tidak bisa membaca SVG/EMF.

## Jebakan pptxgenjs (bila menambah elemen di luar fungsi tema)

`neuron_theme.js` sudah menangani semua ini. Berlaku bila menambah elemen sendiri pada slide hasil fungsi tema (setiap fungsi mengembalikan objek slide pptxgenjs) — tetap pakai token yang diekspor (`T.ACC`, `T.INK`, `T.TYPE`, …):

- Warna hex tanpa `#` dan tanpa 8 digit (`'B3121B'`); keduanya merusak berkas. Transparansi lewat `transparency: 0-100`.
- pptxgenjs mengubah objek opsi di tempat: jangan memakai satu objek `shadow`/opsi untuk dua panggilan.
- Setiap `addText` tanpa placeholder butuh `isTextBox: true`; set `margin: 0` bila teks harus sejajar bentuk lain.
- Daftar: `bullet` hanya pada run pertama tiap butir; `breakLine: true` pada run terakhir kecuali butir terakhir. Satu butir dengan beberapa run (teks tebal di tengah) menghasilkan `<a:pPr>` ganda — `deck.save()` merapikannya; bila menulis berkas sendiri lewat `pres.writeFile`, jalankan `deck.save()` sebagai gantinya.
- `rectRadius` hanya untuk `ROUNDED_RECTANGLE`; gradien tidak didukung; offset bayangan harus ≥ 0 (desain Neuron memang tanpa bayangan dan sudut siku).
- Grafik batang/kolom bertumpuk: `dataLabelPosition` hanya `ctr`, `inEnd`, `inBase` (`outEnd` merusak berkas). Grafik kombinasi dengan sumbu kedua butuh `valAxes` **dan** `catAxes` (dua entri masing-masing).
- Jangan mengurutkan ulang anak `<p:presentation>`.
- `pres.company` dan judul layout ditulis tanpa escape — `&` harus `&amp;` (sudah ditangani untuk `company`).

## Pemeliharaan skill

Folder master: `~/Documents/Neuron/ai-agent/pptx-neuron`. Edit hanya di master, uji dengan `bash tests/uji.sh` (buat, input tepi, validasi, baca, edit, render, pengaman sync), commit, lalu salin ke Claude dengan `bash scripts/sync.sh` (hanya mau menyalin ke folder bernama `pptx-neuron`).
