# pptx-neuron

Skill Claude untuk membuat presentasi PowerPoint (.pptx) dengan **desain Neuron** yang konsisten: logo Neuron, aksen merah `B3121B`, huruf Arial, motif tiga goresan dari simbol logo, judul berupa kalimat kesimpulan, tabel berheader gelap, serta footer berisi label dokumen dan nomor slide. Satu keluarga dengan skill Word `docx-neuron`: warna, huruf, dan logonya sama persis.

Desain Neuron sudah tertanam di skill (token warna, skala huruf, tata letak 10 jenis slide, logo); skill tidak membutuhkan file template eksternal apa pun.

## Jenis slide

| | Jenis | Kapan dipakai |
|---|---|---|
| 1 | Sampul | Halaman pertama: jenis dokumen, judul, penyusun, klasifikasi |
| 2 | Agenda | Daftar bahasan; bisa diulang sebelum tiap bab dengan butir aktif disorot |
| 3 | Pembatas bab | Pemisah bagian, latar gelap |
| 4 | Isi poin | Kesimpulan + 3–5 alasan, panel "Inti" opsional |
| 5 | Dua kolom | Perbandingan (sekarang vs usulan, opsi A vs B) |
| 6 | Tabel | Data atau atribut yang dibandingkan |
| 7 | Grafik | Tren atau perbandingan angka (grafik asli PowerPoint, bisa diedit) |
| 8 | Gambar + keterangan | Tangkapan layar, foto, diagram |
| 9 | Angka penting | 1–3 angka besar beserta artinya |
| 10 | Penutup | Langkah berikutnya dan kontak |

## Prasyarat (preinstalled)

**Yang wajib hanya satu: Node.js 18 atau lebih baru.** Pustaka pembuat .pptx (`pptxgenjs`) sudah dibundel di dalam skill (`scripts/vendor/`), jadi tidak perlu `npm install`.

| Komponen | Status | Fungsi | Pasang |
|---|---|---|---|
| Node.js 18+ | **wajib** | menjalankan skrip pembuat presentasi | `bash scripts/setup.sh` |
| pustaka `pptxgenjs` | sudah bawaan skill | membuat file .pptx | tidak perlu dipasang |
| Python 3 | opsional (biasanya sudah ada) | membaca isi presentasi (`read_pptx.py`) | bawaan macOS / `setup.sh --full` |
| Python `defusedxml`, `lxml` | opsional | validasi dan edit presentasi yang sudah ada | `bash scripts/setup.sh --full` |
| Python `python-pptx`, `Pillow` | opsional | baca lengkap dan grid gambar slide | `bash scripts/setup.sh --full` |
| LibreOffice | opsional | pratinjau visual (.pptx ke PDF) | `bash scripts/setup.sh --full` |
| Poppler | opsional | PDF ke gambar per slide | `bash scripts/setup.sh --full` |

Tanpa komponen opsional, presentasi baru tetap terbuat dengan desain yang sama; yang hilang hanya pratinjau visual dan alat edit presentasi lama.

### Pemasangan untuk karyawan (satu perintah)

macOS, buka Terminal lalu jalankan:

```bash
bash ~/.claude/skills/pptx-neuron/scripts/setup.sh
```

- Jika Node.js sudah ada, skrip hanya melapor siap pakai.
- Jika belum ada dan Homebrew tersedia, Node.js dipasang otomatis.
- Jika Homebrew tidak ada, skrip menampilkan tautan pemasang Node.js LTS (https://nodejs.org, klik-klik biasa) lalu jalankan skrip lagi.
- Tambahkan `--full` untuk pratinjau visual dan alat edit (LibreOffice, Poppler, modul Python `defusedxml`, `lxml`, `python-pptx`, `Pillow`; sekitar 750 MB).

Windows: pasang Node.js dengan `winget install OpenJS.NodeJS.LTS` (skrip `setup.sh` ditulis untuk macOS dan belum diuji di Windows). Di lingkungan sandbox Claude/Cowork, Node.js biasanya sudah tersedia.

Cek kapan saja:

```bash
bash ~/.claude/skills/pptx-neuron/scripts/check_prereqs.sh
```

## Instalasi skill

Folder master: `~/Documents/Neuron/ai-agent/pptx-neuron` (repo git lokal; belum ada repo GitHub). Salin ke Claude dengan:

```bash
bash ~/Documents/Neuron/ai-agent/pptx-neuron/scripts/sync.sh
```

Skrip menyalin skill ke `~/.claude/skills/pptx-neuron` (tanpa `.git`, `tests/`, dan `.DS_Store`). Skill terdeteksi di sesi Claude berikutnya. Letakkan di luar folder `synced/` karena isi `synced/` dikelola aplikasi.

**Satu sumber kebenaran:** edit hanya di folder master, uji (`bash tests/uji.sh`), commit, lalu jalankan `sync.sh` lagi. Jangan mengedit salinan di `~/.claude/skills` langsung, agar keduanya tidak menyimpang.

## Isi folder

```
pptx-neuron/
├── README.md               ← berkas ini
├── SKILL.md                ← instruksi dan aturan desain untuk Claude
├── LICENSE.txt
├── THIRD_PARTY_NOTICES.txt ← lisensi pustaka open-source yang dibundel
├── assets/logo_neuron.png  ← logo resmi (612x176, sama dengan docx-neuron, jangan diregangkan)
├── scripts/
│   ├── neuron_theme.js     ← modul desain Neuron (createDeck: cover, agenda, section, bullets, ...)
│   ├── example.js          ← contoh presentasi lengkap 10 jenis slide (titik awal presentasi baru)
│   ├── read_pptx.py        ← baca isi presentasi tanpa modul Python tambahan
│   ├── render.sh           ← pratinjau: .pptx -> PDF -> JPEG per slide
│   ├── setup.sh            ← pemasang satu perintah (Node.js; --full untuk pratinjau & alat edit)
│   ├── check_prereqs.sh    ← pengecek prasyarat
│   ├── sync.sh             ← salin master ke ~/.claude/skills/pptx-neuron
│   ├── vendor/pptxgenjs.bundle.cjs ← pustaka pptxgenjs dibundel (lisensi: THIRD_PARTY_NOTICES.txt)
│   └── office/, add_slide.py, clean.py, slide_text.py, thumbnail.py  ← alat bawaan skill pptx (edit & validasi)
└── tests/uji.sh            ← uji menyeluruh: buat, validasi, baca, edit, render
```

## Pemakaian singkat

Minta Claude: "buat presentasi Neuron tentang ..." atau "buat slide dengan pptx-neuron untuk ...". Untuk jenis presentasi lain (proposal, laporan, pelatihan) yang berubah hanya label jenis di sampul, label footer, dan isi. Aturan desain lengkap dan contoh kode ada di `SKILL.md`.

## Tips agar optimal

- Selalu buat presentasi baru lewat `neuron_theme.js`; jangan menulis gaya dari nol.
- Coba dulu: `node scripts/example.js` menghasilkan `contoh-neuron.pptx` berisi ke-10 jenis slide.
- Periksa visual setelah membuat presentasi: `bash scripts/render.sh file.pptx render/`.
- Tulis judul slide sebagai kesimpulan ("Selisih stok turun 83%"), bukan topik ("Selisih stok").
- Satu pesan per slide, maksimal 5 poin; penjelasan panjang masuk catatan pembicara.
- Teks **tebal** ditulis `**seperti ini**`. Sumber data selalu dicantumkan di slide.
- Jangan ubah logo, warna, atau ukuran huruf per presentasi; ubah di `neuron_theme.js` (di master), lalu `sync.sh`, agar semua presentasi berikutnya ikut konsisten.
