#!/usr/bin/env bash
# Uji menyeluruh pptx-neuron: buat contoh 10 jenis slide -> validasi -> baca -> edit (duplikasi, ubah teks,
# hapus slide) -> validasi ulang -> render (bila LibreOffice/Poppler ada).
#   bash tests/uji.sh [folder-keluaran]        (bawaan: folder sementara baru)
# Keluar 0 bila semua lulus; render yang dilewati karena alat tidak ada dilaporkan, tidak menggagalkan uji.
set -euo pipefail
export PATH="$PATH:/opt/homebrew/bin:/usr/local/bin"
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
S="$ROOT/scripts"
OUT="${1:-$(mktemp -d "${TMPDIR:-/tmp}/uji-pptx-neuron.XXXXXX")}"
mkdir -p "$OUT"
ok() { printf '  [LULUS] %s\n' "$1"; }
gagal() { printf '  [GAGAL] %s\n' "$1"; exit 1; }

echo "Folder uji: $OUT"
echo "1. Membuat contoh"
node "$S/example.js" "$OUT/contoh.pptx" > "$OUT/build.log" 2>&1 || { cat "$OUT/build.log"; gagal "node example.js"; }
grep -q '\[neuron_theme\]' "$OUT/build.log" && { cat "$OUT/build.log"; gagal "contoh memicu peringatan tema"; }
ok "contoh dibuat tanpa peringatan"

echo "1b. Input tepi & regresi tinjauan"
node "$ROOT/tests/tepi.js" "$OUT/tepi" || gagal "tests/tepi.js"
python3 "$S/office/validate.py" "$OUT/tepi/tepi.pptx" | tail -1 | grep -q PASSED || gagal "validate.py deck tepi"
ok "deck tepi valid"

echo "2. Validasi struktur"
python3 "$S/office/validate.py" "$OUT/contoh.pptx" | tail -1 | grep -q PASSED || gagal "validate.py contoh"
ok "validate.py"
python3 - "$OUT/contoh.pptx" <<'PY' || gagal "XML paragraf (pPr ganda) / jumlah slide"
import re, sys, zipfile
z = zipfile.ZipFile(sys.argv[1])
slides = [n for n in z.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)]
bad = [n for n in slides if re.search(r"</a:r><a:pPr", z.read(n).decode())]
assert not bad, f"pPr di tengah paragraf: {bad}"
assert len(slides) == 12, f"jumlah slide {len(slides)}, harus 12"
theme = z.read("ppt/theme/theme1.xml").decode()
assert 'B3121B' in theme and 'name="Neuron"' in theme, "tema Neuron belum diterapkan"
PY
ok "12 slide, tanpa pPr ganda, tema Neuron terpasang"

echo "3. Baca isi"
python3 "$S/read_pptx.py" "$OUT/contoh.pptx" > "$OUT/isi.md"
for t in "# Modernisasi sistem inventori gudang" "| 1. Fondasi |" "### Grafik" "### Catatan pembicara" "![Tiruan dasbor"; do
  grep -qF "$t" "$OUT/isi.md" || gagal "read_pptx.py tidak memuat: $t"
done
ok "read_pptx.py (judul, tabel, grafik, alt text, catatan)"

echo "4. Edit presentasi yang sudah ada"
rm -rf "$OUT/unpacked"
python3 -c "import sys,zipfile; zipfile.ZipFile(sys.argv[1]).extractall(sys.argv[2])" "$OUT/contoh.pptx" "$OUT/unpacked"
NEW=$(python3 "$S/add_slide.py" "$OUT/unpacked" slide4.xml --after slide4.xml | grep -o 'ppt/slides/slide[0-9]*\.xml' | head -1)
[ -n "$NEW" ] || gagal "add_slide.py"
python3 - "$OUT/unpacked" "$NEW" <<'PY' || gagal "ubah teks / hapus slide"
import re, sys, pathlib
root, new = pathlib.Path(sys.argv[1]), sys.argv[2]
p = root / new
x = p.read_text(encoding="utf-8")
assert "Opname bulanan menghentikan pengiriman" in x
p.write_text(x.replace("Opname bulanan menghentikan pengiriman", "Opname mingguan menghentikan pengiriman"), encoding="utf-8")
# hapus pembatas bab terakhir (slide10.xml) dari urutan slide; clean.py membuang berkasnya
pres = root / "ppt/presentation.xml"
rels = (root / "ppt/_rels/presentation.xml.rels").read_text(encoding="utf-8")
rid = re.search(r'Id="(rId\d+)"[^>]*Target="slides/slide10\.xml"|Target="slides/slide10\.xml"[^>]*Id="(rId\d+)"', rels)
rid = next(g for g in rid.groups() if g)
px = pres.read_text(encoding="utf-8")
px2 = re.sub(r'<p:sldId [^>]*r:id="%s"/>' % rid, "", px)
assert px2 != px
pres.write_text(px2, encoding="utf-8")
PY
python3 "$S/clean.py" "$OUT/unpacked" > /dev/null
(cd "$OUT/unpacked" && rm -f ../edit.pptx && zip -qXr ../edit.pptx .)
python3 "$S/office/validate.py" "$OUT/edit.pptx" --original "$OUT/contoh.pptx" | tail -1 | grep -q PASSED || gagal "validate.py hasil edit"
python3 "$S/read_pptx.py" "$OUT/edit.pptx" > "$OUT/isi-edit.md"
[ "$(grep -c '<!-- Slide number' "$OUT/isi-edit.md")" = 12 ] || gagal "jumlah slide hasil edit harus 12 (12 + 1 duplikat - 1 dihapus)"
grep -q "Opname mingguan" "$OUT/isi-edit.md" || gagal "teks hasil edit tidak terbaca"
[ "$(grep -c 'Opname bulanan menghentikan' "$OUT/isi-edit.md")" = 1 ] || gagal "slide asal ikut berubah"
ok "duplikasi slide, ubah teks, hapus slide, clean, validasi --original"

echo "5. Render"
if bash "$S/render.sh" "$OUT/edit.pptx" "$OUT/render" > "$OUT/render.log" 2>&1; then
  n=$(ls "$OUT"/render/slide-*.jpg | wc -l | tr -d ' ')
  [ "$n" = 12 ] || gagal "render menghasilkan $n gambar, harus 12"
  [ "$OUT/render/edit.pdf" -nt "$OUT/edit.pptx" ] || gagal "PDF render lebih lama dari berkas"
  ok "render 12 slide -> $OUT/render/"
else
  code=$?; [ $code -eq 2 ] && echo "  [LEWAT] LibreOffice/Poppler tidak ada: tampilan belum diperiksa" || { cat "$OUT/render.log"; gagal "render.sh"; }
fi
echo "6. Pengaman sync.sh"
rm -rf "$OUT/skills"; mkdir -p "$OUT/skills/lain"; echo "name: lain" > "$OUT/skills/lain/SKILL.md"
bash "$S/sync.sh" "$OUT/skills" > /dev/null 2>&1 && gagal "sync.sh ke folder induk seharusnya ditolak"
[ -f "$OUT/skills/lain/SKILL.md" ] || gagal "sync.sh menghapus skill lain"
mkdir -p "$OUT/skills2/pptx-neuron"; printf -- "---\nname: bukan-ini\n---\n" > "$OUT/skills2/pptx-neuron/SKILL.md"
bash "$S/sync.sh" "$OUT/skills2/pptx-neuron" > /dev/null 2>&1 && gagal "sync.sh ke skill lain seharusnya ditolak"
bash "$S/sync.sh" "$OUT/skills/pptx-neuron" > /dev/null || gagal "sync.sh ke tujuan sah"
[ -f "$OUT/skills/pptx-neuron/SKILL.md" ] && [ ! -d "$OUT/skills/pptx-neuron/tests" ] && [ -f "$OUT/skills/lain/SKILL.md" ] || gagal "hasil sync.sh"
ok "sync.sh menolak folder induk/skill lain dan menyalin ke tujuan sah"
echo "Semua uji lulus."
