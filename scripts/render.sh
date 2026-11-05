#!/usr/bin/env bash
# Pratinjau visual presentasi: .pptx -> PDF -> JPEG per slide.
#   bash scripts/render.sh presentasi.pptx [folder-keluaran] [dpi]
# Hasil: <folder>/presentasi.pdf dan <folder>/slide-1.jpg, slide-2.jpg, ... (dicetak di akhir).
# Butuh LibreOffice dan Poppler (bash scripts/setup.sh --full).
export PATH="$PATH:/opt/homebrew/bin:/usr/local/bin"
HERE="$(cd "$(dirname "$0")" && pwd)"
IN="$1"; OUT="${2:-$(dirname "$IN")/render}"; DPI="${3:-110}"
[ -f "$IN" ] || { echo "Pemakaian: bash render.sh presentasi.pptx [folder] [dpi]"; exit 1; }
command -v soffice  >/dev/null || { echo "LibreOffice belum ada: tampilan tidak bisa diperiksa (bash scripts/setup.sh --full)."; exit 2; }
command -v pdftoppm >/dev/null || { echo "Poppler belum ada: tampilan tidak bisa diperiksa (bash scripts/setup.sh --full)."; exit 2; }
mkdir -p "$OUT"
PDF="$OUT/$(basename "${IN%.*}").pdf"
rm -f "$OUT"/slide-*.jpg "$PDF"          # jangan sampai memeriksa gambar/PDF dari render sebelumnya
python3 "$HERE/office/soffice.py" --headless --convert-to pdf --outdir "$OUT" "$IN" >/dev/null 2>&1 \
  || { echo "Konversi ke PDF gagal (LibreOffice)."; exit 3; }
[ -s "$PDF" ] || { echo "Konversi ke PDF gagal: $PDF tidak terbentuk."; exit 3; }
pdftoppm -jpeg -r "$DPI" "$PDF" "$OUT/slide" || { echo "PDF ke gambar gagal (pdftoppm)."; exit 3; }
echo "PDF: $PDF"; ls -1 "$OUT"/slide-*.jpg
