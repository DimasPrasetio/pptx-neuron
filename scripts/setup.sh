#!/usr/bin/env bash
# Pemasang satu-perintah untuk pptx-neuron (macOS).
#   bash setup.sh          -> memastikan Node.js (satu-satunya prasyarat wajib)
#   bash setup.sh --full   -> juga LibreOffice + Poppler (pratinjau visual) dan Python defusedxml/lxml/python-pptx/Pillow
#                             (validasi, edit, dan baca presentasi yang sudah ada)
export PATH="$PATH:/opt/homebrew/bin:/usr/local/bin"
HERE="$(cd "$(dirname "$0")" && pwd)"
FULL=0; [ "$1" = "--full" ] && FULL=1

node_ok() { command -v node >/dev/null && [ "$(node -p "process.versions.node.split('.')[0]")" -ge 18 ]; }

if node_ok; then echo "Node.js sudah ada: $(node -v)"
else
  if [ "$(uname)" = "Darwin" ] && command -v brew >/dev/null; then
    echo "Memasang Node.js lewat Homebrew..."; brew install node
  else
    echo "Node.js belum terpasang dan Homebrew tidak ditemukan."
    echo "Pasang Node.js LTS (klik-klik biasa) dari: https://nodejs.org  lalu jalankan skrip ini lagi."
    echo "Windows: winget install OpenJS.NodeJS.LTS"
    exit 1
  fi
fi

if [ $FULL -eq 1 ]; then
  if command -v brew >/dev/null; then
    command -v soffice  >/dev/null || brew install --cask libreoffice
    command -v pdftoppm >/dev/null || brew install poppler
    command -v python3  >/dev/null || brew install python
  else
    echo "Homebrew tidak ditemukan; LibreOffice dan Poppler dilewati (opsional)."
  fi
  # Modul Python untuk validasi & edit (office/validate.py, clean.py, add_slide.py) serta baca lengkap
  # (slide_text.py butuh python-pptx, thumbnail.py butuh Pillow). Dipasang di folder pengguna.
  if command -v python3 >/dev/null && ! python3 -c "import defusedxml, lxml, pptx, PIL" >/dev/null 2>&1; then
    echo "Memasang modul Python defusedxml, lxml, python-pptx, Pillow (folder pengguna)..."
    python3 -m pip install --user -q defusedxml lxml python-pptx Pillow 2>/dev/null \
      || python3 -m pip install --user --break-system-packages -q defusedxml lxml python-pptx Pillow
  fi
fi

echo; bash "$HERE/check_prereqs.sh"
