#!/usr/bin/env bash
# Cek prasyarat pptx-neuron. Yang WAJIB hanya Node.js 18+ (pustaka pptxgenjs sudah dibundel di skill).
# Lainnya opsional: validasi/edit presentasi lama (Python) dan pratinjau visual (LibreOffice + Poppler).
export PATH="$PATH:/opt/homebrew/bin:/usr/local/bin"   # PATH pengguna (nvm) didahulukan
HERE="$(cd "$(dirname "$0")" && pwd)"
miss=0
ok()  { printf "  [OK]       %-22s %s\n" "$1" "$2"; }
bad() { printf "  [HILANG]   %-22s %s\n" "$1" "$2"; miss=1; }
opt() { printf "  [opsional] %-22s %s\n" "$1" "$2"; }

echo "WAJIB"
if command -v node >/dev/null; then
  major=$(node -p "process.versions.node.split('.')[0]")
  if [ "$major" -ge 18 ]; then ok "Node.js" "$(node -v)"; else bad "Node.js" "versi $(node -v) terlalu lama; butuh 18+ (jalankan: bash scripts/setup.sh)"; fi
else bad "Node.js" "belum terpasang (jalankan: bash scripts/setup.sh)"; fi
if command -v node >/dev/null && node -e "const P=require(process.argv[1]); if(!P.JSZip) process.exit(1); new P()" "$HERE/vendor/pptxgenjs.bundle.cjs" >/dev/null 2>&1; then ok "pustaka pptxgenjs" "bawaan skill (tanpa npm)"; else bad "pustaka pptxgenjs" "scripts/vendor/pptxgenjs.bundle.cjs hilang atau rusak; salin ulang folder skill lengkap"; fi
[ -f "$HERE/../assets/logo_neuron.png" ] && ok "logo Neuron" "assets/logo_neuron.png" || bad "logo Neuron" "assets/logo_neuron.png hilang; salin ulang folder skill lengkap"

echo; echo "OPSIONAL: baca, validasi, dan edit presentasi yang sudah ada"
if command -v python3 >/dev/null; then
  ok "Python 3" "$(python3 --version 2>&1) — read_pptx.py jalan tanpa modul tambahan"
  for m in defusedxml lxml pptx PIL; do
    case $m in pptx) n="python-pptx";; PIL) n="Pillow";; *) n=$m;; esac
    python3 -c "import $m" >/dev/null 2>&1 && ok "python: $n" "terpasang" || opt "python: $n" "belum ada -> bash scripts/setup.sh --full"
  done
else opt "Python 3" "belum ada -> bash scripts/setup.sh --full"; fi

echo; echo "OPSIONAL: pratinjau visual"
command -v soffice  >/dev/null && ok "LibreOffice" "$(soffice --version 2>/dev/null | head -1 | cut -c1-40)" || opt "LibreOffice" "belum ada -> bash scripts/setup.sh --full"
command -v pdftoppm >/dev/null && ok "Poppler" "$(pdftoppm -v 2>&1 | head -1)" || opt "Poppler" "belum ada -> bash scripts/setup.sh --full"

echo
if [ $miss -eq 0 ]; then echo "Siap dipakai: presentasi Neuron bisa dibuat."; else echo "Belum siap: perbaiki item WAJIB di atas."; fi
exit $miss
