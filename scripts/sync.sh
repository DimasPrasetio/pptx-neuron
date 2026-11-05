#!/usr/bin/env bash
# Salin skill dari folder master ini ke folder skill Claude (default ~/.claude/skills/pptx-neuron).
#   bash scripts/sync.sh [tujuan]
# Edit selalu di master; jangan edit salinan di ~/.claude/skills langsung agar kedua salinan tidak menyimpang.
SRC="$(cd "$(dirname "$0")/.." && pwd)"
DEST="${1:-$HOME/.claude/skills/pptx-neuron}"
# Pengaman: rsync --delete mengosongkan tujuan. Tujuan harus folder bernama pptx-neuron dan, bila sudah berisi skill,
# skill itu harus pptx-neuron — agar salah ketik (mis. ~/.claude/skills) tidak menghapus skill lain.
[ "$(basename "$DEST")" = "pptx-neuron" ] || { echo "Tujuan harus folder bernama pptx-neuron (mis. ~/.claude/skills/pptx-neuron), bukan: $DEST"; exit 1; }
if [ -f "$DEST/SKILL.md" ] && ! grep -q '^name: pptx-neuron$' "$DEST/SKILL.md"; then echo "Tujuan berisi skill lain ($DEST/SKILL.md); tidak disalin."; exit 1; fi
if [ -d "$DEST" ] && [ ! -f "$DEST/SKILL.md" ] && [ -n "$(ls -A "$DEST" 2>/dev/null)" ]; then echo "Tujuan $DEST tidak kosong dan bukan skill pptx-neuron; tidak disalin."; exit 1; fi
[ "$SRC" = "$(cd "$DEST" 2>/dev/null && pwd)" ] && { echo "Sumber dan tujuan sama; tidak ada yang disalin."; exit 0; }
mkdir -p "$DEST"
rsync -a --delete --delete-excluded --exclude '.DS_Store' --exclude '__pycache__' --exclude 'render/' \
  --exclude '.git' --exclude '.claude/' --exclude 'tests/' "$SRC/" "$DEST/"
echo "Tersinkron: $SRC -> $DEST"
