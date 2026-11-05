#!/usr/bin/env bash
# Salin skill dari folder master ini ke folder skill Claude (default ~/.claude/skills/pptx-neuron).
#   bash scripts/sync.sh [tujuan]
# Edit selalu di master; jangan edit salinan di ~/.claude/skills langsung agar kedua salinan tidak menyimpang.
SRC="$(cd "$(dirname "$0")/.." && pwd)"
DEST="${1:-$HOME/.claude/skills/pptx-neuron}"
[ "$SRC" = "$(cd "$DEST" 2>/dev/null && pwd)" ] && { echo "Sumber dan tujuan sama; tidak ada yang disalin."; exit 0; }
mkdir -p "$DEST"
rsync -a --delete --delete-excluded --exclude '.DS_Store' --exclude '__pycache__' --exclude 'render/' \
  --exclude '.git' --exclude '.claude/' --exclude 'tests/' "$SRC/" "$DEST/"
echo "Tersinkron: $SRC -> $DEST"
