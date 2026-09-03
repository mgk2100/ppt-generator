#!/usr/bin/env bash
# sw-dev-ai-llm 덱 3판(혼합·Claude·Codex 삽화) 빌드 + 렌더 → renders/
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"; REPO="$(cd "$HERE/../.." && pwd)"
source "$REPO/venv/bin/activate"
mkdir -p "$HERE/renders"; TMP="$(mktemp -d)"
declare -A ICONS=([mixed]="$HERE/mixed_icons" [claude]="$HERE/claude_icons/png" [codex]="$HERE/codex_icons/png")
declare -A OUTS=([mixed]="sw-dev-ai-llm" [claude]="sw-dev-ai-llm-claude-icons" [codex]="sw-dev-ai-llm-codex-icons")
for v in mixed claude codex; do
  python3 "$HERE/build_deck.py" "${ICONS[$v]}" "$REPO/output/${OUTS[$v]}.pptx"
  soffice --headless --convert-to pdf --outdir "$TMP" "$REPO/output/${OUTS[$v]}.pptx" >/dev/null 2>&1
  pdftoppm -r 110 -png "$TMP/${OUTS[$v]}.pdf" "$TMP/r"
  mv "$TMP/r-1.png" "$HERE/renders/slide_01_$v.png"
  [ "$v" = mixed ] && mv "$TMP/r-2.png" "$HERE/renders/slide_02.png" || rm -f "$TMP/r-2.png"
done
rm -rf "$TMP"; echo "완료 → $REPO/output/sw-dev-ai-llm*.pptx, 렌더 $HERE/renders/"
