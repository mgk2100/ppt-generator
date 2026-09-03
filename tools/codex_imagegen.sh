#!/usr/bin/env bash
# Codex imagegen 으로 삽화 생성 — 옵션 고정 래퍼 (2026-09-03 운영 규칙)
# 사용: tools/codex_imagegen.sh <prompt.txt> <작업 폴더> [참고 이미지.png]
#   - 프롬프트는 stdin 으로 넘긴다 (-i 가 가변 인자라 인자로 주면 이미지 경로로 삼킴)
#   - 이 워크스테이션은 bubblewrap sandbox 불가 → danger-full-access 고정
#   - 완료 후 <작업 폴더>/out/*.png 를 tools/icon_postprocess.py 로 배경 투명화 → <작업 폴더>/png/
set -euo pipefail
PROMPT="${1:?prompt.txt}"; WORK="${2:?작업 폴더}"; REF="${3:-}"
HERE="$(cd "$(dirname "$0")" && pwd)"
mkdir -p "$WORK/out"
ARGS=(exec --skip-git-repo-check --sandbox danger-full-access -c model_reasoning_effort=medium)
[ -n "$REF" ] && ARGS+=(-i "$REF")
( cd "$WORK" && codex "${ARGS[@]}" - < "$PROMPT" > run.log 2>&1 )
python3 "$HERE/icon_postprocess.py" "$WORK/out" "$WORK/png"
echo "완료: $(ls "$WORK/png" | wc -l) 개 → $WORK/png (로그 $WORK/run.log)"
