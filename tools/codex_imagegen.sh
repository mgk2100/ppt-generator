#!/usr/bin/env bash
# Codex imagegen 으로 삽화 생성 — 옵션 고정 래퍼 (2026-09-03 운영 규칙)
# 사용: tools/codex_imagegen.sh <prompt.txt> <작업 폴더> [참고 이미지.png]
#   - 프롬프트는 stdin 으로 넘긴다 (-i 가 가변 인자라 인자로 주면 이미지 경로로 삼킴)
#   - 이 워크스테이션은 bubblewrap sandbox 불가 → danger-full-access 고정
#   - 추론 수준은 건별로 준다(EFFORT). 기본 = xhigh (2026-09-09 사용자 지시로 ultra 에서 내렸다).
#     화질은 세기에 안 달렸다 — 눈가림 채점 실측에서 max 43.7 대 xhigh 44.0(카드 배경 6장),
#     max 42.2 대 xhigh 43.3(영상 판 12장)으로 차이가 없고 시간은 32% 짧았다
#     (~/Share/issue-blog/docs/design/IMAGE-EFFORT-AB-2026-09-08.md). 픽셀을 만드는 것은
#     image_gen 이고 세기가 붙는 것은 지시문을 옮기고 몇 번 더 확인하느냐뿐이다.
#     ultra 는 특히 나쁘다 — 왕복 4회 이상이 ultra 40.8% · xhigh 2.9% 이고 그 세션들이 토큰의 87%를 먹는다.
#   - **다만 이 실측은 한 번에 한 장을 만드는 자리에서 쟀다.** 여기서 ultra 를 쓰던 이유는 화질이
#     아니라 **삽화 여러 종을 한 번에 만들 때의 파일명·스타일 일관성**이다(2026-09-03 파일명 뒤섞임
#     사고 4종). 그 실패 모드는 위 실측이 재지 않았다 — 여러 종을 한 번에 만들 때 파일명이 섞이면
#     그 호출만 EFFORT=ultra 로 올린다. 단건 재생성·사진 분석은 EFFORT=medium.
#   - 완료 후 <작업 폴더>/out/*.png 를 tools/icon_postprocess.py 로 배경 투명화 → <작업 폴더>/png/
set -euo pipefail
PROMPT="${1:?prompt.txt}"; WORK="${2:?작업 폴더}"; REF="${3:-}"
HERE="$(cd "$(dirname "$0")" && pwd)"
mkdir -p "$WORK/out"
EFFORT="${EFFORT:-xhigh}"
ARGS=(exec --skip-git-repo-check --sandbox danger-full-access -c model_reasoning_effort="$EFFORT")
[ -n "$REF" ] && ARGS+=(-i "$REF")
( cd "$WORK" && codex "${ARGS[@]}" - < "$PROMPT" > run.log 2>&1 )
python3 "$HERE/icon_postprocess.py" "$WORK/out" "$WORK/png"
echo "완료: $(ls "$WORK/png" | wc -l) 개 → $WORK/png (추론 수준 $EFFORT, 로그 $WORK/run.log)"
