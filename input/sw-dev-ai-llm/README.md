# sw-dev-ai-llm — "SW 개발 AI" LLM 개발 파이프라인 덱 (임원 보고용)

- 산출물: `output/sw-dev-ai-llm.pptx` (1p 6단계 파이프라인 · 2p 팀장 대응 예상 Q&A). 비교판 `-claude-icons` / `-codex-icons`.
- **재생성**: `./rebuild.sh` — 3판 빌드 + 렌더(`renders/slide_01_{mixed,claude,codex}.png`, `slide_02.png`). 단일 빌드는 `python3 build_deck.py [아이콘 폴더] [출력 pptx]`.
- 마스터: `base_master.pptx` = `로보틱스 비즈니스 방향성_v3.2_260901.pptx` 에서 슬라이드를 모두 제거한 파일(레이아웃 `1_제목 슬라이드`·SL 로고 포함, 144KB). 146MB 원본은 git 에 올리지 않음.
- 삽화: `claude_icons/` (SVG 코드 `make_icons.py` → `png/`), `codex_icons/png/` (Codex imagegen → `tools/icon_postprocess.py` 후처리), `mixed_icons/` (채택: 신경망·압축·대화 5종 Claude + 사물 12종 Codex).
- 문서: `docs/qa-briefing.md`(팀장 Q&A 상세·출처), `docs/claude-vs-codex.md`(도구 비교), `docs/analysis_{claude,codex}.md`(참고 슬라이드 분석).
- 실화면: `sources/sl-sw-agent/assets/ca_web_assist_conv.png` 사용.
