# sw-dev-ai-llm — "SL SW Agent" 임원 보고 덱

## 산출물
| 파일 | 내용 |
|---|---|
| `output/sw-dev-ai-llm.pptx` | **본 덱 1장** — 6단계 지식 축적 파이프라인. ① 박스 우상단에 첨부 1 이 **PPT 파일 아이콘 OLE 개체로 임베드**되어 있다 |
| `output/sw-dev-ai-llm-attach1.pptx` | **첨부 1 (별도 2장)** — 데이터 취득·자산화 상세(구조 + 확보 현황 통합) / 예상 질의 대응(팀장용) |
| `output/sw-dev-ai-llm-compare.pptx` | **대조 자료 (별도 2장)** — ① SW Agent = AI **활용**(상용 모델 도입) vs 영상·LiDAR 인지 / 로봇 강화학습 = AI **개발**(모델 직접 제작) 6개 항목 대조 ② 앞단 데이터 확보 방식(annotation 인력 투입 vs 설계자 사용량 축적). 총괄센터장 질의 대응용(2026-09-17) |

(평면 아이콘 시절 Claude/Codex 비교판 `-claude-icons`·`-codex-icons` 는 2026-09-03 삭제 — 필요하면 커밋 `0a6bba9` 에서 꺼낸다)

**첨부 1 임베드 방식**: PowerPoint 의 `삽입 → 개체 → 파일로부터 만들기 → 파일에 포함` 과 동일하다. 첨부 PPT 원본이 본 덱 안(`ppt/embeddings/attach1.pptx`)에 통째로 들어가므로 **다른 PC 로 본 덱 하나만 보내도 첨부가 열린다**(연결이 아니라 포함). ① 박스 우상단의 PPT 파일 아이콘을 더블클릭하면 열린다.

## 재생성
```bash
./rebuild.sh                       # 첨부 1 → 미리보기 렌더 → 본 덱(개체 임베드) → 렌더까지 한 번에
ICON_DIR=./mixed_icons ./rebuild.sh # 다른 아이콘 세트로 빌드
python3 compare/build_compare.py   # 대조 자료 2장
```
개별 실행: `python3 build_attach1.py [아이콘 폴더] [출력]`, `python3 build_deck.py [아이콘 폴더] [출력]`.
본 덱은 `output/sw-dev-ai-llm-attach1.pptx` 와 `assets/attach1_preview.png` 가 있어야 개체를 넣는다(없으면 경고 후 개체 생략).

## 디자인 기준 (2026-09-03 사용자 확정)
- **색은 파랑으로 통일**. 구간별 색 구분(민트/파랑/노랑)은 폐기했다 — 배경·박스·패널·화살표 전부 파랑 계열(`#0E5DB7` · `#1B2E53` · 틴트 `#F2F6FC`). 금색 `#D9A93E` 는 헤어라인과 다크 밴드 강조에만 소량.
- **색의 변화는 아이콘 안에서** 준다: 3D 아이소메트릭 아이콘 14종(`iso_icons/`)이 밝은 하늘색~진네이비 여러 톤 + 금색 포인트를 층층이 써서 아이콘마다 표정이 생긴다. Codex imagegen 으로 제작(프롬프트 = `iso_icons/PROMPT_*.txt`, 원본 = `iso_icons/raw/`).
- **아이콘 크게 + 라벨 간결**: 단계당 아이콘 0.96", 라벨은 2~7자.
- **실사 사진**: 첨부 1 히어로 밴드·서버랙·도면 + 본 덱 1p 하단 밴드. 딥네이비 반투명 오버레이 위에 흰/금색 텍스트.
- **고급스러움**: 금색 헤어라인, 소프트 그림자, 라운드 사진 모서리, 다크 슬래브.
- 사내 LLM 학습 계획은 제외(올해 도입 계획 없음). 구버전 스크립트 = `build_deck_v1_local_llm.py.bak`.

## 구성 파일
- `deck_common.py` 팔레트·도형·사진·**OLE 임베드** 헬퍼 / `build_deck.py` 본 덱 / `build_attach1.py` 첨부 1 / `prep_assets.py` 사진·캡처 자르기 / `rebuild.sh`
- `base_master.pptx` 로보틱스 덱에서 슬라이드를 모두 제거한 마스터(144KB, 146MB 원본 불필요)
- `photos/` Codex imagegen 실사 원본 4장 + 프롬프트 / `assets/` 잘라 놓은 배치용 이미지
- `iso_icons/`(**현행 3D 아이소메트릭 14종** + `raw/` 원본 + 프롬프트) · `mixed_icons/` · `claude_icons/` · `codex_icons/`(구 평면 세트)
- `docs/` qa-briefing(팀장 Q&A 상세) · claude-vs-codex(도구 비교) · analysis_{claude,codex}
- `compare/build_compare.py` 대조 자료 2장 — **전부 네이티브 도형·텍스트**(이미지 0개)라 PowerPoint 에서 문구·색·위치를 그대로 고칠 수 있다. 좌표는 파일 상단 상수(`X0/Y0/W/H`·열 폭·행 비율)에서 계산으로만 만든다. 색은 `deck_common` 팔레트 + 이 자료 전용 색(SLATE 계열) 상단 선언
