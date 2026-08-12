# smart-factory-briefing 디자인 시스템 (sl-sw-agent-v2 계승) (getdesign.md/hp 기반 — 전 슬라이드 공통, 위반 금지)

출처: `npx getdesign add hp` DESIGN.md 원문 토큰. 철학 = **"화이트 캔버스, 색 2개(Electric Blue + Ink)가 일의 90%를 한다"**.
다이어그램 문법은 Jay Alammar(The Illustrated Transformer)식 블록 다이어그램을 이 팔레트로 재해석.

## 1. 색상 토큰 (이 목록 밖 색 사용 금지)

```python
PRIMARY        = RGBColor(0x02, 0x4A, 0xD8)  # Electric Blue — 유일한 신호색
PRIMARY_BRIGHT = RGBColor(0x29, 0x6E, 0xF9)  # 다크 슬랩 위 링크/강조 전용
PRIMARY_DEEP   = RGBColor(0x0E, 0x31, 0x91)  # pressed/보조 강조
PRIMARY_SOFT   = RGBColor(0xC9, 0xE0, 0xFC)  # 옅은 파랑 표면 (강조 블록 배경)
INK            = RGBColor(0x1A, 0x1A, 0x1A)  # 제목·본문 기본, 다크 슬랩 배경
CHARCOAL       = RGBColor(0x3D, 0x3D, 0x3D)  # 보조 설명
GRAPHITE       = RGBColor(0x63, 0x63, 0x63)  # 캡션·메타(9pt 계열)
CANVAS         = RGBColor(0xFF, 0xFF, 0xFF)  # 배경/카드 표면
CLOUD          = RGBColor(0xF7, 0xF7, 0xF7)  # 회색 섹션 밴드
FOG            = RGBColor(0xE8, 0xE8, 0xE8)  # 짙은 밴드/hairline
STEEL          = RGBColor(0xC2, 0xC2, 0xC2)  # 강조 보더
ON_INK         = RGBColor(0xFF, 0xFF, 0xFF)  # 다크 슬랩 위 텍스트
```

- **단일 accent 원칙**: 파랑 계열(PRIMARY/BRIGHT/DEEP/SOFT) 외 유채색 절대 금지. 기능 구분은 번호(01/02/03)·타이포 위계·회색 밴드로.
- PRIMARY 사용처: 번호·키워드 텍스트, 주 데이터 흐름 화살표, chevron 장식, 강조 블록 보더. **넓은 면적 fill 금지** (SOFT만 옅은 배경 허용).
- 다크 슬랩: INK 배경 + ON_INK 텍스트 + PRIMARY_BRIGHT 강조 — 슬라이드 하단 마감 밴드 1곳에만.

## 2. 타이포그래피 (맑은 고딕 단일 — 2026-08-11 개정, 사용자 지시)

**전 슬라이드 단일 폰트 = "맑은 고딕"** (한국 보고용 표준). 표지는 마스터 상속 — 대상 외.
각 build_slide 함수 **마지막에 `force_font(slide, "맑은 고딕")` 호출 의무** — 전 run 의
라틴+EA+CS 를 일괄 교체하고, 기존 'SemiBold/Bold' 패밀리는 bold=True 로 자동 변환된다.

| 역할 | 크기 | weight | 색 |
|---|---|---|---|
| 슬라이드 제목 | 20pt | bold | INK |
| 리드(부제) | 15pt | bold | INK (키워드만 PRIMARY) |
| 카드/블록 헤더 | 12-13pt | bold | INK |
| 본문 | 10.5-12pt | regular | CHARCOAL |
| 캡션·메타 | 9pt | regular | GRAPHITE |
| 큰 번호(01/02/03) | 20-28pt | bold | PRIMARY |
| 수치·식별자 | 본문과 동일 (모노 폰트 금지 — 단일 폰트 원칙) | bold 로 강조 | INK |

- 자간·행간: line_spacing = 폰트×1.35, 문단 간격 여유 있게. 텍스트를 꽉 채우지 말고 호흡.

### 2-1. 제목·리드 표준 (전 콘텐츠 슬라이드 강제 통일 — 위치·크기·스타일 편차 금지)

- **제목**: `set_title(slide, "...", font_size=20, bold=True)` — 추가 장식 없음
- **리드(부제)**: 제목 바로 아래 **고정 규격** —
  `add_textbox(slide, CONTENT_SAFE.left, Inches(0.72), CONTENT_SAFE.width, Inches(0.35), ...)`
  15pt bold INK, 핵심 키워드 1~2개만 PRIMARY (`add_rich_text` segments), 좌측 정렬
- 리드 아래 **hairline·장식 금지** (마스터 제목 구분선이 이미 존재 — 중복 방지)
- 본문 콘텐츠 시작 y ≥ **1.18"** (리드와 간격 확보)

### 2-2. 각주·단위 표기 표준 (2026-08-13 사용자 지시)

- **각주는 항목당 1줄** — 서로 다른 용어·내용을 한 줄에 이어 붙이지 않는다 (add_footnote 를 항목 수만큼 호출하거나 개행으로 분리)
- **약어 각주는 풀네임 병기** — `MCP(Model Context Protocol): 설명` 형식. 제품명(NVLink·pgvector 등)은 대상 아님
- **표 안 수치의 단위·정의는 '헤더*+하단 각주' 방식 금지** — 단위는 열 헤더에 직접 표기
  (예: "속도 (tok/s)"), 정의·기준 설명은 **해당 표 바로 아래 캡션**(9pt GRAPHITE)으로 밀착 배치
- 슬라이드 공통 배경 설명만 하단 ※ 각주 사용

## 3. 표면·보더 (그림자 대신 hairline)

- **add_shadow() 사용 금지.** 카드 lift 는 1px hairline 보더(FOG, 강조 시 STEEL)로.
- 카드: CANVAS 배경 + FOG 1pt 보더 + ROUNDED_RECTANGLE(radius 소 — adjustment ≈0.06~0.10).
- 섹션 밴드: CLOUD 배경 사각형(보더 없음)으로 영역 구분 — "회색이 숨을 쉬게 한다".
- 강조 블록: PRIMARY_SOFT 배경 + PRIMARY 1pt 보더 (다이어그램 핵심 노드 전용).
- 구분선: add_accent_bar 두께 1px, FOG 색.

## 4. 금지 목록 (기존 덱의 'AI스러움' 제거)

- ❌ 이모지 전면 금지 (배지·헤더·플로우 라벨 포함 전부)
- ❌ add_shadow, 그라디언트, 파스텔 다색 카드(#FFF5F5/#F0FFF4 등)
- ❌ 기능별 색상 구분(그린/레드/퍼플) — 파랑+무채색만
- ❌ CHEVRON/GEAR/CAN/CUBE 등 장식성 의미 도형 — 블록은 전부 ROUNDED_RECTANGLE(소 radius) 또는 RECTANGLE
- ❌ 도형 안 장문 — 블록 안에는 라벨 1-2줄만, 설명은 블록 밖 캡션으로

## 5. 시그니처 모티프

- **브랜드성 장식 금지**: 사선(PARALLELOGRAM)·워드마크류 모티프를 쓰지 않는다 — 특정 업체를 연상시키는 장식 없이 일반적 미니멀 유지. 리드 문장은 장식 없이 타이포만.
- **다이어그램 문법 (Alammar × HP)**: 흰 블록 + hairline 보더 + 블록 위 라벨(INK 10-11pt) + 블록 아래 캡션(GRAPHITE 9pt, JetBrains Mono 혼용). 주 흐름 화살표 = PRIMARY 1.5pt STRAIGHT/ELBOW + add_arrowhead, 보조 = STEEL 1pt. 커넥터 스타일 2종 이내.
- **표**: add_grid_table — 헤더 행 CLOUD 배경+INK SemiBold, 본문 CANVAS, 보더는 수평 hairline(FOG) 위주(수직선 최소), zebra 금지.

## 6. 레이아웃

- 회사 마스터(로고·상하 구분선·푸터)는 그대로 — CONTENT_SAFE(0.28~10.56 / 0.68~7.02) 안에서만 작업.
- 여백 우선: 요소 간 간격 0.15" 이상, 카드 내부 inset 0.15" 이상. 밀도는 장식이 아니라 정보 구조(표·다이어그램)로.
- 정렬 원칙: 좌측 정렬 기본, 그리드 라인 준수(calc_grid/VFlow/HFlow), 중앙 정렬은 번호·다크 슬랩 문구만.
