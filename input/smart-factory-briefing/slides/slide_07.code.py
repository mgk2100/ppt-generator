"""slide_07 — 로컬 LLM 실측 ① : 후보 모델 테스트 결과 (실측 로그 기반).

좌측 (~62%): 실측 모델 표 add_grid_table 5열 × (헤더+12행)
  — Gemma-4-31B(채택)·GLM-5.2 행 PRIMARY_SOFT 강조, 헤더 CLOUD+bold,
    수평 hairline 위주(수직선 없음), 행 높이 텍스트 기반 산출.
    속도 단위는 헤더 직접 표기 "속도 (tok/s)" + 표 바로 아래 밀착 캡션
    (9pt GRAPHITE, design_system 2-2 — '헤더*+하단 각주' 방식 금지).
우측 (~35%): 카드 2개 VFlow
  — '4bit 압축의 한계' (CANVAS + STEEL 보더)
  — 'GLM-5.2 (754B) 1bit 압축 결과' (PRIMARY_SOFT + PRIMARY 보더)
하단: 전체 폭 다크 슬랩 (INK, '상용 LLM' PRIMARY_BRIGHT bold)
  + ※ 각주 1줄 (4bit/1bit 양자화 정의만 — tok/s 설명은 표 캡션으로 이동).

spec:   input/smart-factory-briefing/slides/slide_07.spec.yaml
design: input/smart-factory-briefing/design_system.md
  — 맑은 고딕 단일 (build 마지막 force_font 의무), Electric Blue + Ink,
    이모지·add_shadow·CHEVRON류 의미 도형·zebra 금지.
수치는 전부 실측 로그 기반 — 임의 변경 금지.
"""

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from ppt_utils import (
    set_title, clear_placeholders,
    add_textbox, add_bullet_list, add_footnote,
    add_grid_table,
    estimate_container_height, VFlow,
    force_font,
)
from template_contract import CONTENT_SAFE

# ---- 디자인 토큰 (design_system.md — 이 목록 밖 색 사용 금지) ----
PRIMARY        = RGBColor(0x02, 0x4A, 0xD8)
PRIMARY_BRIGHT = RGBColor(0x29, 0x6E, 0xF9)
PRIMARY_SOFT   = RGBColor(0xC9, 0xE0, 0xFC)
INK            = RGBColor(0x1A, 0x1A, 0x1A)
CHARCOAL       = RGBColor(0x3D, 0x3D, 0x3D)
GRAPHITE       = RGBColor(0x63, 0x63, 0x63)
CANVAS         = RGBColor(0xFF, 0xFF, 0xFF)
CLOUD          = RGBColor(0xF7, 0xF7, 0xF7)
FOG            = RGBColor(0xE8, 0xE8, 0xE8)
STEEL          = RGBColor(0xC2, 0xC2, 0xC2)
ON_INK         = RGBColor(0xFF, 0xFF, 0xFF)

# ---- 콘텐츠 데이터 (spec.content_blocks — 수치 임의 변경 금지) ----
LEAD_SEGS = [                   # rev12 — 성능 관점 메시지는 이 페이지가 담당
    ("30B급 ~ 754B급 ", INK),
    ("15종 이상", PRIMARY),
    (" 테스트 — 보유 서버 구동 가능 모델은 상용 LLM 대비 ", INK),
    ("성능 부족", PRIMARY),
]

TABLE_HEADER = ["모델", "규모", "압축", "속도 (tok/s)", "판정"]
TABLE_ROWS = [
    ["Gemma-4-31B", "31B", "4bit", "172", "채택 — 현 라이브"],
    ["gpt-oss-120b", "117B", "4bit", "148~172",
     "채택 이력 → 품질 열세로 교체"],
    ["Qwen3-Next-80B", "80B", "4bit", "~160", "채택 이력 → 품질 열세로 교체"],
    ["Qwen2.5-Coder-32B", "32B", "bf16", "21",
     "채택 이력 → 품질 우수하나 속도 한계로 교체"],
    ["Qwen3-Coder-30B", "30B", "bf16", "136",
     "탈락 — 품질 미달"],
    ["Gemma-3-27B", "27B", "4bit", "65", "탈락 — 품질 미달"],
    ["Llama-3.3-70B", "70B", "4bit", "35", "탈락 — 코딩 비특화"],
    ["GLM-4.5-Air", "106B", "4bit", "—", "탈락 — 품질 미달"],
    ["Devstral-2-123B", "123B", "4bit·GPU 2장", "~31", "탈락 — 품질·속도 미달"],
    ["Qwen3.5-122B", "122B", "4bit·GPU 2장", "~103",
     "탈락 — 범용(비전·언어) 모델 · 코딩 성능 미달"],
    ["MiniMax-M2", "230B", "4bit·GPU 2장", "—",
     "탈락 — 출력 불안정 (코드 미생성)"],
    ["GLM-5.2", "754B", "1bit·217GB", "4.9", "탈락 — 속도·품질 모두 미달"],
]
HILITE_ROWS = {0}              # Gemma-4-31B(채택 라이브)만 강조 (rev11 — GLM-5.2 강조 해제)
COL_ALIGNS = ["l", "c", "c", "c", "l"]
# 근사 비율 — rev3: 속도 열 확대 (헤더 "속도 (tok/s)" 1줄 수용),
# 판정 열은 2줄 허용 (긴 교체 사유 — 행 높이 텍스트 기반 재산출)
COL_W_IN = [1.50, 0.52, 1.04, 1.04, 2.25]

CARD1_TITLE = "4bit 압축의 한계"
CARD1_ITEMS = [
    "대형 모델을 96GB GPU에 올리려면 강한 압축이 필수 — 이 과정의 품질 손실이 "
    "모델 크기를 키운 효과를 상쇄",
    "압축한 대형 모델이 중형(32B) 모델보다 나은 결과를 주지 못한 사례를 실측으로 확인",
    "일부 대형 모델은 간단한 질문에도 답을 정리하지 못해 사용 중단",
]
CARD2_TITLE = "GLM-5.2 (754B) 1bit 압축 결과"
CARD2_ITEMS = [
    "1bit로 압축해도 217GB — GPU 96GB 1장을 넘는 약 200GB는 CPU 메모리에 적재",
    "생성 속도 4.9 tok/s — 1명이 쓰기에도 느림 (CPU 적재 구간이 병목)",
    "코딩 벤치 88.75% — 당시 채택 32B 모델(95.0%)보다 낮은 품질",
    "극압축의 품질 저하가 754B 규모의 이점을 상쇄 → 도입 불가 판정",
]

SLAB_SEGS = [
    ("현재 운영 — 로컬 = 분석 서술 · 검색 임베딩 등 한정 용도 / 실제 업무 = ",
     ON_INK, False),
    ("상용 LLM", PRIMARY_BRIGHT, True),
]

# rev3 (design_system 2-2) — tok/s 정의는 표 밀착 캡션, 각주는 양자화 1줄만
TABLE_CAPTION = ("tok/s: 초당 생성 토큰 수 ≈ 답변 속도 — 대화형 실사용 기준 "
                 "30~50 이상 필요 · 표는 대표 12종")
FOOTNOTE = "4bit/1bit 양자화: 모델 용량을 1/4~1/16로 압축하는 기법 (품질 손실 동반)"

# ---- 표 서식 파라미터 ----
_TBL_PT = 9.5
_CELL_MARGINS = (Inches(0.05), Inches(0.02), Inches(0.05), Inches(0.02))
_CELL_PAD_W = int(Inches(0.10))      # 좌우 여백 합 — 줄바꿈 폭 계산용
_CARD_TITLE_PT = 12.5
_CARD_BODY_PT = 10
_CARD_INSET = int(Inches(0.15))
_BULLET = "▸"                   # ▸


def _cell_min_h(text, font_pt, cell_w, floor_in=0.30):
    """셀 텍스트 기반 최소 행 높이 (넘침 금지) — 바닥값과 비교해 큰 쪽.

    rev6: 상하 버퍼 0.07→0.045in (실제 셀 상하 margin 0.02in 의 2배 여유 유지).
    판정 열 2줄 행이 4개로 늘어 — 과잉 버퍼를 줄여 표 전체가 본문 영역
    (body_bottom) 안에 들어오게 한다. add_grid_table 은 row_heights 를
    비율로 쓰므로 sum(row_min) ≤ 가용 높이면 행이 비례 확대만 된다 (축소 없음).
    """
    need = estimate_container_height(
        text, font_pt,
        max_width=max(int(cell_w) - _CELL_PAD_W, int(Inches(0.4))),
        padding_top=Inches(0.045), padding_bottom=Inches(0.045),
    )
    return max(int(Inches(floor_in)), int(need))


def _verdict_style(text):
    """판정 열 색/굵기 — 채택 = PRIMARY bold, 채택 이력 = INK, 탈락 = CHARCOAL."""
    if text.startswith("채택 —"):
        return PRIMARY, True
    if text.startswith("채택 이력"):
        return INK, False
    return CHARCOAL, False


def _card_metrics(title, items, w):
    """카드 자연 높이 산출 — (title_h, items_h, natural_h). 전부 텍스트 기반."""
    inner_w = int(w) - 2 * _CARD_INSET
    title_h = int(estimate_container_height(title, _CARD_TITLE_PT,
                                            max_width=inner_w))
    joined = "\n".join(f"{_BULLET} {it}" for it in items)
    items_h = (int(estimate_container_height(joined, _CARD_BODY_PT,
                                             max_width=inner_w))
               + int(Pt(4)) * (len(items) - 1))
    nat_h = int(Inches(0.10)) + title_h + int(Inches(0.02)) + items_h \
        + int(Inches(0.12))
    return title_h, items_h, nat_h


def _draw_card(slide, x, y, w, h, title, items, fill, border_color,
               title_h, items_h):
    """CANVAS/PRIMARY_SOFT 카드 — hairline 보더 (그림자 금지)."""
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, int(x), int(y), int(w), int(h))
    try:
        card.adjustments[0] = 0.055          # 소 radius
    except (IndexError, ValueError):
        pass
    card.fill.solid()
    card.fill.fore_color.rgb = fill
    card.line.color.rgb = border_color
    card.line.width = Pt(1.0)
    card.shadow.inherit = False              # 그림자 금지

    inner_x = int(x) + _CARD_INSET
    inner_w = int(w) - 2 * _CARD_INSET
    ty = int(y) + int(Inches(0.10))
    add_textbox(slide, inner_x, ty, inner_w, title_h, title,
                font_size=_CARD_TITLE_PT, bold=True, color=INK)
    add_bullet_list(slide, inner_x, ty + title_h + int(Inches(0.02)),
                    inner_w, items, font_size=_CARD_BODY_PT, color=CHARCOAL,
                    bullet_char=_BULLET, item_spacing=Pt(4), h=items_h)
    return card


def build_slide_07(slide):
    set_title(slide, "로컬 LLM 실측 ① — 후보 모델 테스트 결과",
              font_size=20, bold=True)
    clear_placeholders(slide, keep=[0])

    left = int(CONTENT_SAFE.left)
    width = int(CONTENT_SAFE.width)

    # ---- 하단 고정 요소 역산: 각주 → 다크 슬랩 → 본문 가용 높이 ----
    foot_h = int(estimate_container_height("※ " + FOOTNOTE, 9, max_width=width))
    foot_top = int(CONTENT_SAFE.bottom) - foot_h    # add_footnote 배치와 동일
    slab_h = int(Inches(0.50))
    slab_top = foot_top - int(Inches(0.10)) - slab_h

    # ---- 리드 (표준 규격: y=0.72" h=0.35", 15pt bold INK, 키워드만 PRIMARY) ----
    lead_box = slide.shapes.add_textbox(
        left, int(Inches(0.72)), width, int(Inches(0.35)))
    tf = lead_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    for text, color in LEAD_SEGS:
        run = p.add_run()
        run.text = text
        run.font.size = Pt(15)
        run.font.bold = True
        run.font.color.rgb = color

    body_top = int(Inches(1.18))
    body_bottom = slab_top - int(Inches(0.16))

    # ---- 좌측 (~62%) 실측 모델 표 — add_grid_table 5열 × 13행 ----
    table_w = int(Inches(6.35))
    tot_in = sum(COL_W_IN)
    col_w = [int(table_w * c / tot_in) for c in COL_W_IN]

    # 표 밀착 캡션 (design_system 2-2 — 단위 정의는 표 바로 아래 9pt GRAPHITE)
    cap_gap = int(Inches(0.02))
    caption_h = int(estimate_container_height(
        TABLE_CAPTION, 9, max_width=table_w,
        padding_top=Inches(0.02), padding_bottom=Inches(0.02)))

    # 행 높이 — 텍스트 기반 산출 (셀 넘침 금지, 판정 열 2줄 허용)
    row_min = [max(_cell_min_h(TABLE_HEADER[c], _TBL_PT, col_w[c])
                   for c in range(5))]
    for row in TABLE_ROWS:
        row_min.append(max(_cell_min_h(row[c], _TBL_PT, col_w[c])
                           for c in range(5)))
    table_h = max(body_bottom - body_top - caption_h - cap_gap,
                  sum(row_min))                           # 넘침 방지 가드

    cells = {}
    for c, text in enumerate(TABLE_HEADER):
        cells[(0, c)] = dict(
            text=text, fill=CLOUD, font_color=INK,
            font_size=_TBL_PT, bold=True, align=COL_ALIGNS[c], anchor="ctr",
            border_edges="b", border_color=STEEL, border_width_pt=1.0,
            margins=_CELL_MARGINS,
        )
    last_r = len(TABLE_ROWS)
    for i, row in enumerate(TABLE_ROWS):
        r = i + 1
        hilite = i in HILITE_ROWS
        row_fill = PRIMARY_SOFT if hilite else CANVAS
        b_color = STEEL if r == last_r else FOG        # 마지막 행만 마감선
        b_width = 0.75
        for c, text in enumerate(row):
            if c == 0:                                  # 모델명 — 식별자 INK
                f_color, f_bold = INK, (i == 0)
            elif c == 4:                                # 판정
                f_color, f_bold = _verdict_style(text)
            else:                                       # 규모/압축/속도 — 수치
                f_color, f_bold = (INK, i == 0) if hilite else (CHARCOAL, False)
            cells[(r, c)] = dict(
                text=text, fill=row_fill, font_color=f_color,
                font_size=_TBL_PT, bold=f_bold,
                align=COL_ALIGNS[c], anchor="ctr",
                border_edges="b", border_color=b_color, border_width_pt=b_width,
                margins=_CELL_MARGINS,
            )

    _, tbl = add_grid_table(
        slide, left, body_top, table_w, table_h,
        nrows=1 + len(TABLE_ROWS), ncols=5, cells=cells,
        col_widths=col_w, row_heights=row_min,
        gridlines=False,               # 수직선 없음 — 수평 hairline만 셀별 지정
        default_font_size=_TBL_PT,
    )
    # 판정 열(2줄 허용) — 한글 글자 단위 개행 금지 (eaLnBrk=0):
    # 조사·단어 중간에서 끊기지 않고 공백(단어) 경계에서만 줄바꿈된다.
    for r in range(1 + len(TABLE_ROWS)):
        for p in tbl.cell(r, 4).text_frame.paragraphs:
            p._p.get_or_add_pPr().set("eaLnBrk", "0")

    # ---- 표 바로 아래 밀착 캡션 (tok/s 정의 — 9pt GRAPHITE 1줄) ----
    add_textbox(slide, left, body_top + table_h + cap_gap, table_w, caption_h,
                TABLE_CAPTION, font_size=9, color=GRAPHITE)

    # ---- 우측 (~35%) 카드 2개 — VFlow ----
    right_x = left + table_w + int(Inches(0.18))
    right_w = int(CONTENT_SAFE.right) - right_x         # ≈ 3.75"
    gap = int(Inches(0.18))

    t1h, i1h, h1 = _card_metrics(CARD1_TITLE, CARD1_ITEMS, right_w)
    t2h, i2h, h2 = _card_metrics(CARD2_TITLE, CARD2_ITEMS, right_w)
    extra = (body_bottom - body_top - gap) - h1 - h2
    if extra > 0:                       # 잔여 공간 → 두 카드에 균등 배분
        h1 += extra // 2
        h2 += extra - extra // 2

    rflow = VFlow(x=right_x, w=right_w, y_start=body_top,
                  y_max=body_bottom, gap=gap)
    r1 = rflow.reserve(h1)
    _draw_card(slide, r1.left, r1.top, r1.width, r1.height,
               CARD1_TITLE, CARD1_ITEMS, CANVAS, STEEL, t1h, i1h)
    r2 = rflow.reserve(h2)
    _draw_card(slide, r2.left, r2.top, r2.width, r2.height,
               CARD2_TITLE, CARD2_ITEMS, PRIMARY_SOFT, PRIMARY, t2h, i2h)

    # ---- 하단 전체 폭 다크 슬랩 (INK — 마감 밴드 1곳) ----
    slab = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, left, slab_top, width, slab_h)
    slab.fill.solid()
    slab.fill.fore_color.rgb = INK
    slab.line.fill.background()
    slab.shadow.inherit = False
    stf = slab.text_frame
    stf.word_wrap = True
    stf.vertical_anchor = MSO_ANCHOR.MIDDLE
    sp = stf.paragraphs[0]
    sp.alignment = PP_ALIGN.CENTER      # 중앙 정렬은 다크 슬랩 문구만 허용
    for text, color, bold in SLAB_SEGS:
        run = sp.add_run()
        run.text = text
        run.font.size = Pt(11)
        run.font.bold = bold
        run.font.color.rgb = color

    # ---- 하단 ※ 각주 — 양자화 정의 1줄만 (design_system 2-2) ----
    add_footnote(slide, FOOTNOTE, font_size=9, color=GRAPHITE)

    # ---- 폰트 일괄 통일 (design_system.md 2절 — 맑은 고딕 단일) ----
    force_font(slide, "맑은 고딕")
