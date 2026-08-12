"""slide_08 — 로컬 LLM 실측 ② : 동시 처리 한계 · 프론티어급 구동 조건.

design_system.md (getdesign hp) 강제:
- 색상은 토큰만 (PRIMARY 계열 + 무채색), 이모지·그림자·그라디언트 금지
- 맑은 고딕 단일 (함수 말미 force_font 일괄 강제), 카드 lift 는 hairline, 표는 수평 FOG hairline 위주
- 좌측: 수제 수평 막대 3개 (30:2:1 선형 비례, 최소 폭 0.09in) — add_chart 미사용,
  섹션 제목 아래 기준 모델 부제 1줄 (Gemma-4-31B 4bit · GPU 96GB 1장, 10pt GRAPHITE),
  카테고리 라벨 2줄 허용 (rev4: 문서 분량 수치 병기, 2줄째 8.5pt GRAPHITE, 라벨 열 1.7in)
- 우측: Kimi K3 자체 구동 산정 add_grid_table (헤더 + 5행)
- 하단: INK 다크 슬랩 2줄 + NVLink 각주
"""

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from ppt_utils import (
    set_title, clear_placeholders,
    add_textbox, add_rich_text, add_accent_bar,
    add_grid_table, add_footnote,
    set_body_anchor, set_text_inset,
    estimate_container_height, force_font,
)
from template_contract import CONTENT_SAFE

# ---------------- 디자인 토큰 (design_system.md — 이 목록 밖 색 금지) ----------------
PRIMARY        = RGBColor(0x02, 0x4A, 0xD8)
PRIMARY_BRIGHT = RGBColor(0x29, 0x6E, 0xF9)
INK            = RGBColor(0x1A, 0x1A, 0x1A)
CHARCOAL       = RGBColor(0x3D, 0x3D, 0x3D)
GRAPHITE       = RGBColor(0x63, 0x63, 0x63)
CANVAS         = RGBColor(0xFF, 0xFF, 0xFF)
CLOUD          = RGBColor(0xF7, 0xF7, 0xF7)
FOG            = RGBColor(0xE8, 0xE8, 0xE8)
STEEL          = RGBColor(0xC2, 0xC2, 0xC2)
ON_INK         = RGBColor(0xFF, 0xFF, 0xFF)

F_REG  = "Pretendard"
F_SEMI = "Pretendard SemiBold"

IN = Inches


def _rich_box(slide, x, y, w, h, segments, align=PP_ALIGN.LEFT,
              line_spacing=None):
    """혼합 서식 단락 1개짜리 텍스트박스 (선두 빈 단락 제거)."""
    tb = slide.shapes.add_textbox(int(x), int(y), int(w), int(h))
    tf = tb.text_frame
    tf.word_wrap = True
    add_rich_text(tf, segments, align=align, line_spacing=line_spacing)
    p0 = tf.paragraphs[0]._p
    p0.getparent().remove(p0)
    return tb


def _flat_rect(slide, shape_type, x, y, w, h, fill):
    """보더 없는 단색 면 (그림자 금지 — hairline/밴드 문법)."""
    shp = slide.shapes.add_shape(shape_type, int(x), int(y), int(w), int(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def build_slide_8(slide):
    set_title(slide, "로컬 LLM 실측 ② — 동시 처리 한계 · 프론티어급 구동 조건",
              font_size=20, bold=True)
    clear_placeholders(slide, keep=[0])

    # ---------------- 수직 배분 (spec notes: 리드 / 본문 2열 / 슬랩 / 각주) ----------------
    LEAD_Y      = IN(0.72)                                 # 표준 리드 y (design 2-1)
    LEAD_H      = IN(0.35)
    BODY_TOP    = 1.18                                     # 본문 시작 y ≥ 1.18"
    SLAB_TOP    = 5.94
    SLAB_H      = 0.72
    BODY_BOTTOM = SLAB_TOP - 0.15                          # 5.79
    BODY_H      = BODY_BOTTOM - BODY_TOP                   # 4.63

    # ---------------- 리드 (SemiBold 15pt INK, 키워드 PRIMARY) ----------------
    lead = _rich_box(
        slide, CONTENT_SAFE.left, LEAD_Y, CONTENT_SAFE.width, LEAD_H,
        [
            {"text": "긴 문서일수록 동시 사용자 ", "font_name": F_SEMI,
             "font_size": 15, "color": INK},
            {"text": "급감", "font_name": F_SEMI, "font_size": 15,
             "color": PRIMARY},
            {"text": " — 프론티어급 자체 구동은 ", "font_name": F_SEMI,
             "font_size": 15, "color": INK},
            {"text": "데이터센터급", "font_name": F_SEMI, "font_size": 15,
             "color": PRIMARY},
            {"text": " GPU 영역", "font_name": F_SEMI, "font_size": 15,
             "color": INK},
        ],
    )
    set_text_inset(lead, 0, 0, 0, 0)

    # ---------------- 2열 분할 (좌 ~48% / 우 ~52%) ----------------
    COL_GAP = 0.22
    left_x  = 0.28
    left_w  = round((10.28 - COL_GAP) * 0.48, 2)           # 4.83
    right_x = left_x + left_w + COL_GAP                    # 5.33
    right_w = 10.56 - right_x                              # 5.23

    # ============ 좌측 — 동시 처리 인원 수제 막대 (CLOUD 섹션 밴드 위) ============
    _flat_rect(slide, MSO_SHAPE.RECTANGLE,
               IN(left_x), IN(BODY_TOP), IN(left_w), IN(BODY_H), CLOUD)

    pad     = 0.18
    inner_x = left_x + pad                                 # 0.46
    inner_w = left_w - 2 * pad                             # 4.47

    hdr = add_textbox(slide, IN(inner_x), IN(BODY_TOP + 0.16),
                      IN(inner_w), IN(0.30),
                      "사용 형태별 동시 처리 인원",
                      font_name=F_SEMI, font_size=12.5, color=INK)
    set_text_inset(hdr, 0, 0, 0, 0)

    # 기준 모델 부제 (10pt GRAPHITE) — rev4: '자사 라이브 모델' 문구 삭제
    sub_text = "기준: Gemma-4-31B (4bit) · GPU 96GB 1장 서빙 실측"
    sub_h_emu = estimate_container_height(sub_text, 10, max_width=IN(inner_w))
    sub = add_textbox(slide, IN(inner_x), IN(BODY_TOP + 0.16 + 0.32),
                      IN(inner_w), sub_h_emu, sub_text,
                      font_name=F_REG, font_size=10, color=GRAPHITE)
    set_text_inset(sub, 0, 0, 0, 0)

    # 하단 캡션 (9.5pt GRAPHITE, 2줄) — 밴드 바닥 기준 역산
    cap_text = ("GPU 메모리에서 모델을 제외한 남은 공간에 대화 내용을 담는 구조 — "
                "문서가 길수록 이 공간이 빨리 소진 · GPU 를 4장으로 늘리면 동시 처리 인원도 약 4배")
    cap_h_emu = estimate_container_height(cap_text, 9.5, max_width=IN(inner_w))
    cap_y_emu = IN(BODY_BOTTOM - pad) - cap_h_emu
    cap = add_textbox(slide, IN(inner_x), cap_y_emu, IN(inner_w), cap_h_emu,
                      cap_text, font_name=F_REG, font_size=9.5, color=GRAPHITE)
    set_text_inset(cap, 0, 0, 0, 0)

    # 막대 영역 — 부제 높이만큼 아래로 (캡션·경계 역산이라 겹침 없음)
    bars_top    = BODY_TOP + 0.16 + 0.32 + sub_h_emu / 914400.0 + 0.14
    bars_bottom = BODY_BOTTOM - pad - cap_h_emu / 914400.0 - 0.18
    row_h       = (bars_bottom - bars_top) / 3.0

    LABEL_W = 1.70                                         # rev4: 2줄 라벨 수용 위해 확대
    VAL_W   = 0.82
    bar_x   = inner_x + LABEL_W + 0.10                     # 막대 시작 x 동기 이동
    BAR_MAX = inner_w - LABEL_W - 0.10 - 0.06 - VAL_W      # ≈ 1.79
    MIN_BAR = 0.09
    BAR_H   = 0.34
    MAX_VAL = 30.0

    # rev4: 문서 분량 수치 병기 (2줄째 8.5pt GRAPHITE)
    bars = [
        ("짧은 질의응답 (챗봇)", None,                          30, "30명 이상"),
        ("중간 길이 문서",              "(약 1.6만 토큰 · A4 약 10장)",  2, "약 2명"),
        ("긴 문서 · 이력 분석",         "(약 3.2만 토큰 · A4 약 20장)",  1, "약 1명"),
    ]

    # 기준선 — FOG hairline 세로 1개 (막대 시작선)
    add_accent_bar(slide, IN(bar_x), IN(bars_top + 0.04), IN(0.01),
                   IN(bars_bottom - bars_top - 0.08), FOG)

    for i, (label, sub_label, value, val_text) in enumerate(bars):
        row_top = bars_top + i * row_h
        row_cy  = row_top + row_h / 2.0

        # 카테고리 라벨 (좌측, 1줄째 10pt CHARCOAL / 2줄째 8.5pt GRAPHITE)
        lbl_h_emu = estimate_container_height(label, 10, max_width=IN(LABEL_W))
        if sub_label:
            lbl_h_emu += estimate_container_height(
                sub_label, 8.5, max_width=IN(LABEL_W))
        lbl = _rich_box(
            slide, IN(inner_x), IN(row_cy) - int(lbl_h_emu / 2),
            IN(LABEL_W), lbl_h_emu,
            [{"text": label, "font_name": F_REG, "font_size": 10,
              "color": CHARCOAL}],
            align=PP_ALIGN.RIGHT,
        )
        if sub_label:
            add_rich_text(
                lbl.text_frame,
                [{"text": sub_label, "font_name": F_REG, "font_size": 8.5,
                  "color": GRAPHITE}],
                align=PP_ALIGN.RIGHT, space_before=Pt(1),
            )
        set_body_anchor(lbl, "ctr")
        set_text_inset(lbl, 0, 0, IN(0.02), 0)

        # 막대 — 선형 비례 30:2:1 (스케일 왜곡 금지), 최소 가시 폭 보장
        bar_len = max(value / MAX_VAL * BAR_MAX, MIN_BAR)
        bar_y   = row_cy - BAR_H / 2.0
        bar = _flat_rect(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                         IN(bar_x), IN(bar_y), IN(bar_len), IN(BAR_H), PRIMARY)
        try:
            bar.adjustments[0] = 0.10                      # radius 소
        except Exception:
            pass

        # 값 라벨 — 막대 끝 우측 SemiBold 11pt INK (시리즈색 텍스트 금지)
        val = add_textbox(slide, IN(bar_x + bar_len + 0.06),
                          IN(row_cy - 0.13), IN(VAL_W), IN(0.26), val_text,
                          font_name=F_SEMI, font_size=11, color=INK)
        set_text_inset(val, 0, 0, 0, 0)

    # ============ 우측 — Kimi K3 자체 구동 산정 (add_grid_table) ============
    rt = add_textbox(slide, IN(right_x), IN(BODY_TOP), IN(right_w), IN(0.30),
                     "프론티어급 오픈 모델 자체 구동 산정 — Kimi K3 기준",
                     font_name=F_SEMI, font_size=12.5, color=INK)
    set_text_inset(rt, 0, 0, 0, 0)

    rs = add_textbox(slide, IN(right_x), IN(BODY_TOP + 0.36),
                     IN(right_w), IN(0.24),
                     "2.8조 파라미터 — 상용 프론티어(Claude · GPT)급 성능의 최상위 오픈 모델",
                     font_name=F_REG, font_size=9.5, color=GRAPHITE)
    set_text_inset(rs, 0, 0, 0, 0)

    TBL_Y = BODY_TOP + 0.36 + 0.24 + 0.14                  # 1.92

    # 실측 근거 캡션 (frontier_evidence_caption) 공간 — 표 높이를 줄여 확보
    ev_h_emu = estimate_container_height(
        ("실측 — 프론티어급 GLM-5.2(754B)를 1bit 극압축(217GB)으로 구동 시 "
         "4.9 tok/s, 1명 사용도 곤란 → 산정이 아닌 실측으로 확인된 한계"),
        9.5, max_width=IN(right_w))
    TBL_H = BODY_BOTTOM - TBL_Y - ev_h_emu / 914400.0 - 0.12

    rows = [
        ("모델 용량 (4bit 압축)",   "약 1.4TB",                    False),
        ("필요 GPU (96GB 기준)",    "모델 적재 약 16장 → 대화 메모리 포함 약 20장", False),
        ("서버 환산 (1대 = 4장)",   "약 4~5대",                    False),
        ("시스템 메모리",           "서버당 512GB × 4~5대 = 총 2TB 이상", False),
        ("압축 없이 (bf16) 구동",   "GPU 약 58장",                 True),   # 값 PRIMARY
    ]

    M = (IN(0.12), IN(0.04), IN(0.08), IN(0.04))
    cells = {
        (0, 0): dict(text="항목", fill=CLOUD, font_color=INK,
                     font_name=F_SEMI, font_size=11, align="l",
                     border_edges="b", border_color=STEEL,
                     border_width_pt=1.0, margins=M),
        (0, 1): dict(text="산정 결과", fill=CLOUD, font_color=INK,
                     font_name=F_SEMI, font_size=11, align="l",
                     border_edges="b", border_color=STEEL,
                     border_width_pt=1.0, margins=M),
    }
    for i, (k, v, hot) in enumerate(rows, start=1):
        cells[(i, 0)] = dict(text=k, font_color=CHARCOAL, font_name=F_REG,
                             font_size=10.5, align="l",
                             border_edges="b", border_color=FOG,
                             border_width_pt=0.75, margins=M)
        cells[(i, 1)] = dict(text=v, font_color=(PRIMARY if hot else INK),
                             font_name=F_SEMI, font_size=11, align="l",
                             border_edges="b", border_color=FOG,
                             border_width_pt=0.75, margins=M)

    add_grid_table(
        slide, IN(right_x), IN(TBL_Y), IN(right_w), IN(TBL_H),
        nrows=6, ncols=2, cells=cells,
        col_widths=[0.95, 1.05],
        row_heights=[0.75, 1, 1, 1, 1, 1],
        default_fill=CANVAS, gridlines=False,
    )

    # 실측 근거 캡션 — 표 바로 아래 2줄 (9.5pt GRAPHITE, '4.9 tok/s' 만 bold INK)
    ev = _rich_box(
        slide, IN(right_x), IN(TBL_Y + TBL_H + 0.12), IN(right_w), ev_h_emu,
        [
            {"text": ("실측 — 프론티어급 GLM-5.2(754B)를 1bit 극압축(217GB)으로 "
                      "구동 시 "),
             "font_name": F_REG, "font_size": 9.5, "color": GRAPHITE},
            {"text": "4.9 tok/s", "font_name": F_SEMI, "font_size": 9.5,
             "color": INK, "bold": True},
            {"text": ", 1명 사용도 곤란 → 산정이 아닌 실측으로 확인된 한계",
             "font_name": F_REG, "font_size": 9.5, "color": GRAPHITE},
        ],
    )
    set_text_inset(ev, 0, 0, 0, 0)

    # ============ 하단 — 다크 슬랩 (INK 배경, 2줄) ============
    _flat_rect(slide, MSO_SHAPE.RECTANGLE,
               IN(0.28), IN(SLAB_TOP), IN(10.28), IN(SLAB_H), INK)

    slab_tb = _rich_box(
        slide, IN(0.48), IN(SLAB_TOP + 0.02), IN(9.88), IN(SLAB_H - 0.04),
        [
            {"text": ("RTX PRO 6000 은 NVLink 미지원 — 4장 초과 분산 시 통신 병목, "
                      "프론티어급은 H200 / B200 등 데이터센터급 GPU 영역"),
             "font_name": F_REG, "font_size": 10.5, "color": ON_INK},
        ],
        align=PP_ALIGN.CENTER, line_spacing=Pt(14),
    )
    add_rich_text(
        slab_tb.text_frame,
        [
            {"text": "GPU 4장 규모의 현실적 운용 = ", "font_name": F_SEMI,
             "font_size": 12, "color": ON_INK},
            {"text": "30B급 전후", "font_name": F_SEMI, "font_size": 12,
             "color": PRIMARY_BRIGHT},
            {"text": " 오픈소스 모델", "font_name": F_SEMI, "font_size": 12,
             "color": ON_INK},
        ],
        align=PP_ALIGN.CENTER, space_before=Pt(3), line_spacing=Pt(16),
    )
    set_body_anchor(slab_tb, "ctr")
    set_text_inset(slab_tb, IN(0.1), IN(0.02), IN(0.1), IN(0.02))

    # ---------------- 각주 (NVLink) ----------------
    fn = add_footnote(
        slide, "NVLink: GPU 간 고속 직결 인터커넥트 — PCIe 대비 수 배 대역폭",
        color=GRAPHITE)
    for p in fn.text_frame.paragraphs:
        for r in p.runs:
            r.font.name = F_REG

    # ---------------- 전 run 맑은 고딕 일괄 강제 (design 2절 — 함수 마지막) ----------------
    force_font(slide, "맑은 고딕")
