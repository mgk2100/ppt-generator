from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from ppt_utils import (
    set_title, clear_placeholders,
    add_textbox, add_para, add_rich_text,
    add_accent_bar, make_icon_circle,
    add_shadow,
    set_body_anchor,
)
from template_contract import CONTENT_SAFE, LAYOUT_CONTENT


# ============================ 팔레트 ============================
C_NAVY   = RGBColor(0x1F, 0x49, 0x7D)
C_BLUE   = RGBColor(0x4F, 0x81, 0xBD)
C_GREEN  = RGBColor(0x2E, 0x8B, 0x57)
C_YELLOW = RGBColor(0xD4, 0xA0, 0x17)
C_RED    = RGBColor(0xC0, 0x50, 0x4D)
C_INK    = RGBColor(0x1A, 0x1F, 0x2E)
C_GRAY   = RGBColor(0x66, 0x66, 0x66)
C_LGRAY  = RGBColor(0xE0, 0xE4, 0xEA)
C_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
BG_GREEN  = RGBColor(0xE8, 0xF5, 0xEC)
BG_YELLOW = RGBColor(0xFF, 0xF7, 0xD9)
BG_RED    = RGBColor(0xFC, 0xE8, 0xE6)


def build_slide_3(slide):
    set_title(slide, "영역별 예상 효과 및 투자 회수  ·  100명 기준")
    clear_placeholders(slide, keep=[0])

    # ========== CONTENT_SAFE 경계 (inches) ==========
    # left=0.28", top=0.68", right=10.56", bottom=7.02"
    # width = 10.28", height = 6.34"
    SAFE_LEFT   = Inches(0.28)
    SAFE_TOP    = Inches(0.68)
    SAFE_RIGHT  = Inches(10.56)
    SAFE_BOTTOM = Inches(7.02)
    SAFE_WIDTH  = SAFE_RIGHT - SAFE_LEFT   # 10.28"
    # SAFE_HEIGHT = SAFE_BOTTOM - SAFE_TOP  # 6.34"

    # ---------------------------------------------------------------
    # 1. Key Message Bar (상단)
    # ---------------------------------------------------------------
    km_top = SAFE_TOP
    km_height = Inches(0.48)

    # 좌측 accent vertical bar
    add_accent_bar(
        slide,
        x=SAFE_LEFT,
        y=km_top,
        w=Inches(0.08),
        h=km_height,
        color=C_NAVY,
    )

    # KeyMessage 배경 박스 (아주 연한 navy 느낌 — 여기선 white + shadow)
    km_bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        SAFE_LEFT + Inches(0.08),
        km_top,
        SAFE_WIDTH - Inches(0.08),
        km_height,
    )
    km_bg.fill.solid()
    km_bg.fill.fore_color.rgb = RGBColor(0xF4, 0xF7, 0xFB)
    km_bg.line.fill.background()

    # KeyMessage 텍스트
    km_tb = add_textbox(
        slide,
        x=SAFE_LEFT + Inches(0.22),
        y=km_top,
        w=SAFE_WIDTH - Inches(0.32),
        h=km_height,
        text="",
        font_size=12,
    )
    km_tf = km_tb.text_frame
    km_tf.word_wrap = True
    # 첫 단락 비우기 & 대체
    km_p0 = km_tf.paragraphs[0]
    km_p0.alignment = PP_ALIGN.LEFT
    # 기존 빈 run은 그대로 두고 rich_text 추가 — 대신 여기서는 add_rich_text로 새 단락 추가 후
    # 첫 단락 삭제 대신, paragraphs[0]에 직접 run 추가
    r0 = km_p0.add_run()
    r0.text = "Key Message   "
    r0.font.size = Pt(10)
    r0.font.bold = True
    r0.font.color.rgb = C_NAVY

    r1 = km_p0.add_run()
    r1.text = "영역별 효과 3~5배 차이 · 선별 적용이 핵심 · 투자 회수 1~2개월"
    r1.font.size = Pt(12)
    r1.font.bold = True
    r1.font.color.rgb = C_INK

    set_body_anchor(km_tb, 'ctr')

    # ---------------------------------------------------------------
    # 2. 신호등 테이블 (중단)
    # ---------------------------------------------------------------
    # 영역: top = km_bottom + gap, height = 공간의 ~55%
    tbl_top = km_top + km_height + Inches(0.15)
    # 하단 카드 영역 확보: 카드 높이 1.55" + 상단 소제목 0.28" + gap 0.18"
    # 카드 영역 전체 ≈ 2.0"
    cards_block_h = Inches(2.10)
    cards_top = SAFE_BOTTOM - cards_block_h  # = 7.02 - 2.10 = 4.92"

    tbl_bottom = cards_top - Inches(0.15)   # 4.77"
    tbl_height = tbl_bottom - tbl_top       # 4.77 - (0.68+0.48+0.15) = 4.77 - 1.31 = 3.46"

    # Stoplight 소제목
    st_title_h = Inches(0.28)
    add_textbox(
        slide,
        x=SAFE_LEFT,
        y=tbl_top,
        w=SAFE_WIDTH,
        h=st_title_h,
        text="① 영역별 효과 (신호등 평가)",
        font_size=12,
        bold=True,
        color=C_NAVY,
    )

    # 헤더 + 4행 데이터 영역
    tbl_body_top = tbl_top + st_title_h + Inches(0.04)
    tbl_body_h = tbl_height - st_title_h - Inches(0.04)  # ~3.14"
    # 헤더 0.30", 각 행 = (3.14 - 0.30 - 3*0.06) / 4
    header_h = Inches(0.30)
    row_gap = Inches(0.06)
    rows_n = 4
    total_rows_h = tbl_body_h - header_h - Inches(0.04) - row_gap * (rows_n - 1)
    row_h = int(total_rows_h / rows_n)

    # 열 너비 (총 SAFE_WIDTH = 10.28")
    # dot + 영역(5.2") / 효과(1.8") / 해석(3.28")
    col_dot_w = Inches(0.34)       # dot 영역
    col_area_w = Inches(4.96)      # 영역명
    col_effect_w = Inches(1.80)    # 효과%
    col_interp_w = Inches(3.18)    # 해석
    # 합: 0.34 + 4.96 + 1.80 + 3.18 = 10.28 ✓

    # 각 열의 left
    col_dot_x = SAFE_LEFT + Inches(0.02)
    col_area_x = SAFE_LEFT + col_dot_w
    col_effect_x = col_area_x + col_area_w
    col_interp_x = col_effect_x + col_effect_w

    # ---- 헤더 행 ----
    header_bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        SAFE_LEFT,
        tbl_body_top,
        SAFE_WIDTH,
        header_h,
    )
    header_bg.fill.solid()
    header_bg.fill.fore_color.rgb = C_NAVY
    header_bg.line.fill.background()

    # 헤더 텍스트
    add_textbox(
        slide,
        x=col_area_x + Inches(0.10),
        y=tbl_body_top,
        w=col_area_w - Inches(0.10),
        h=header_h,
        text="영역",
        font_size=10,
        bold=True,
        color=C_WHITE,
        align=PP_ALIGN.LEFT,
    )
    add_textbox(
        slide,
        x=col_effect_x,
        y=tbl_body_top,
        w=col_effect_w,
        h=header_h,
        text="효과",
        font_size=10,
        bold=True,
        color=C_WHITE,
        align=PP_ALIGN.CENTER,
    )
    add_textbox(
        slide,
        x=col_interp_x + Inches(0.10),
        y=tbl_body_top,
        w=col_interp_w - Inches(0.10),
        h=header_h,
        text="해석",
        font_size=10,
        bold=True,
        color=C_WHITE,
        align=PP_ALIGN.LEFT,
    )

    # ---- 데이터 행들 ----
    rows_data = [
        {
            "dot": C_GREEN,
            "bg": BG_GREEN,
            "area": "선행 R&D · 도구 · 테스트 자동화",
            "effect": "+20 ~ +35%",
            "effect_color": C_GREEN,
            "interp": "즉시 도입 권장 (위험 낮음 · 효과 큼)",
        },
        {
            "dot": C_GREEN,
            "bg": BG_GREEN,
            "area": "양산 일반 SW (비안전)",
            "effect": "+10 ~ +20%",
            "effect_color": C_GREEN,
            "interp": "코드 검사 도구 통합 후 도입",
        },
        {
            "dot": C_YELLOW,
            "bg": BG_YELLOW,
            "area": "양산 안전 코드 (차량 안전등급 높음↑)",
            "effect": "-5 ~ +10%",
            "effect_color": C_YELLOW,
            "interp": "본전 영역 · 시범 측정 필요",
        },
        {
            "dot": C_RED,
            "bg": BG_RED,
            "area": "품질검증 단계 (인력 관점)",
            "effect": "-15 ~ -30%",
            "effect_color": C_RED,
            "interp": "코드 리뷰 부담 증가 · 인력 보강 필수",
        },
    ]

    row_y = tbl_body_top + header_h + Inches(0.04)
    for row in rows_data:
        # 행 배경
        bg_shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            SAFE_LEFT,
            row_y,
            SAFE_WIDTH,
            row_h,
        )
        bg_shape.fill.solid()
        bg_shape.fill.fore_color.rgb = row["bg"]
        bg_shape.line.fill.background()

        # dot (OVAL) - 세로 중앙 정렬
        dot_size = Inches(0.18)
        dot_x = SAFE_LEFT + Inches(0.08)
        dot_y = row_y + int((row_h - dot_size) / 2)
        make_icon_circle(
            slide,
            x=dot_x,
            y=dot_y,
            size=dot_size,
            fill_color=row["dot"],
            text="",
        )

        # 영역명
        area_tb = add_textbox(
            slide,
            x=col_area_x + Inches(0.02),
            y=row_y,
            w=col_area_w - Inches(0.08),
            h=row_h,
            text=row["area"],
            font_size=12,
            bold=True,
            color=C_INK,
            align=PP_ALIGN.LEFT,
        )
        set_body_anchor(area_tb, 'ctr')

        # 효과%
        effect_tb = add_textbox(
            slide,
            x=col_effect_x,
            y=row_y,
            w=col_effect_w,
            h=row_h,
            text=row["effect"],
            font_size=14,
            bold=True,
            color=row["effect_color"],
            align=PP_ALIGN.CENTER,
        )
        set_body_anchor(effect_tb, 'ctr')

        # 해석
        interp_tb = add_textbox(
            slide,
            x=col_interp_x + Inches(0.10),
            y=row_y,
            w=col_interp_w - Inches(0.18),
            h=row_h,
            text=row["interp"],
            font_size=11,
            color=C_INK,
            align=PP_ALIGN.LEFT,
        )
        set_body_anchor(interp_tb, 'ctr')

        row_y += row_h + row_gap

    # ---------------------------------------------------------------
    # 3. 큰 숫자 카드 (하단, 1x3)
    # ---------------------------------------------------------------
    cards_title_h = Inches(0.28)
    add_textbox(
        slide,
        x=SAFE_LEFT,
        y=cards_top,
        w=SAFE_WIDTH,
        h=cards_title_h,
        text="② 투자 회수 효과 (100명 · 1년 기준)",
        font_size=12,
        bold=True,
        color=C_NAVY,
    )

    card_block_top = cards_top + cards_title_h + Inches(0.04)
    card_block_h = SAFE_BOTTOM - card_block_top  # ~ 2.10 - 0.28 - 0.04 = 1.78"
    # 안전 여유: 카드 높이 조금 줄이기
    card_h = card_block_h - Inches(0.02)  # ~1.76"

    card_gap = Inches(0.22)
    card_w = int((SAFE_WIDTH - card_gap * 2) / 3)

    card_defs = [
        {
            "label": "연 절감 효과",
            "value": "약 15억 원",
            "sub": "21,000시간 × 시간당 7.2만",
            "accent": C_GREEN,
        },
        {
            "label": "초기 투자",
            "value": "약 1.2억 원",
            "sub": "도구 3천 + 운영 9천만원",
            "accent": C_BLUE,
        },
        {
            "label": "손익분기점",
            "value": "1 ~ 2개월",
            "sub": "보수 보정 시 2~3개월",
            "accent": C_NAVY,
        },
    ]

    # 3단 레이아웃: label 0.32 / value 중앙 0.80 / sub 하단 0.32
    label_h = Inches(0.34)
    sub_h = Inches(0.34)
    accent_bar_h = Inches(0.08)

    cx = SAFE_LEFT
    for cd in card_defs:
        # 카드 배경 (white)
        card_bg = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            cx,
            card_block_top,
            card_w,
            card_h,
        )
        card_bg.adjustments[0] = 0.08
        card_bg.fill.solid()
        card_bg.fill.fore_color.rgb = C_WHITE
        card_bg.line.color.rgb = C_LGRAY
        card_bg.line.width = Pt(0.75)
        add_shadow(card_bg, blur_pt=3, dist_pt=2)

        # 상단 accent 바
        add_accent_bar(
            slide,
            x=cx + Inches(0.10),
            y=card_block_top + Inches(0.10),
            w=card_w - Inches(0.20),
            h=accent_bar_h,
            color=cd["accent"],
        )

        # label (상단)
        label_top = card_block_top + Inches(0.10) + accent_bar_h + Inches(0.02)
        label_tb = add_textbox(
            slide,
            x=cx + Inches(0.12),
            y=label_top,
            w=card_w - Inches(0.24),
            h=label_h,
            text=cd["label"],
            font_size=10,
            bold=True,
            color=C_GRAY,
            align=PP_ALIGN.CENTER,
        )
        set_body_anchor(label_tb, 'ctr')

        # value (중앙, 큰 글자)
        value_top = label_top + label_h
        # card 밑여유: 하단 sub 높이 + 여백
        value_h = card_h - (value_top - card_block_top) - sub_h - Inches(0.06)
        value_tb = add_textbox(
            slide,
            x=cx + Inches(0.10),
            y=value_top,
            w=card_w - Inches(0.20),
            h=value_h,
            text=cd["value"],
            font_size=28,
            bold=True,
            color=cd["accent"],
            align=PP_ALIGN.CENTER,
        )
        set_body_anchor(value_tb, 'ctr')

        # sub (하단)
        sub_top = card_block_top + card_h - sub_h - Inches(0.04)
        sub_tb = add_textbox(
            slide,
            x=cx + Inches(0.10),
            y=sub_top,
            w=card_w - Inches(0.20),
            h=sub_h,
            text=cd["sub"],
            font_size=9,
            color=C_GRAY,
            align=PP_ALIGN.CENTER,
        )
        set_body_anchor(sub_tb, 'ctr')

        cx += card_w + card_gap
