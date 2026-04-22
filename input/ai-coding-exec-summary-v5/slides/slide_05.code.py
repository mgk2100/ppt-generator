"""Slide 05 — 실행 로드맵 및 권고 사항

본문: y=0.68"~6.55"
- Key Message Bar
- CHEVRON 3단계 (Phase 1 단기 green / Phase 2 중기 blue / Phase 3 장기 navy)
- 하단 2열: ② 자사 시범 사업 (info_card, BG_BLUE) / ③ 업계 권고 사항 (recommendation_card, BG_TEAL)

각주: y=6.60"~7.00" — EU AI Act 용어 1줄 (9pt gray)

제약: "승인 요청" 금지. "추정 ROI 11배" 금지. Phase 기간 단기/중기/장기만.
"""

from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from ppt_utils import (
    set_title, clear_placeholders,
    add_textbox, add_para, add_rich_text,
    add_accent_bar, make_icon_circle, make_icon_badge,
    add_shadow, set_shape_opacity,
    set_body_anchor, set_text_inset,
)
from template_contract import CONTENT_SAFE, LAYOUT_CONTENT


# 색상 팔레트
C_NAVY  = RGBColor(0x1F, 0x49, 0x7D)
C_BLUE  = RGBColor(0x4F, 0x81, 0xBD)
C_GREEN = RGBColor(0x2E, 0x8B, 0x57)
C_TEAL  = RGBColor(0x2C, 0x7F, 0x94)
C_INK   = RGBColor(0x1A, 0x1F, 0x2E)
C_GRAY  = RGBColor(0x66, 0x66, 0x66)
C_LGRAY = RGBColor(0xE0, 0xE4, 0xEA)
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BG_BLUE = RGBColor(0xE7, 0xF0, 0xFA)
BG_TEAL = RGBColor(0xE3, 0xF2, 0xF7)


def build_slide_5(slide):
    set_title(slide, "실행 로드맵 및 권고 사항")
    clear_placeholders(slide, keep=[0])

    # ---------------------------------------------------------------
    # 0) 영역 상수
    # ---------------------------------------------------------------
    LEFT   = Inches(0.28)
    RIGHT  = Inches(10.56)
    WIDTH  = RIGHT - LEFT                       # 10.28"
    BODY_TOP    = Inches(0.68)
    BODY_BOTTOM = Inches(6.55)
    FOOT_TOP    = Inches(6.60)
    FOOT_BOTTOM = Inches(7.00)

    # ---------------------------------------------------------------
    # 1) Key Message Bar
    # ---------------------------------------------------------------
    km_y = BODY_TOP
    km_h = Inches(0.52)
    km_bar_w = Inches(0.09)

    # 좌측 accent bar (teal)
    add_accent_bar(slide, LEFT, km_y, km_bar_w, km_h, C_TEAL)
    # 배경
    km_bg_x = LEFT + km_bar_w
    km_bg_w = WIDTH - km_bar_w
    km_bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, km_bg_x, km_y, km_bg_w, km_h
    )
    km_bg.fill.solid()
    km_bg.fill.fore_color.rgb = BG_TEAL
    km_bg.line.fill.background()

    # 본문 텍스트
    km_tb = add_textbox(
        slide,
        km_bg_x + Inches(0.14), km_y,
        km_bg_w - Inches(0.28), km_h,
        "", font_size=12,
    )
    km_tf = km_tb.text_frame
    km_tf.word_wrap = True
    km_tf.margin_left = Inches(0.02)
    km_tf.margin_right = Inches(0.02)
    km_tf.margin_top = Inches(0.04)
    km_tf.margin_bottom = Inches(0.04)
    set_body_anchor(km_tb, "ctr")
    # 기존 paragraph 사용
    p0 = km_tf.paragraphs[0]
    p0.alignment = PP_ALIGN.LEFT
    # 기본 run 비움
    for _r in list(p0.runs):
        _r.text = ""
    # rich text
    r1 = p0.add_run()
    r1.text = "단계적 도입 + 자사 시범"
    r1.font.size = Pt(13)
    r1.font.bold = True
    r1.font.color.rgb = C_NAVY
    r2 = p0.add_run()
    r2.text = "으로 실측  ·  "
    r2.font.size = Pt(12)
    r2.font.color.rgb = C_INK
    r3 = p0.add_run()
    r3.text = "선행 영역 우선 착수"
    r3.font.size = Pt(13)
    r3.font.bold = True
    r3.font.color.rgb = C_TEAL
    r4 = p0.add_run()
    r4.text = "  (업계 권고)"
    r4.font.size = Pt(11)
    r4.font.color.rgb = C_GRAY

    # ---------------------------------------------------------------
    # 2) Chevron 3단계 (섹션 제목 + 3 chevron)
    # ---------------------------------------------------------------
    sec1_y = km_y + km_h + Inches(0.14)
    # 섹션 제목
    sec1_title_h = Inches(0.28)
    add_textbox(
        slide, LEFT, sec1_y, WIDTH, sec1_title_h,
        "① 단계적 도입 로드맵 (업계 일반 권고)",
        font_size=12, bold=True, color=C_NAVY, align=PP_ALIGN.LEFT,
    )

    # Chevron 영역
    chev_y = sec1_y + sec1_title_h + Inches(0.04)
    chev_h = Inches(1.02)
    # 3 chevron 균등 배치. 슬라이더처럼 겹쳐진 효과를 주기 위해 약간 오버랩
    chev_gap = Inches(0.06)
    chev_count = 3
    chev_w = (WIDTH - chev_gap * (chev_count - 1)) // chev_count

    phases = [
        {
            "title": "Phase 1  ·  단기",
            "body": "도구 · 테스트 자동화 파일럿\n위험 낮음 · 효과 큼",
            "color": C_GREEN,
        },
        {
            "title": "Phase 2  ·  중기",
            "body": "선행 R&D · 응용SW 확대\n사내 AI 학습 · 자동 검사 통합",
            "color": C_BLUE,
        },
        {
            "title": "Phase 3  ·  장기",
            "body": "양산 비안전 영역 단계 적용\n양산 안전은 시범 측정 후",
            "color": C_NAVY,
        },
    ]

    cx = LEFT
    for ph in phases:
        chev = slide.shapes.add_shape(
            MSO_SHAPE.CHEVRON, cx, chev_y, chev_w, chev_h
        )
        chev.fill.solid()
        chev.fill.fore_color.rgb = ph["color"]
        chev.line.fill.background()
        add_shadow(chev, blur_pt=4, dist_pt=2, opacity_pct=22)

        tf = chev.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.22)
        tf.margin_right = Inches(0.40)  # chevron 꼬리 여유
        tf.margin_top = Inches(0.10)
        tf.margin_bottom = Inches(0.10)
        # 기존 paragraph 제목
        p_t = tf.paragraphs[0]
        p_t.alignment = PP_ALIGN.LEFT
        rt = p_t.add_run()
        rt.text = ph["title"]
        rt.font.size = Pt(13)
        rt.font.bold = True
        rt.font.color.rgb = C_WHITE

        # 본문
        body_lines = ph["body"].split("\n")
        for i, line in enumerate(body_lines):
            add_para(
                tf, line,
                font_size=10, color=C_WHITE, bold=False,
                align=PP_ALIGN.LEFT,
                space_before=Pt(4) if i == 0 else Pt(2),
                space_after=Pt(0),
            )
        set_body_anchor(chev, "ctr")
        cx += chev_w + chev_gap

    # ---------------------------------------------------------------
    # 3) 하단 2열 (② 좌 info_card / ③ 우 recommendation_card)
    # ---------------------------------------------------------------
    lower_y = chev_y + chev_h + Inches(0.20)
    lower_h = BODY_BOTTOM - lower_y            # ~3.01"
    col_gap = Inches(0.20)
    col_w = (WIDTH - col_gap) // 2

    # ---- 좌측 info_card ----
    left_x = LEFT
    _build_info_card(
        slide,
        x=left_x, y=lower_y, w=col_w, h=lower_h,
        title="② 자사 시범 사업 (실측 프로토콜)",
        accent=C_BLUE, bg=BG_BLUE,
        items=[
            ("대상",       "영역별 대표 SW 모듈"),
            ("기간",       "별도 산정"),
            ("규모·예산",  "별도 산정 (파일럿 계획 수립 단계)"),
            ("측정 지표",  "공수 · 결함밀도 · 코드 리뷰 부담"),
            ("목적",       "자사 환경에서 영역별 실효성 실측"),
        ],
    )

    # ---- 우측 recommendation_card ----
    right_x = left_x + col_w + col_gap
    _build_recommendation_card(
        slide,
        x=right_x, y=lower_y, w=col_w, h=lower_h,
        title="③ 업계 권고 사항",
        accent=C_TEAL, bg=BG_TEAL,
        items=[
            ("①", "선행 · 테스트 자동화 영역 먼저 검토",
                  "경쟁사 모두 이 영역부터 착수"),
            ("②", "자사 시범 사업으로 실효성 실측",
                  "양산 안전 영역은 공개 데이터 부재"),
            ("③", "품질검증 인력·도구 동시 보강",
                  "AI 코드량 증가는 리뷰 부담 증가로 직결"),
            ("④", "사내 AI 사용 가이드라인 정비",
                  "EU AI Act 2027.08 자동차 완전 적용"),
        ],
    )

    # ---------------------------------------------------------------
    # 4) 각주 — EU AI Act 용어 (y 6.60~7.00)
    # ---------------------------------------------------------------
    foot_h = FOOT_BOTTOM - FOOT_TOP
    add_textbox(
        slide,
        LEFT, FOOT_TOP, WIDTH, foot_h,
        "※ EU AI Act: EU 인공지능 규제법. 2024.08 발효 · 2027.08 자동차 등 "
        "embedded high-risk AI 완전 적용 D-Day",
        font_size=9, color=C_GRAY, align=PP_ALIGN.LEFT,
    )


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _build_info_card(slide, x, y, w, h, title, accent, bg, items):
    """② 자사 시범 사업 — 상단 타이틀 바 + key:value 목록."""
    # 배경
    bg_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = bg
    bg_shape.line.color.rgb = accent
    bg_shape.line.width = Pt(0.5)
    add_shadow(bg_shape, blur_pt=4, dist_pt=2, opacity_pct=14)

    # 상단 accent bar (좌측 세로)
    bar_w = Inches(0.08)
    add_accent_bar(slide, x, y, bar_w, h, accent)

    inner_x = x + bar_w + Inches(0.14)
    inner_w = w - bar_w - Inches(0.28)
    inner_y = y + Inches(0.14)

    # 타이틀
    title_h = Inches(0.34)
    add_textbox(
        slide, inner_x, inner_y, inner_w, title_h,
        title, font_size=12, bold=True, color=accent, align=PP_ALIGN.LEFT,
    )

    # 구분선 (얇은 라인)
    sep_y = inner_y + title_h + Inches(0.02)
    add_accent_bar(
        slide, inner_x, sep_y, inner_w, Emu(9525), C_LGRAY
    )

    # items — key:value
    items_y = sep_y + Inches(0.10)
    bottom_limit = y + h - Inches(0.14)
    avail_h = bottom_limit - items_y
    row_h = avail_h // len(items)
    if row_h < Inches(0.38):
        row_h = Inches(0.38)
    key_w = Inches(1.20)
    val_gap = Inches(0.10)

    cy = items_y
    for key, val in items:
        # 작은 bullet dot
        dot_size = Inches(0.10)
        dot_y = cy + (row_h - dot_size) // 2
        dot = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, inner_x, dot_y, dot_size, dot_size
        )
        dot.fill.solid()
        dot.fill.fore_color.rgb = accent
        dot.line.fill.background()

        # key
        kx = inner_x + dot_size + Inches(0.08)
        add_textbox(
            slide, kx, cy, key_w, row_h,
            key, font_size=11, bold=True, color=C_NAVY, align=PP_ALIGN.LEFT,
        )
        # value
        vx = kx + key_w + val_gap
        vw = inner_w - (vx - inner_x)
        add_textbox(
            slide, vx, cy, vw, row_h,
            val, font_size=10.5, color=C_INK, align=PP_ALIGN.LEFT,
        )
        cy += row_h


def _build_recommendation_card(slide, x, y, w, h, title, accent, bg, items):
    """③ 업계 권고 사항 — 번호 배지 + 라벨 + 보조 note."""
    # 배경
    bg_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = bg
    bg_shape.line.color.rgb = accent
    bg_shape.line.width = Pt(0.5)
    add_shadow(bg_shape, blur_pt=4, dist_pt=2, opacity_pct=14)

    # 좌측 accent bar
    bar_w = Inches(0.08)
    add_accent_bar(slide, x, y, bar_w, h, accent)

    inner_x = x + bar_w + Inches(0.14)
    inner_w = w - bar_w - Inches(0.28)
    inner_y = y + Inches(0.14)

    # 타이틀
    title_h = Inches(0.34)
    add_textbox(
        slide, inner_x, inner_y, inner_w, title_h,
        title, font_size=12, bold=True, color=accent, align=PP_ALIGN.LEFT,
    )
    # 구분선
    sep_y = inner_y + title_h + Inches(0.02)
    add_accent_bar(
        slide, inner_x, sep_y, inner_w, Emu(9525), C_LGRAY
    )

    items_y = sep_y + Inches(0.10)
    bottom_limit = y + h - Inches(0.14)
    avail_h = bottom_limit - items_y
    row_h = avail_h // len(items)
    if row_h < Inches(0.42):
        row_h = Inches(0.42)

    # 각 행: 번호 원 + (라벨 / note)
    badge = Inches(0.30)
    label_gap = Inches(0.10)

    cy = items_y
    for num, label, note in items:
        # 번호 원
        b_y = cy + Inches(0.04)
        make_icon_circle(
            slide, inner_x, b_y, badge, accent,
            text=num, font_size=11, font_color=C_WHITE,
        )

        # 텍스트 영역
        tx = inner_x + badge + label_gap
        tw = inner_w - (tx - inner_x)
        tb = add_textbox(
            slide, tx, cy, tw, row_h,
            "", font_size=11,
        )
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.02)
        tf.margin_right = Inches(0.02)
        tf.margin_top = Inches(0.02)
        tf.margin_bottom = Inches(0.02)
        p1 = tf.paragraphs[0]
        p1.alignment = PP_ALIGN.LEFT
        for _r in list(p1.runs):
            _r.text = ""
        r_lbl = p1.add_run()
        r_lbl.text = label
        r_lbl.font.size = Pt(11)
        r_lbl.font.bold = True
        r_lbl.font.color.rgb = C_INK

        add_para(
            tf, note,
            font_size=9, color=C_GRAY, bold=False,
            align=PP_ALIGN.LEFT,
            space_before=Pt(2), space_after=Pt(0),
        )

        cy += row_h
