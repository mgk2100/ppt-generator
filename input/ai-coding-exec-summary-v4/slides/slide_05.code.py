from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from ppt_utils import (
    set_title, clear_placeholders,
    add_textbox, add_para, add_rich_text,
    add_accent_bar, make_icon_circle, make_icon_badge,
    add_shadow, set_body_anchor, set_text_inset,
    calc_grid,
)
from template_contract import CONTENT_SAFE


# ---------- 색상 ----------
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

    # ================================================================
    # 레이아웃 계획 (CONTENT_SAFE: left=0.28, top=0.68, w=10.28, h=6.34)
    # - Key Message Bar:   top 0.72, height 0.48
    # - Chevron Row (x3):  top ~1.30, height 1.15
    # - Bottom 2-col:      top ~2.62, height ~4.35 (bottom 6.97)
    # ================================================================

    safe_left  = CONTENT_SAFE.left
    safe_top   = CONTENT_SAFE.top
    safe_w     = CONTENT_SAFE.width
    safe_right = CONTENT_SAFE.right
    safe_bot   = CONTENT_SAFE.bottom

    # ---------------- 1) Key Message Bar ----------------
    kmb_x = safe_left
    kmb_y = safe_top + Inches(0.04)
    kmb_w = safe_w
    kmb_h = Inches(0.48)

    # 좌측 액센트 바 (teal)
    accent_bar_w = Inches(0.10)
    add_accent_bar(slide, kmb_x, kmb_y, accent_bar_w, kmb_h, C_TEAL)

    # 배경 박스 (연한 회색)
    kmb_bg_x = kmb_x + accent_bar_w
    kmb_bg_w = kmb_w - accent_bar_w
    kmb_bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, kmb_bg_x, kmb_y, kmb_bg_w, kmb_h
    )
    kmb_bg.fill.solid()
    kmb_bg.fill.fore_color.rgb = C_LGRAY
    kmb_bg.line.fill.background()

    # Key Message 텍스트
    kmb_text = add_textbox(
        slide,
        kmb_bg_x + Inches(0.15), kmb_y,
        kmb_bg_w - Inches(0.30), kmb_h,
        "Key Message  |  단계적 도입 + 자사 시범으로 실측 · 선행 영역 우선 착수 (업계 권고)",
        font_size=13, bold=True, color=C_INK, align=PP_ALIGN.LEFT,
    )
    set_body_anchor(kmb_text, 'ctr')

    # ---------------- 2) Chevron 3단계 Row ----------------
    chev_title_y = kmb_y + kmb_h + Inches(0.12)
    chev_title_h = Inches(0.30)
    chev_title = add_textbox(
        slide,
        safe_left, chev_title_y,
        safe_w, chev_title_h,
        "① 단계적 도입 로드맵 (업계 일반 권고)",
        font_size=12, bold=True, color=C_INK, align=PP_ALIGN.LEFT,
    )
    set_body_anchor(chev_title, 'ctr')

    chev_row_y = chev_title_y + chev_title_h + Inches(0.04)
    chev_row_h = Inches(1.10)

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

    # Chevron 3개 가로 배치 (약간 겹치게 해서 흐름을 표현)
    n_phases = len(phases)
    chev_gap = Inches(-0.05)   # 약간 겹침 — 화살표 느낌 강화
    total_gap = int(chev_gap) * (n_phases - 1)
    chev_w_each = (int(safe_w) - total_gap) // n_phases

    for i, ph in enumerate(phases):
        cx = int(safe_left) + i * (chev_w_each + int(chev_gap))
        cy = int(chev_row_y)
        cw = chev_w_each
        chv = slide.shapes.add_shape(
            MSO_SHAPE.CHEVRON, cx, cy, cw, int(chev_row_h)
        )
        chv.fill.solid()
        chv.fill.fore_color.rgb = ph["color"]
        chv.line.fill.background()

        # 텍스트: 제목(bold, 큰) + 본문(작은)
        tf = chv.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.alignment = PP_ALIGN.CENTER
        run = p1.add_run()
        run.text = ph["title"]
        run.font.size = Pt(13)
        run.font.bold = True
        run.font.color.rgb = C_WHITE

        # 본문 (두 줄)
        body_lines = ph["body"].split("\n")
        for line in body_lines:
            p = tf.add_paragraph()
            p.alignment = PP_ALIGN.CENTER
            p.space_before = Pt(2)
            r = p.add_run()
            r.text = line
            r.font.size = Pt(10)
            r.font.color.rgb = C_WHITE

        set_body_anchor(chv, 'ctr')
        set_text_inset(
            chv,
            left=Inches(0.18), right=Inches(0.32),
            top=Inches(0.08), bottom=Inches(0.08),
        )

    # ---------------- 3) 하단 2열 ----------------
    bot_y = chev_row_y + chev_row_h + Inches(0.22)
    bot_h = safe_bot - bot_y - Inches(0.02)

    grid = calc_grid(
        1, 2,
        area=(safe_left, bot_y, safe_w, bot_h),
        gap=Inches(0.20),
    )

    left_cell = grid[0][0]
    right_cell = grid[0][1]

    _draw_info_card(slide, left_cell)
    _draw_recommendation_card(slide, right_cell)


def _draw_info_card(slide, cell):
    """좌측: ② 자사 시범 사업 — info_card (BG_BLUE, C_BLUE accent, key:value 5)."""
    cx, cy, cw, ch = cell.left, cell.top, cell.width, cell.height

    # 상단 accent 바
    top_bar_h = Inches(0.08)
    add_accent_bar(slide, cx, cy, cw, top_bar_h, C_BLUE)

    # 카드 배경 (연 파랑)
    body_y = cy + top_bar_h
    body_h = ch - top_bar_h
    body = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx, body_y, cw, body_h)
    body.fill.solid()
    body.fill.fore_color.rgb = BG_BLUE
    body.line.color.rgb = C_LGRAY
    body.line.width = Pt(0.5)
    add_shadow(body, blur_pt=3, dist_pt=2, opacity_pct=18)

    # 카드 제목
    title_h = Inches(0.38)
    title_tb = add_textbox(
        slide,
        cx + Inches(0.20), body_y + Inches(0.08),
        cw - Inches(0.40), title_h,
        "② 자사 시범 사업 (실측 프로토콜)",
        font_size=12, bold=True, color=C_BLUE, align=PP_ALIGN.LEFT,
    )
    set_body_anchor(title_tb, 'ctr')

    # 제목 아래 얇은 구분선
    div_y = body_y + Inches(0.08) + title_h + Inches(0.02)
    add_accent_bar(
        slide,
        cx + Inches(0.20), div_y,
        cw - Inches(0.40), Inches(0.01),
        C_LGRAY,
    )

    # key:value 리스트
    items = [
        ("대상",        "영역별 대표 SW 모듈\n(선행 / 양산 일반 / 양산 안전 / 검증 자동화)"),
        ("기간",        "별도 산정"),
        ("규모·예산",    "별도 산정 (파일럿 계획 수립 단계)"),
        ("측정 지표",    "공수 · 결함밀도 · 코드 리뷰 부담"),
        ("목적",        "자사 환경에서 영역별 실효성 실측"),
    ]

    list_x = cx + Inches(0.22)
    list_y = div_y + Inches(0.10)
    list_w = cw - Inches(0.44)
    list_bottom = body_y + body_h - Inches(0.12)
    list_h_total = list_bottom - list_y
    per_h = list_h_total / len(items)

    key_w = Inches(1.15)

    for i, (key, value) in enumerate(items):
        row_y = list_y + int(per_h * i)
        # key 박스 (좌측, bold, BLUE)
        key_tb = add_textbox(
            slide,
            list_x, row_y,
            key_w, int(per_h),
            key,
            font_size=10, bold=True, color=C_BLUE, align=PP_ALIGN.LEFT,
        )
        set_body_anchor(key_tb, 't')
        set_text_inset(
            key_tb,
            left=Inches(0.02), right=Inches(0.02),
            top=Inches(0.04), bottom=Inches(0.02),
        )

        # value 박스 (우측)
        val_x = list_x + key_w
        val_w = list_w - key_w
        val_tb = add_textbox(
            slide,
            val_x, row_y,
            val_w, int(per_h),
            value,
            font_size=10, color=C_INK, align=PP_ALIGN.LEFT,
        )
        set_body_anchor(val_tb, 't')
        set_text_inset(
            val_tb,
            left=Inches(0.02), right=Inches(0.02),
            top=Inches(0.04), bottom=Inches(0.02),
        )


def _draw_recommendation_card(slide, cell):
    """우측: ③ 업계 권고 사항 — recommendation_card (BG_TEAL, 4 번호 항목)."""
    cx, cy, cw, ch = cell.left, cell.top, cell.width, cell.height

    # 상단 accent 바
    top_bar_h = Inches(0.08)
    add_accent_bar(slide, cx, cy, cw, top_bar_h, C_TEAL)

    # 카드 배경 (연 teal)
    body_y = cy + top_bar_h
    body_h = ch - top_bar_h
    body = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx, body_y, cw, body_h)
    body.fill.solid()
    body.fill.fore_color.rgb = BG_TEAL
    body.line.color.rgb = C_LGRAY
    body.line.width = Pt(0.5)
    add_shadow(body, blur_pt=3, dist_pt=2, opacity_pct=18)

    # 카드 제목
    title_h = Inches(0.38)
    title_tb = add_textbox(
        slide,
        cx + Inches(0.20), body_y + Inches(0.08),
        cw - Inches(0.40), title_h,
        "③ 업계 권고 사항",
        font_size=12, bold=True, color=C_TEAL, align=PP_ALIGN.LEFT,
    )
    set_body_anchor(title_tb, 'ctr')

    # 제목 아래 얇은 구분선
    div_y = body_y + Inches(0.08) + title_h + Inches(0.02)
    add_accent_bar(
        slide,
        cx + Inches(0.20), div_y,
        cw - Inches(0.40), Inches(0.01),
        C_LGRAY,
    )

    # 4 권고 항목 (번호원 + label bold + note)
    items = [
        {
            "num": "1",
            "label": "선행 · 테스트 자동화 영역 먼저 검토",
            "note": "경쟁사 모두 이 영역부터 착수",
        },
        {
            "num": "2",
            "label": "자사 시범 사업으로 실효성 실측",
            "note": "양산 안전 영역은 공개 데이터 부재",
        },
        {
            "num": "3",
            "label": "품질검증 인력·도구 동시 보강",
            "note": "AI 코드량 증가는 리뷰 부담 증가로 직결",
        },
        {
            "num": "4",
            "label": "사내 AI 사용 가이드라인 정비",
            "note": "EU AI Act 2027.08 자동차 완전 적용",
        },
    ]

    list_x = cx + Inches(0.22)
    list_y = div_y + Inches(0.10)
    list_w = cw - Inches(0.44)
    list_bottom = body_y + body_h - Inches(0.12)
    list_h_total = list_bottom - list_y
    per_h = list_h_total / len(items)

    circle_size = Inches(0.32)

    for i, it in enumerate(items):
        row_y = list_y + int(per_h * i)

        # 번호 원
        circle_x = list_x
        circle_y = row_y + Inches(0.04)
        make_icon_circle(
            slide,
            circle_x, circle_y,
            circle_size,
            C_TEAL,
            text=it["num"],
            font_size=11,
            font_color=C_WHITE,
        )

        # 텍스트 영역 (원 오른쪽)
        text_x = circle_x + circle_size + Inches(0.10)
        text_w = list_w - (circle_size + Inches(0.10))

        text_tb = slide.shapes.add_textbox(
            text_x, row_y, text_w, int(per_h)
        )
        tf = text_tb.text_frame
        tf.word_wrap = True

        # 1행: label (bold, ink)
        p1 = tf.paragraphs[0]
        p1.alignment = PP_ALIGN.LEFT
        p1.space_after = Pt(2)
        r1 = p1.add_run()
        r1.text = it["label"]
        r1.font.size = Pt(10.5)
        r1.font.bold = True
        r1.font.color.rgb = C_INK

        # 2행: note (small, gray)
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.LEFT
        r2 = p2.add_run()
        r2.text = it["note"]
        r2.font.size = Pt(9)
        r2.font.color.rgb = C_GRAY

        set_body_anchor(text_tb, 't')
        set_text_inset(
            text_tb,
            left=Inches(0.02), right=Inches(0.02),
            top=Inches(0.02), bottom=Inches(0.02),
        )
