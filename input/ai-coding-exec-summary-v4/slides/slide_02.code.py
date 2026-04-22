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
C_GREEN = RGBColor(0x2E, 0x8B, 0x57)
C_BLUE  = RGBColor(0x4F, 0x81, 0xBD)
C_RED   = RGBColor(0xC0, 0x50, 0x4D)
C_INK   = RGBColor(0x1A, 0x1F, 0x2E)
C_GRAY  = RGBColor(0x66, 0x66, 0x66)
C_LGRAY = RGBColor(0xE0, 0xE4, 0xEA)
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)


def build_slide_2(slide):
    set_title(slide, "핵심 요약")
    clear_placeholders(slide, keep=[0])

    # ================================================================
    # 레이아웃 계획 (CONTENT_SAFE: left=0.28, top=0.68, w=10.28, h=6.34)
    # - Key Message Bar:  top 0.72, height 0.50
    # - Card Grid (1x3):  top 1.34, height 3.30
    # - Takeaway Bar:     top 4.78, height 2.20
    # ================================================================

    safe_left  = CONTENT_SAFE.left
    safe_top   = CONTENT_SAFE.top
    safe_w     = CONTENT_SAFE.width
    safe_right = CONTENT_SAFE.right

    # ---------------- 1) Key Message Bar ----------------
    kmb_x = safe_left
    kmb_y = safe_top + Inches(0.04)
    kmb_w = safe_w
    kmb_h = Inches(0.50)

    # 좌측 액센트 바
    accent_bar_w = Inches(0.10)
    add_accent_bar(slide, kmb_x, kmb_y, accent_bar_w, kmb_h, C_NAVY)

    # 배경 박스 (연한 회색)
    kmb_bg_x = kmb_x + accent_bar_w
    kmb_bg_w = kmb_w - accent_bar_w
    kmb_bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, kmb_bg_x, kmb_y, kmb_bg_w, kmb_h
    )
    kmb_bg.fill.solid()
    kmb_bg.fill.fore_color.rgb = C_LGRAY
    kmb_bg.line.fill.background()

    # Key Message 텍스트 (박스 안)
    kmb_text = add_textbox(
        slide,
        kmb_bg_x + Inches(0.15), kmb_y,
        kmb_bg_w - Inches(0.3), kmb_h,
        "Key Message  |  경쟁사는 이미 도입 단계 · 영역별 효과 차이 크므로 선별 적용 필수",
        font_size=13, bold=True, color=C_INK, align=PP_ALIGN.LEFT,
    )
    set_body_anchor(kmb_text, 'ctr')

    # ---------------- 2) 3 Message Cards (1x3) ----------------
    cards_area_top = kmb_y + kmb_h + Inches(0.20)
    cards_area_h   = Inches(3.30)
    cards_area = (safe_left, cards_area_top, safe_w, cards_area_h)

    grid = calc_grid(1, 3, area=cards_area, gap=Inches(0.20))

    card_defs = [
        {
            "accent": C_GREEN,
            "icon": "✓",
            "title": "경쟁사 이미 도입 중",
            "lines": [
                ("Mercedes",   "5,000명+ Copilot 전사 사용"),
                ("BMW",        "사내 파일럿 · SPACE 5축 개선"),
                ("현대모비스",  "Mobis Development Studio"),
                ("현대오토에버", "H-Chat 그룹사 전사"),
            ],
        },
        {
            "accent": C_BLUE,
            "icon": "△",
            "title": "영역별 효과 차이 (업계 일반)",
            "lines": [
                ("선행 R&D · 도구",   "+20~35%"),
                ("양산 일반 SW",     "+10~20%"),
                ("양산 안전 코드",    "±5% (본전)"),
                ("→ 선별 적용이 핵심", ""),
            ],
        },
        {
            "accent": C_RED,
            "icon": "⚠",
            "title": "품질검증 부담 증가",
            "lines": [
                ("코드 리뷰 시간",       "+91% (Faros 2025)"),
                ("코드 리뷰 요청",       "+98% (Faros 2025)"),
                ("→ 인력·도구 동시 보강 필수", ""),
                ("(가장 자주 누락되는 투자)", ""),
            ],
        },
    ]

    for idx, cell in enumerate(grid[0]):
        _draw_card(slide, cell, card_defs[idx])

    # ---------------- 3) Takeaway Bar (하단) ----------------
    tb_x = safe_left
    tb_y = cards_area_top + cards_area_h + Inches(0.15)
    tb_w = safe_w
    tb_bottom = CONTENT_SAFE.bottom - Inches(0.05)
    tb_h = tb_bottom - tb_y

    # 배경 박스 (네이비)
    tb_bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, tb_x, tb_y, tb_w, tb_h)
    tb_bg.fill.solid()
    tb_bg.fill.fore_color.rgb = C_NAVY
    tb_bg.line.fill.background()

    # 제목
    tb_title_h = Inches(0.40)
    tb_title = add_textbox(
        slide,
        tb_x + Inches(0.25), tb_y + Inches(0.08),
        tb_w - Inches(0.5), tb_title_h,
        "▶ 업계 주요 시사점",
        font_size=13, bold=True, color=C_WHITE, align=PP_ALIGN.LEFT,
    )
    set_body_anchor(tb_title, 'ctr')

    # 3개 bullet — 1줄씩, 얇은 흰 구분선 없이 텍스트만
    bullets = [
        "도구: GitHub Copilot · Azure OpenAI · Claude · 사내 LLM 프록시 (H-Chat) 4 가지 패턴",
        "측정: SPACE 5축 · 자체 설문 · Git 로그 분석 등 다각도 조합",
        "2027.08 EU AI Act 자동차 완전 적용 D-Day → 준비 시점",
    ]

    bullet_area_y = tb_y + tb_title_h + Inches(0.10)
    bullet_area_h = tb_bottom - bullet_area_y - Inches(0.08)
    per_bullet_h  = bullet_area_h / len(bullets)

    for i, text in enumerate(bullets):
        by = bullet_area_y + int(per_bullet_h * i)
        bb = add_textbox(
            slide,
            tb_x + Inches(0.30), by,
            tb_w - Inches(0.60), int(per_bullet_h),
            "• " + text,
            font_size=11, color=C_WHITE, align=PP_ALIGN.LEFT,
        )
        set_body_anchor(bb, 'ctr')


def _draw_card(slide, cell, card_def):
    """단일 메시지 카드: 상단 액센트 바 + 아이콘 + 제목 + label:value 리스트."""
    cx, cy, cw, ch = cell.left, cell.top, cell.width, cell.height

    # 상단 accent 바
    top_bar_h = Inches(0.08)
    add_accent_bar(slide, cx, cy, cw, top_bar_h, card_def["accent"])

    # 카드 배경
    body_y = cy + top_bar_h
    body_h = ch - top_bar_h
    body = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx, body_y, cw, body_h)
    body.fill.solid()
    body.fill.fore_color.rgb = C_WHITE
    body.line.color.rgb = C_LGRAY
    body.line.width = Pt(0.75)
    add_shadow(body, blur_pt=3, dist_pt=2, opacity_pct=20)

    # 아이콘 배지 (좌상단)
    icon_size = Inches(0.40)
    icon_x = cx + Inches(0.15)
    icon_y = body_y + Inches(0.18)
    make_icon_circle(
        slide, icon_x, icon_y, icon_size,
        card_def["accent"], text=card_def["icon"],
        font_size=14, font_color=C_WHITE,
    )

    # 카드 제목 (아이콘 우측)
    title_x = icon_x + icon_size + Inches(0.10)
    title_w = cw - (title_x - cx) - Inches(0.15)
    title_tb = add_textbox(
        slide, title_x, icon_y,
        title_w, icon_size,
        card_def["title"],
        font_size=12, bold=True, color=card_def["accent"],
        align=PP_ALIGN.LEFT,
    )
    set_body_anchor(title_tb, 'ctr')

    # label:value 리스트
    list_x = cx + Inches(0.18)
    list_y = icon_y + icon_size + Inches(0.20)
    list_w = cw - Inches(0.36)
    list_bottom = body_y + body_h - Inches(0.10)
    list_h_total = list_bottom - list_y

    lines = card_def["lines"]
    per_h = list_h_total / len(lines)

    for i, (label, value) in enumerate(lines):
        row_y = list_y + int(per_h * i)
        row_tb = slide.shapes.add_textbox(list_x, row_y, list_w, int(per_h))
        tf = row_tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT

        # label
        run_label = p.add_run()
        run_label.text = label
        run_label.font.size = Pt(10.5)
        run_label.font.bold = True
        run_label.font.color.rgb = C_INK

        if value:
            # 구분자
            run_sep = p.add_run()
            run_sep.text = "  "
            run_sep.font.size = Pt(10.5)

            # value (accent 색)
            run_value = p.add_run()
            run_value.text = value
            run_value.font.size = Pt(10.5)
            run_value.font.color.rgb = card_def["accent"]
            run_value.font.bold = False

        set_body_anchor(row_tb, 'ctr')
        set_text_inset(row_tb, left=Inches(0.02), right=Inches(0.02),
                       top=Inches(0.02), bottom=Inches(0.02))
