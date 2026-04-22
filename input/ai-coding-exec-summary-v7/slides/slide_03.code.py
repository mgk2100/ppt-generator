"""Slide 3 — 경쟁사 상세 · 도구 · 규모 · 측정

2x3 competitor detail grid with tool/scale/measurement/result/source rows.
Navy vs Teal badges distinguish global (Mercedes/BMW/Bosch) vs Asian (Mobis/Autoever/Geely).
Bottom 4-line glossary footnote.
"""

from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from ppt_utils import (
    set_title, clear_placeholders,
    add_textbox, add_para, add_rich_text,
    add_accent_bar, make_icon_badge,
    set_body_anchor, set_text_inset,
    calc_grid,
)
from template_contract import CONTENT_SAFE, LAYOUT_CONTENT


# Palette
C_NAVY = RGBColor(0x1F, 0x49, 0x7D)
C_BLUE = RGBColor(0x4F, 0x81, 0xBD)
C_TEAL = RGBColor(0x2C, 0x7F, 0x94)
C_INK = RGBColor(0x1A, 0x1F, 0x2E)
C_GRAY = RGBColor(0x66, 0x66, 0x66)
C_LGRAY = RGBColor(0xE0, 0xE4, 0xEA)
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
C_SUBTLE_BG = RGBColor(0xF6, 0xF8, 0xFB)
C_LABEL = RGBColor(0x88, 0x90, 0x9B)


def _hex_to_rgb(hex_str):
    """Convert '#RRGGBB' to RGBColor."""
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _draw_card(slide, cell, card, accent):
    """Render a single competitor card within cell rect.

    Card structure (5 rows):
      Row 0: Company name (bold) + badge (top-right)
      Row 1: 도구: <tool>
      Row 2: 규모: <scale>
      Row 3: 측정: <measurement>
      Row 4: 결과: <result>   (bold value, accent color)
      Row 5: 출처: <source>   (gray, italic)
    """
    x, y, w, h = cell.left, cell.top, cell.width, cell.height

    # --- Card background (very light, accent left bar) ---
    bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    bg.fill.solid()
    bg.fill.fore_color.rgb = C_SUBTLE_BG
    bg.line.color.rgb = C_LGRAY
    bg.line.width = Pt(0.5)
    bg.adjustments[0] = 0.04
    bg.text_frame.word_wrap = True
    # Clear auto-created paragraph (we don't use this tf for text)
    bg.text_frame.text = ""

    # Left accent bar (3 pt wide)
    bar_w = Inches(0.05)
    add_accent_bar(slide, x, y, bar_w, h, accent)

    # --- Header: company name (left) + badge (right) ---
    header_y = y + Inches(0.08)
    header_h = Inches(0.34)

    # Badge dimensions — right-aligned
    badge_text = card["badge"]
    # approximate badge width from text length
    badge_chars = len(badge_text)
    badge_w = Inches(0.60) + Inches(0.055) * max(0, badge_chars - 5)
    if badge_w > Inches(1.05):
        badge_w = Inches(1.05)
    badge_h = Inches(0.26)
    badge_x = x + w - badge_w - Inches(0.10)
    badge_y = header_y + (header_h - badge_h) // 2

    # Company name textbox (between left-bar and badge)
    name_x = x + int(bar_w) + Inches(0.10)
    name_w = badge_x - name_x - Inches(0.05)
    name_tb = add_textbox(
        slide, name_x, header_y, name_w, header_h,
        card["name"], font_size=12, bold=True, color=C_INK,
    )
    set_text_inset(name_tb, left=Inches(0.02), right=Inches(0.02),
                   top=Inches(0.02), bottom=Inches(0.02))
    set_body_anchor(name_tb, "ctr")

    # Badge
    badge_color = _hex_to_rgb(card["badge_color"])
    badge = make_icon_badge(
        slide, badge_x, badge_y, badge_w, badge_h,
        badge_text, badge_color,
        font_size=9, font_color=C_WHITE, corner_radius=0.25,
    )
    set_text_inset(badge, left=Inches(0.04), right=Inches(0.04),
                   top=Inches(0.01), bottom=Inches(0.01))

    # --- Thin divider under header ---
    div_y = y + Inches(0.46)
    div = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        x + int(bar_w) + Inches(0.10), div_y,
        w - int(bar_w) - Inches(0.20), Emu(6350),  # ~0.5pt tall
    )
    div.fill.solid()
    div.fill.fore_color.rgb = C_LGRAY
    div.line.fill.background()

    # --- Body rows (single textbox, rich paragraphs) ---
    body_x = x + int(bar_w) + Inches(0.10)
    body_y = y + Inches(0.52)
    body_w = w - int(bar_w) - Inches(0.20)
    body_h = y + h - body_y - Inches(0.08)

    tb = slide.shapes.add_textbox(body_x, body_y, body_w, body_h)
    tf = tb.text_frame
    tf.word_wrap = True
    set_text_inset(tb, left=Inches(0.02), right=Inches(0.02),
                   top=Inches(0.02), bottom=Inches(0.02))

    # Remove initial empty paragraph's default run
    tf.paragraphs[0]._p.clear()
    # But keep its pPr; we'll set first via add_rich_text manually.

    def _row(label, value, value_bold=False, value_color=None, first=False,
             value_italic=False, extra_space=0):
        # Build a paragraph with label + value styles
        segments = [
            {"text": f"{label}:  ", "font_size": 9, "color": C_LABEL, "bold": False},
            {"text": value, "font_size": 9,
             "color": value_color if value_color else C_INK,
             "bold": value_bold, "italic": value_italic},
        ]
        # Always use add_rich_text (adds a new paragraph). If first, remove
        # the default paragraph's empty state by leaving pristine.
        p = add_rich_text(
            tf, segments,
            align=PP_ALIGN.LEFT,
            space_before=Pt(0 if first else 2 + extra_space),
            space_after=Pt(0),
            line_spacing=Pt(11.5),
        )
        return p

    _row("도구", card["tool"], first=True)
    _row("규모", card["scale"])
    _row("측정", card["measurement"])
    _row("결과", card["result"], value_bold=True, value_color=accent)
    _row("출처", card["source"], value_color=C_GRAY, value_italic=True)


def build_slide_3(slide):
    set_title(slide, "경쟁사 상세  ·  도구 · 규모 · 측정")
    clear_placeholders(slide, keep=[0])

    # ================================================================
    # Region definitions
    # ================================================================
    # Body: y=0.68" ~ 6.50"   (h = 5.82")
    # Footnote: y=6.55" ~ 7.00" (h = 0.45")
    body_top = Inches(0.68)
    body_bottom = Inches(6.50)
    body_left = CONTENT_SAFE.left
    body_right = CONTENT_SAFE.right
    body_w = body_right - body_left

    foot_top = Inches(6.55)
    foot_bottom = Inches(7.00)

    # ================================================================
    # 1. Key Message Bar (top of body)
    # ================================================================
    km_h = Inches(0.38)
    km_y = body_top
    km_bar = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        body_left, km_y, body_w, km_h,
    )
    km_bar.fill.solid()
    km_bar.fill.fore_color.rgb = C_NAVY
    km_bar.line.fill.background()
    km_bar.adjustments[0] = 0.18

    # Left accent (brighter blue strip)
    add_accent_bar(slide, body_left, km_y, Inches(0.07), km_h, C_BLUE)

    # Key message text inside bar
    km_tb = add_textbox(
        slide,
        body_left + Inches(0.18), km_y,
        body_w - Inches(0.22), km_h,
        "주요 OEM·Tier-1 6사 도입 현황 (공식 공표 기준)   ·   2027.08 EU AI Act 자동차 D-Day",
        font_size=11, bold=True, color=C_WHITE, align=PP_ALIGN.LEFT,
    )
    set_body_anchor(km_tb, "ctr")
    set_text_inset(km_tb, left=Inches(0.10), right=Inches(0.10),
                   top=Inches(0.04), bottom=Inches(0.04))

    # ================================================================
    # 2. 2x3 Card Grid
    # ================================================================
    grid_top = km_y + km_h + Inches(0.15)
    grid_bottom = body_bottom
    grid_h = grid_bottom - grid_top
    grid_area = (body_left, grid_top, body_w, grid_h)

    grid = calc_grid(
        rows=2, cols=3,
        area=grid_area,
        gap=Inches(0.18),
    )

    cards = [
        {"name": "Mercedes-Benz",   "badge": "2023.07~",   "badge_color": "#1F497D",
         "tool": "GitHub Copilot",
         "scale": "5,000명+ · 115k repos",
         "measurement": "개발자 자체 설문 + 흐름 상태 보고",
         "result": "주당 30분+ 절감 · 누적 200만 라인 수락",
         "source": "GitHub 공식 case study"},
        {"name": "BMW Group",       "badge": "2024",       "badge_color": "#1F497D",
         "tool": "GitHub Copilot (사내 파일럿)",
         "scale": "사내 파일럿",
         "measurement": "SPACE 프레임워크 5축",
         "result": "5축 전항목 개선 · 결함 감소",
         "source": "AMCIS 2024 논문"},
        {"name": "Bosch",           "badge": "진행 중",      "badge_color": "#1F497D",
         "tool": "GitHub Copilot (bosch-copilot org)",
         "scale": "사내 org 운영",
         "measurement": "비공개",
         "result": "단계적 확산 중",
         "source": "GitHub 공식 org 페이지"},
        {"name": "현대모비스",       "badge": "2025.09",    "badge_color": "#2C7F94",
         "tool": "Mobis Development Studio",
         "scale": "Wind River 협업 · SDV 개발환경",
         "measurement": "CI/CD/CT 자동화 지표",
         "result": "차세대 개발시스템 확장",
         "source": "Wind River 보도자료"},
        {"name": "현대오토에버",     "badge": "2024~",       "badge_color": "#2C7F94",
         "tool": "H-Chat (Azure/Gemini/Claude 프록시)",
         "scale": "그룹사 전사 배포",
         "measurement": "비공개",
         "result": "코드 보조·문서 작성 지원",
         "source": "현대오토에버 공식 페이지"},
        {"name": "Geely Auto",      "badge": "2025.07",    "badge_color": "#2C7F94",
         "tool": "AI 안전 프로세스 (코딩 도구 아님)",
         "scale": "차량 기능 AI 전반",
         "measurement": "ISO/PAS 8800 인증 심사",
         "result": "세계 최초 AI 안전 인증",
         "source": "SGS 보도자료"},
    ]

    for i, cell in enumerate(grid.flat):
        card = cards[i]
        accent = _hex_to_rgb(card["badge_color"])
        _draw_card(slide, cell, card, accent)

    # ================================================================
    # 3. Footnote (4 lines, 9pt gray, y=6.55~7.00)
    # ================================================================
    # SPACE 는 S2에서 먼저 해설 — 이 슬라이드에서 중복 각주 생략
    footnote_entries = [
        "AMCIS 2024: Americas Conference on Information Systems (미국정보시스템학회) — BMW 측정 논문 공개",
        "SDV (Software Defined Vehicle): SW 업데이트로 차량 기능을 확장·변경하는 차세대 차량 아키텍처",
        "ISO/PAS 8800:2024: AI 시스템의 차량 안전 표준 (2024.12 신규 발행). Geely 세계 최초 인증",
    ]

    foot_h = foot_bottom - foot_top
    foot_w = CONTENT_SAFE.right - CONTENT_SAFE.left
    foot_tb = slide.shapes.add_textbox(
        CONTENT_SAFE.left, foot_top, foot_w, foot_h,
    )
    foot_tf = foot_tb.text_frame
    foot_tf.word_wrap = True
    set_text_inset(foot_tb, left=Inches(0.04), right=Inches(0.04),
                   top=Inches(0.02), bottom=Inches(0.02))
    # Clear default paragraph
    foot_tf.paragraphs[0]._p.clear()

    for i, entry in enumerate(footnote_entries):
        segments = [
            {"text": "※  ", "font_size": 8, "color": C_BLUE, "bold": True},
            {"text": entry, "font_size": 8, "color": C_GRAY, "bold": False},
        ]
        add_rich_text(
            foot_tf, segments,
            align=PP_ALIGN.LEFT,
            space_before=Pt(0 if i == 0 else 0.5),
            space_after=Pt(0),
            line_spacing=Pt(10),
        )
