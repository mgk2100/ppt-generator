"""Slide 2 — 왜 지금 검토가 필요한가 · 산업 동향 및 경쟁사 현황.

구성:
1) 상단: Key Message Bar (EU AI Act D-Day + 주요 OEM 도입 시점)
2) 중단: CHEVRON 5단계 타임라인 (2023.07 Mercedes → 2024 BMW → 2025.09 현대모비스
   → 2026.08 EU 일반 → 2027.08 자동차 완전 적용)
3) 하단: 2x3 경쟁사 카드 그리드 (Mercedes/BMW/Bosch/현대모비스/현대오토에버/Geely)
"""

from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from ppt_utils import (
    set_title,
    clear_placeholders,
    add_textbox,
    add_para,
    add_rich_text,
    add_accent_bar,
    make_icon_badge,
    add_shadow,
    set_body_anchor,
    set_text_inset,
    calc_grid,
)
from template_contract import CONTENT_SAFE, LAYOUT_CONTENT


# ============================ 색상 팔레트 ============================
C_NAVY   = RGBColor(0x1F, 0x49, 0x7D)
C_BLUE   = RGBColor(0x4F, 0x81, 0xBD)
C_GREEN  = RGBColor(0x2E, 0x8B, 0x57)
C_YELLOW = RGBColor(0xD4, 0xA0, 0x17)
C_RED    = RGBColor(0xC0, 0x50, 0x4D)
C_TEAL   = RGBColor(0x2C, 0x7F, 0x94)
C_INK    = RGBColor(0x1A, 0x1F, 0x2E)
C_GRAY   = RGBColor(0x66, 0x66, 0x66)
C_LGRAY  = RGBColor(0xE0, 0xE4, 0xEA)
C_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
C_OFFWHITE = RGBColor(0xF7, 0xF9, 0xFC)


# ============================ 헬퍼 ============================

def _hex_to_rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _set_center_anchor(shape):
    """텍스트 프레임을 세로 가운데 정렬."""
    set_body_anchor(shape, "ctr")


def _add_card_container(slide, x, y, w, h, top_bar_color):
    """경쟁사 카드 — 상단 컬러 바 + 흰 본체(연회색 테두리)."""
    # 본체 (흰 배경, 연한 테두리)
    body = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    body.adjustments[0] = 0.06
    body.fill.solid()
    body.fill.fore_color.rgb = C_WHITE
    body.line.color.rgb = C_LGRAY
    body.line.width = Pt(0.75)
    add_shadow(body, blur_pt=5, dist_pt=2, opacity_pct=22, color=C_GRAY)

    # 상단 컬러 바 (카드 폭의 전체)
    bar_h = Inches(0.08)
    bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, x + Inches(0.02), y + Inches(0.02),
        w - Inches(0.04), bar_h,
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = top_bar_color
    bar.line.fill.background()

    return body


def _add_chevron_step(slide, x, y, w, h, period, label, fill_color, is_last=False):
    """CHEVRON 한 단계: 상단 period(볼드) + 하단 label."""
    shape = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    add_shadow(shape, blur_pt=4, dist_pt=2, opacity_pct=28, color=C_GRAY)

    tf = shape.text_frame
    tf.word_wrap = True
    set_text_inset(shape, left=Inches(0.12), top=Inches(0.06),
                   right=Inches(0.22), bottom=Inches(0.06))
    _set_center_anchor(shape)

    # 첫 번째 단락: period (볼드 10pt)
    p0 = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.CENTER
    run0 = p0.add_run()
    run0.text = period
    run0.font.size = Pt(10.5)
    run0.font.bold = True
    run0.font.color.rgb = C_WHITE

    # 두 번째 단락: label (9pt)
    add_para(
        tf, label,
        font_size=9, color=C_WHITE, bold=False,
        align=PP_ALIGN.CENTER, space_before=Pt(1),
    )
    return shape


# ============================ 메인 빌더 ============================

def build_slide_3(slide):
    set_title(slide, "왜 지금 검토가 필요한가  ·  산업 동향 및 경쟁사 현황")
    clear_placeholders(slide, keep=[0])

    # CONTENT_SAFE 좌표 기준
    safe_left = CONTENT_SAFE.left
    safe_top = CONTENT_SAFE.top
    safe_w = CONTENT_SAFE.width
    safe_bottom = CONTENT_SAFE.bottom

    # ------------------------------------------------------------------
    # 1) Key Message Bar  (상단)
    # ------------------------------------------------------------------
    km_x = safe_left
    km_y = safe_top + Inches(0.02)           # 0.70"
    km_w = safe_w                             # 10.28"
    km_h = Inches(0.48)

    # 좌측 accent 세로 바
    accent_bar_w = Inches(0.08)
    add_accent_bar(slide, km_x, km_y, accent_bar_w, km_h, C_BLUE)

    # 본체 박스 (연한 배경)
    km_body_x = km_x + accent_bar_w
    km_body_w = km_w - accent_bar_w
    km_body = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, km_body_x, km_y, km_body_w, km_h,
    )
    km_body.fill.solid()
    km_body.fill.fore_color.rgb = C_OFFWHITE
    km_body.line.color.rgb = C_LGRAY
    km_body.line.width = Pt(0.5)

    # 아이콘 배지 (KEY MESSAGE 라벨)
    km_badge_w = Inches(1.10)
    km_badge_h = Inches(0.32)
    km_badge_x = km_body_x + Inches(0.10)
    km_badge_y = km_y + (km_h - km_badge_h) // 2
    make_icon_badge(
        slide, km_badge_x, km_badge_y, km_badge_w, km_badge_h,
        "KEY MESSAGE", C_NAVY, font_size=9, font_color=C_WHITE,
        corner_radius=0.25,
    )

    # 메시지 텍스트 (배지 오른쪽)
    km_text_x = km_badge_x + km_badge_w + Inches(0.15)
    km_text_w = km_body_x + km_body_w - km_text_x - Inches(0.10)
    km_text = slide.shapes.add_textbox(
        km_text_x, km_y, km_text_w, km_h,
    )
    km_tf = km_text.text_frame
    km_tf.word_wrap = True
    set_text_inset(km_text, left=Inches(0.02), top=Inches(0.04),
                   right=Inches(0.02), bottom=Inches(0.04))
    _set_center_anchor(km_text)
    # 첫 단락 삭제 대신 paragraphs[0] 에 직접 rich text
    p_km = km_tf.paragraphs[0]
    p_km.alignment = PP_ALIGN.LEFT
    r1 = p_km.add_run()
    r1.text = "주요 OEM은 "
    r1.font.size = Pt(11); r1.font.color.rgb = C_INK
    r2 = p_km.add_run()
    r2.text = "2023~2025년 이미 도입 완료"
    r2.font.size = Pt(11); r2.font.bold = True; r2.font.color.rgb = C_GREEN
    r3 = p_km.add_run()
    r3.text = "  ·  "
    r3.font.size = Pt(11); r3.font.color.rgb = C_GRAY
    r4 = p_km.add_run()
    r4.text = "2027.08 EU AI Act 자동차 D-Day"
    r4.font.size = Pt(11); r4.font.bold = True; r4.font.color.rgb = C_RED

    # ------------------------------------------------------------------
    # 2) CHEVRON 5단계 타임라인 (중단)
    # ------------------------------------------------------------------
    tl_section_y = km_y + km_h + Inches(0.14)    # 1.32"

    # 섹션 소제목
    tl_title_h = Inches(0.28)
    tl_title = add_textbox(
        slide, safe_left, tl_section_y, Inches(4.5), tl_title_h,
        "① 산업 흐름 타임라인",
        font_size=11, color=C_NAVY, bold=True,
    )

    # 우측 보조 캡션 (Today 표시)
    cap_w = Inches(3.4)
    cap_x = safe_left + safe_w - cap_w
    cap = add_textbox(
        slide, cap_x, tl_section_y, cap_w, tl_title_h,
        "긍정(2023~2025)  ·  임박(2026)  ·  D-Day(2027)",
        font_size=9, color=C_GRAY, bold=False, align=PP_ALIGN.RIGHT,
    )

    # CHEVRON 배치 영역
    ch_y = tl_section_y + tl_title_h + Inches(0.04)   # 1.64"
    ch_h = Inches(0.70)
    ch_area_x = safe_left
    ch_area_w = safe_w
    overlap = Inches(0.09)

    steps = [
        ("2023.07", "Mercedes 도입",      C_GREEN),
        ("2024",    "BMW 파일럿",         C_GREEN),
        ("2025.09", "현대모비스",         C_GREEN),
        ("2026.08", "EU AI Act 일반",     C_YELLOW),
        ("2027.08", "자동차 완전 적용",   C_RED),
    ]
    n_steps = len(steps)
    # 5개 chevron 배치: 총 폭 = n*w - (n-1)*overlap  → w = (total + (n-1)*overlap) / n
    total_w = ch_area_w
    step_w = (total_w + (n_steps - 1) * overlap) // n_steps

    for i, (period, label, color) in enumerate(steps):
        cx = ch_area_x + i * (step_w - overlap)
        _add_chevron_step(
            slide, cx, ch_y, step_w, ch_h,
            period, label, color,
            is_last=(i == n_steps - 1),
        )

    # ------------------------------------------------------------------
    # 3) 2x3 경쟁사 카드 그리드 (하단)
    # ------------------------------------------------------------------
    grid_top = ch_y + ch_h + Inches(0.20)   # 2.54"
    grid_bottom = safe_bottom - Inches(0.04)  # 6.98"
    grid_h = grid_bottom - grid_top          # 4.44"
    grid_area = (safe_left, grid_top, safe_w, grid_h)

    # 섹션 소제목을 별도 행으로 두지 않고 카드 그리드로 바로 진입 (공간 최적화)
    grid = calc_grid(
        rows=2, cols=3, area=grid_area,
        gap=Inches(0.18),
    )

    cards = [
        {
            "name": "Mercedes-Benz",
            "badge": "2023.07",
            "badge_color": C_NAVY,
            "highlight": "5,000명+ 개발자",
            "body": "주당 30분+ 절감\n누적 200만 라인 수락",
            "top_bar": C_NAVY,
        },
        {
            "name": "BMW Group",
            "badge": "2024",
            "badge_color": C_NAVY,
            "highlight": "SPACE 5축 전항목 개선",
            "body": "결함 감소 보고\nAMCIS 2024 공개",
            "top_bar": C_NAVY,
        },
        {
            "name": "Bosch",
            "badge": "진행 중",
            "badge_color": C_NAVY,
            "highlight": "bosch-copilot 사내 org",
            "body": "액세스 관리\n단계적 확산",
            "top_bar": C_NAVY,
        },
        {
            "name": "현대모비스",
            "badge": "2025.09",
            "badge_color": C_TEAL,
            "highlight": "Mobis Development Studio",
            "body": "Wind River 협업\nSW 중심 차량 개발환경",
            "top_bar": C_TEAL,
        },
        {
            "name": "현대오토에버",
            "badge": "2024~",
            "badge_color": C_TEAL,
            "highlight": "H-Chat 그룹사 전사",
            "body": "Azure/Gemini/Claude\n사내 LLM 프록시",
            "top_bar": C_TEAL,
        },
        {
            "name": "Geely",
            "badge": "2025.07",
            "badge_color": C_TEAL,
            "highlight": "세계 최초 AI 안전 인증",
            "body": "ISO/PAS 8800:2024\nSGS-TÜV Saar 발행",
            "top_bar": C_TEAL,
        },
    ]

    for i, card in enumerate(cards):
        r, c = divmod(i, 3)
        cell = grid[r][c]
        _build_competitor_card(
            slide,
            cell.left, cell.top, cell.width, cell.height,
            card,
        )


def _build_competitor_card(slide, x, y, w, h, card):
    """경쟁사 카드: 상단 컬러 바 + 회사명 + 우상단 배지 + 하이라이트 + 본문."""
    # 컨테이너 + 상단 바
    _add_card_container(slide, x, y, w, h, card["top_bar"])

    # 내부 여백
    pad_l = Inches(0.16)
    pad_r = Inches(0.16)
    pad_t_from_bar = Inches(0.18)   # 상단 바(0.08)+여유
    inner_top = y + Inches(0.12) + pad_t_from_bar   # 상단 바 아래

    inner_w = w - pad_l - pad_r

    # 우상단 배지 (timestamp)
    badge_w = Inches(0.82)
    badge_h = Inches(0.26)
    badge_x = x + w - pad_r - badge_w
    badge_y = inner_top
    make_icon_badge(
        slide, badge_x, badge_y, badge_w, badge_h,
        card["badge"], card["badge_color"],
        font_size=8.5, font_color=C_WHITE, corner_radius=0.3,
    )

    # 회사명 (좌측)
    name_x = x + pad_l
    name_y = inner_top - Inches(0.02)
    name_w = inner_w - badge_w - Inches(0.08)
    name_h = Inches(0.32)
    name_box = add_textbox(
        slide, name_x, name_y, name_w, name_h,
        card["name"],
        font_size=13, color=C_INK, bold=True, align=PP_ALIGN.LEFT,
    )
    set_text_inset(name_box, left=Inches(0.02), top=Inches(0.02),
                   right=Inches(0.02), bottom=Inches(0.02))

    # 하이라이트 (볼드)
    hl_x = x + pad_l
    hl_y = name_y + name_h + Inches(0.04)
    hl_w = inner_w
    hl_h = Inches(0.30)
    hl_box = add_textbox(
        slide, hl_x, hl_y, hl_w, hl_h,
        card["highlight"],
        font_size=10.5, color=card["top_bar"], bold=True, align=PP_ALIGN.LEFT,
    )
    set_text_inset(hl_box, left=Inches(0.02), top=Inches(0.02),
                   right=Inches(0.02), bottom=Inches(0.02))

    # 얇은 구분선
    sep_x = x + pad_l
    sep_y = hl_y + hl_h + Inches(0.02)
    sep_w = inner_w
    sep_h = Inches(0.015)
    add_accent_bar(slide, sep_x, sep_y, sep_w, sep_h, C_LGRAY)

    # 본문 (2줄)
    body_x = x + pad_l
    body_y = sep_y + sep_h + Inches(0.05)
    body_w = inner_w
    # 카드 바닥까지 남은 높이
    body_h = (y + h) - body_y - Inches(0.12)
    if body_h < Inches(0.4):
        body_h = Inches(0.4)

    body_box = slide.shapes.add_textbox(body_x, body_y, body_w, body_h)
    body_tf = body_box.text_frame
    body_tf.word_wrap = True
    set_text_inset(body_box, left=Inches(0.02), top=Inches(0.02),
                   right=Inches(0.02), bottom=Inches(0.02))

    lines = card["body"].split("\n")
    for i, line in enumerate(lines):
        if i == 0:
            p = body_tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            run = p.add_run()
            run.text = line
            run.font.size = Pt(9.5)
            run.font.color.rgb = C_GRAY
        else:
            add_para(
                body_tf, line,
                font_size=9.5, color=C_GRAY, align=PP_ALIGN.LEFT,
                space_before=Pt(2),
            )
