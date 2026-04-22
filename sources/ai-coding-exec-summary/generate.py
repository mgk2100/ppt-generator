#!/usr/bin/env python3
"""AI 코딩 어시스턴트 도입 영향 분석 — 임원용 4페이지 요약본.

대상: 팀장 > 실장 > 센터장
스코프: 자동차 OEM · Tier-1 벤더사
구성: 4장 (표지 없음). 각 슬라이드 = 완결된 메시지 1개.
"""

import sys
from pathlib import Path

sys.path.insert(0, '/home/ubuntu/Share/ppt-generator')

from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from ppt_utils import (
    ensure_fonts, load_template, get_layout, clear_placeholders,
    set_title, add_textbox, add_para, add_rich_text,
    add_accent_bar, make_icon_circle, make_icon_badge,
    add_shadow, set_shape_opacity, add_styled_table,
    calc_grid, set_body_anchor, set_text_inset,
    CONTENT_SAFE,
)

OUTPUT = Path("/home/ubuntu/Share/ppt-generator/output/ai-coding-exec-summary.pptx")
LAYOUT_NAME = "제목 및 내용 (페이지 번호 삭제)"

# ============================ 색상 팔레트 ============================
C_NAVY   = RGBColor(0x1F, 0x49, 0x7D)
C_BLUE   = RGBColor(0x4F, 0x81, 0xBD)
C_GREEN  = RGBColor(0x2E, 0x8B, 0x57)
C_YELLOW = RGBColor(0xD4, 0xA0, 0x17)
C_RED    = RGBColor(0xC0, 0x50, 0x4D)
C_TEAL   = RGBColor(0x2C, 0x7F, 0x94)
C_GRAY   = RGBColor(0x66, 0x66, 0x66)
C_LGRAY  = RGBColor(0xE0, 0xE4, 0xEA)
C_DGRAY  = RGBColor(0x33, 0x33, 0x33)
C_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
C_INK    = RGBColor(0x1A, 0x1F, 0x2E)

BG_GREEN  = RGBColor(0xE8, 0xF5, 0xEC)
BG_YELLOW = RGBColor(0xFF, 0xF7, 0xD9)
BG_RED    = RGBColor(0xFC, 0xE8, 0xE6)
BG_BLUE   = RGBColor(0xE7, 0xF0, 0xFA)
BG_TEAL   = RGBColor(0xE3, 0xF2, 0xF7)
BG_NAVY   = RGBColor(0xEE, 0xF1, 0xF7)


# ============================ 공통 헬퍼 ============================

def add_key_message_bar(slide, message, accent_color):
    """상단 Key Message Bar: 좌측 세로 바 + 연한 배경 + bold 메시지."""
    y = CONTENT_SAFE.top + Inches(0.02)
    h = Inches(0.50)
    x = CONTENT_SAFE.left
    w = CONTENT_SAFE.width

    bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    bg.fill.solid()
    bg.fill.fore_color.rgb = accent_color
    set_shape_opacity(bg, 10)
    bg.line.fill.background()

    add_accent_bar(slide, x, y, Inches(0.06), h, accent_color)

    tb = add_textbox(slide, x + Inches(0.22), y + Inches(0.04),
                     w - Inches(0.32), h - Inches(0.08),
                     message, font_size=13, bold=True,
                     color=accent_color, align=PP_ALIGN.LEFT)
    set_body_anchor(tb, 'ctr')
    return y + h  # 다음 요소의 top 기준점 반환


def add_card(slide, x, y, w, h, accent_color, bg_color=None, bar_h=None):
    """일반 카드: 상단 accent 바 + 배경 + 엷은 테두리 + 그림자."""
    if bg_color is None:
        bg_color = C_WHITE
    if bar_h is None:
        bar_h = Inches(0.06)

    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = C_LGRAY
    card.line.width = Pt(0.5)
    add_shadow(card, blur_pt=3, dist_pt=1)

    add_accent_bar(slide, x, y, w, bar_h, accent_color)
    return card


def add_chevron_step(slide, x, y, w, h, text_top, text_bottom,
                     fill_color, text_color=None):
    """CHEVRON 1단계 도형 + 상/하 텍스트."""
    if text_color is None:
        text_color = C_WHITE

    ch = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, x, y, w, h)
    ch.fill.solid()
    ch.fill.fore_color.rgb = fill_color
    ch.line.fill.background()

    tf = ch.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.08)
    tf.margin_right = Inches(0.20)
    tf.margin_top = Inches(0.06)
    tf.margin_bottom = Inches(0.06)

    p1 = tf.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    r1 = p1.add_run()
    r1.text = str(text_top)
    r1.font.size = Pt(10)
    r1.font.bold = True
    r1.font.color.rgb = text_color

    if text_bottom:
        add_para(tf, str(text_bottom), font_size=9, color=text_color,
                 align=PP_ALIGN.CENTER, space_before=Pt(2))

    set_body_anchor(ch, 'ctr')
    return ch


# ============================ Slide 1: 핵심 요약 ============================

def build_slide_1(prs):
    layout = get_layout(prs, LAYOUT_NAME)
    slide = prs.slides.add_slide(layout)
    clear_placeholders(slide, keep=[0])
    set_title(slide, "AI 코딩 어시스턴트 도입 영향 분석  ·  핵심 요약")

    km_bottom = add_key_message_bar(slide,
        "영역별 차등 적용 시 6개월 내 투자 회수 · 100명 기준 연 15억 원 절감 가능",
        C_NAVY)

    # --- 3대 메시지 카드 ---
    card_top = km_bottom + Inches(0.25)
    card_h = Inches(2.35)
    grid = calc_grid(1, 3,
                     area=(CONTENT_SAFE.left, card_top, CONTENT_SAFE.width, card_h),
                     gap=Inches(0.18))

    messages = [
        {
            "accent": C_GREEN, "icon": "✓",
            "title": "경쟁사 이미 도입 중",
            "lines": [
                ("Mercedes", " 5,000명+ 개발자"),
                ("BMW", " 사내 파일럿 (SPACE 개선)"),
                ("Bosch · 현대모비스 · 현대오토에버", ""),
                ("2027.08", " EU AI Act D-Day"),
            ],
        },
        {
            "accent": C_BLUE, "icon": "△",
            "title": "영역별 효과 3~5배 차이",
            "lines": [
                ("선행 R&D · 도구", "  +20~35%"),
                ("양산 일반 SW", "  +10~20%"),
                ("양산 안전 코드", "  ±5% (본전)"),
                ("→ 선별 적용이 핵심", ""),
            ],
        },
        {
            "accent": C_RED, "icon": "⚠",
            "title": "품질검증 부담 증가",
            "lines": [
                ("코드 리뷰 시간", "  +91%"),
                ("코드 리뷰 요청", "  +98%"),
                ("→ 인력·도구 동시 보강 필수", ""),
                ("(가장 자주 누락되는 투자)", ""),
            ],
        },
    ]

    for i, msg in enumerate(messages):
        cell = grid[0][i]
        add_card(slide, cell.left, cell.top, cell.width, cell.height,
                 msg["accent"], bg_color=C_WHITE)

        # 아이콘
        icon_size = Inches(0.52)
        icon_x = cell.left + Inches(0.22)
        icon_y = cell.top + Inches(0.24)
        make_icon_circle(slide, icon_x, icon_y, icon_size, msg["accent"],
                         text=msg["icon"], font_size=18, font_color=C_WHITE)

        # 카드 제목
        title_x = icon_x + icon_size + Inches(0.15)
        title_y = cell.top + Inches(0.30)
        title_w = cell.width - (title_x - cell.left) - Inches(0.20)
        add_textbox(slide, title_x, title_y, title_w, Inches(0.40),
                    msg["title"], font_size=14, bold=True,
                    color=msg["accent"])

        # 본문 (rich text로 label+value 대비)
        body_x = cell.left + Inches(0.28)
        body_y = cell.top + Inches(1.00)
        body_w = cell.width - Inches(0.56)
        body_h = cell.height - Inches(1.10)
        body_tb = slide.shapes.add_textbox(body_x, body_y, body_w, body_h)
        body_tf = body_tb.text_frame
        body_tf.word_wrap = True

        for idx, (label, value) in enumerate(msg["lines"]):
            if idx == 0:
                p = body_tf.paragraphs[0]
            else:
                p = body_tf.add_paragraph()
                p.space_before = Pt(4)
            p.alignment = PP_ALIGN.LEFT
            r1 = p.add_run()
            r1.text = str(label)
            r1.font.size = Pt(11)
            r1.font.bold = True
            r1.font.color.rgb = C_INK
            if value:
                r2 = p.add_run()
                r2.text = str(value)
                r2.font.size = Pt(11)
                r2.font.color.rgb = msg["accent"]
                r2.font.bold = True

    # --- 승인 요청 박스 ---
    box_top = card_top + card_h + Inches(0.20)
    box_h = CONTENT_SAFE.bottom - box_top - Inches(0.02)
    box_x = CONTENT_SAFE.left
    box_w = CONTENT_SAFE.width

    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 box_x, box_top, box_w, box_h)
    box.fill.solid()
    box.fill.fore_color.rgb = C_NAVY
    set_shape_opacity(box, 7)
    box.line.color.rgb = C_NAVY
    box.line.width = Pt(1.0)

    add_textbox(slide, box_x + Inches(0.30), box_top + Inches(0.10),
                box_w - Inches(0.60), Inches(0.32),
                "▶  즉시 결정 요청 사항",
                font_size=13, bold=True, color=C_NAVY)

    items = [
        "□  선행 · 테스트 자동화 영역  즉시 도입 승인",
        "□  4개월 시범 사업 예산 승인  (약 3천만원)",
        "□  품질검증 인력 보강 계획  착수",
        "□  사내 AI 사용 가이드라인  제정",
    ]
    item_grid_top = box_top + Inches(0.48)
    item_grid_h = box_h - Inches(0.52)
    item_grid = calc_grid(2, 2,
        area=(box_x + Inches(0.30), item_grid_top,
              box_w - Inches(0.60), item_grid_h),
        gap=Inches(0.12))
    for idx, item in enumerate(items):
        cell = item_grid[idx // 2][idx % 2]
        tb = add_textbox(slide, cell.left, cell.top, cell.width, cell.height,
                         item, font_size=12, bold=True, color=C_INK)
        set_body_anchor(tb, 'ctr')


# ============================ Slide 2: 왜 지금인가 ============================

def build_slide_2(prs):
    layout = get_layout(prs, LAYOUT_NAME)
    slide = prs.slides.add_slide(layout)
    clear_placeholders(slide, keep=[0])
    set_title(slide, "왜 지금 검토가 필요한가  ·  산업 동향 및 경쟁사 현황")

    km_bottom = add_key_message_bar(slide,
        "주요 OEM은 2023~2025년 이미 도입 완료 · 2027.08 EU AI Act 자동차 D-Day",
        C_BLUE)

    # --- 산업 타임라인 (CHEVRON 5단계) ---
    tl_top = km_bottom + Inches(0.25)
    tl_h = Inches(0.70)
    tl_x = CONTENT_SAFE.left
    tl_w = CONTENT_SAFE.width

    # 섹션 제목
    add_textbox(slide, tl_x, tl_top, Inches(5.0), Inches(0.28),
                "① 산업 흐름 타임라인",
                font_size=11, bold=True, color=C_BLUE)

    chev_top = tl_top + Inches(0.32)
    chev_h = tl_h - Inches(0.32)
    steps = [
        ("2023.07", "Mercedes 도입", C_GREEN),
        ("2024", "BMW 파일럿", C_GREEN),
        ("2025.09", "현대모비스", C_GREEN),
        ("2026.08", "EU AI Act 일반", C_YELLOW),
        ("2027.08", "자동차 완전 적용", C_RED),
    ]
    n_steps = len(steps)
    overlap = Inches(0.08)
    avail_w = tl_w + overlap * (n_steps - 1)
    step_w = int(avail_w / n_steps)
    cx = tl_x
    for period, label, color in steps:
        add_chevron_step(slide, cx, chev_top, step_w, chev_h,
                         period, label, color, text_color=C_WHITE)
        cx += step_w - int(overlap)

    # --- 경쟁사 도입 현황 (2x3 카드) ---
    cg_top = tl_top + tl_h + Inches(0.30)
    cg_bottom = CONTENT_SAFE.bottom - Inches(0.02)
    cg_h = cg_bottom - cg_top

    add_textbox(slide, tl_x, cg_top, Inches(5.0), Inches(0.28),
                "② 경쟁사 도입 현황",
                font_size=11, bold=True, color=C_BLUE)

    grid_top = cg_top + Inches(0.32)
    grid_h = cg_h - Inches(0.32)
    grid = calc_grid(2, 3,
        area=(CONTENT_SAFE.left, grid_top, CONTENT_SAFE.width, grid_h),
        gap=Inches(0.15))

    competitors = [
        {
            "accent": C_NAVY, "name": "Mercedes-Benz",
            "badge": "2023.07",
            "highlight": "5,000명+ 개발자",
            "body": "주당 30분+ 절감\n누적 200만 라인 수락",
        },
        {
            "accent": C_NAVY, "name": "BMW Group",
            "badge": "2024",
            "highlight": "SPACE 5축 전항목 개선",
            "body": "결함 감소 보고\nAMCIS 2024 공개",
        },
        {
            "accent": C_NAVY, "name": "Bosch",
            "badge": "진행 중",
            "highlight": "bosch-copilot 사내 org",
            "body": "액세스 관리\n단계적 확산",
        },
        {
            "accent": C_TEAL, "name": "현대모비스",
            "badge": "2025.09",
            "highlight": "Mobis Development Studio",
            "body": "Wind River 협업\nSW 중심 차량 개발환경",
        },
        {
            "accent": C_TEAL, "name": "현대오토에버",
            "badge": "2024~",
            "highlight": "H-Chat 그룹사 전사",
            "body": "Azure/Gemini/Claude\n사내 LLM 프록시",
        },
        {
            "accent": C_TEAL, "name": "Geely",
            "badge": "2025.07",
            "highlight": "세계 최초 AI 안전 인증",
            "body": "ISO/PAS 8800:2024\nSGS-TÜV Saar 발행",
        },
    ]

    for idx, co in enumerate(competitors):
        cell = grid[idx // 3][idx % 3]
        add_card(slide, cell.left, cell.top, cell.width, cell.height,
                 co["accent"])

        # 회사명 + 시점 배지
        name_y = cell.top + Inches(0.18)
        add_textbox(slide, cell.left + Inches(0.22), name_y,
                    cell.width - Inches(1.25), Inches(0.30),
                    co["name"], font_size=13, bold=True, color=C_INK)

        badge_w = Inches(0.95)
        badge_x = cell.left + cell.width - badge_w - Inches(0.18)
        make_icon_badge(slide, badge_x, name_y + Inches(0.02),
                        badge_w, Inches(0.26),
                        co["badge"], co["accent"], font_size=9,
                        corner_radius=0.40)

        # 하이라이트
        add_textbox(slide, cell.left + Inches(0.22),
                    cell.top + Inches(0.56),
                    cell.width - Inches(0.40), Inches(0.30),
                    co["highlight"], font_size=11, bold=True, color=co["accent"])

        # 본문
        body_tb = add_textbox(slide, cell.left + Inches(0.22),
                              cell.top + Inches(0.88),
                              cell.width - Inches(0.40),
                              cell.height - Inches(1.00),
                              co["body"].split("\n")[0],
                              font_size=10, color=C_GRAY)
        for line in co["body"].split("\n")[1:]:
            add_para(body_tb.text_frame, line, font_size=10, color=C_GRAY,
                     space_before=Pt(2))


# ============================ Slide 3: 예상 효과 ============================

def build_slide_3(prs):
    layout = get_layout(prs, LAYOUT_NAME)
    slide = prs.slides.add_slide(layout)
    clear_placeholders(slide, keep=[0])
    set_title(slide, "영역별 예상 효과 및 투자 회수  ·  100명 기준")

    km_bottom = add_key_message_bar(slide,
        "영역별 효과 3~5배 차이 · 선별 적용이 핵심 · 투자 회수 1~2개월",
        C_NAVY)

    # --- 신호등 효과 표 (수동 테이블) ---
    tbl_top = km_bottom + Inches(0.25)
    tbl_x = CONTENT_SAFE.left
    tbl_w = CONTENT_SAFE.width

    add_textbox(slide, tbl_x, tbl_top, Inches(5.0), Inches(0.28),
                "① 영역별 효과 (신호등 평가)",
                font_size=11, bold=True, color=C_NAVY)

    # 수동 카드 테이블 (색상 제어 용이)
    row_top = tbl_top + Inches(0.32)
    row_h = Inches(0.58)
    col_areas = [
        ("영역",     tbl_x,                           Inches(4.30)),
        ("효과",     tbl_x + Inches(4.30) + Inches(0.10), Inches(1.80)),
        ("해석",     tbl_x + Inches(4.30) + Inches(1.80) + Inches(0.20),
                     tbl_w - Inches(4.30) - Inches(1.80) - Inches(0.30)),
    ]

    # 헤더
    for title, cx, cw in col_areas:
        hdr = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx, row_top, cw, Inches(0.32))
        hdr.fill.solid()
        hdr.fill.fore_color.rgb = C_NAVY
        hdr.line.fill.background()
        tb = add_textbox(slide, cx + Inches(0.10), row_top + Inches(0.03),
                         cw - Inches(0.20), Inches(0.26),
                         title, font_size=10, bold=True, color=C_WHITE,
                         align=PP_ALIGN.CENTER)
        set_body_anchor(tb, 'ctr')

    rows = [
        {"dot_color": C_GREEN, "area": "선행 R&D · 도구 · 테스트 자동화",
         "effect": "+20 ~ +35%", "effect_color": C_GREEN,
         "bg": BG_GREEN,
         "reading": "즉시 도입 권장 (위험 낮음 · 효과 큼)"},
        {"dot_color": C_GREEN, "area": "양산 일반 SW (비안전)",
         "effect": "+10 ~ +20%", "effect_color": C_GREEN,
         "bg": BG_GREEN,
         "reading": "코드 검사 도구 통합 후 도입"},
        {"dot_color": C_YELLOW, "area": "양산 안전 코드 (차량 안전등급 높음↑)",
         "effect": "-5 ~ +10%", "effect_color": C_YELLOW,
         "bg": BG_YELLOW,
         "reading": "본전 영역 · 시범 측정 필요"},
        {"dot_color": C_RED, "area": "품질검증 단계 (인력 관점)",
         "effect": "-15 ~ -30%", "effect_color": C_RED,
         "bg": BG_RED,
         "reading": "코드 리뷰 부담 증가 · 인력 보강 필수"},
    ]

    cur_y = row_top + Inches(0.32)
    for r in rows:
        # 배경 박스 (전체 행)
        bg_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                          tbl_x, cur_y, tbl_w, row_h)
        bg_shape.fill.solid()
        bg_shape.fill.fore_color.rgb = r["bg"]
        bg_shape.line.color.rgb = C_LGRAY
        bg_shape.line.width = Pt(0.5)

        # 신호등 컬러 닷 (OVAL)
        dot_size = Inches(0.22)
        dot_x = col_areas[0][1] + Inches(0.20)
        dot_y = cur_y + (row_h - dot_size) // 2
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                     dot_x, dot_y, dot_size, dot_size)
        dot.fill.solid()
        dot.fill.fore_color.rgb = r["dot_color"]
        dot.line.color.rgb = r["dot_color"]
        dot.line.width = Pt(0.5)

        # 영역 열 (이름만)
        area_tb = add_textbox(slide,
            dot_x + dot_size + Inches(0.15), cur_y,
            col_areas[0][2] - (dot_x + dot_size + Inches(0.15) - col_areas[0][1]) - Inches(0.10),
            row_h,
            r["area"], font_size=12, bold=True, color=C_INK)
        set_body_anchor(area_tb, 'ctr')

        # 효과 열
        eff_tb = add_textbox(slide,
            col_areas[1][1], cur_y,
            col_areas[1][2], row_h,
            r["effect"], font_size=14, bold=True, color=r["effect_color"],
            align=PP_ALIGN.CENTER)
        set_body_anchor(eff_tb, 'ctr')

        # 해석 열
        rd_tb = add_textbox(slide,
            col_areas[2][1] + Inches(0.10), cur_y,
            col_areas[2][2] - Inches(0.15), row_h,
            r["reading"], font_size=11, color=C_DGRAY)
        set_body_anchor(rd_tb, 'ctr')

        cur_y += row_h + Inches(0.04)

    # --- ROI 큰 숫자 3개 카드 ---
    roi_top = cur_y + Inches(0.20)
    roi_bottom = CONTENT_SAFE.bottom - Inches(0.02)
    roi_h = roi_bottom - roi_top  # 남은 하단 공간 전체 활용

    add_textbox(slide, tbl_x, roi_top, Inches(6.0), Inches(0.28),
                "② 투자 회수 효과 (100명 · 1년 기준)",
                font_size=11, bold=True, color=C_NAVY)

    roi_card_top = roi_top + Inches(0.32)
    roi_h_card = roi_h - Inches(0.32)  # 충분한 높이 확보
    roi_grid = calc_grid(1, 3,
        area=(tbl_x, roi_card_top, tbl_w, roi_h_card),
        gap=Inches(0.18))

    roi_cards = [
        {"label": "연 절감 효과", "value": "약 15억 원",
         "sub": "21,000시간 × 시간당 7.2만", "accent": C_GREEN},
        {"label": "초기 투자",   "value": "약 1.2억 원",
         "sub": "도구 3천 + 운영 9천만원", "accent": C_BLUE},
        {"label": "손익분기점",  "value": "1 ~ 2개월",
         "sub": "보수 보정 시 2~3개월", "accent": C_NAVY},
    ]

    # 3단 레이아웃: 상(label) / 중(value 큰 숫자) / 하(sub)
    label_h = Inches(0.30)
    sub_h = Inches(0.28)
    value_margin = Inches(0.08)
    value_y_offset = label_h + value_margin
    value_h = roi_h_card - label_h - sub_h - value_margin * 2

    for i, rc in enumerate(roi_cards):
        cell = roi_grid[0][i]
        add_card(slide, cell.left, cell.top, cell.width, cell.height,
                 rc["accent"], bg_color=C_WHITE, bar_h=Inches(0.05))
        # 상단 라벨
        add_textbox(slide, cell.left + Inches(0.15),
                    cell.top + Inches(0.12),
                    cell.width - Inches(0.30), label_h,
                    rc["label"], font_size=11, bold=True, color=C_GRAY)
        # 중앙 큰 값
        vtb = add_textbox(slide, cell.left + Inches(0.15),
                    cell.top + value_y_offset,
                    cell.width - Inches(0.30), value_h,
                    rc["value"], font_size=28, bold=True, color=rc["accent"],
                    align=PP_ALIGN.CENTER)
        set_body_anchor(vtb, 'ctr')
        # 하단 부연
        add_textbox(slide, cell.left + Inches(0.15),
                    cell.top + cell.height - sub_h - Inches(0.06),
                    cell.width - Inches(0.30), sub_h,
                    rc["sub"], font_size=9, color=C_GRAY,
                    align=PP_ALIGN.CENTER)


# ============================ Slide 4: 실행 계획 ============================

def build_slide_4(prs):
    layout = get_layout(prs, LAYOUT_NAME)
    slide = prs.slides.add_slide(layout)
    clear_placeholders(slide, keep=[0])
    set_title(slide, "실행 로드맵 및 승인 요청 사항")

    km_bottom = add_key_message_bar(slide,
        "3단계 도입 + 4개월 시범 사업으로 리스크 최소화 · 선행 영역 즉시 착수",
        C_TEAL)

    # --- 3단계 로드맵 (CHEVRON) ---
    rm_top = km_bottom + Inches(0.25)
    rm_x = CONTENT_SAFE.left
    rm_w = CONTENT_SAFE.width

    add_textbox(slide, rm_x, rm_top, Inches(5.0), Inches(0.28),
                "① 3단계 도입 로드맵",
                font_size=11, bold=True, color=C_TEAL)

    chev_top = rm_top + Inches(0.32)
    chev_h = Inches(0.90)

    phases = [
        ("Phase 1  ·  0~3개월",
         "도구 · 테스트 자동화 파일럿\n위험 낮음 · 효과 큼", C_GREEN),
        ("Phase 2  ·  3~9개월",
         "선행 R&D · 응용SW 확대\n사내 AI 학습 · 자동 검사 통합", C_BLUE),
        ("Phase 3  ·  9~18개월",
         "양산 비안전 영역 단계 적용\n양산 안전은 시범 측정 후", C_NAVY),
    ]
    n = len(phases)
    overlap = Inches(0.10)
    avail_w = rm_w + overlap * (n - 1)
    step_w = int(avail_w / n)
    cx = rm_x
    for title_t, body_t, color in phases:
        add_chevron_step(slide, cx, chev_top, step_w, chev_h,
                         title_t, body_t, color, text_color=C_WHITE)
        cx += step_w - int(overlap)

    # --- 하단 2열: 시범 사업 | 승인 요청 ---
    bot_top = chev_top + chev_h + Inches(0.30)
    bot_bottom = CONTENT_SAFE.bottom - Inches(0.02)
    bot_h = bot_bottom - bot_top

    bot_grid = calc_grid(1, 2,
        area=(rm_x, bot_top, rm_w, bot_h),
        gap=Inches(0.22))

    # --- [좌] 4개월 시범 사업 ---
    left_cell = bot_grid[0][0]
    add_card(slide, left_cell.left, left_cell.top,
             left_cell.width, left_cell.height, C_BLUE,
             bg_color=BG_BLUE)

    add_textbox(slide, left_cell.left + Inches(0.25),
                left_cell.top + Inches(0.16),
                left_cell.width - Inches(0.50), Inches(0.32),
                "② 4개월 시범 사업 제안",
                font_size=13, bold=True, color=C_BLUE)

    pilot_items = [
        ("대상",   "SW 모듈 4개  (선행 1 · 양산 일반 1 · 양산 안전 1 · 검증 자동화 1)"),
        ("투입",   "20명 × 4개월  (숙련도·경력 균형 배정)"),
        ("예산",   "약 3천만원  (연간 절감가치의 2%)"),
        ("측정",   "공수 · 결함밀도 · 코드 리뷰 부담"),
        ("목적",   "양산 안전 영역 효과 실증  (가장 불확실한 영역)"),
    ]
    items_y = left_cell.top + Inches(0.62)
    items_h = left_cell.height - Inches(0.75)
    items_tb = slide.shapes.add_textbox(
        left_cell.left + Inches(0.25), items_y,
        left_cell.width - Inches(0.50), items_h)
    items_tf = items_tb.text_frame
    items_tf.word_wrap = True

    for idx, (k, v) in enumerate(pilot_items):
        if idx == 0:
            p = items_tf.paragraphs[0]
        else:
            p = items_tf.add_paragraph()
            p.space_before = Pt(6)
        p.alignment = PP_ALIGN.LEFT
        r1 = p.add_run()
        r1.text = f"{k}  "
        r1.font.size = Pt(11)
        r1.font.bold = True
        r1.font.color.rgb = C_BLUE
        r2 = p.add_run()
        r2.text = v
        r2.font.size = Pt(11)
        r2.font.color.rgb = C_INK

    # --- [우] 승인 요청 사항 ---
    right_cell = bot_grid[0][1]
    add_card(slide, right_cell.left, right_cell.top,
             right_cell.width, right_cell.height, C_TEAL,
             bg_color=BG_TEAL)

    add_textbox(slide, right_cell.left + Inches(0.25),
                right_cell.top + Inches(0.16),
                right_cell.width - Inches(0.50), Inches(0.32),
                "③ 승인 요청 사항",
                font_size=13, bold=True, color=C_TEAL)

    approvals = [
        ("①", "선행 · 테스트 자동화 영역 즉시 도입",  "추정 ROI 11배"),
        ("②", "4개월 시범 사업 예산  3천만원 승인",    "의사결정 근거 확보"),
        ("③", "품질검증 인력 보강 계획 수립",          "도입 효과 유지"),
        ("④", "사내 AI 사용 가이드라인 제정 착수",     "EU AI Act 대응"),
    ]
    ap_y = right_cell.top + Inches(0.62)
    ap_h_total = right_cell.height - Inches(0.75)
    ap_grid = calc_grid(4, 1,
        area=(right_cell.left + Inches(0.25), ap_y,
              right_cell.width - Inches(0.50), ap_h_total),
        gap=Inches(0.06))

    for idx, (num, label, note) in enumerate(approvals):
        cell = ap_grid[idx][0]
        # 번호 배지
        make_icon_circle(slide, cell.left, cell.top + Inches(0.03),
                         Inches(0.30), C_TEAL,
                         text=num, font_size=11, font_color=C_WHITE)
        # 라벨
        tb = add_textbox(slide, cell.left + Inches(0.38),
                         cell.top,
                         cell.width - Inches(0.40), Inches(0.26),
                         label, font_size=11, bold=True, color=C_INK)
        # 주석
        add_textbox(slide, cell.left + Inches(0.38),
                    cell.top + Inches(0.24),
                    cell.width - Inches(0.40), Inches(0.22),
                    f"└  {note}", font_size=9, color=C_GRAY, bold=False)


# ============================ main ============================

def main():
    ensure_fonts()
    prs = load_template()

    build_slide_1(prs)
    build_slide_2(prs)
    build_slide_3(prs)
    build_slide_4(prs)

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(OUTPUT))
    print(f"✓ Saved: {OUTPUT}")
    print(f"  Slides: {len(prs.slides)}")


if __name__ == "__main__":
    main()
