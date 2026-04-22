"""Slide 05 — 자동차 업계 LLM/AI 활용 확장 (차량 탑재 · R&D · 자율주행)

본문 y=0.68~6.55
- Key Message Bar
- 3열 카드:
  ① 차량 탑재 LLM 에이전트 (BLUE)
  ② R&D·디자인 생성형 AI (GREEN)
  ③ 자율주행 End-to-End AI (NAVY)

각주 y=6.60~7.00 — 새로 등장한 용어만 (End-to-End 뉴럴넷 · 디지털 트윈)
"""

from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from ppt_utils import (
    set_title, clear_placeholders,
    add_textbox, add_para, add_rich_text,
    add_accent_bar, make_icon_badge,
    add_shadow,
    set_body_anchor,
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
BG_BLUE  = RGBColor(0xE7, 0xF0, 0xFA)
BG_GREEN = RGBColor(0xE6, 0xF4, 0xEC)
BG_NAVY  = RGBColor(0xE3, 0xE9, 0xF4)


def build_slide_5(slide):
    set_title(slide, "자동차 업계 LLM/AI 활용 확장  ·  차량 탑재 · R&D · 자율주행")
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

    add_accent_bar(slide, LEFT, km_y, km_bar_w, km_h, C_NAVY)
    km_bg_x = LEFT + km_bar_w
    km_bg_w = WIDTH - km_bar_w
    km_bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, km_bg_x, km_y, km_bg_w, km_h
    )
    km_bg.fill.solid()
    km_bg.fill.fore_color.rgb = BG_NAVY
    km_bg.line.fill.background()

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

    p0 = km_tf.paragraphs[0]
    p0.alignment = PP_ALIGN.LEFT
    for _r in list(p0.runs):
        _r.text = ""
    r1 = p0.add_run()
    r1.text = "코딩 어시스턴트는 자동차 AI 활용의 한 축"
    r1.font.size = Pt(13)
    r1.font.bold = True
    r1.font.color.rgb = C_NAVY
    r2 = p0.add_run()
    r2.text = "  ·  "
    r2.font.size = Pt(12)
    r2.font.color.rgb = C_GRAY
    r3 = p0.add_run()
    r3.text = "차량 탑재 LLM  ·  R&D 생성형 AI  ·  자율주행 뉴럴넷"
    r3.font.size = Pt(12)
    r3.font.bold = True
    r3.font.color.rgb = C_INK
    r4 = p0.add_run()
    r4.text = "  까지 영역 확장 중"
    r4.font.size = Pt(11)
    r4.font.color.rgb = C_GRAY

    # ---------------------------------------------------------------
    # 2) 3열 카드
    # ---------------------------------------------------------------
    cards_y = km_y + km_h + Inches(0.18)
    cards_h = BODY_BOTTOM - cards_y            # ~5.17"
    col_gap = Inches(0.18)
    col_w = (WIDTH - col_gap * 2) // 3

    categories = [
        {
            "num": "①",
            "title": "차량 탑재 LLM 에이전트",
            "subtitle": "IN-VEHICLE",
            "accent": C_BLUE,
            "bg": BG_BLUE,
            "items": [
                {
                    "headline": "Volkswagen Group",
                    "year": "2024.Q2",
                    "body": "MY2024 ID.7 / Passat / Tiguan / Golf 에 ChatGPT 기반 Cerence Chat Pro 탑재",
                },
                {
                    "headline": "Mercedes-Benz MBUX",
                    "year": "CES 2025",
                    "body": "Google Cloud 파트너십 → MB.OS Virtual Assistant (Gemini 기반) 시연",
                },
                {
                    "headline": "Stellantis × Mistral AI",
                    "year": "2024.02",
                    "body": "차량 챗봇 + 엔지니어링 Copilot 전략적 파트너십 (프랑스 AI 공동투자)",
                },
                {
                    "headline": "GM × Google Cloud",
                    "year": "2024",
                    "body": "OnStar · 인포테인먼트에 Gemini 기반 음성 에이전트 단계적 통합",
                },
            ],
        },
        {
            "num": "②",
            "title": "R&D · 디자인 생성형 AI",
            "subtitle": "DESIGN / R&D",
            "accent": C_GREEN,
            "bg": BG_GREEN,
            "items": [
                {
                    "headline": "Toyota Research Institute",
                    "year": "2023.06",
                    "body": "차량 디자인 Generative AI 공식 공개 — 엔지니어 제약 반영 스케치 자동 생성",
                },
                {
                    "headline": "BYD Xuanji Architecture",
                    "year": "CES 2024",
                    "body": "차량 전역 AI 아키텍처 공식 발표 — 차체 · 파워트레인 · ADAS 통합",
                },
                {
                    "headline": "NVIDIA Omniverse 자동차",
                    "year": "2024~",
                    "body": "Mercedes · JLR · Lucid 디지털 트윈 · 시뮬레이션 파트너십 공식화",
                },
            ],
        },
        {
            "num": "③",
            "title": "자율주행 End-to-End AI",
            "subtitle": "AUTONOMOUS",
            "accent": C_NAVY,
            "bg": BG_NAVY,
            "items": [
                {
                    "headline": "Tesla FSD v12",
                    "year": "2024.03",
                    "body": "수만 줄 C++ 제어 로직을 End-to-End 뉴럴넷 1개로 전환 (공식 릴리스)",
                },
                {
                    "headline": "Waymo EMMA",
                    "year": "2024.10",
                    "body": "Multimodal Foundation Model (arXiv 2410.23262) — 센서 · 언어 · 플래닝 통합",
                },
                {
                    "headline": "Wayve GAIA-1",
                    "year": "2023.06",
                    "body": "생성형 세계모델 (Generative World Model) 공개 — Microsoft 파트너십",
                },
            ],
        },
    ]

    cx = LEFT
    for cat in categories:
        _draw_category_card(
            slide,
            x=cx, y=cards_y, w=col_w, h=cards_h,
            cat=cat,
        )
        cx += col_w + col_gap

    # ---------------------------------------------------------------
    # 3) 각주 — 새 용어만 (SPACE · Faros AI · EU AI Act 는 선행 슬라이드에서 해설)
    # ---------------------------------------------------------------
    foot_h = FOOT_BOTTOM - FOOT_TOP
    foot_tb = add_textbox(
        slide,
        LEFT, FOOT_TOP, WIDTH, foot_h,
        "", font_size=9,
    )
    foot_tf = foot_tb.text_frame
    foot_tf.word_wrap = True
    foot_tf.margin_left = Inches(0)
    foot_tf.margin_right = Inches(0)
    foot_tf.margin_top = Inches(0)
    foot_tf.margin_bottom = Inches(0)

    foot_lines = [
        "※ End-to-End 뉴럴넷: 센서 입력 → 조향·가속 출력을 단일 신경망으로 학습. "
        "기존 Perception→Planning→Control 모듈 체인을 통합",
        "※ 디지털 트윈 (Digital Twin): 실제 공장·차량을 가상에서 1:1 복제해 "
        "시뮬레이션·학습 데이터 생성에 활용",
    ]
    for i, line in enumerate(foot_lines):
        if i == 0:
            p = foot_tf.paragraphs[0]
        else:
            p = foot_tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        p.line_spacing = 1.1
        run = p.add_run()
        run.text = line
        run.font.size = Pt(9)
        run.font.color.rgb = C_GRAY


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _draw_category_card(slide, x, y, w, h, cat):
    """카테고리 카드 — 상단 컬러 헤더 + 케이스 목록."""
    accent = cat["accent"]

    # 배경 박스 (전체 카드)
    bg_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    bg_shape.adjustments[0] = 0.04
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = C_WHITE
    bg_shape.line.color.rgb = C_LGRAY
    bg_shape.line.width = Pt(0.5)
    add_shadow(bg_shape, blur_pt=5, dist_pt=2, opacity_pct=18)

    # 상단 컬러 헤더
    header_h = Inches(0.70)
    header = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, header_h)
    header.fill.solid()
    header.fill.fore_color.rgb = accent
    header.line.fill.background()

    # 헤더 내 텍스트: 번호 원 + 제목 + 영문 서브
    num_size = Inches(0.42)
    num_x = x + Inches(0.14)
    num_y = y + (header_h - num_size) // 2
    num_shape = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, num_x, num_y, num_size, num_size
    )
    num_shape.fill.solid()
    num_shape.fill.fore_color.rgb = C_WHITE
    num_shape.line.fill.background()
    n_tf = num_shape.text_frame
    n_tf.margin_left = Inches(0)
    n_tf.margin_right = Inches(0)
    n_tf.margin_top = Inches(0)
    n_tf.margin_bottom = Inches(0)
    set_body_anchor(num_shape, "ctr")
    np0 = n_tf.paragraphs[0]
    np0.alignment = PP_ALIGN.CENTER
    nr = np0.add_run()
    nr.text = cat["num"]
    nr.font.size = Pt(16)
    nr.font.bold = True
    nr.font.color.rgb = accent

    # 제목 + 서브타이틀
    title_x = num_x + num_size + Inches(0.12)
    title_w = x + w - title_x - Inches(0.14)
    title_tb = add_textbox(
        slide, title_x, y + Inches(0.08), title_w, header_h - Inches(0.16),
        "", font_size=13,
    )
    t_tf = title_tb.text_frame
    t_tf.word_wrap = True
    t_tf.margin_left = Inches(0.02)
    t_tf.margin_right = Inches(0.02)
    t_tf.margin_top = Inches(0)
    t_tf.margin_bottom = Inches(0)

    tp0 = t_tf.paragraphs[0]
    tp0.alignment = PP_ALIGN.LEFT
    tp0.space_before = Pt(0)
    tp0.space_after = Pt(0)
    tr = tp0.add_run()
    tr.text = cat["title"]
    tr.font.size = Pt(13)
    tr.font.bold = True
    tr.font.color.rgb = C_WHITE

    add_para(
        t_tf, cat["subtitle"],
        font_size=9, color=RGBColor(0xE0, 0xE8, 0xF0), bold=False,
        align=PP_ALIGN.LEFT,
        space_before=Pt(1), space_after=Pt(0),
    )

    # ---------------------------------------------------------------
    # 아이템 목록
    # ---------------------------------------------------------------
    items = cat["items"]
    items_top = y + header_h + Inches(0.16)
    items_bottom = y + h - Inches(0.14)
    avail_h = items_bottom - items_top
    row_h = avail_h // len(items)

    inner_x = x + Inches(0.18)
    inner_w = w - Inches(0.36)

    cy = items_top
    for itm in items:
        _draw_case_item(
            slide,
            x=inner_x, y=cy, w=inner_w, h=row_h,
            accent=accent,
            headline=itm["headline"],
            year=itm["year"],
            body=itm["body"],
        )
        cy += row_h


def _draw_case_item(slide, x, y, w, h, accent, headline, year, body):
    """케이스 한 개 — 좌 accent dot + 헤드라인/연도 + 본문."""
    # 좌측 작은 dot
    dot_size = Inches(0.12)
    dot_y = y + Inches(0.06)
    dot = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, x, dot_y, dot_size, dot_size
    )
    dot.fill.solid()
    dot.fill.fore_color.rgb = accent
    dot.line.fill.background()

    # 텍스트 영역
    txt_x = x + dot_size + Inches(0.10)
    txt_w = w - (txt_x - x)
    tb = add_textbox(
        slide, txt_x, y, txt_w, h,
        "", font_size=10,
    )
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0)
    tf.margin_right = Inches(0)
    tf.margin_top = Inches(0)
    tf.margin_bottom = Inches(0)

    p1 = tf.paragraphs[0]
    p1.alignment = PP_ALIGN.LEFT
    p1.space_before = Pt(0)
    p1.space_after = Pt(0)
    for _r in list(p1.runs):
        _r.text = ""
    r_h = p1.add_run()
    r_h.text = headline
    r_h.font.size = Pt(10.5)
    r_h.font.bold = True
    r_h.font.color.rgb = C_INK
    r_y = p1.add_run()
    r_y.text = f"   {year}"
    r_y.font.size = Pt(9)
    r_y.font.bold = False
    r_y.font.color.rgb = accent

    # 본문
    add_para(
        tf, body,
        font_size=9.5, color=C_GRAY, bold=False,
        align=PP_ALIGN.LEFT,
        space_before=Pt(2), space_after=Pt(0),
    )
