#!/usr/bin/env python3
"""배터리 3사 AI 활용 현황 분석 보고서 PPT 생성기 - 템플릿 기반 고품질 버전"""

import os
import copy
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.chart.data import CategoryChartData
from pptx.oxml.ns import qn
from lxml import etree

TEMPLATE = "/home/ubuntu/Share/ppt-generator/ref/표지.pptx"
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
OUTPUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "report_v2.pptx")

# Template dimensions (EMU)
SW = 9906000   # slide width
SH = 6858000   # slide height

# Content area from Layout 1 placeholder idx=10
CL = 252001    # content left
CT = 622038    # content top
CW = 9396000   # content width
CH = 5795294   # content height
CB = CT + CH   # content bottom

# Colors
SL_BLUE      = RGBColor(0x00, 0x34, 0x78)
SL_BLUE2     = RGBColor(0x00, 0x50, 0x9E)
SL_BLUE_LT   = RGBColor(0xE0, 0xEC, 0xF8)
SL_BLUE_BG   = RGBColor(0xF5, 0xF8, 0xFC)
WHITE        = RGBColor(0xFF, 0xFF, 0xFF)
BLACK        = RGBColor(0x22, 0x22, 0x22)
GRAY_D       = RGBColor(0x55, 0x55, 0x55)
GRAY         = RGBColor(0x88, 0x88, 0x88)
GRAY_L       = RGBColor(0xAA, 0xAA, 0xAA)
GRAY_BG      = RGBColor(0xFA, 0xFA, 0xFA)
GRAY_BD      = RGBColor(0xDD, 0xDD, 0xDD)
GREEN        = RGBColor(0x27, 0xAE, 0x60)
GREEN_BG     = RGBColor(0xF0, 0xFF, 0xF4)
RED          = RGBColor(0xC0, 0x39, 0x2B)
RED_BG       = RGBColor(0xFF, 0xF5, 0xF5)
LG_RED       = RGBColor(0xA5, 0x00, 0x34)
SS_BLUE      = RGBColor(0x14, 0x28, 0xA0)
SK_YELLOW    = RGBColor(0xF7, 0xB5, 0x00)
SK_DARK      = RGBColor(0x7A, 0x63, 0x00)

FONT = "맑은 고딕"
TOTAL_SLIDES = 15


# ─── helpers ───

def _emu(inches):
    return int(inches * 914400)

def _shape(slide, shape_type, l, t, w, h):
    return slide.shapes.add_shape(shape_type, l, t, w, h)

def _rect(slide, l, t, w, h, fill, line=None, radius=None):
    st = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    s = _shape(slide, st, l, t, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line:
        s.line.color.rgb = line; s.line.width = Pt(0.75)
    else:
        s.line.fill.background()
    if radius:
        s.adjustments[0] = radius
    return s

def _line(slide, l, t, w, color=GRAY_BD, weight=Pt(0.75)):
    s = _shape(slide, MSO_SHAPE.RECTANGLE, l, t, w, weight)
    s.fill.solid(); s.fill.fore_color.rgb = color; s.line.fill.background()
    return s

def _circle(slide, l, t, size, fill):
    s = _shape(slide, MSO_SHAPE.OVAL, l, t, size, size)
    s.fill.solid(); s.fill.fore_color.rgb = fill; s.line.fill.background()
    return s

def _tf(shape, text, size, color=BLACK, bold=False, align=PP_ALIGN.LEFT, spacing=1.2):
    """Set text on an existing shape's text frame"""
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text; p.font.size = Pt(size); p.font.color.rgb = color
    p.font.bold = bold; p.font.name = FONT; p.alignment = align
    p.line_spacing = Pt(int(size * spacing))
    p.space_before = Pt(0); p.space_after = Pt(0)
    return tf

def _tb(slide, l, t, w, h, text, size, color=BLACK, bold=False, align=PP_ALIGN.LEFT, spacing=1.2):
    """Add textbox and return it"""
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text; p.font.size = Pt(size); p.font.color.rgb = color
    p.font.bold = bold; p.font.name = FONT; p.alignment = align
    p.line_spacing = Pt(int(size * spacing))
    p.space_before = Pt(0); p.space_after = Pt(0)
    return tb

def _mtb(slide, l, t, w, h, lines):
    """Multi-line textbox. lines = [(text, size, color, bold, align), ...]"""
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    for i, ld in enumerate(lines):
        text = ld[0]; sz = ld[1] if len(ld)>1 else 10; col = ld[2] if len(ld)>2 else BLACK
        bld = ld[3] if len(ld)>3 else False; al = ld[4] if len(ld)>4 else PP_ALIGN.LEFT
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = text; p.font.size = Pt(sz); p.font.color.rgb = col
        p.font.bold = bld; p.font.name = FONT; p.alignment = al
        p.line_spacing = Pt(int(sz * 1.35)); p.space_before = Pt(0); p.space_after = Pt(2)
    return tb

def _arrow(slide, l, t, w=_emu(0.22), h=_emu(0.2)):
    s = _shape(slide, MSO_SHAPE.RIGHT_ARROW, l, t, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = GRAY_L; s.line.fill.background()

def _flow_box(slide, l, t, w, h, text, fill, tc=WHITE, sz=8, border_color=None):
    """HTML-style flow step: light bg + colored border"""
    s = _shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if border_color:
        s.line.color.rgb = border_color; s.line.width = Pt(1.5)
    else:
        s.line.fill.background()
    s.adjustments[0] = 0.12
    _tf(s, text, sz, tc, True, PP_ALIGN.CENTER, 1.25)

def _bullet_list(slide, l, t, w, h, items, sz=12, color=BLACK, marker="\u25B8 "):
    """HTML info-block li style: ▸ marker in #aaa, text in body color"""
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    marker_color = RGBColor(0xAA, 0xAA, 0xAA)
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        # Marker run (▸ in #aaa)
        run_m = p.add_run()
        run_m.text = marker
        run_m.font.size = Pt(sz); run_m.font.color.rgb = marker_color; run_m.font.name = FONT
        # Text run (body color)
        run_t = p.add_run()
        run_t.text = item
        run_t.font.size = Pt(sz); run_t.font.color.rgb = color; run_t.font.name = FONT
        p.line_spacing = Pt(int(sz * 1.8))
        p.space_before = Pt(2); p.space_after = Pt(2)
    return tb


# ─── slide creation using template ───

def new_content_slide(prs, title_text):
    """Create slide using Layout 1 (제목 및 내용), set title, return slide"""
    layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(layout)
    # Set title placeholder
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0:
            ph.text = title_text
            for p in ph.text_frame.paragraphs:
                p.font.size = Pt(24); p.font.bold = True
                p.font.color.rgb = SL_BLUE; p.font.name = FONT
            break
    # Remove the content placeholder (idx=10) so we can place our own shapes
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 10:
            sp = ph._element
            sp.getparent().remove(sp)
            break
    return slide


def new_cover_slide(prs, title, subtitle, extra_lines=None):
    """Create slide using Layout 0 (제목 슬라이드)"""
    layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(layout)
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0:  # CENTER_TITLE
            ph.text = title
            for p in ph.text_frame.paragraphs:
                p.font.size = Pt(36); p.font.bold = True
                p.font.color.rgb = WHITE; p.font.name = FONT
                p.alignment = PP_ALIGN.CENTER
        elif ph.placeholder_format.idx == 1:  # SUBTITLE
            ph.text = subtitle
            for p in ph.text_frame.paragraphs:
                p.font.size = Pt(24); p.font.color.rgb = WHITE; p.font.name = FONT
                p.alignment = PP_ALIGN.CENTER
    return slide


# ─── card / table helpers ───

def make_card(slide, l, t, w, h, title, body_lines, title_color=GRAY,
              bg=GRAY_BG, border=None, title_sz=12, body_sz=12):
    """Info block card (HTML info-block style: #fafafa bg, gray uppercase title)"""
    _rect(slide, l, t, w, h, bg, line=border, radius=0.06)
    _tb(slide, l + _emu(0.18), t + _emu(0.12), w - _emu(0.36), _emu(0.25),
        title, title_sz, title_color, True)
    _bullet_list(slide, l + _emu(0.18), t + _emu(0.45), w - _emu(0.36),
                 h - _emu(0.52), body_lines, body_sz)


def make_summary_card(slide, l, t, w, h, company, product, desc, bg_color, text_color=WHITE):
    """Colored summary card with decorative circle"""
    s = _rect(slide, l, t, w, h, bg_color, radius=0.05)
    # Decorative circle (top-right, semi-transparent white)
    circle_size = _emu(0.75)
    c = _shape(slide, MSO_SHAPE.OVAL,
               l + w - circle_size + _emu(0.15), t - _emu(0.15),
               circle_size, circle_size)
    c.fill.solid()
    c.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    # Set 15% opacity via XML
    sp_elem = c._element.spPr.solidFill
    srgb = sp_elem.find(qn('a:srgbClr'))
    if srgb is not None:
        a = srgb.makeelement(qn('a:alpha'), {'val': '15000'})
        srgb.append(a)
    c.line.fill.background()
    # co-name: font-weight 600 (semi-bold), opacity 0.85
    co_color = RGBColor(0xDD, 0xDD, 0xDD) if text_color == WHITE else RGBColor(0x55, 0x55, 0x55)
    _tb(slide, l+_emu(0.18), t+_emu(0.12), w-_emu(0.36), _emu(0.2),
        company, 12, co_color, True)
    # product: font-weight 800 (extra-bold)
    _tb(slide, l+_emu(0.18), t+_emu(0.35), w-_emu(0.36), _emu(0.35),
        product, 18, text_color, True)
    # desc: font-weight normal, opacity 0.9
    desc_color = RGBColor(0xEE, 0xEE, 0xEE) if text_color == WHITE else RGBColor(0x44, 0x44, 0x44)
    _tb(slide, l+_emu(0.18), t+_emu(0.85), w-_emu(0.36), h-_emu(0.95),
        desc, 12, desc_color, spacing=1.45)


def make_table(slide, l, t, col_widths, rows, header_bg=SL_BLUE, alt_bg=SL_BLUE_BG):
    """Create a styled table"""
    n_rows = len(rows); n_cols = len(col_widths)
    total_w = sum(col_widths)
    row_h = _emu(0.5)
    tbl_h = n_rows * row_h
    tbl_shape = slide.shapes.add_table(n_rows, n_cols, l, t, total_w, tbl_h)
    tbl = tbl_shape.table

    for ci, w in enumerate(col_widths):
        tbl.columns[ci].width = w

    for ri, row_data in enumerate(rows):
        for ci, cell_text in enumerate(row_data):
            cell = tbl.cell(ri, ci)
            # Handle multiline
            lines_list = cell_text.split("\n") if "\n" in cell_text else [cell_text]
            cell.text = ""
            for li, line in enumerate(lines_list):
                p = cell.text_frame.paragraphs[0] if li == 0 else cell.text_frame.add_paragraph()
                p.text = line; p.font.name = FONT
                p.alignment = PP_ALIGN.CENTER if ci > 0 else PP_ALIGN.LEFT

                if ri == 0:
                    p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = WHITE
                else:
                    p.font.size = Pt(12)
                    p.font.color.rgb = BLACK
                    if ci == 0: p.font.bold = True
                p.line_spacing = Pt(16); p.space_before = Pt(0); p.space_after = Pt(0)

            cell.fill.solid()
            if ri == 0:
                cell.fill.fore_color.rgb = header_bg
            elif ci == 0:
                cell.fill.fore_color.rgb = GRAY_BG
            else:
                cell.fill.fore_color.rgb = WHITE if ri % 2 == 1 else alt_bg

            cell.margin_left = _emu(0.08); cell.margin_right = _emu(0.08)
            cell.margin_top = _emu(0.04); cell.margin_bottom = _emu(0.04)

    return tbl_shape


# ===================================================================
# SLIDES
# ===================================================================

def slide_01_cover(prs):
    slide = new_cover_slide(prs,
        "국내 배터리 3사의 AI 활용 현황 분석",
        "2026. 03. 31")
    # Top-right textbox: document type checkboxes
    _tb(slide, 6531701, 360000, 3097103, 272758,
        "□ 의사결정    □ 보고    ■ 정보공유", 13, WHITE,
        align=PP_ALIGN.LEFT)
    # Bottom textbox: author info
    _tb(slide, 1609310, 5985284, 6681380, 442035,
        "미래융합설계센터 알고리즘개발팀 강민규 선임", 24, BLACK,
        bold=True, align=PP_ALIGN.LEFT)


def slide_02_toc(prs):
    slide = new_content_slide(prs, "목차")
    items = [
        ("01", "핵심 요약"),
        ("02", "산업 배경"),
        ("03", "LG CNS"),
        ("04", "Samsung SDI"),
        ("05", "SK ON"),
        ("06", "비교 분석"),
        ("07", "시사점"),
        ("08", "부록"),
    ]
    y = CT + _emu(0.2)
    for num, title in items:
        c = _circle(slide, CL + _emu(0.5), y, _emu(0.45), SL_BLUE)
        _tf(c, num, 11, WHITE, True, PP_ALIGN.CENTER)
        _tb(slide, CL + _emu(1.15), y + _emu(0.05), _emu(3.5), _emu(0.35),
            title, 14, BLACK, True)
        _line(slide, CL + _emu(5.5), y + _emu(0.2), CW - _emu(6.0))
        y += _emu(0.65)


def slide_03_summary(prs):
    slide = new_content_slide(prs, "배터리 3사 AI 핵심 전략")
    # Subtitle
    _tb(slide, CL, CT + _emu(0.05), _emu(8), _emu(0.3),
        "3사의 AI 활용 방식과 핵심 성과", 18, RGBColor(0x66, 0x66, 0x66))

    # 3 summary cards
    cw = _emu(3.15); gap = _emu(0.3); cy = CT + _emu(0.55)
    cards = [
        ("LG CNS", "AI 탑재 장비", "배터리 시험 장비에 AI 탑재\n실험 설계부터 폐배터리 처리까지 자동화", LG_RED),
        ("Samsung SDI", "배터리 진단", "전 세계 1,400개 배터리 현장 AI 24시간 감시\n이상 징후 사전 감지로 사고 예방", SS_BLUE),
        ("SK ON", "AI 연구원", "AI 연구원이 배터리 설계~원가 산출 자동 수행\n설계 기간 1/3로 단축", SK_YELLOW),
    ]
    for i, (co, prod, desc, color) in enumerate(cards):
        tc = WHITE if color != SK_YELLOW else BLACK
        make_summary_card(slide, CL + i*(cw+gap), cy, cw, _emu(2.0), co, prod, desc, color, tc)

    # Why box (HTML style: gradient bg + left border)
    wy = CT + _emu(2.8)
    _rect(slide, CL, wy, CW, _emu(3.1), SL_BLUE_BG, radius=0.02)
    _rect(slide, CL, wy, _emu(0.05), _emu(3.1), SL_BLUE)
    _tb(slide, CL + _emu(0.3), wy + _emu(0.12), _emu(6), _emu(0.3),
        "왜 배터리 업계가 AI를 도입하고 있는가?", 16, SL_BLUE, True)

    # HTML icons: ⏱(9201) ⚠(9888) 📊(128202) 🌐(127760)
    reasons = [
        ("\u23F1", "개발 기간 압박", "전기차 시장의 폭발적 성장으로 신제품\n출시 주기가 크게 단축"),
        ("\u26A0", "안전성 강화", "ESS 화재 등 안전 사고 예방에 대한\n사회적 요구가 높아짐"),
        ("\U0001F4CA", "데이터 폭증", "제조 공정과 운영 현장에서\n사람이 처리할 수 없는 수준의 데이터 발생"),
        ("\U0001F310", "글로벌 경쟁", "중국 CATL은 5,000만 건 이상의\n데이터로 AI를 운영 중"),
    ]
    for idx, (icon, title, desc) in enumerate(reasons):
        col = idx % 2
        row = idx // 2
        rx = CL + _emu(0.3) + col * _emu(5.0)
        ry = wy + _emu(0.55) + row * _emu(1.25)
        # Icon box (HTML: 36x36 rounded square, #003478 bg, white icon)
        icon_s = _emu(0.36)
        ib = _rect(slide, rx, ry, icon_s, icon_s, SL_BLUE, radius=0.15)
        _tf(ib, icon, 12, WHITE, False, PP_ALIGN.CENTER)
        # Title (bold) - 14pt (177800 EMU)
        _tb(slide, rx + _emu(0.48), ry - _emu(0.02), _emu(3.5), _emu(0.22),
            title, 14, BLACK, True)
        # Description - 12pt (152400 EMU)
        _tb(slide, rx + _emu(0.48), ry + _emu(0.24), _emu(4.2), _emu(0.65),
            desc, 12, GRAY, spacing=1.4)


def slide_04_context(prs):
    slide = new_content_slide(prs, "산업 배경")
    _tb(slide, CL, CT, _emu(6), _emu(0.2),
        "인터배터리 2026 전시회 & AI 도입 배경", 18, RGBColor(0x66, 0x66, 0x66))

    # Conference box (left)
    bx = CL; by = CT + _emu(0.35)
    _rect(slide, bx, by, _emu(4.8), _emu(2.3), WHITE, GRAY_BD, 0.03)
    _tb(slide, bx+_emu(0.18), by+_emu(0.1), _emu(4), _emu(0.25),
        "인터배터리 2026 전시회", 14, SL_BLUE, True)
    _mtb(slide, bx+_emu(0.18), by+_emu(0.4), _emu(4.3), _emu(0.45), [
        ("2026년 3월 11~13일, 서울 코엑스", 12, GRAY),
        ("국내 최대 배터리 산업 전시회  |  핵심 키워드: AI, 로봇, ESS", 12, GRAY),
    ])
    stats = [("667", "참가 기업"), ("14", "참가 국가"), ("7.7만", "총 참관객"), ("3", "AI 핵심 기업")]
    for i, (num, label) in enumerate(stats):
        sx = bx + _emu(0.18) + i * _emu(1.1)
        _rect(slide, sx, by + _emu(1.0), _emu(1.0), _emu(1.05), SL_BLUE_BG, radius=0.08)
        _tb(slide, sx, by+_emu(1.08), _emu(1.0), _emu(0.4), num, 20, SL_BLUE, True, PP_ALIGN.CENTER)
        _tb(slide, sx, by+_emu(1.58), _emu(1.0), _emu(0.2), label, 12, GRAY, align=PP_ALIGN.CENTER)

    # AI explainer (right)
    ax = CL + _emu(5.1)
    _rect(slide, ax, by, _emu(5.1), _emu(2.3), WHITE, GRAY_BD, 0.03)
    _tb(slide, ax+_emu(0.18), by+_emu(0.1), _emu(4), _emu(0.25),
        "AI(인공지능)란 무엇인가?", 14, SL_BLUE, True)

    # analogy
    _rect(slide, ax+_emu(0.18), by+_emu(0.42), _emu(4.7), _emu(0.58),
          RGBColor(0xF0,0xF8,0xFF), radius=0.05)
    _rect(slide, ax+_emu(0.18), by+_emu(0.42), _emu(0.04), _emu(0.58),
          RGBColor(0x34,0x98,0xDB))
    _mtb(slide, ax+_emu(0.4), by+_emu(0.46), _emu(4.3), _emu(0.5), [
        ('"AI는 잠들지 않는 경험 많은 엔지니어와 같습니다.', 12, RGBColor(0x33,0x33,0x33)),
        ('수백만 개의 데이터를 동시에 읽고, 과거에서 배우며, 더 정확해집니다."', 12, RGBColor(0x33,0x33,0x33)),
    ])
    terms = [("\U0001F9E0", "머신러닝","컴퓨터가 데이터에서\n규칙과 패턴을 학습"),
             ("\U0001F4DA", "빅데이터","수작업으로 처리 불가능한\n방대한 양의 데이터"),
             ("\U0001F4E1", "실시간 모니터링","데이터 수집 즉시 분석,\n이상 발생 시 바로 알림")]
    for i, (icon, nm, desc) in enumerate(terms):
        tx = ax + _emu(0.18) + i * _emu(1.6)
        _rect(slide, tx, by+_emu(1.12), _emu(1.48), _emu(1.1), GRAY_BG, radius=0.06)
        _tb(slide, tx, by+_emu(1.16), _emu(1.48), _emu(0.25), icon, 16, BLACK, False, PP_ALIGN.CENTER)
        _tb(slide, tx, by+_emu(1.42), _emu(1.48), _emu(0.2), nm, 12, SL_BLUE, True, PP_ALIGN.CENTER)
        _tb(slide, tx+_emu(0.05), by+_emu(1.65), _emu(1.38), _emu(0.45), desc, 12, GRAY, align=PP_ALIGN.CENTER, spacing=1.35)

    # Bottom: AI adoption background
    aby = CT + _emu(2.95)
    _tb(slide, CL, aby, _emu(6), _emu(0.25), "배터리 산업의 AI 도입 배경", 14, SL_BLUE, True)
    reasons = [
        ("개발 기간", "전기차 시장 폭발적 성장으로 신제품 출시 주기가 수년에서 수개월로 단축"),
        ("안전 규제", "ESS 화재 사고 이후 배터리 안전성에 대한 규제와 사회적 요구 급격히 강화"),
        ("데이터 활용", "제조 공정의 센서 데이터가 사람이 분석할 수 없는 수준으로 증가"),
        ("글로벌 경쟁", "중국 CATL(세계 1위, 점유율 39%)은 5,000만 건 데이터로 AI 운영 중"),
    ]
    for idx, (title, desc) in enumerate(reasons):
        col = idx % 2; row = idx // 2
        rx = CL + _emu(0.1) + col * _emu(5.1)
        ry = aby + _emu(0.4) + row * _emu(1.05)
        _circle(slide, rx, ry + _emu(0.02), _emu(0.1), SL_BLUE)
        _tb(slide, rx+_emu(0.2), ry-_emu(0.02), _emu(1.8), _emu(0.22), title, 12, BLACK, True)
        _tb(slide, rx+_emu(0.2), ry+_emu(0.22), _emu(4.5), _emu(0.55), desc, 12, GRAY, spacing=1.35)


def slide_company_combined(prs, title, logo_text, logo_color, product, tagline,
                           ai_tasks, data_used, key_role,
                           flow_steps, before_items, after_items,
                           kpis=None, extra=None, footnotes=None):
    """Combined company slide: overview + flow on one page (no booth photo)"""
    slide = new_content_slide(prs, title)

    has_extra = extra is not None and len(extra) > 0

    # Logo + product + tagline
    # Detect multi-line product (e.g. SBI\nSBB)
    prod_lines = product.count('\n') + 1
    prod_h = _emu(0.3) * prod_lines
    logo_h = max(_emu(0.6), prod_h + _emu(0.1))
    logo = _rect(slide, CL, CT + _emu(0.02), _emu(0.6), logo_h, logo_color, radius=0.12)
    tc = WHITE if logo_color != SK_YELLOW else BLACK
    _tf(logo, logo_text, 12, tc, True, PP_ALIGN.CENTER)
    _tb(slide, CL + _emu(0.75), CT + _emu(0.02), _emu(8), prod_h, product, 20, BLACK, True)
    tag_y = CT + _emu(0.02) + prod_h + _emu(0.06)
    _tb(slide, CL + _emu(0.75), tag_y, _emu(8), _emu(0.22), tagline, 12, GRAY)

    # AI Tasks + Data cards
    card_top = tag_y + _emu(0.3)
    card_h = _emu(1.45) if has_extra else _emu(1.75)
    make_card(slide, CL, card_top, _emu(4.75), card_h,
              "AI가 하는 일", ai_tasks, logo_color, body_sz=12)
    make_card(slide, CL + _emu(5.25), card_top, _emu(5.0), card_h,
              "사용되는 데이터", data_used, logo_color, body_sz=12)

    # Extra info blocks (compact)
    next_y = card_top + card_h + _emu(0.08)
    if has_extra:
        for i, (etitle, eitems) in enumerate(extra):
            ex = CL + i * _emu(5.1)
            ew = _emu(4.85) if i == 0 else _emu(5.15)
            make_card(slide, ex, next_y, ew, _emu(1.15), etitle, eitems, logo_color, title_sz=12, body_sz=12)
        next_y += _emu(1.2)

    # Flow diagram (HTML style: #fafafa boundary box + icon steps + text arrows)
    FLOW_HUMAN_BG = RGBColor(0xE8, 0xF4, 0xFD)
    FLOW_HUMAN_BD = RGBColor(0x85, 0xC1, 0xE9)
    FLOW_HUMAN_TC = RGBColor(0x1A, 0x52, 0x76)
    FLOW_RESULT_BG = RGBColor(0xE8, 0xF5, 0xE9)
    FLOW_RESULT_BD = RGBColor(0x66, 0xBB, 0x6A)
    FLOW_RESULT_TC = RGBColor(0x1B, 0x5E, 0x20)
    AI_STYLES = {
        LG_RED: (RGBColor(0xFC,0xE4,0xEC), LG_RED, LG_RED),
        SS_BLUE: (RGBColor(0xE8,0xEA,0xF6), SS_BLUE, SS_BLUE),
        SK_YELLOW: (RGBColor(0xFF,0xF8,0xE1), SK_DARK, SK_YELLOW),
    }
    ai_bg, ai_tc, ai_bd = AI_STYLES.get(logo_color, (GRAY_BG, BLACK, GRAY))

    # "AI 작업 흐름" label (HTML h4 style: #888, uppercase)
    fy = next_y + _emu(0.05)
    _tb(slide, CL, fy, _emu(2), _emu(0.2), "AI 작업 흐름", 12, GRAY, True)

    # Outer boundary box (full width matching AI가 하는 일 + 사용되는 데이터 cards)
    box_x = CL
    box_w = CW
    n = len(flow_steps)
    step_h = _emu(0.75)
    flow_pad = _emu(0.15)
    box_h = step_h + flow_pad * 2
    step_w = _emu(1.1) if n <= 5 else _emu(0.88)
    arrow_gap = _emu(0.28)
    total_inner = n * step_w + (n - 1) * arrow_gap
    fy = fy + _emu(0.22)
    _rect(slide, box_x, fy, box_w, box_h, RGBColor(0xFA, 0xFA, 0xFA), radius=0.06)

    # Steps inside boundary (centered)
    sx = box_x + (box_w - total_inner) // 2
    sy = fy + flow_pad
    fsz = 12
    for i, (text, fill, ftc) in enumerate(flow_steps):
        fx = sx + i * (step_w + arrow_gap)
        if fill == RGBColor(0xE8, 0xF4, 0xFD):  # human
            _flow_box(slide, fx, sy, step_w, step_h, text,
                      FLOW_HUMAN_BG, FLOW_HUMAN_TC, fsz, FLOW_HUMAN_BD)
        elif fill == GREEN:  # result
            _flow_box(slide, fx, sy, step_w, step_h, text,
                      FLOW_RESULT_BG, FLOW_RESULT_TC, fsz, FLOW_RESULT_BD)
        elif fill == RED:  # alert
            _flow_box(slide, fx, sy, step_w, step_h, text,
                      RED_BG, RED, fsz, RED)
        else:  # AI
            _flow_box(slide, fx, sy, step_w, step_h, text,
                      ai_bg, ai_tc, fsz, ai_bd)
        # Text arrow between steps (HTML .flow-arrow: ➔ #aaa)
        if i < n - 1:
            ax_mid = fx + step_w
            _tb(slide, ax_mid, sy + _emu(0.08), arrow_gap, _emu(0.25),
                "\u27A1", 12, GRAY_L, False, PP_ALIGN.CENTER)

    # Before / After (HTML style: #fff5f5 / #f0fff4, ❌/✅ markers)
    BA_BEFORE_BG = RGBColor(0xFF, 0xF5, 0xF5)
    BA_AFTER_BG = RGBColor(0xF0, 0xFF, 0xF4)
    bay = fy + box_h + _emu(0.2)
    baw = _emu(4.6)
    bah = _emu(1.2)
    _rect(slide, CL, bay, baw, bah, BA_BEFORE_BG, radius=0.06)
    _tb(slide, CL + _emu(0.12), bay + _emu(0.05), _emu(3), _emu(0.18),
        "BEFORE AI (기존)", 12, RED, True)
    for i, item in enumerate(before_items):
        _tb(slide, CL + _emu(0.15), bay + _emu(0.3) + i * _emu(0.28), baw - _emu(0.3), _emu(0.2),
            f"\u274C  {item}", 12, BLACK)
    # Arrow box (HTML .ba-arrow: #f0f0f0 bg, centered arrow)
    arrow_x = CL + baw
    arrow_w = _emu(0.5)
    _rect(slide, arrow_x, bay, arrow_w, bah, RGBColor(0xF0, 0xF0, 0xF0))
    _arrow(slide, arrow_x + _emu(0.14), bay + bah // 2 - _emu(0.08), _emu(0.22), _emu(0.18))
    ax = CL + baw + arrow_w
    aw = CW - baw - arrow_w
    _rect(slide, ax, bay, aw, bah, BA_AFTER_BG, radius=0.06)
    _tb(slide, ax + _emu(0.12), bay + _emu(0.05), _emu(3), _emu(0.18),
        "AFTER AI (현재)", 12, GREEN, True)
    for i, item in enumerate(after_items):
        _tb(slide, ax + _emu(0.15), bay + _emu(0.3) + i * _emu(0.28), aw - _emu(0.3), _emu(0.2),
            f"\u2705  {item}", 12, BLACK)

    # KPIs
    if kpis:
        ky = bay + bah + _emu(0.1)
        kpi_w = _emu(3.25); kpi_gap = _emu(0.15)
        for i, (num, unit, sub) in enumerate(kpis):
            kx = CL + i * (kpi_w + kpi_gap)
            _rect(slide, kx, ky, kpi_w, _emu(0.8), GRAY_BG, radius=0.05)
            _tb(slide, kx, ky + _emu(0.02), kpi_w, _emu(0.3), num, 22, logo_color, True, PP_ALIGN.CENTER)
            _tb(slide, kx, ky + _emu(0.35), kpi_w, _emu(0.13), unit, 12, GRAY_D, align=PP_ALIGN.CENTER)
            _tb(slide, kx, ky + _emu(0.55), kpi_w, _emu(0.1), sub, 12, GRAY_L, align=PP_ALIGN.CENTER)

    # Footnotes (bottom-left, ※ style)
    if footnotes:
        fn_text = "    ".join([f"※ {fn}" for fn in footnotes])
        _tb(slide, CL, CB - _emu(0.28), CW, _emu(0.25), fn_text, 10, GRAY)


def slide_09_comparison_matrix(prs):
    slide = new_content_slide(prs, "비교 분석 — AI 적용 영역")
    _tb(slide, CL, CT, _emu(8), _emu(0.2),
        "배터리 밸류체인의 어느 단계에 AI를 적용했는지 비교", 18, RGBColor(0x66, 0x66, 0x66))

    rows = [
        ("밸류체인 단계", "LG CNS", "Samsung SDI", "SK ON"),
        ("R&D / 설계", "—", "○ 부분", "● 핵심"),
        ("시험 / 분석", "● 핵심", "—", "○ 부분"),
        ("생산 / 품질", "○ 부분", "● 핵심", "● 핵심"),
        ("운영 / 모니터링", "—", "● 핵심", "—"),
        ("안전 / 예방", "○ 부분", "● 핵심", "○ 부분"),
    ]
    cws = [_emu(2.2), _emu(2.7), _emu(2.7), _emu(2.7)]
    tbl_shape = make_table(slide, CL, CT + _emu(0.35), cws, rows)

    # Color the check/partial cells
    tbl = tbl_shape.table
    for ri in range(1, len(rows)):
        for ci in range(1, 4):
            cell = tbl.cell(ri, ci)
            txt = cell.text_frame.paragraphs[0].text
            if "핵심" in txt:
                cell.text_frame.paragraphs[0].font.color.rgb = GREEN
                cell.text_frame.paragraphs[0].font.bold = True
            elif "부분" in txt:
                cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xF3,0x9C,0x12)
            elif txt == "—":
                cell.text_frame.paragraphs[0].font.color.rgb = GRAY_L

    _tb(slide, CL, CT + _emu(2.95), _emu(8), _emu(0.2),
        "● 핵심 = AI 적용 주력 영역    ○ 부분 = 보조적 활용    — = 미적용 또는 미공개", 12, GRAY)

    # Insight box
    iy = CT + _emu(3.3)
    _rect(slide, CL, iy, CW, _emu(2.2), SL_BLUE_BG, radius=0.02)
    _rect(slide, CL, iy, _emu(0.04), _emu(2.2), SL_BLUE)
    _tb(slide, CL+_emu(0.25), iy+_emu(0.1), _emu(4), _emu(0.25), "핵심 발견", 14, SL_BLUE, True)
    findings = [
        "3사의 AI 적용 영역이 확연히 다름 — 경쟁이 아닌 상호보완적 구조",
        "LG CNS = 시험/분석 + 폐배터리  |  Samsung SDI = 운영/안전/생산  |  SK ON = R&D/설계/생산",
        "설계 → 시험 → 생산 → 운영 → 안전까지, 밸류체인 전 단계에서 AI 도입이 가속화 중",
    ]
    _bullet_list(slide, CL+_emu(0.25), iy+_emu(0.4), CW-_emu(0.5), _emu(1.6), findings, 12, RGBColor(0x33, 0x33, 0x33), "▸  ")


def slide_10_chart(prs):
    slide = new_content_slide(prs, "비교 분석 — AI 활용 수준")
    _tb(slide, CL, CT, _emu(8), _emu(0.3), "레이더 차트 및 종합 비유", 18, RGBColor(0x66, 0x66, 0x66))

    # Radar chart
    chart_data = CategoryChartData()
    chart_data.categories = ['자동화 수준', '데이터 활용', '실시간성', 'AI 제품 다양성', '성과 입증도']
    chart_data.add_series('LG CNS', (8, 7, 7, 8, 7))
    chart_data.add_series('Samsung SDI', (7, 9, 9, 8, 7))
    chart_data.add_series('SK ON', (9, 7, 5, 8, 9))
    cf = slide.shapes.add_chart(XL_CHART_TYPE.RADAR_FILLED,
                                CL, CT+_emu(0.35), _emu(5.0), _emu(4.5), chart_data)
    chart = cf.chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.font.size = Pt(12); chart.legend.font.name = FONT
    s0 = chart.series[0]; s0.format.fill.solid(); s0.format.fill.fore_color.rgb = LG_RED
    s1 = chart.series[1]; s1.format.fill.solid(); s1.format.fill.fore_color.rgb = SS_BLUE
    s2 = chart.series[2]; s2.format.fill.solid(); s2.format.fill.fore_color.rgb = SK_YELLOW
    _tb(slide, CL, CT+_emu(4.95), _emu(5.0), _emu(0.2),
        "* 인터배터리 2026 전시 내용 기반의 상대적 평가", 12, GRAY, align=PP_ALIGN.CENTER)

    # Role comparison (right)
    cx = CL + _emu(5.3)
    _tb(slide, cx, CT+_emu(3.7), _emu(4), _emu(0.25), "종합 비유", 14, SL_BLUE, True)
    roles = [("LG CNS", "배터리 자동 실험실", LG_RED),
             ("Samsung SDI", "24시간 배터리 전담의", SS_BLUE),
             ("SK ON", "AI 동료 연구원", SK_YELLOW)]
    for i, (co, role, color) in enumerate(roles):
        ry = CT + _emu(4.05) + i * _emu(0.38)
        _circle(slide, cx+_emu(0.08), ry+_emu(0.03), _emu(0.12), color)
        tc = BLACK if color != SK_YELLOW else SK_DARK
        _tb(slide, cx+_emu(0.3), ry, _emu(1.5), _emu(0.2), co, 12, tc, True)
        _tb(slide, cx+_emu(1.85), ry, _emu(2.5), _emu(0.2), f"=  {role}", 12, GRAY)


def slide_11_comparison_table(prs):
    slide = new_content_slide(prs, "비교 분석 — 종합 비교표")
    _tb(slide, CL, CT, _emu(8), _emu(0.3), "3사의 AI 전략을 한눈에 비교", 18, RGBColor(0x66, 0x66, 0x66))

    rows = [
        ("항목", "LG CNS", "Samsung SDI", "SK ON"),
        ("회사 유형", "IT 서비스", "배터리 제조", "배터리 제조"),
        ("AI 핵심 제품", "AX 사이클러\n포메이션\nFactova HED", "SBI\nAI BMS\nVision AI", "ADAM\n비전 AI\n수명 예측 AI"),
        ("AI가 하는 일", "실험 자동 설계\n결과 분석·보고서", "24시간 배터리 감시\n수명 예측·사고 예방", "셀 설계 자동화\n성능·원가 예측"),
        ("주요 성과", "생산성 30%↑\n장비 20대 납품", "1,400개 현장 관리\nSBI 10월 상용화", "설계 기간 1/3\nAI Day 대상 수상"),
        ("적용 단계", "시험·분석\n폐배터리 처리", "운영·안전\n생산·품질", "R&D·설계\n생산·품질"),
        ("비유적 역할", "배터리 자동 실험실", "24시간 배터리 전담의", "AI 동료 연구원"),
    ]
    cws = [_emu(1.8), _emu(2.8), _emu(2.8), _emu(2.8)]
    make_table(slide, CL, CT + _emu(0.35), cws, rows)


def slide_12_insights(prs):
    slide = new_content_slide(prs, "시사점")
    _tb(slide, CL, CT, _emu(6), _emu(0.2),
        "3사의 AI 활용 분석을 통해 도출한 핵심 시사점", 18, RGBColor(0x66, 0x66, 0x66))

    insights = [
        ("01", "밸류체인 전 단계에서\nAI 도입 가속화",
         "설계(SK ON) → 시험(LG CNS) → 생산(삼성SDI·SK ON)\n→ 운영·안전(삼성SDI) → 폐배터리(LG CNS)까지\n밸류체인 전 단계에서 AI 도입 진행 중.\n배터리 산업에서 AI는 선택이 아닌 필수.",
         LG_RED),
        ("02", "각사의 차별화된\nAI 전략",
         "LG CNS = 실험 자동화 (IT기업의 AI 소프트웨어 강점)\nSamsung SDI = 배터리 진단 (1,400개 현장 데이터)\nSK ON = AI 연구원 (설계·예측 자동화)\n같은 'AI'라도 접근 방식이 완전히 다름.",
         SS_BLUE),
        ("03", 'AI는 사람을\n"대체"하지 않는다',
         "SK ON: AI 연구원은 '함께 일하는 동료'\nAI가 반복 작업 대신 수행.\n사람은 최종 판단·창의적 결정에 집중.\n3사 모두 AI를 '도구'가 아닌 '파트너'로 활용.",
         SK_YELLOW),
    ]
    cw = _emu(3.2); gap = _emu(0.15)
    for i, (num, title, desc, color) in enumerate(insights):
        cx = CL + i * (cw + gap)
        cy = CT + _emu(0.4)
        _rect(slide, cx, cy, cw, _emu(5.2), WHITE, GRAY_BD, 0.03)
        _rect(slide, cx, cy, cw, _emu(0.04), color)  # top accent
        _tb(slide, cx+_emu(0.18), cy+_emu(0.15), _emu(1), _emu(0.4), num, 26, GRAY_BD, True)
        _tb(slide, cx+_emu(0.18), cy+_emu(0.55), cw-_emu(0.36), _emu(0.6), title, 15, SL_BLUE, True, spacing=1.3)
        _tb(slide, cx+_emu(0.18), cy+_emu(1.3), cw-_emu(0.36), _emu(3.0), desc, 12, RGBColor(0x55, 0x55, 0x55), spacing=1.5)


def slide_13_sl_implications(prs):
    slide = new_content_slide(prs, "SL에 주는 시사점")
    _tb(slide, CL, CT, _emu(6), _emu(0.2), "배터리 3사의 AI 활용이 SL에 시사하는 바", 18, RGBColor(0x66, 0x66, 0x66))

    _rect(slide, CL, CT+_emu(0.35), CW, _emu(5.0), SL_BLUE, radius=0.02)

    items = [
        ("1", "글로벌 AI 경쟁 심화",
         "중국 CATL(세계 1위, 점유율 39%), 5,000만 건 이상 데이터로 AI 운영 중.\n한국 3사 합산 점유율 약 16%. AI 도입 지연 시 격차 확대."),
        ("2", "데이터 확보가 AI 성공의 핵심",
         "삼성SDI 1,400개 현장, SK ON 설계 DB, LG CNS 실험 데이터를 AI에 활용.\n체계적 데이터 관리가 AI 도입의 첫걸음."),
        ("3", "작은 것부터 시작하는 것이 현실적",
         "처음부터 대규모 AI가 아닌, 특정 업무(시험 자동화, 데이터 분석 등)부터\n시작하여 점진적 확대. 3사의 공통된 접근 방식."),
    ]
    for idx, (num, title, desc) in enumerate(items):
        iy = CT + _emu(0.6) + idx * _emu(1.35)
        c = _circle(slide, CL + _emu(0.3), iy, _emu(0.32), RGBColor(0x1A, 0x4E, 0x8A))
        _tf(c, num, 12, WHITE, True, PP_ALIGN.CENTER)
        _tb(slide, CL + _emu(0.8), iy - _emu(0.02), _emu(7), _emu(0.25), title, 14, WHITE, True)
        _tb(slide, CL + _emu(0.8), iy + _emu(0.28), _emu(8.5), _emu(0.7), desc, 12,
            WHITE, spacing=1.45)


def slide_14_glossary(prs):
    slide = new_content_slide(prs, "부록 — AI 용어집")
    _tb(slide, CL, CT, _emu(6), _emu(0.2), "보고서에 등장하는 주요 용어 정리", 18, RGBColor(0x66, 0x66, 0x66))

    terms = [
        ("AI (인공지능)", "사람의 학습·판단 능력을 컴퓨터로 구현한 기술.\n데이터에서 패턴을 찾고 예측."),
        ("SBI", "Samsung Battery Intelligence.\n삼성SDI의 AI 배터리 진단 소프트웨어."),
        ("ADAM", "AI-Based Design & Analysis Machine.\nSK ON의 AI 배터리 설계 시스템."),
        ("BMS", "Battery Management System.\n배터리 충방전과 온도를 관리하는 시스템."),
        ("ESS", "Energy Storage System.\n대용량 에너지를 저장하는 장치."),
        ("에이전틱 AI", "사람의 지시 없이도 스스로 판단하고\n실행하는 AI. LG CNS가 배터리 장비에 적용."),
        ("RFQ", "Request for Quotation.\n고객이 원하는 배터리 사양을 적어 보내는 요청서."),
        ("비전 AI", "카메라 이미지를 AI가 분석,\n불량품을 자동 검출하는 기술."),
        ("밸류체인", "설계→시험→생산→운영의 전체 과정.\n제품이 만들어지는 모든 단계."),
    ]
    for i, (term, desc) in enumerate(terms):
        col = i % 2; row = i // 2
        tx = CL + col * _emu(5.15)
        ty = CT + _emu(0.35) + row * _emu(1.05)
        tw = _emu(4.95)
        _rect(slide, tx, ty, tw, _emu(0.92), GRAY_BG, radius=0.04)
        _tb(slide, tx+_emu(0.15), ty+_emu(0.08), tw-_emu(0.3), _emu(0.22), term, 12, SL_BLUE, True)
        _tb(slide, tx+_emu(0.15), ty+_emu(0.35), tw-_emu(0.3), _emu(0.35), desc, 12, RGBColor(0x66, 0x66, 0x66), spacing=1.3)


def slide_15_sources(prs):
    slide = new_content_slide(prs, "부록 — 출처")
    _tb(slide, CL, CT, _emu(6), _emu(0.2), "본 보고서 작성에 활용된 자료", 18, RGBColor(0x66, 0x66, 0x66))

    sources = [
        "인터배터리 2026 전시회 현장 조사 (2026.3.11~13, 서울 코엑스)",
        "LG CNS — AI로 배터리 공정 혁신 (스마트비즈, kidd, ZDNet, 2026.3)",
        "삼성SDI — AI 기반 ESS 화재 예방 SW 'SBI' (서울경제, 헤럴드경제, Korea Herald)",
        "SK ON — AI 연구원이 배터리 개발 참여 (ASK inno, 오늘경제, 2026.3)",
        "배터리 3사 기술 수장 AI R&D 혁신 발표 (헤럴드경제, 2026.3)",
        "IDTechEx — AI-Driven Battery Technology 2025-2035",
        "CATL AI Strategy Analysis (SCMP, Klover.ai, 2026.3)",
    ]
    for i, src in enumerate(sources):
        sy = CT + _emu(0.4) + i * _emu(0.48)
        _tb(slide, CL + _emu(0.15), sy, _emu(0.25), _emu(0.25), "\U0001F4C4", 12, BLACK, False, PP_ALIGN.CENTER)
        _tb(slide, CL + _emu(0.45), sy, _emu(9), _emu(0.25), src, 12, RGBColor(0x55, 0x55, 0x55))

    # Closing
    cy = CT + _emu(4.0)
    _rect(slide, CL, cy, CW, _emu(1.2), SL_BLUE, radius=0.02)
    _tb(slide, CL, cy + _emu(0.15), CW, _emu(0.4),
        "감사합니다", 22, WHITE, True, PP_ALIGN.CENTER)
    _tb(slide, CL, cy + _emu(0.6), CW, _emu(0.3),
        "국내 배터리 3사 AI 활용 현황 분석  |  인터배터리 2026  |  SL  |  2026년 3월",
        12, RGBColor(0xBB,0xCC,0xEE), align=PP_ALIGN.CENTER)


def slide_05_summary_compare(prs):
    """3사 핵심 성과 & 미래 계획 한눈에 비교"""
    slide = new_content_slide(prs, "3사 핵심 성과 & 미래 계획")
    _tb(slide, CL, CT + _emu(0.05), _emu(8), _emu(0.3),
        "배터리 3사의 AI 도입 성과와 향후 방향", 18, RGBColor(0x66, 0x66, 0x66))

    companies = [
        ("LG CNS", LG_RED, WHITE, "LLM·에이전틱 AI",
         ["생산성 30% 이상 향상 (AI 도입 시)", "LG에너지솔루션에 장비 20대 납품", "AI 탑재 장비 3종 라인업 구축"],
         ["R&D~생산~재활용 전 과정 AI 확대", "배터리 산업 데이터 통합 플랫폼 구축", "AI 장비 라인업 지속 확대"]),
        ("Samsung SDI", SS_BLUE, WHITE, "AI 기반",
         ["전 세계 1,400개 현장 AI 관리", "SBI 2026년 10월 상용화 예정", "제조 공정 500가지 품질 항목 AI 체크"],
         ["SBI 2026년 10월 상용화", "차세대 전고체 배터리 2027 양산", "로봇용 AI 배터리 사업 확대"]),
        ("SK ON", SK_YELLOW, BLACK, "AI 기반",
         ["배터리 설계 기간 1/3로 단축", "SKI AI Day·AIDT Awards 대상 수상", "소재 개발 AI 연구원 별도 개발 중"],
         ["2028년 배터리 전용 AI 구축", "AI로 배터리 소재 개발 자동화", "설계~영업까지 전사적 AI 확대"]),
    ]

    label_w = _emu(0.8)
    content_x = CL + label_w + _emu(0.15)
    col_w = _emu(2.8)
    col_gap = _emu(0.25)

    # Company header bars + AI tech type
    hy = CT + _emu(0.5)
    for i, (name, color, tc, ai_tech, _, _) in enumerate(companies):
        cx = content_x + i * (col_w + col_gap)
        _rect(slide, cx, hy, col_w, _emu(0.35), color, radius=0.06)
        _tb(slide, cx, hy + _emu(0.02), col_w, _emu(0.3), name, 14, tc, True, PP_ALIGN.CENTER)
        # AI technology type label
        _tb(slide, cx, hy + _emu(0.38), col_w, _emu(0.2), ai_tech, 12, GRAY, False, PP_ALIGN.CENTER)

    # ── 핵심 성과 ──
    sy = CT + _emu(1.2)
    box_h = _emu(2.0)
    # Left label
    label_bg = _rect(slide, CL, sy, label_w, box_h, SL_BLUE, radius=0.06)
    _tf(label_bg, "핵심\n성과", 14, WHITE, True, PP_ALIGN.CENTER)
    # Big content box
    total_content_w = col_w * 3 + col_gap * 2
    _rect(slide, content_x, sy, total_content_w, box_h, GRAY_BG, radius=0.04)
    # Column dividers
    for d in range(1, 3):
        div_x = content_x + d * (col_w + col_gap) - col_gap // 2
        _line(slide, div_x, sy + _emu(0.1), Pt(1), GRAY_L)
        # Vertical line using thin rect
        _rect(slide, div_x, sy + _emu(0.1), Pt(1), box_h - _emu(0.2), GRAY_L)
    for i, (_, _, _, _, achievements, _) in enumerate(companies):
        cx = content_x + i * (col_w + col_gap)
        for j, item in enumerate(achievements):
            _tb(slide, cx + _emu(0.1), sy + _emu(0.15) + j * _emu(0.55),
                col_w - _emu(0.2), _emu(0.5), f"\u25B8 {item}", 12, BLACK)

    # ── 미래 계획 ──
    py = sy + box_h + _emu(0.35)
    # Left label
    label_bg2 = _rect(slide, CL, py, label_w, box_h, SL_BLUE2, radius=0.06)
    _tf(label_bg2, "미래\n계획", 14, WHITE, True, PP_ALIGN.CENTER)
    # Big content box
    _rect(slide, content_x, py, total_content_w, box_h, SL_BLUE_BG, radius=0.04)
    # Column dividers
    for d in range(1, 3):
        div_x = content_x + d * (col_w + col_gap) - col_gap // 2
        _rect(slide, div_x, py + _emu(0.1), Pt(1), box_h - _emu(0.2), GRAY_L)
    for i, (_, _, _, _, _, plans) in enumerate(companies):
        cx = content_x + i * (col_w + col_gap)
        for j, item in enumerate(plans):
            _tb(slide, cx + _emu(0.1), py + _emu(0.15) + j * _emu(0.55),
                col_w - _emu(0.2), _emu(0.5), f"\u25B8 {item}", 12, BLACK)

    # Footnote
    _tb(slide, CL, CB - _emu(0.28), CW, _emu(0.25),
        "※ 전고체 배터리: 액체 전해질을 고체로 대체한 차세대 배터리. 화재 위험 대폭 감소, 에너지 밀도 40% 향상",
        10, GRAY)


# ===================================================================
# MAIN
# ===================================================================

def main():
    prs = Presentation(TEMPLATE)
    # Remove existing 2 slides from template
    while len(prs.slides) > 0:
        rId = prs.slides._sldIdLst[0].get(qn('r:id'))
        prs.part.drop_rel(rId)
        prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])

    # 1. Cover
    slide_01_cover(prs)

    # 2. Summary
    slide_03_summary(prs)

    # 3. LG CNS (통합)
    slide_company_combined(prs,
        title="LG CNS", logo_text="LG\nCNS", logo_color=LG_RED,
        product="AI 배터리 장비 솔루션",
        tagline="에이전틱 AI 기반 — R&D 시험부터 폐배터리 처리까지 전 과정 자동화",
        ai_tasks=[
            "충방전 시험 장비에 AI 탑재, 실험 설계·분석·보고서 자동화",
            "배터리 활성화 장비에 AI 탑재, 공정 시간 단축",
            "폐배터리 방전 장비에 AI 탑재, 안전한 방전 조건 자동 계산",
        ],
        data_used=[
            "충방전 사이클, 전압·전류·온도 측정값",
            "배터리 활성화 공정의 전류 패턴 데이터",
            "폐배터리의 전압, 온도, 전류 실시간 데이터",
        ],
        key_role="IT 서비스 기업(SW)과 에이테크놀로지(HW) 공동 개발",
        flow_steps=[
            ("\U0001F468\u200D\U0001F52C 연구원\n명령 입력", RGBColor(0xE8,0xF4,0xFD), BLACK),
            ("\U0001F916 AI\n실험 설계", LG_RED, WHITE),
            ("\u2699 장비\n자동 실행", RGBColor(0xE8,0xF4,0xFD), BLACK),
            ("\U0001F916 AI\n실시간 분석", LG_RED, WHITE),
            ("\U0001F4C4 보고서\n자동 생성", GREEN, WHITE),
        ],
        before_items=["연구원이 수작업으로 실험 조건 계산",
                      "엑셀로 수동 분석 (수 시간 소요)",
                      "폐배터리 방전 시 화재 위험에 노출"],
        after_items=["AI가 최적 실험 조건 즉시 생성",
                     "실시간 자동 분석 및 보고서 작성",
                     "AI가 안전한 방전 조건 자동 계산"],
        footnotes=["에이전틱 AI(Agentic AI): 사람의 지시 없이도 스스로 판단·실행하는 AI"],
    )

    # 6. Samsung SDI (통합)
    slide_company_combined(prs,
        title="Samsung SDI", logo_text="SDI", logo_color=SS_BLUE,
        product="SBI (Samsung Battery Intelligence)\nSBB (Samsung Battery Box)",
        tagline="AI 기반 — 전 세계 1,400개 배터리 현장 24시간 진단, 사고 예방",
        ai_tasks=[
            "전 세계 1,400개+ 배터리 현장 AI 24시간 감시·진단",
            "이상 징후 사전 감지·수명 예측 후 진단 리포트 자동 생성",
            "SBI가 SBB 내부 배터리를 원격 제어, 위험 시 자동 전력 차단",
            "Vision AI로 제조 공정 불량 자동 검출",
        ],
        data_used=[
            "전 세계 1,400개 배터리 현장의 운영 데이터",
            "전압·온도·건강상태 센서 데이터, 성능 변화 추이",
            "SBB 내부 배터리 상태 실시간 제어 데이터",
            "제조 공정 검사 이미지, X-ray 데이터",
        ],
        key_role="슬로건: 'AI thinks, Battery enables'",
        flow_steps=[
            ("\U0001F50B 현장 데이터\n수집", RGBColor(0xE8,0xF4,0xFD), BLACK),
            ("\U0001F916 SBI\n24시간 감시", SS_BLUE, WHITE),
            ("\U0001F916 이상 징후\n감지·수명 예측", SS_BLUE, WHITE),
            ("\U0001F916 SBB\n원격 제어·차단", SS_BLUE, WHITE),
            ("\U0001F916 Vision AI\n불량 검출", SS_BLUE, WHITE),
            ("\U0001F4C4 진단 리포트\n생성", GREEN, WHITE),
        ],
        before_items=["사후 대응: 문제 발생 후 감지",
                      "사람이 수동 데이터 분석",
                      "배터리 교체 시기 경험에 의존"],
        after_items=["사전 예방: AI가 이상 징후 미리 감지",
                     "전 세계 1,400개 현장 AI 24시간 자동 감시",
                     "정확한 수명 예측으로 최적 교체 시기 결정"],
        footnotes=[
            "SBI(Samsung Battery Intelligence): AI 배터리 진단 소프트웨어",
            "SBB(Samsung Battery Box): SBI가 AI로 관리하는 대용량 배터리 저장 장치",
            "Vision AI: 카메라·X-ray 이미지를 AI가 분석, 불량 검출",
        ],
    )

    # 7. SK ON (통합)
    slide_company_combined(prs,
        title="SK ON", logo_text="SK\nON", logo_color=SK_YELLOW,
        product="ADAM & AI 연구원 체계",
        tagline="AI 기반 — AI 연구원이 배터리 설계~성능 예측을 자동 수행",
        ai_tasks=[
            "셀 설계 AI: 고객 요청서 수신 시 다수 설계안 자동 생성",
            "성능 예측 AI: 실제 배터리 제작 없이 성능 사전 예측",
            "원가 산출 AI: 재료비·공정비 등 제조 비용 자동 계산",
            "비전 AI로 불량 검출, 초기 데이터로 수명 예측",
        ],
        data_used=[
            "과거 배터리 셀 설계 데이터베이스",
            "배터리 소재 특성·성능 시험 결과 데이터",
            "원가 데이터 (재료비·공정비)",
            "제조 공정 이미지 데이터",
        ],
        key_role="'AI 연구원'은 사람을 대체하는 것이 아닌, 함께 일하는 AI 동료",
        flow_steps=[
            ("\U0001F468\u200D\U0001F52C 고객 요청서\n입력", RGBColor(0xE8,0xF4,0xFD), BLACK),
            ("\U0001F916 셀 설계\nAI", SK_YELLOW, BLACK),
            ("\U0001F916 성능\n예측 AI", SK_YELLOW, BLACK),
            ("\U0001F916 원가\n산출 AI", SK_YELLOW, BLACK),
            ("\U0001F916 분석 결과\n자동 정리", SK_YELLOW, BLACK),
            ("\U0001F468\u200D\U0001F52C 연구원\n최종 판단", RGBColor(0xE8,0xF4,0xFD), BLACK),
            ("\u2705 최종\n설계안", GREEN, WHITE),
        ],
        before_items=["연구원 수작업 셀 설계 (시행착오 반복)",
                      "한 번에 하나의 설계안만 검토 가능",
                      "원가 계산에 별도 팀과 수 주 소요"],
        after_items=["AI가 다수 설계안 동시 자동 생성",
                     "실제 제작 없이 성능·원가 예측",
                     "연구원은 안전성 판단·최종 결정에 집중"],
        footnotes=[
            "ADAM(AI-Based Design & Analysis Machine): AI 배터리 설계 시스템",
            "RFQ(Request for Quotation): 고객이 원하는 배터리 사양 요청서",
        ],
    )

    # 8. 3사 핵심 성과 & 미래 계획
    slide_05_summary_compare(prs)

    prs.save(OUTPUT)
    size_mb = os.path.getsize(OUTPUT) / (1024 * 1024)
    print(f"PPT generated: {OUTPUT} ({size_mb:.1f} MB, {len(prs.slides)} slides)")


if __name__ == "__main__":
    main()
