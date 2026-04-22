"""Slide 2 — 핵심 요약 (Executive Summary)

- Key Message Bar (상단)
- 3 메시지 카드 (1x3) — 경쟁사 / 영역별 효과 / 품질검증 부담
- 승인 요청 박스 (2x2)

자사 예측 수치 없음. 외부 수치만 출처 병기.
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
    make_icon_circle,
    make_icon_badge,
    add_shadow,
    set_shape_opacity,
    set_body_anchor,
    set_text_inset,
    calc_grid,
)
from template_contract import CONTENT_SAFE, LAYOUT_CONTENT


# ----------------- 팔레트 -----------------
C_NAVY  = RGBColor(0x1F, 0x49, 0x7D)
C_GREEN = RGBColor(0x2E, 0x8B, 0x57)
C_BLUE  = RGBColor(0x4F, 0x81, 0xBD)
C_RED   = RGBColor(0xC0, 0x50, 0x4D)
C_INK   = RGBColor(0x1A, 0x1F, 0x2E)
C_GRAY  = RGBColor(0x66, 0x66, 0x66)
C_LGRAY = RGBColor(0xE0, 0xE4, 0xEA)
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
C_BG    = RGBColor(0xF7, 0xF9, 0xFC)   # 카드 배경
C_APBG  = RGBColor(0xEE, 0xF3, 0xF9)   # 승인 박스 배경


def _add_filled_rect(slide, x, y, w, h, fill, line=None, line_w_pt=None):
    """직각 사각형 배경 패널."""
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        if line_w_pt is not None:
            shp.line.width = Pt(line_w_pt)
    return shp


def _build_key_message_bar(slide, x, y, w, h):
    """상단 Key Message Bar."""
    # 배경 패널
    panel = _add_filled_rect(slide, x, y, w, h, C_NAVY)

    # 좌측 텍스트 (흰색, Key Message)
    pad = Inches(0.18)
    tb = add_textbox(
        slide,
        x + pad,
        y,
        w - pad * 2,
        h,
        "",
        font_size=14,
    )
    tf = tb.text_frame
    tf.word_wrap = True
    # 첫 단락 지우기 대용 — 기본 단락 제거하고 rich 삽입
    p0 = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.LEFT
    # 기본 run 제거
    for r in list(p0.runs):
        r.text = ""

    # Label
    add_rich_text(
        tf,
        [
            {"text": "KEY MESSAGE  ", "font_size": 9, "color": C_LGRAY, "bold": True},
            {"text": "영역별 차등 적용 필수", "font_size": 15, "color": C_WHITE, "bold": True},
            {"text": "  ·  ", "font_size": 15, "color": C_LGRAY, "bold": False},
            {"text": "경쟁사는 이미 도입 단계", "font_size": 15, "color": C_WHITE, "bold": True},
        ],
        align=PP_ALIGN.LEFT,
        space_before=Pt(0),
        space_after=Pt(0),
    )
    set_body_anchor(tb, "ctr")
    set_text_inset(tb, left=Inches(0.04), right=Inches(0.04),
                   top=Inches(0.02), bottom=Inches(0.02))


def _build_card(slide, rect, accent, icon, title, lines):
    """메시지 카드 1개 — 상단 accent bar + 아이콘 + 제목 + 라인 리스트."""
    x, y, w, h = rect.left, rect.top, rect.width, rect.height

    # 카드 배경 (연회색 패널)
    card = _add_filled_rect(slide, x, y, w, h, C_BG, line=C_LGRAY, line_w_pt=0.75)

    # 상단 accent bar
    bar_h = Inches(0.08)
    add_accent_bar(slide, x, y, w, bar_h, accent)

    # 아이콘 circle (좌상단)
    pad = Inches(0.18)
    icon_sz = Inches(0.42)
    icon_y = y + bar_h + Inches(0.14)
    make_icon_circle(
        slide,
        x + pad,
        icon_y,
        icon_sz,
        accent,
        text=icon,
        font_size=14,
        font_color=C_WHITE,
    )

    # 카드 제목 (아이콘 우측)
    title_x = x + pad + icon_sz + Inches(0.12)
    title_w = w - (title_x - x) - pad
    title_h = icon_sz
    title_tb = add_textbox(
        slide,
        title_x,
        icon_y,
        title_w,
        title_h,
        title,
        font_size=12,
        color=C_INK,
        bold=True,
        align=PP_ALIGN.LEFT,
    )
    set_body_anchor(title_tb, "ctr")
    set_text_inset(title_tb, left=Inches(0.02), right=Inches(0.02),
                   top=Inches(0.01), bottom=Inches(0.01))

    # 구분선
    sep_y = icon_y + icon_sz + Inches(0.10)
    add_accent_bar(slide, x + pad, sep_y, w - pad * 2, Emu(9525), C_LGRAY)

    # 라인 리스트 (label + value)
    list_top = sep_y + Inches(0.10)
    list_h = (y + h) - list_top - Inches(0.12)
    list_x = x + pad
    list_w = w - pad * 2

    lb = slide.shapes.add_textbox(list_x, list_top, list_w, list_h)
    tf = lb.text_frame
    tf.word_wrap = True

    for i, ln in enumerate(lines):
        label = ln.get("label", "")
        value = ln.get("value", "")
        segs = []
        # 앞에 작은 bullet
        segs.append({"text": "•  ", "font_size": 10, "color": accent, "bold": True})
        # label (강조 볼드)
        segs.append({"text": str(label), "font_size": 10, "color": C_INK, "bold": True})
        if value:
            segs.append({"text": "   " + str(value), "font_size": 10, "color": C_GRAY, "bold": False})

        if i == 0:
            # 첫 단락: 기본 p 재활용
            p0 = tf.paragraphs[0]
            p0.alignment = PP_ALIGN.LEFT
            for r in list(p0.runs):
                r.text = ""
        add_rich_text(
            tf,
            segs,
            align=PP_ALIGN.LEFT,
            space_before=Pt(2),
            space_after=Pt(2),
            line_spacing=Pt(14),
        )
    set_body_anchor(lb, "t")


def _build_approval_box(slide, x, y, w, h, accent, title, items):
    """하단 승인 요청 박스 — 제목 헤더 + 2x2 체크리스트."""
    # 외곽 패널
    _add_filled_rect(slide, x, y, w, h, C_APBG, line=accent, line_w_pt=1.0)

    # 헤더 바
    hdr_h = Inches(0.36)
    _add_filled_rect(slide, x, y, w, hdr_h, accent)

    # 헤더 텍스트
    hdr_tb = add_textbox(
        slide,
        x + Inches(0.18),
        y,
        w - Inches(0.36),
        hdr_h,
        title,
        font_size=12,
        color=C_WHITE,
        bold=True,
        align=PP_ALIGN.LEFT,
    )
    set_body_anchor(hdr_tb, "ctr")
    set_text_inset(hdr_tb, left=Inches(0.04), right=Inches(0.04),
                   top=Inches(0.02), bottom=Inches(0.02))

    # 2x2 체크박스 영역
    body_top = y + hdr_h + Inches(0.08)
    body_h = (y + h) - body_top - Inches(0.08)
    body_area = (x + Inches(0.12), body_top, w - Inches(0.24), body_h)
    grid = calc_grid(2, 2, area=body_area, gap=Inches(0.12))

    positions = [(0, 0), (0, 1), (1, 0), (1, 1)]
    for idx, (r, c) in enumerate(positions):
        cell = grid[r][c]
        # 체크박스 사각형
        cb_sz = Inches(0.22)
        cb_x = cell.left + Inches(0.04)
        cb_y = cell.top + (cell.height - cb_sz) // 2
        cb = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, cb_x, cb_y, cb_sz, cb_sz)
        cb.fill.solid()
        cb.fill.fore_color.rgb = C_WHITE
        cb.line.color.rgb = accent
        cb.line.width = Pt(1.25)

        # 라벨 텍스트
        label_x = cb_x + cb_sz + Inches(0.10)
        label_w = cell.left + cell.width - label_x
        lbl = add_textbox(
            slide,
            label_x,
            cell.top,
            label_w,
            cell.height,
            items[idx],
            font_size=11,
            color=C_INK,
            bold=False,
            align=PP_ALIGN.LEFT,
        )
        set_body_anchor(lbl, "ctr")
        set_text_inset(lbl, left=Inches(0.02), right=Inches(0.02),
                       top=Inches(0.02), bottom=Inches(0.02))


def build_slide_2(slide):
    set_title(slide, "핵심 요약")
    clear_placeholders(slide, keep=[0])

    # ---------- 레이아웃 영역 계산 ----------
    # 제목 플레이스홀더 아래부터 시작. CONTENT_SAFE 안쪽.
    safe_left = CONTENT_SAFE.left
    safe_top = CONTENT_SAFE.top
    safe_right = CONTENT_SAFE.right
    safe_bottom = CONTENT_SAFE.bottom
    safe_w = CONTENT_SAFE.width

    # 타이틀 공간은 placeholder(idx=0)가 차지하므로 본문은 더 아래에서 시작.
    # 제목 placeholder는 마스터에서 정의됨 — 보통 상단 ~0.6" 사용.
    # Body 시작 y (CONTENT_SAFE.top=0.68")
    body_top = safe_top  # Inches(0.68)

    # 1) Key Message Bar
    km_h = Inches(0.58)
    km_x = safe_left
    km_y = body_top
    km_w = safe_w
    _build_key_message_bar(slide, km_x, km_y, km_w, km_h)

    # 2) 3-card row
    gap_v = Inches(0.18)
    approval_h = Inches(1.70)
    cards_top = km_y + km_h + gap_v
    cards_bottom = safe_bottom - approval_h - gap_v
    cards_h = cards_bottom - cards_top

    cards_area = (safe_left, cards_top, safe_w, cards_h)
    cards_grid = calc_grid(1, 3, area=cards_area, gap=Inches(0.18))

    cards = [
        {
            "accent": C_GREEN,
            "icon": "\u2713",  # ✓
            "title": "경쟁사 이미 도입 중",
            "lines": [
                {"label": "Mercedes", "value": "5,000명+ 개발자 (공식)"},
                {"label": "BMW", "value": "사내 파일럿 · AMCIS 2024"},
                {"label": "Bosch · 현대모비스 · 현대오토에버", "value": ""},
                {"label": "2027.08", "value": "EU AI Act 자동차 D-Day"},
            ],
        },
        {
            "accent": C_BLUE,
            "icon": "\u25B3",  # △
            "title": "영역별 효과 차이 (업계 일반)",
            "lines": [
                {"label": "선행 R&D · 도구", "value": "+20~35%"},
                {"label": "양산 일반 SW", "value": "+10~20%"},
                {"label": "양산 안전 코드", "value": "±5% (본전)"},
                {"label": "→ 선별 적용이 핵심", "value": ""},
            ],
        },
        {
            "accent": C_RED,
            "icon": "\u26A0",  # ⚠
            "title": "품질검증 부담 증가",
            "lines": [
                {"label": "코드 리뷰 시간", "value": "+91% (Faros 2025)"},
                {"label": "코드 리뷰 요청", "value": "+98% (Faros 2025)"},
                {"label": "→ 인력·도구 동시 보강 필수", "value": ""},
                {"label": "(가장 자주 누락되는 투자)", "value": ""},
            ],
        },
    ]

    for i, card in enumerate(cards):
        cell = cards_grid[0][i]
        _build_card(
            slide,
            cell,
            accent=card["accent"],
            icon=card["icon"],
            title=card["title"],
            lines=card["lines"],
        )

    # 3) Approval Box (하단)
    ap_x = safe_left
    ap_y = cards_bottom + gap_v
    ap_w = safe_w
    ap_h = approval_h

    approval_items = [
        "선행 · 테스트 자동화 영역  도입 검토 착수",
        "자사 시범 사업 추진  (규모·예산 별도 산정)",
        "품질검증 인력 보강 계획  수립",
        "사내 AI 사용 가이드라인  제정",
    ]
    _build_approval_box(
        slide,
        ap_x,
        ap_y,
        ap_w,
        ap_h,
        accent=C_NAVY,
        title="\u25B6  즉시 결정 요청 사항",
        items=approval_items,
    )
