from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from ppt_utils import (
    set_title, clear_placeholders,
    add_textbox, add_para, add_rich_text,
    add_accent_bar, make_icon_circle,
    set_body_anchor, set_text_inset,
    calc_grid,
)
from template_contract import CONTENT_SAFE, LAYOUT_CONTENT


# ============================ 색상 팔레트 ============================
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


def build_slide_4(slide):
    set_title(slide, "영역별 예상 효과  ·  외부 공식 벤치마크")
    clear_placeholders(slide, keep=[0])

    # ============================ 영역 분할 ============================
    # safe_left = 0.28" , safe_right = 10.56" → width = 10.28"
    # safe_top  = 0.68" , safe_bottom = 7.02" → height = 6.34"
    safe_left = int(CONTENT_SAFE.left)
    safe_top = int(CONTENT_SAFE.top)
    safe_w = int(CONTENT_SAFE.width)
    safe_bot = int(CONTENT_SAFE.bottom)

    # ------------------------ 1) Key Message Bar ------------------------
    km_x = safe_left
    km_y = safe_top
    km_w = safe_w
    km_h = Inches(0.50)

    # 좌측 accent 세로 바
    add_accent_bar(slide, km_x, km_y, Inches(0.08), km_h, C_NAVY)

    # 메시지 박스 (연한 배경 + 볼드 본문)
    km_bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        km_x + Inches(0.08), km_y,
        km_w - Inches(0.08), km_h,
    )
    km_bg.fill.solid()
    km_bg.fill.fore_color.rgb = RGBColor(0xF2, 0xF5, 0xFA)
    km_bg.line.fill.background()

    km_tb = add_textbox(
        slide,
        km_x + Inches(0.20), km_y,
        km_w - Inches(0.20), km_h,
        "",
        font_size=12,
    )
    # 내용 초기화 후 rich text 작성
    km_tf = km_tb.text_frame
    km_tf.word_wrap = True
    # 기본 run 제거
    p0 = km_tf.paragraphs[0]
    for _r in list(p0.runs):
        _r._r.getparent().remove(_r._r)
    p0.alignment = PP_ALIGN.LEFT
    r1 = p0.add_run()
    r1.text = "Key Message  "
    r1.font.size = Pt(11)
    r1.font.bold = True
    r1.font.color.rgb = C_NAVY
    r2 = p0.add_run()
    r2.text = "영역별 효과는 3~5배 차이 · 선행·도구 영역이 즉시 도입 후보"
    r2.font.size = Pt(13)
    r2.font.bold = True
    r2.font.color.rgb = C_INK
    set_body_anchor(km_tb, 'ctr')
    set_text_inset(km_tb, left=Inches(0.12), right=Inches(0.10),
                   top=Inches(0.05), bottom=Inches(0.05))

    # ------------------------ 2) 신호등 테이블 ------------------------
    # 테이블 영역: Key Message 아래부터 시작
    tbl_top = km_y + int(km_h) + Inches(0.15)
    # 하단 카드 영역 확보 (카드 높이 약 1.55" + gap 0.18" + footnote 여유 0.00" = 1.73")
    # 카드 구간 시작 좌표를 먼저 정함
    card_h = Inches(1.55)
    card_gap = Inches(0.18)  # 테이블 <-> 카드 사이 gap
    card_top_est = safe_bot - int(card_h)
    tbl_bot = card_top_est - int(card_gap)
    tbl_h = tbl_bot - tbl_top

    # 섹션 소제목 (① 영역별 효과)
    sub1_h = Inches(0.26)
    add_textbox(
        slide,
        safe_left, tbl_top,
        safe_w, sub1_h,
        "① 영역별 효과  (업계 일반 보정 추정)",
        font_size=11, bold=True, color=C_NAVY, align=PP_ALIGN.LEFT,
    )

    # 테이블 본체 영역
    t_top = tbl_top + int(sub1_h) + Inches(0.02)
    t_h = tbl_h - int(sub1_h) - Inches(0.02)

    # 헤더 + 4행 = 5행. 헤더 비율 0.7, 데이터 행 각 1.0
    header_h = int(t_h * 0.18)
    row_h = (t_h - header_h) // 4

    # 열 비율
    # dot(0.45") | 영역(3.5) | 효과(1.8) | 해석(4.0) — total 약 9.75, dot은 고정
    dot_w = Inches(0.45)
    remaining = safe_w - int(dot_w)
    # area : effect : interp = 32 : 18 : 40
    area_w = int(remaining * 32 / 90)
    effect_w = int(remaining * 18 / 90)
    interp_w = remaining - area_w - effect_w

    col_x_dot = safe_left
    col_x_area = col_x_dot + int(dot_w)
    col_x_effect = col_x_area + area_w
    col_x_interp = col_x_effect + effect_w

    # ---- 헤더 행 ----
    header_y = t_top
    # 헤더 배경
    hd_bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        safe_left, header_y,
        safe_w, header_h,
    )
    hd_bg.fill.solid()
    hd_bg.fill.fore_color.rgb = C_NAVY
    hd_bg.line.fill.background()

    def _header_label(text, x, w):
        tb = add_textbox(
            slide, x, header_y, w, header_h,
            text, font_size=11, bold=True, color=C_WHITE,
            align=PP_ALIGN.CENTER if x != col_x_area else PP_ALIGN.LEFT,
        )
        set_body_anchor(tb, 'ctr')
        set_text_inset(tb, left=Inches(0.10), right=Inches(0.06),
                       top=Inches(0.02), bottom=Inches(0.02))
        return tb

    _header_label("", col_x_dot, int(dot_w))
    _header_label("영역", col_x_area, area_w)
    _header_label("효과", col_x_effect, effect_w)
    _header_label("해석", col_x_interp, interp_w)

    # ---- 4개 데이터 행 ----
    rows_data = [
        {
            "dot": C_GREEN, "bg": BG_GREEN,
            "area": "선행 R&D · 도구 · 테스트 자동화",
            "effect": "+20 ~ +35%", "effect_color": C_GREEN,
            "interp": "즉시 도입 후보 (위험 낮음 · 효과 큼)",
        },
        {
            "dot": C_GREEN, "bg": BG_GREEN,
            "area": "양산 일반 SW (비안전)",
            "effect": "+10 ~ +20%", "effect_color": C_GREEN,
            "interp": "코드 검사 도구 통합 후 도입",
        },
        {
            "dot": C_YELLOW, "bg": BG_YELLOW,
            "area": "양산 안전 코드 (차량 안전등급 높음↑)",
            "effect": "-5 ~ +10%", "effect_color": C_YELLOW,
            "interp": "본전 영역 · 자사 시범 측정 필요",
        },
        {
            "dot": C_RED, "bg": BG_RED,
            "area": "품질검증 단계 (인력 관점)",
            "effect": "-15 ~ -30%", "effect_color": C_RED,
            "interp": "코드 리뷰 부담 증가 · 인력 보강 필수",
        },
    ]

    for i, rd in enumerate(rows_data):
        ry = header_y + header_h + row_h * i
        # 행 배경
        row_bg = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            safe_left, ry, safe_w, row_h,
        )
        row_bg.fill.solid()
        row_bg.fill.fore_color.rgb = rd["bg"]
        row_bg.line.color.rgb = C_LGRAY
        row_bg.line.width = Pt(0.5)

        # dot (OVAL)
        dot_size = Inches(0.22)
        dot_cx = col_x_dot + int(dot_w) // 2 - int(dot_size) // 2
        dot_cy = ry + row_h // 2 - int(dot_size) // 2
        make_icon_circle(
            slide, dot_cx, dot_cy, dot_size,
            fill_color=rd["dot"], text="",
        )

        # 영역 텍스트
        area_tb = add_textbox(
            slide, col_x_area, ry, area_w, row_h,
            rd["area"], font_size=11, bold=True, color=C_INK,
            align=PP_ALIGN.LEFT,
        )
        set_body_anchor(area_tb, 'ctr')
        set_text_inset(area_tb, left=Inches(0.08), right=Inches(0.06),
                       top=Inches(0.02), bottom=Inches(0.02))

        # 효과 텍스트 (bold accent)
        effect_tb = add_textbox(
            slide, col_x_effect, ry, effect_w, row_h,
            rd["effect"], font_size=14, bold=True, color=rd["effect_color"],
            align=PP_ALIGN.CENTER,
        )
        set_body_anchor(effect_tb, 'ctr')
        set_text_inset(effect_tb, left=Inches(0.04), right=Inches(0.04),
                       top=Inches(0.02), bottom=Inches(0.02))

        # 해석 텍스트
        interp_tb = add_textbox(
            slide, col_x_interp, ry, interp_w, row_h,
            rd["interp"], font_size=10, bold=False, color=C_INK,
            align=PP_ALIGN.LEFT,
        )
        set_body_anchor(interp_tb, 'ctr')
        set_text_inset(interp_tb, left=Inches(0.10), right=Inches(0.08),
                       top=Inches(0.02), bottom=Inches(0.02))

    # ------------------------ 3) 외부 공식 출처 카드 (1x3) ------------------------
    card_top = card_top_est
    # 섹션 소제목은 공간이 부족하므로 카드 상단에 살짝 얹기 대신 생략하고,
    # 대신 각 카드가 "외부 공식 벤치마크" 성격을 드러내도록 구성.
    # 그러나 요구사항 충족을 위해 카드 구간 위에 짧은 소제목 라인을 얹는다.
    # card_top_est는 safe_bot - card_h 이고, card_gap이 0.18"이므로
    # card_top 바로 위 0.18" gap 안에 소제목을 넣기에는 빡빡 → 카드 내부 상단 대신
    # 카드들과 간섭 없이 그려지도록 bottom-aligned 카드 영역 유지.

    # 3개 카드 grid
    grid = calc_grid(
        rows=1, cols=3,
        area=(safe_left, card_top, safe_w, int(card_h)),
        gap=Inches(0.18),
    )

    cards_data = [
        {
            "source": "Mercedes-Benz",
            "source_sub": "GitHub 공식 case study (2023~)",
            "value": "주당 30분+ 절감",
            "detail": "5,000명+ 개발자 · 누적 200만 라인 수락",
            "accent": C_NAVY,
        },
        {
            "source": "BMW Group",
            "source_sub": "AMCIS 2024 논문",
            "value": "SPACE 5축 전항목 개선",
            "detail": "결함 감소 보고 (정량치 비공개)",
            "accent": C_NAVY,
        },
        {
            "source": "Faros AI (2025)",
            "source_sub": "다수 엔터프라이즈 분석",
            "value": "리뷰 시간 +91% · 요청 +98%",
            "detail": "AI 생성 코드의 검증 부담 증가 근거",
            "accent": C_BLUE,
        },
    ]

    for i, cd in enumerate(cards_data):
        cell = grid[0][i]
        cx = cell.left
        cy = cell.top
        cw = cell.width
        ch = cell.height

        # 카드 배경 (화이트 + 미세 테두리)
        card_bg = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            cx, cy, cw, ch,
        )
        card_bg.fill.solid()
        card_bg.fill.fore_color.rgb = C_WHITE
        card_bg.line.color.rgb = C_LGRAY
        card_bg.line.width = Pt(0.75)

        # 좌측 accent bar
        add_accent_bar(slide, cx, cy, Inches(0.08), ch, cd["accent"])

        # 내부 패딩 기준
        inner_x = cx + Inches(0.22)
        inner_w = cw - Inches(0.32)

        # 3단 분할: 상단(출처 + sub) / 중앙(값) / 하단(상세)
        top_h = int(ch * 0.34)
        mid_h = int(ch * 0.36)
        bot_h = ch - top_h - mid_h

        top_y = cy + Inches(0.08)
        top_h_use = top_h - Inches(0.08)

        # 상단: 출처명 (bold 14) + sub (9, gray)
        src_tb = add_textbox(
            slide, inner_x, top_y, inner_w, top_h_use,
            "", font_size=10,
        )
        src_tf = src_tb.text_frame
        src_tf.word_wrap = True
        # 기존 기본 run 제거
        p0 = src_tf.paragraphs[0]
        for _r in list(p0.runs):
            _r._r.getparent().remove(_r._r)
        p0.alignment = PP_ALIGN.LEFT
        rs = p0.add_run()
        rs.text = cd["source"]
        rs.font.size = Pt(14)
        rs.font.bold = True
        rs.font.color.rgb = cd["accent"]
        add_para(
            src_tf, cd["source_sub"],
            font_size=9, color=C_GRAY, align=PP_ALIGN.LEFT,
            space_before=Pt(2), space_after=Pt(0),
        )
        set_body_anchor(src_tb, 't')
        set_text_inset(src_tb, left=Inches(0.02), right=Inches(0.02),
                       top=Inches(0.02), bottom=Inches(0.02))

        # 중앙: 핵심 수치 (bold 14 accent, center)
        mid_y = cy + top_h
        val_tb = add_textbox(
            slide, inner_x, mid_y, inner_w, mid_h,
            cd["value"], font_size=14, bold=True, color=cd["accent"],
            align=PP_ALIGN.CENTER,
        )
        set_body_anchor(val_tb, 'ctr')
        set_text_inset(val_tb, left=Inches(0.04), right=Inches(0.04),
                       top=Inches(0.02), bottom=Inches(0.02))

        # 하단: 상세 설명 (10pt)
        bot_y = cy + top_h + mid_h
        detail_tb = add_textbox(
            slide, inner_x, bot_y, inner_w, bot_h,
            cd["detail"], font_size=10, color=C_INK,
            align=PP_ALIGN.LEFT,
        )
        set_body_anchor(detail_tb, 't')
        set_text_inset(detail_tb, left=Inches(0.02), right=Inches(0.02),
                       top=Inches(0.02), bottom=Inches(0.04))
