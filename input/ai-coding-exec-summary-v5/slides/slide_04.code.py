"""Slide 4 — 영역별 예상 효과 · 외부 공식 벤치마크

본문 y=0.68~6.55, 각주 y=6.60~7.00.
① 신호등 테이블 4행 (영역/효과/해석) — OVAL dot + 배경색
② 벤치마크 상세 카드 3개 (1x3)
각주 3줄 (9pt gray)
"""

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
C_NAVY = RGBColor(0x1F, 0x49, 0x7D)
C_BLUE = RGBColor(0x4F, 0x81, 0xBD)
C_GREEN = RGBColor(0x2E, 0x8B, 0x57)
C_YELLOW = RGBColor(0xD4, 0xA0, 0x17)
C_RED = RGBColor(0xC0, 0x50, 0x4D)
C_INK = RGBColor(0x1A, 0x1F, 0x2E)
C_GRAY = RGBColor(0x66, 0x66, 0x66)
C_LGRAY = RGBColor(0xE0, 0xE4, 0xEA)
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BG_GREEN = RGBColor(0xE8, 0xF5, 0xEC)
BG_YELLOW = RGBColor(0xFF, 0xF7, 0xD9)
BG_RED = RGBColor(0xFC, 0xE8, 0xE6)
C_FOOT = RGBColor(0x99, 0x99, 0x99)


def build_slide_4(slide):
    set_title(slide, "영역별 예상 효과  ·  외부 공식 벤치마크")
    clear_placeholders(slide, keep=[0])

    safe_left = int(CONTENT_SAFE.left)
    safe_width = int(CONTENT_SAFE.width)
    safe_right = int(CONTENT_SAFE.right)

    # ============================================================
    # Key Message Bar  y=0.68~1.00 (0.32")
    # ============================================================
    kmb_y = Inches(0.68)
    kmb_h = Inches(0.32)
    # 좌측 accent 세로 바
    add_accent_bar(
        slide,
        safe_left, kmb_y,
        Inches(0.06), kmb_h,
        C_NAVY,
    )
    # 배경 연한 회색 바
    kmb_bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        safe_left + Inches(0.06), kmb_y,
        safe_width - Inches(0.06), kmb_h,
    )
    kmb_bg.fill.solid()
    kmb_bg.fill.fore_color.rgb = RGBColor(0xF4, 0xF6, 0xFA)
    kmb_bg.line.fill.background()
    # 텍스트 (rich)
    kmb_tb = slide.shapes.add_textbox(
        safe_left + Inches(0.18), kmb_y,
        safe_width - Inches(0.24), kmb_h,
    )
    kmb_tf = kmb_tb.text_frame
    kmb_tf.word_wrap = True
    set_text_inset(kmb_tb, left=Inches(0.08), right=Inches(0.08),
                   top=Inches(0.04), bottom=Inches(0.04))
    set_body_anchor(kmb_tb, 'ctr')
    # 첫 단락 비우고 첫 para에 추가
    p0 = kmb_tf.paragraphs[0]
    p0.alignment = PP_ALIGN.LEFT
    r0 = p0.add_run()
    r0.text = "Key Message  |  "
    r0.font.size = Pt(11)
    r0.font.bold = True
    r0.font.color.rgb = C_NAVY
    r1 = p0.add_run()
    r1.text = "영역별 효과 "
    r1.font.size = Pt(11)
    r1.font.color.rgb = C_INK
    r2 = p0.add_run()
    r2.text = "3~5배 차이"
    r2.font.size = Pt(11)
    r2.font.bold = True
    r2.font.color.rgb = C_NAVY
    r3 = p0.add_run()
    r3.text = " · 선행·도구 영역이 "
    r3.font.size = Pt(11)
    r3.font.color.rgb = C_INK
    r4 = p0.add_run()
    r4.text = "즉시 도입 후보"
    r4.font.size = Pt(11)
    r4.font.bold = True
    r4.font.color.rgb = C_GREEN

    # ============================================================
    # ① 신호등 테이블 섹션  y=1.08~2.95
    # ============================================================
    sec1_title_y = Inches(1.08)
    # 섹션 ① 타이틀
    sec1_tb = add_textbox(
        slide,
        safe_left, sec1_title_y,
        safe_width, Inches(0.24),
        "① 영역별 효과 (업계 일반 보정 추정)",
        font_size=11, bold=True, color=C_NAVY,
    )

    # 헤더 바  y=1.34
    hdr_y = Inches(1.34)
    hdr_h = Inches(0.26)
    # 컬럼 레이아웃: dot(0.35) | 영역(3.2) | 효과(1.8) | 해석(remaining)
    col_dot_w = Inches(0.38)
    col_area_w = Inches(3.7)
    col_effect_w = Inches(1.9)
    col_interp_w = safe_width - col_dot_w - col_area_w - col_effect_w

    # 헤더 배경 (dot 컬럼은 비움)
    hdr_bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        safe_left + col_dot_w, hdr_y,
        safe_width - col_dot_w, hdr_h,
    )
    hdr_bg.fill.solid()
    hdr_bg.fill.fore_color.rgb = C_NAVY
    hdr_bg.line.fill.background()

    # 헤더 텍스트
    hx = safe_left + col_dot_w
    hdr_area = add_textbox(
        slide, hx + Inches(0.10), hdr_y,
        col_area_w - Inches(0.10), hdr_h,
        "영역", font_size=10, bold=True, color=C_WHITE,
        align=PP_ALIGN.LEFT,
    )
    set_body_anchor(hdr_area, 'ctr')
    hdr_eff = add_textbox(
        slide, hx + col_area_w, hdr_y,
        col_effect_w, hdr_h,
        "효과", font_size=10, bold=True, color=C_WHITE,
        align=PP_ALIGN.CENTER,
    )
    set_body_anchor(hdr_eff, 'ctr')
    hdr_int = add_textbox(
        slide, hx + col_area_w + col_effect_w + Inches(0.08), hdr_y,
        col_interp_w - Inches(0.08), hdr_h,
        "해석", font_size=10, bold=True, color=C_WHITE,
        align=PP_ALIGN.LEFT,
    )
    set_body_anchor(hdr_int, 'ctr')

    # 데이터 행 4개  y=1.62~2.94  → 각 행 0.33"
    rows_data = [
        {
            "dot": C_GREEN, "bg": BG_GREEN,
            "area": "선행 R&D · 도구 · 테스트 자동화",
            "effect": "+20 ~ +35%", "effect_color": C_GREEN,
            "interp": "즉시 도입 후보",
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

    row_y = Inches(1.62)
    row_h = Inches(0.33)
    row_gap = Inches(0.01)

    for row in rows_data:
        # 행 배경 (bg_color 연한 색)
        row_bg = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            safe_left + col_dot_w, row_y,
            safe_width - col_dot_w, row_h,
        )
        row_bg.fill.solid()
        row_bg.fill.fore_color.rgb = row["bg"]
        row_bg.line.color.rgb = C_LGRAY
        row_bg.line.width = Pt(0.25)

        # OVAL dot  (col_dot 중앙)
        dot_size = Inches(0.18)
        dot_x = safe_left + (col_dot_w - int(dot_size)) // 2
        dot_y = row_y + (int(row_h) - int(dot_size)) // 2
        make_icon_circle(
            slide, dot_x, dot_y, dot_size,
            row["dot"],
        )

        # 영역 (좌측 정렬)
        area_tb = add_textbox(
            slide, safe_left + col_dot_w + Inches(0.10), row_y,
            col_area_w - Inches(0.10), row_h,
            row["area"], font_size=10, color=C_INK,
            align=PP_ALIGN.LEFT,
        )
        set_body_anchor(area_tb, 'ctr')

        # 효과 (가운데 정렬, bold, 컬러)
        eff_tb = add_textbox(
            slide, safe_left + col_dot_w + col_area_w, row_y,
            col_effect_w, row_h,
            row["effect"], font_size=11, bold=True,
            color=row["effect_color"],
            align=PP_ALIGN.CENTER,
        )
        set_body_anchor(eff_tb, 'ctr')

        # 해석 (좌측 정렬)
        interp_tb = add_textbox(
            slide, safe_left + col_dot_w + col_area_w + col_effect_w + Inches(0.08), row_y,
            col_interp_w - Inches(0.08), row_h,
            row["interp"], font_size=10, color=C_INK,
            align=PP_ALIGN.LEFT,
        )
        set_body_anchor(interp_tb, 'ctr')

        row_y += int(row_h) + int(row_gap)

    # ============================================================
    # ② 벤치마크 상세 카드 섹션  y=3.06~6.55
    # ============================================================
    sec2_title_y = Inches(3.06)
    add_textbox(
        slide,
        safe_left, sec2_title_y,
        safe_width, Inches(0.24),
        "② 외부 공식 벤치마크 (상세)",
        font_size=11, bold=True, color=C_NAVY,
    )

    cards_area_top = Inches(3.34)
    cards_area_h = Inches(3.20)  # 3.34~6.54

    cards_data = [
        {
            "source": "Mercedes-Benz",
            "source_sub": "GitHub 공식 case study",
            "tool": "GitHub Copilot",
            "scale": "5,000명+ · 115k repos",
            "measurement": "개발자 자체 설문",
            "headline": "주당 30분+ 절감",
            "detail": "누적 200만 라인 수락",
            "accent": C_NAVY,
        },
        {
            "source": "BMW Group",
            "source_sub": "AMCIS 2024 논문",
            "tool": "GitHub Copilot (사내 파일럿)",
            "scale": "사내 파일럿",
            "measurement": "SPACE 프레임워크 5축",
            "headline": "5축 전항목 개선",
            "detail": "결함 감소 보고",
            "accent": C_NAVY,
        },
        {
            "source": "Faros AI (2025)",
            "source_sub": "다수 엔터프라이즈 분석",
            "tool": "Copilot / Cursor 등 혼합",
            "scale": "규모 비공개",
            "measurement": "Git 로그·리뷰 메트릭 분석",
            "headline": "리뷰 시간 +91% · 요청 +98%",
            "detail": "작업 처리량 +21%",
            "accent": C_BLUE,
        },
    ]

    grid = calc_grid(
        rows=1, cols=3,
        area=(safe_left, cards_area_top, safe_width, cards_area_h),
        gap=Inches(0.18),
    )

    for i, card in enumerate(cards_data):
        cell = grid[0][i]
        cx = cell.left
        cy = cell.top
        cw = cell.width
        ch = cell.height

        # 상단 accent 바
        bar_h = Inches(0.07)
        add_accent_bar(slide, cx, cy, cw, bar_h, card["accent"])

        # 카드 배경 (흰색) + 얇은 테두리
        bg = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            cx, cy + int(bar_h),
            cw, ch - int(bar_h),
        )
        bg.fill.solid()
        bg.fill.fore_color.rgb = C_WHITE
        bg.line.color.rgb = C_LGRAY
        bg.line.width = Pt(0.5)

        # 내부 레이아웃 (y 누적)
        pad_x = Inches(0.14)
        inner_x = cx + int(pad_x)
        inner_w = cw - 2 * int(pad_x)
        y = cy + int(bar_h) + Inches(0.10)

        # 출처 + 부제 (하나의 textbox, 2 paragraph)
        src_total_h = Inches(0.50)
        src_tb = slide.shapes.add_textbox(inner_x, y, inner_w, src_total_h)
        stf = src_tb.text_frame
        stf.word_wrap = True
        stf.margin_left = Inches(0)
        stf.margin_right = Inches(0)
        stf.margin_top = Inches(0)
        stf.margin_bottom = Inches(0)
        sp1 = stf.paragraphs[0]
        sp1.alignment = PP_ALIGN.LEFT
        sp1.space_after = Pt(2)
        sr1 = sp1.add_run()
        sr1.text = card["source"]
        sr1.font.size = Pt(12)
        sr1.font.bold = True
        sr1.font.color.rgb = card["accent"]
        sp2 = stf.add_paragraph()
        sp2.alignment = PP_ALIGN.LEFT
        sr2 = sp2.add_run()
        sr2.text = card["source_sub"]
        sr2.font.size = Pt(9)
        sr2.font.color.rgb = C_GRAY
        y += int(src_total_h) + Inches(0.02)

        # 구분선
        sep = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            inner_x, y, inner_w, Emu(9525),  # ~1pt
        )
        sep.fill.solid()
        sep.fill.fore_color.rgb = C_LGRAY
        sep.line.fill.background()
        y += Inches(0.08)

        # 메타 3줄: 도구 / 규모 / 측정 (하나의 textbox, 3 paragraph)
        meta_entries = [
            ("도구", card["tool"]),
            ("규모", card["scale"]),
            ("측정", card["measurement"]),
        ]
        meta_h = Inches(0.63)  # 0.21 x 3
        meta_tb = slide.shapes.add_textbox(inner_x, y, inner_w, meta_h)
        mtf = meta_tb.text_frame
        mtf.word_wrap = True
        mtf.margin_left = Inches(0)
        mtf.margin_right = Inches(0)
        mtf.margin_top = Inches(0.01)
        mtf.margin_bottom = Inches(0.01)
        for mi, (label, value) in enumerate(meta_entries):
            if mi == 0:
                mp = mtf.paragraphs[0]
            else:
                mp = mtf.add_paragraph()
            mp.alignment = PP_ALIGN.LEFT
            mp.space_before = Pt(0)
            mp.space_after = Pt(2)
            mr1 = mp.add_run()
            mr1.text = f"{label}  "
            mr1.font.size = Pt(9)
            mr1.font.bold = True
            mr1.font.color.rgb = C_GRAY
            mr2 = mp.add_run()
            mr2.text = value
            mr2.font.size = Pt(9)
            mr2.font.color.rgb = C_INK
        y += int(meta_h)

        # headline + detail (하나의 textbox)
        y += Inches(0.08)
        hl_total_h = Inches(0.70)
        hl_tb = slide.shapes.add_textbox(inner_x, y, inner_w, hl_total_h)
        htf = hl_tb.text_frame
        htf.word_wrap = True
        htf.margin_left = Inches(0)
        htf.margin_right = Inches(0)
        htf.margin_top = Inches(0)
        htf.margin_bottom = Inches(0)
        hp1 = htf.paragraphs[0]
        hp1.alignment = PP_ALIGN.LEFT
        hp1.space_after = Pt(3)
        hr1 = hp1.add_run()
        hr1.text = card["headline"]
        hr1.font.size = Pt(14)
        hr1.font.bold = True
        hr1.font.color.rgb = card["accent"]
        hp2 = htf.add_paragraph()
        hp2.alignment = PP_ALIGN.LEFT
        hr2 = hp2.add_run()
        hr2.text = card["detail"]
        hr2.font.size = Pt(9)
        hr2.font.color.rgb = C_GRAY

    # ============================================================
    # 각주 (3줄, 9pt gray) y=6.60~7.00
    # ============================================================
    foot_y = Inches(6.60)
    foot_h = Inches(0.40)

    foot_tb = slide.shapes.add_textbox(
        safe_left, foot_y, safe_width, foot_h,
    )
    ftf = foot_tb.text_frame
    ftf.word_wrap = True
    ftf.margin_left = Inches(0)
    ftf.margin_right = Inches(0)
    ftf.margin_top = Inches(0)
    ftf.margin_bottom = Inches(0)

    foot_lines = [
        "※ SPACE (Satisfaction·Performance·Activity·Communication·Efficiency): 개발자 생산성 5축",
        "※ Faros AI (2025): AI 도구 다수 엔터프라이즈 분석 리포트",
        "※ 차량 안전등급 (ASIL): ISO 26262의 Automotive Safety Integrity Level (A 낮음 → D 최고)",
    ]

    for i, line in enumerate(foot_lines):
        if i == 0:
            p = ftf.paragraphs[0]
        else:
            p = ftf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        p.line_spacing = 1.1
        run = p.add_run()
        run.text = line
        run.font.size = Pt(9)
        run.font.color.rgb = C_FOOT
