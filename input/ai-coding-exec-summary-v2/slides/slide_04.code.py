"""Slide 4 — 실행 로드맵 및 승인 요청 사항

상단 Key Message Bar -> 가로 CHEVRON 3단계 로드맵 -> 하단 2열 분할
(좌: 시범사업 정보 카드, 우: 승인 요청 번호 리스트).
"""

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from ppt_utils import (
    set_title, clear_placeholders,
    add_textbox, add_para, add_rich_text,
    add_accent_bar, make_icon_circle, make_icon_badge,
    add_shadow, set_shape_opacity,
    calc_grid, set_body_anchor,
)
from template_contract import CONTENT_SAFE, LAYOUT_CONTENT


# ---- 색상 팔레트 ----
C_NAVY   = RGBColor(0x1F, 0x49, 0x7D)
C_BLUE   = RGBColor(0x4F, 0x81, 0xBD)
C_GREEN  = RGBColor(0x2E, 0x8B, 0x57)
C_TEAL   = RGBColor(0x2C, 0x7F, 0x94)
C_INK    = RGBColor(0x1A, 0x1F, 0x2E)
C_GRAY   = RGBColor(0x66, 0x66, 0x66)
C_LGRAY  = RGBColor(0xE0, 0xE4, 0xEA)
C_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
BG_BLUE  = RGBColor(0xE7, 0xF0, 0xFA)
BG_TEAL  = RGBColor(0xE3, 0xF2, 0xF7)


def build_slide_4(slide):
    set_title(slide, "실행 로드맵 및 승인 요청 사항")
    clear_placeholders(slide, keep=[0])

    # =========================================================
    # 1) Key Message Bar (상단)
    # =========================================================
    km_left = Inches(0.30)
    km_top = Inches(0.78)
    km_w = Inches(10.20)
    km_h = Inches(0.46)

    # 좌측 accent bar
    add_accent_bar(slide, km_left, km_top, Inches(0.08), km_h, C_TEAL)

    # 배경 박스
    km_bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        km_left + Inches(0.08), km_top,
        km_w - Inches(0.08), km_h,
    )
    km_bg.fill.solid()
    km_bg.fill.fore_color.rgb = BG_TEAL
    km_bg.line.fill.background()

    # Key Message 텍스트
    km_tb = add_textbox(
        slide,
        km_left + Inches(0.22), km_top + Inches(0.02),
        km_w - Inches(0.30), km_h - Inches(0.04),
        "",
        font_size=12,
        color=C_INK,
        bold=False,
        align=PP_ALIGN.LEFT,
    )
    # 기존 단락 초기화 후 rich_text로 구성
    km_tf = km_tb.text_frame
    km_tf.word_wrap = True
    # 첫 단락에 덮어쓰기 위해 기존 runs 제거
    p0 = km_tf.paragraphs[0]
    p0.alignment = PP_ALIGN.LEFT
    # 기존 run 제거
    for r in list(p0.runs):
        r._r.getparent().remove(r._r)
    # 라벨 run
    run_label = p0.add_run()
    run_label.text = "Key Message   "
    run_label.font.size = Pt(11)
    run_label.font.bold = True
    run_label.font.color.rgb = C_TEAL
    # 메시지 run
    run_msg = p0.add_run()
    run_msg.text = "3단계 도입 + 4개월 시범 사업으로 리스크 최소화 · 선행 영역 즉시 착수"
    run_msg.font.size = Pt(12)
    run_msg.font.bold = True
    run_msg.font.color.rgb = C_INK
    set_body_anchor(km_bg, "ctr")

    # =========================================================
    # 2) CHEVRON 3단계 로드맵
    # =========================================================
    rm_top = Inches(1.42)
    rm_h = Inches(0.90)
    rm_left = Inches(0.30)
    rm_total_w = Inches(10.20)

    phases = [
        {
            "title": "Phase 1  ·  0~3개월",
            "body_l1": "도구 · 테스트 자동화 파일럿",
            "body_l2": "위험 낮음 · 효과 큼",
            "color": C_GREEN,
        },
        {
            "title": "Phase 2  ·  3~9개월",
            "body_l1": "선행 R&D · 응용SW 확대",
            "body_l2": "사내 AI 학습 · 자동 검사 통합",
            "color": C_BLUE,
        },
        {
            "title": "Phase 3  ·  9~18개월",
            "body_l1": "양산 비안전 영역 단계 적용",
            "body_l2": "양산 안전은 시범 측정 후",
            "color": C_NAVY,
        },
    ]

    overlap = Inches(0.10)
    n = len(phases)
    # 전체 폭 = n * chev_w - (n-1) * overlap  =>  chev_w = (total + (n-1)*overlap) / n
    chev_w = int((rm_total_w + overlap * (n - 1)) / n)

    for i, ph in enumerate(phases):
        cx = rm_left + (chev_w - overlap) * i
        chev = slide.shapes.add_shape(
            MSO_SHAPE.CHEVRON,
            cx, rm_top, chev_w, rm_h,
        )
        chev.fill.solid()
        chev.fill.fore_color.rgb = ph["color"]
        chev.line.fill.background()

        # 텍스트 (2단: 상단 title bold 10pt, 하단 body 9pt 2줄)
        tf = chev.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.35) if i > 0 else Inches(0.18)
        tf.margin_right = Inches(0.35)
        tf.margin_top = Inches(0.08)
        tf.margin_bottom = Inches(0.06)

        # 첫 단락: title
        p0 = tf.paragraphs[0]
        p0.alignment = PP_ALIGN.CENTER
        for r in list(p0.runs):
            r._r.getparent().remove(r._r)
        run_t = p0.add_run()
        run_t.text = ph["title"]
        run_t.font.size = Pt(11)
        run_t.font.bold = True
        run_t.font.color.rgb = C_WHITE

        # 본문 라인 1
        add_para(
            tf, ph["body_l1"],
            font_size=9, color=C_WHITE, bold=False,
            align=PP_ALIGN.CENTER, space_before=Pt(2),
        )
        # 본문 라인 2
        add_para(
            tf, ph["body_l2"],
            font_size=9, color=C_WHITE, bold=False,
            align=PP_ALIGN.CENTER,
        )
        set_body_anchor(chev, "ctr")

    # =========================================================
    # 3) 하단 2열 분할 — 정보 카드(좌) / 승인 요청(우)
    # =========================================================
    bottom_top = Inches(2.52)
    bottom_h = Inches(4.40)   # 2.52 + 4.40 = 6.92  (≤ 7.02)
    bottom_area = (
        Inches(0.30),
        bottom_top,
        Inches(10.20),
        bottom_h,
    )
    grid = calc_grid(1, 2, area=bottom_area, gap=Inches(0.22))
    left_cell = grid[0][0]
    right_cell = grid[0][1]

    # ----------- 좌측: 시범사업 정보 카드 -----------
    _draw_info_card(
        slide,
        left=left_cell.left,
        top=left_cell.top,
        width=left_cell.width,
        height=left_cell.height,
        accent=C_BLUE,
        bg=BG_BLUE,
        title="② 4개월 시범 사업 제안",
        items=[
            ("대상", "SW 모듈 4개  (선행 1 · 양산 일반 1 · 양산 안전 1 · 검증 자동화 1)"),
            ("투입", "20명 × 4개월  (숙련도·경력 균형 배정)"),
            ("예산", "약 3천만원  (연간 절감가치의 2%)"),
            ("측정", "공수 · 결함밀도 · 코드 리뷰 부담"),
            ("목적", "양산 안전 영역 효과 실증  (가장 불확실한 영역)"),
        ],
    )

    # ----------- 우측: 승인 요청 카드 -----------
    _draw_approval_card(
        slide,
        left=right_cell.left,
        top=right_cell.top,
        width=right_cell.width,
        height=right_cell.height,
        accent=C_TEAL,
        bg=BG_TEAL,
        title="③ 승인 요청 사항",
        items=[
            ("①", "선행 · 테스트 자동화 영역 즉시 도입", "추정 ROI 11배"),
            ("②", "4개월 시범 사업 예산  3천만원 승인", "의사결정 근거 확보"),
            ("③", "품질검증 인력 보강 계획 수립", "도입 효과 유지"),
            ("④", "사내 AI 사용 가이드라인 제정 착수", "EU AI Act 대응"),
        ],
    )


# =============================================================
# 내부 헬퍼: 정보 카드 (좌측)
# =============================================================
def _draw_info_card(slide, left, top, width, height,
                    accent, bg, title, items):
    # 카드 배경
    card_bg = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        left, top, width, height,
    )
    card_bg.adjustments[0] = 0.04
    card_bg.fill.solid()
    card_bg.fill.fore_color.rgb = bg
    card_bg.line.fill.background()

    # 좌측 세로 accent bar
    add_accent_bar(slide, left, top, Inches(0.08), height, accent)

    # 제목 영역
    title_h = Inches(0.44)
    title_left = left + Inches(0.20)
    title_w = width - Inches(0.32)
    add_textbox(
        slide,
        title_left, top + Inches(0.08),
        title_w, title_h,
        title,
        font_size=13,
        color=accent,
        bold=True,
        align=PP_ALIGN.LEFT,
    )

    # 타이틀 아래 얇은 디바이더
    div_y = top + Inches(0.52)
    add_accent_bar(
        slide,
        title_left, div_y,
        title_w, Inches(0.015),
        RGBColor(0xC8, 0xD5, 0xE5),
    )

    # 항목 리스트
    items_top = div_y + Inches(0.14)
    items_area_h = height - (items_top - top) - Inches(0.14)
    row_h = int(items_area_h / max(len(items), 1))

    key_col_w = Inches(0.78)
    gap_kv = Inches(0.08)

    for i, (k, v) in enumerate(items):
        row_y = items_top + row_h * i
        inner_h = row_h - Inches(0.02)

        # 키 배지
        key_shape = make_icon_badge(
            slide,
            title_left, row_y + Inches(0.02),
            key_col_w, Inches(0.30),
            k,
            fill_color=accent,
            font_size=10,
            font_color=C_WHITE,
            corner_radius=0.30,
        )

        # 값 텍스트
        val_left = title_left + key_col_w + gap_kv
        val_w = title_w - key_col_w - gap_kv
        val_tb = add_textbox(
            slide,
            val_left, row_y,
            val_w, inner_h,
            v,
            font_size=10,
            color=C_INK,
            bold=False,
            align=PP_ALIGN.LEFT,
        )
        set_body_anchor(val_tb, "ctr")


# =============================================================
# 내부 헬퍼: 승인 요청 카드 (우측)
# =============================================================
def _draw_approval_card(slide, left, top, width, height,
                        accent, bg, title, items):
    # 카드 배경
    card_bg = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        left, top, width, height,
    )
    card_bg.adjustments[0] = 0.04
    card_bg.fill.solid()
    card_bg.fill.fore_color.rgb = bg
    card_bg.line.fill.background()

    # 좌측 세로 accent bar
    add_accent_bar(slide, left, top, Inches(0.08), height, accent)

    # 제목
    title_left = left + Inches(0.20)
    title_w = width - Inches(0.32)
    add_textbox(
        slide,
        title_left, top + Inches(0.08),
        title_w, Inches(0.44),
        title,
        font_size=13,
        color=accent,
        bold=True,
        align=PP_ALIGN.LEFT,
    )

    # 디바이더
    div_y = top + Inches(0.52)
    add_accent_bar(
        slide,
        title_left, div_y,
        title_w, Inches(0.015),
        RGBColor(0xB8, 0xD6, 0xDF),
    )

    # 항목
    items_top = div_y + Inches(0.14)
    items_area_h = height - (items_top - top) - Inches(0.14)
    row_h = int(items_area_h / max(len(items), 1))

    circle_size = Inches(0.44)
    gap_cn = Inches(0.14)

    for i, (num, label, note) in enumerate(items):
        row_y = items_top + row_h * i

        # 번호 원
        make_icon_circle(
            slide,
            title_left, row_y + Inches(0.04),
            circle_size,
            fill_color=accent,
            text=num,
            font_size=13,
            font_color=C_WHITE,
        )

        # 라벨 + 주석 (오른쪽)
        text_left = title_left + circle_size + gap_cn
        text_w = title_w - circle_size - gap_cn

        # 라벨 (bold)
        label_h = Inches(0.28)
        add_textbox(
            slide,
            text_left, row_y,
            text_w, label_h,
            label,
            font_size=11,
            color=C_INK,
            bold=True,
            align=PP_ALIGN.LEFT,
        )
        # note (작은 회색)
        add_textbox(
            slide,
            text_left, row_y + label_h - Inches(0.02),
            text_w, Inches(0.24),
            note,
            font_size=9,
            color=C_GRAY,
            bold=False,
            align=PP_ALIGN.LEFT,
        )
