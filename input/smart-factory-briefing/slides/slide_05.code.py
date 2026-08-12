"""slide_05 — AI 활용 전략 — 3단계 전환 (smart-factory-briefing).

spec: input/smart-factory-briefing/slides/slide_05.spec.yaml
design: input/smart-factory-briefing/design_system.md (맑은 고딕 단일 — force_font,
        단일 accent #024AD8, 이모지·add_shadow·CHEVRON 금지, hairline 보더)
구성: 리드 / 상단 3단계 전환 블록(ROUNDED_RECTANGLE + STRAIGHT 화살표) /
      하단 콜아웃 2열(좌 GPU 비용 grid_table · 우 세로 미니 플로우) /
      다크 슬랩 결론 밴드 / ※ 각주(지식증류)
"""

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR

from ppt_utils import (
    set_title, clear_placeholders,
    add_para, add_rich_text, add_accent_bar,
    add_arrowhead, add_smart_connector, align_shapes,
    add_footnote, set_text_inset, set_body_anchor,
    calc_grid, VFlow, add_grid_table, force_font,
)
from template_contract import CONTENT_SAFE


# ---------------- 디자인 토큰 (design_system.md — 이 목록 밖 색 사용 금지) ----------------
PRIMARY        = RGBColor(0x02, 0x4A, 0xD8)  # Electric Blue — 유일한 신호색
PRIMARY_BRIGHT = RGBColor(0x29, 0x6E, 0xF9)  # 다크 슬랩 위 강조 전용
PRIMARY_SOFT   = RGBColor(0xC9, 0xE0, 0xFC)  # 강조 블록 배경
INK            = RGBColor(0x1A, 0x1A, 0x1A)
CHARCOAL       = RGBColor(0x3D, 0x3D, 0x3D)
GRAPHITE       = RGBColor(0x63, 0x63, 0x63)
CANVAS         = RGBColor(0xFF, 0xFF, 0xFF)
CLOUD          = RGBColor(0xF7, 0xF7, 0xF7)
FOG            = RGBColor(0xE8, 0xE8, 0xE8)
STEEL          = RGBColor(0xC2, 0xC2, 0xC2)
ON_INK         = RGBColor(0xFF, 0xFF, 0xFF)

F_SEMI = "Pretendard SemiBold"   # font_name 자체가 SemiBold → bold=False
F_REG  = "Pretendard"


# ---------------- 로컬 헬퍼 (ppt_utils 재구현 아님 — 조합만) ----------------

def _strip_lead_para(tf):
    """add_rich_text 는 항상 새 단락을 추가하므로, 최초의 빈 단락을 제거한다."""
    first = tf.paragraphs[0]
    if not first.runs and len(tf.paragraphs) > 1:
        first._p.getparent().remove(first._p)


def _phase_block(slide, cell, when, name, body, current=False):
    """3단계 전환 블록 1개 — ROUNDED_RECTANGLE(소 radius) + hairline 보더.

    상단 라벨(시기 9pt GRAPHITE) + 단계명(SemiBold 12.5pt INK) + 요약(10.5pt CHARCOAL).
    현재 단계만 PRIMARY_SOFT 배경 + PRIMARY 1pt 보더.
    """
    shp = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, cell.left, cell.top, cell.width, cell.height)
    shp.adjustments[0] = 0.08
    shp.fill.solid()
    shp.fill.fore_color.rgb = PRIMARY_SOFT if current else CANVAS
    shp.line.color.rgb = PRIMARY if current else FOG
    shp.line.width = Pt(1)
    set_text_inset(shp, left=Inches(0.15), right=Inches(0.15),
                   top=Inches(0.10), bottom=Inches(0.10))
    set_body_anchor(shp, "ctr")

    tf = shp.text_frame
    tf.word_wrap = True
    p0 = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.LEFT
    run = p0.add_run()
    run.text = when
    run.font.name = F_REG
    run.font.size = Pt(9)
    run.font.color.rgb = GRAPHITE

    p1 = add_para(tf, name, font_name=F_SEMI, font_size=12.5, color=INK,
                  space_before=Pt(3), space_after=Pt(2))
    p1.line_spacing = Pt(12.5 * 1.35)
    for line in body.split("\n"):
        p2 = add_para(tf, line, font_name=F_REG, font_size=10.5,
                      color=CHARCOAL, space_before=Pt(1))
        p2.line_spacing = Pt(10.5 * 1.35)
    return shp


def _callout_shell(slide, cell, fill, border_color, divider_color, header_text):
    """콜아웃 카드 셸 — 카드 배경 + 헤더(SemiBold 12pt INK) + hairline 구분선.

    본문 배치용 내부 VFlow 를 반환한다 (카드마다 본문 구성이 다름).
    """
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, cell.left, cell.top, cell.width, cell.height)
    card.adjustments[0] = 0.05
    card.fill.solid()
    card.fill.fore_color.rgb = fill
    card.line.color.rgb = border_color
    card.line.width = Pt(1)

    inset = Inches(0.18)
    inner = VFlow(x=cell.left + inset, w=cell.width - 2 * inset,
                  y_start=cell.top + Inches(0.16),
                  y_max=cell.top + cell.height - Inches(0.14),
                  gap=Inches(0.08))
    header_h = Inches(0.28) + Inches(0.24) * header_text.count("\n")
    inner.textbox(slide, header_text, font_size=12, h=header_h,
                  font_name=F_SEMI, color=INK)
    bar = inner.reserve(Inches(0.014), gap=Inches(0.10))
    add_accent_bar(slide, bar.left, bar.top, bar.width, bar.height, divider_color)
    return card, inner


def _flow_block(slide, rect, marker, label, body):
    """세로 미니 플로우 블록 — CANVAS + FOG hairline, 라벨 1줄 + 본문 1줄."""
    shp = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, rect.left, rect.top, rect.width, rect.height)
    shp.adjustments[0] = 0.10
    shp.fill.solid()
    shp.fill.fore_color.rgb = CANVAS
    shp.line.color.rgb = FOG
    shp.line.width = Pt(1)
    set_text_inset(shp, left=Inches(0.14), right=Inches(0.14),
                   top=Inches(0.06), bottom=Inches(0.06))
    set_body_anchor(shp, "ctr")

    tf = shp.text_frame
    tf.word_wrap = True
    p0 = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.LEFT
    r0 = p0.add_run()
    r0.text = marker
    r0.font.name = F_SEMI
    r0.font.size = Pt(10.5)
    r0.font.color.rgb = PRIMARY
    r1 = p0.add_run()
    r1.text = label
    r1.font.name = F_SEMI
    r1.font.size = Pt(10.5)
    r1.font.color.rgb = INK

    p1 = add_para(tf, body, font_name=F_REG, font_size=10,
                  color=CHARCOAL, space_before=Pt(2))
    p1.line_spacing = Pt(10 * 1.3)
    return shp


# ---------------- 슬라이드 빌드 ----------------

def build_slide_05(slide):
    set_title(slide, "AI 활용 전략 — 3단계 전환", font_size=20, bold=True)
    clear_placeholders(slide, keep=[0])

    # 본문 콘텐츠는 리드 아래 y ≥ 1.18" 부터 시작 (design_system 2-1)
    flow = VFlow(x=CONTENT_SAFE.left, w=CONTENT_SAFE.width,
                 y_start=Inches(1.18), y_max=CONTENT_SAFE.bottom)

    # --- 리드 문장 (표준 규격 고정: y=0.72, h=0.35 — 15pt bold INK, 키워드만 PRIMARY) ---
    lead_box = slide.shapes.add_textbox(
        CONTENT_SAFE.left, Inches(0.72), CONTENT_SAFE.width, Inches(0.35))
    lead_tf = lead_box.text_frame
    lead_tf.word_wrap = True
    add_rich_text(lead_tf, [
        {"text": "로컬 sLLM 출발 → ", "font_size": 15,
         "color": INK, "bold": True},
        {"text": "실측 한계", "font_size": 15,
         "color": PRIMARY, "bold": True},
        {"text": "로 상용 LLM 전환 → 지식증류로 상용 + 로컬 ",
         "font_size": 15, "color": INK, "bold": True},
        {"text": "병행", "font_size": 15,
         "color": PRIMARY, "bold": True},
    ], line_spacing=Pt(15 * 1.35))
    _strip_lead_para(lead_tf)

    # --- 상단 3단계 전환 블록 (블록 3개 + PRIMARY 1.5pt STRAIGHT 화살표 2개) ---
    band = flow.reserve(Inches(1.52), gap=Inches(0.18))
    pg = calc_grid(1, 3, area=band, gap=Inches(0.64))
    phases = [
        ("2026.04", "초기", "코드 유출 차단 · 비용 0원 기대\n로컬 sLLM 채택", False),
        ("2026.07", "중간 — 현재", "로컬 sLLM 실측 한계 확인\n상용 LLM 전환", True),
        ("2026 하반기", "최종 목표", "3축 통합 AI Agent\n문서 · 분석은 상용 + 로컬 병행", False),
    ]
    blocks = [
        _phase_block(slide, pg[0][i], when, name, body, current=cur)
        for i, (when, name, body, cur) in enumerate(phases)
    ]
    align_shapes(*blocks, axis="h")  # 수평 정렬 후 연결 → 완벽한 수평 화살표

    for a, b in zip(blocks, blocks[1:]):
        conn = add_smart_connector(slide, a, b,
                                   connector_type=MSO_CONNECTOR.STRAIGHT,
                                   direction="LR", arrow=False)
        conn.line.color.rgb = PRIMARY
        conn.line.width = Pt(1.5)
        add_arrowhead(conn)

    # '현재' 태그 — 현재 단계 블록 우상단 소형 pill (좁은 면적만 PRIMARY fill)
    cur_blk = blocks[1]
    tag_w, tag_h = Inches(0.58), Inches(0.22)
    tag = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        cur_blk.left + cur_blk.width - tag_w - Inches(0.10),
        cur_blk.top + Inches(0.09), tag_w, tag_h)
    tag.adjustments[0] = 0.5
    tag.fill.solid()
    tag.fill.fore_color.rgb = PRIMARY
    tag.line.fill.background()
    set_text_inset(tag, left=Inches(0.02), top=Inches(0.0),
                   right=Inches(0.02), bottom=Inches(0.0))
    tag_p = tag.text_frame.paragraphs[0]
    tag_p.alignment = PP_ALIGN.CENTER
    tag_run = tag_p.add_run()
    tag_run.text = "현재"
    tag_run.font.name = F_SEMI
    tag_run.font.size = Pt(8)
    tag_run.font.color.rgb = ON_INK
    set_body_anchor(tag, "ctr")

    # --- 하단 콜아웃 2열 + 다크 슬랩 + 각주 공간 배분 ---
    slab_h = Inches(0.44)
    footnote_zone = Inches(0.38)   # add_footnote 자동 배치 영역 (9pt 1줄)
    callout_h = flow.remaining - int(slab_h) - int(footnote_zone) - int(Inches(0.14))
    row = flow.reserve(callout_h, gap=Inches(0.14))
    cg = calc_grid(1, 2, area=row, gap=Inches(0.30))

    # 좌측 카드: 실측 결론 1줄 + GPU 비용 grid_table(모델/필요 GPU/총 비용) + 기준 캡션
    _, lf = _callout_shell(
        slide, cg[0][0], fill=CANVAS, border_color=STEEL, divider_color=FOG,
        header_text="로컬 sLLM 의 한계 — HW 증축 비현실적")
    lf.textbox(slide, "▸ 상용 수준 성능 확보에는 프론티어급 모델 자체 구동 필요",
               font_size=10.5, h=Inches(0.26), gap=Inches(0.10),
               font_name=F_REG, color=CHARCOAL)

    cap_h = Inches(0.22)
    tbl_h = lf.remaining - int(cap_h) - int(Inches(0.08))
    trect = lf.reserve(tbl_h, gap=Inches(0.08))
    _hdr = {"fill": CLOUD, "font_name": F_SEMI, "font_size": 10, "font_color": INK,
            "border_edges": "b", "border_color": STEEL, "border_width_pt": 0.75}
    _mdl = {"font_name": F_REG, "font_size": 10.5, "font_color": CHARCOAL,
            "border_edges": "b", "border_color": FOG, "border_width_pt": 0.75}
    _num = {"font_name": F_SEMI, "font_size": 10.5, "font_color": INK,
            "border_edges": "b", "border_color": FOG, "border_width_pt": 0.75}
    add_grid_table(
        slide, trect.left, trect.top, trect.width, trect.height, 3, 3,
        col_widths=[1.9, 0.8, 1.0], row_heights=[0.85, 1, 1],
        font_name=F_REG, default_font_size=10.5, gridlines=False,
        cells={
            (0, 0): dict(_hdr, text="기준 (Kimi K3)", align="l"),
            (0, 1): dict(_hdr, text="필요 GPU", align="c"),
            (0, 2): dict(_hdr, text="총 비용", align="r"),
            (1, 0): dict(_mdl, text="모델 적재만", align="l"),
            (1, 1): dict(_num, text="약 16장", align="c"),
            (1, 2): dict(_num, text="≈ 3.0억 원", align="r"),
            (2, 0): dict(_mdl, text="실사용 (대화 메모리 포함)", align="l"),
            (2, 1): dict(_num, text="약 20장", align="c"),
            (2, 2): dict(_num, text="≈ 3.7억 원", align="r"),
        })
    lf.textbox(slide, "Kimi K3 (2.8조 파라미터) · RTX PRO 6000 1,850만원/장 기준 — 상세 산정은 '로컬 LLM 실측 ②'",
               font_size=9, h=cap_h, font_name=F_REG, color=GRAPHITE)

    # 우측 카드: ①②③ 세로 미니 플로우 (블록 3개 + PRIMARY 1.5pt 수직 화살표 2개)
    # 헤더 1행 (rev6) — estimate_text_size 실측 12pt 폭 3.95in < 내폭 4.63in
    _, rf = _callout_shell(
        slide, cg[0][1], fill=PRIMARY_SOFT, border_color=PRIMARY,
        divider_color=STEEL,
        header_text="로컬 sLLM 전략 — 데이터 축적으로 성능 확보")
    steps = [
        ("① ", "초기", "최신 상용 모델로 AI Agent 성능 확보"),
        ("② ", "성능 확보 후", "지식증류로 로컬 sLLM 학습"),
        ("③ ", "로컬 성능 확보 시", "상용 + 로컬 병행 → 상용 비용 완화"),
    ]
    conn_gap = Inches(0.24)
    blk_h = (rf.remaining - 2 * int(conn_gap)) // 3
    flow_blocks = []
    for i, (marker, label, body) in enumerate(steps):
        rect = rf.reserve(
            blk_h, gap=(conn_gap if i < len(steps) - 1 else Inches(0)))
        flow_blocks.append(_flow_block(slide, rect, marker, label, body))
    for a, b in zip(flow_blocks, flow_blocks[1:]):
        conn = add_smart_connector(slide, a, b,
                                   connector_type=MSO_CONNECTOR.STRAIGHT,
                                   direction="TB", arrow=False)
        conn.line.color.rgb = PRIMARY
        conn.line.width = Pt(1.5)
        add_arrowhead(conn)

    # --- 다크 슬랩 결론 밴드 (INK 배경 + ON_INK + PRIMARY_BRIGHT 강조 — 하단 1곳) ---
    slab = flow.reserve(slab_h)
    slab_shp = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, slab.left, slab.top, slab.width, slab.height)
    slab_shp.fill.solid()
    slab_shp.fill.fore_color.rgb = INK
    slab_shp.line.fill.background()
    set_text_inset(slab_shp, left=Inches(0.15), right=Inches(0.15),
                   top=Inches(0.02), bottom=Inches(0.02))
    set_body_anchor(slab_shp, "ctr")
    slab_tf = slab_shp.text_frame
    slab_tf.word_wrap = True
    slab_p = slab_tf.paragraphs[0]
    slab_p.alignment = PP_ALIGN.CENTER
    for text, color in [
        ("지식증류", PRIMARY_BRIGHT),
        ("로 로컬 sLLM 재투입 — ", ON_INK),
        ("상용 + 로컬 병행", PRIMARY_BRIGHT),
        ("으로 상용 비용 완화", ON_INK),
    ]:
        r = slab_p.add_run()
        r.text = text
        r.font.name = F_SEMI
        r.font.size = Pt(12)
        r.font.color.rgb = color

    # (지식증류 각주는 3페이지(아키텍처)에 있음 — rev12 중복 삭제)

    # --- 전 run 폰트 일괄 교체 (design_system 2절 — SemiBold 는 bold=True 로 변환) ---
    force_font(slide, "맑은 고딕")


# 하네스 호환 별칭 (spec.idx = 4)
build_slide_5 = build_slide_05
