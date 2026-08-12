"""slide_02 — 통합 플랫폼 개요 — SL SW Agent. (rev4)

구성(위→아래, VFlow): 리드(표준 규격, 키워드 PRIMARY) / PRIMARY_SOFT 배너 /
3열 기능 카드(번호+헤더+흐름 요약+현황 소표 [항목|현황|목표]) — 세로 확장 /
※ 각주 3줄(RAG·MISRA·MCP 항목당 1줄).
rev3: 웹 캡처(add_picture)+캡션 삭제 — 실서비스 라이브 시연으로 대체.
rev4: 마일스톤 다크 슬랩 삭제 — 3열 카드가 각주 위 0.12"까지 세로 확장 흡수.
design_system.md 토큰 강제 — 이모지·그림자·CHEVRON 금지.
"""

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from ppt_utils import (
    set_title, clear_placeholders,
    add_para,
    add_grid_table,
    set_body_anchor, set_text_inset,
    calc_grid, estimate_container_height, VFlow,
    force_font,
)
from template_contract import CONTENT_SAFE

# ---- 디자인 토큰 (design_system.md — 이 목록 밖 색 사용 금지) ----
PRIMARY        = RGBColor(0x02, 0x4A, 0xD8)
PRIMARY_DEEP   = RGBColor(0x0E, 0x31, 0x91)
PRIMARY_SOFT   = RGBColor(0xC9, 0xE0, 0xFC)
INK            = RGBColor(0x1A, 0x1A, 0x1A)
CHARCOAL       = RGBColor(0x3D, 0x3D, 0x3D)
GRAPHITE       = RGBColor(0x63, 0x63, 0x63)
CANVAS         = RGBColor(0xFF, 0xFF, 0xFF)
CLOUD          = RGBColor(0xF7, 0xF7, 0xF7)
FOG            = RGBColor(0xE8, 0xE8, 0xE8)

FONT_R  = "Pretendard"
FONT_SB = "Pretendard SemiBold"

# 각주 — 항목당 1줄 (design_system 2-2)
FOOTNOTES = [
    "RAG: 사내 문서·코드를 검색해 답변 근거로 활용",
    "MISRA: 차량 SW C 코딩 안전 규칙",
    "MCP: AI 도구가 사내 시스템 기능을 호출하는 표준 프로토콜",
]

CARDS = [
    {
        "num": "01",
        "header": "AI 코딩 지원",
        "flow": "질문 · 코드 입력 → 사내 문서 · 코드 검색(RAG) → LLM 답변 · 리뷰 · 수정안",
        "rows": [["코드 질의응답", "웹 라이브 동작", "사내 스킬 연계 고도화"],
                 ["Claude Code 플러그인", "MCP 연동 구축", "skill 자동 실행 확대"]],
    },
    {
        "num": "02",
        "header": "코드 심층 분석",
        "flow": "코드베이스 업로드 → 파싱 · 정적 검증(MISRA) → LLM 취약점 진단 · 수정안 → 보고서",
        "rows": [["정적 검증", "컴파일 · MISRA 구축", "다언어 확장"],
                 ["분석 자료 작성", "AI 진단 · 수정안 구현", "자료 자동 작성"]],
    },
    {
        "num": "03",
        "header": "SW 설계 문서 생성",
        "flow": "소스코드 업로드 → 함수 · 호출 관계 파싱 → LLM 본문 · 다이어그램 → SAD · SDD 출력",
        "rows": [["SAD 생성", "파이프라인 검증", "생성 품질 고도화"],
                 ["SDD 생성", "파이프라인 구축", "산출물 정식 적용"]],
    },
]


def _set_runs(paragraph, segments):
    """paragraph(기존 첫 단락)에 서식 run 들을 직접 구성 — 빈 선행 단락 방지."""
    for seg in segments:
        run = paragraph.add_run()
        run.text = seg["text"]
        run.font.size = Pt(seg["size"])
        run.font.name = seg["font"]
        run.font.bold = seg.get("bold", False)
        run.font.color.rgb = seg["color"]


def _status_table(slide, x, y, w, h, rows):
    """카드 하단 현황 소표 — [항목|현황|목표] 헤더 CLOUD + 본문 CANVAS, 수평 hairline 위주."""
    margins = (Inches(0.06), Inches(0.01), Inches(0.03), Inches(0.01))
    header = {"fill": CLOUD, "font_color": INK, "font_name": FONT_SB,
              "font_size": 9, "align": "l", "anchor": "ctr",
              "border_edges": "tb", "border_color": FOG,
              "border_width_pt": 0.75, "margins": margins}
    body0 = {"fill": CANVAS, "font_color": INK, "font_name": FONT_R,
             "font_size": 9, "align": "l", "anchor": "ctr",
             "border_edges": "b", "border_color": FOG,
             "border_width_pt": 0.75, "margins": margins}
    body1 = dict(body0, font_color=CHARCOAL)

    cells = {
        (0, 0): dict(header, text="항목"),
        (0, 1): dict(header, text="현황"),
        (0, 2): dict(header, text="목표"),
    }
    for r, (item, status, goal) in enumerate(rows, start=1):
        cells[(r, 0)] = dict(body0, text=item)
        cells[(r, 1)] = dict(body1, text=status)
        cells[(r, 2)] = dict(body1, text=goal)

    add_grid_table(slide, x, y, w, h, nrows=3, ncols=3, cells=cells,
                   col_widths=[1.05, 1.0, 1.0], row_heights=[0.6, 1, 1],
                   font_name=FONT_R, default_font_size=9, gridlines=False)


def build_slide_02(slide):
    set_title(slide, "통합 플랫폼 개요 — SL SW Agent",
              font_size=20, color=INK, bold=True)
    clear_placeholders(slide, keep=[0])

    L = int(CONTENT_SAFE.left)
    W = int(CONTENT_SAFE.width)
    B = int(CONTENT_SAFE.bottom)

    # ---- 1) 리드 — 표준 규격 (y=0.72", h=0.35") 15pt bold INK, '통합 제공'만 PRIMARY ----
    lead_box = slide.shapes.add_textbox(L, Inches(0.72), W, Inches(0.35))
    lead_tf = lead_box.text_frame
    lead_tf.word_wrap = True
    set_text_inset(lead_box, left=0, top=Inches(0.02), right=0, bottom=Inches(0.02))
    lead_p = lead_tf.paragraphs[0]
    lead_p.alignment = PP_ALIGN.LEFT
    _set_runs(lead_p, [
        {"text": "코딩 지원 · 코드 심층 분석 · SW 설계 문서 생성 — 하나의 웹 플랫폼에서 ",
         "size": 15, "font": FONT_SB, "bold": True, "color": INK},
        {"text": "통합 제공", "size": 15, "font": FONT_SB, "bold": True, "color": PRIMARY},
    ])

    # ---- 본문 콘텐츠 시작 y = 1.18" (리드 표준 간격) ----
    flow = VFlow(x=L, w=W, y_start=Inches(1.18), gap=Inches(0.10))

    # ---- 2) agent 배너 — PRIMARY_SOFT 배경 + PRIMARY_DEEP SemiBold 11.5pt ----
    banner_h = Inches(0.36)
    y = flow.advance(banner_h, gap=Inches(0.14))
    banner = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, L, y, W, banner_h)
    banner.adjustments[0] = 0.10
    banner.fill.solid()
    banner.fill.fore_color.rgb = PRIMARY_SOFT
    banner.line.fill.background()
    banner.text_frame.word_wrap = True
    set_text_inset(banner, left=Inches(0.16), top=Inches(0.02),
                   right=Inches(0.16), bottom=Inches(0.02))
    set_body_anchor(banner, "ctr")
    banner_p = banner.text_frame.paragraphs[0]
    banner_p.alignment = PP_ALIGN.LEFT
    _set_runs(banner_p, [
        {"text": "SL SW Agent — SW 개발 핵심 영역을 통합 관리하는 사내 AI agent",
         "size": 11.5, "font": FONT_SB, "color": PRIMARY_DEEP},
    ])

    # ---- 하단 고정 요소 기하 선산출 — 각주 3줄 높이 → 카드 하단 y 역산 ----
    fn_lines = ["※ " + t for t in FOOTNOTES]
    fn_text = "\n".join(fn_lines)
    fn_h = int(estimate_container_height(fn_text, 9, max_width=W))
    fn_top = B - fn_h

    # ---- 3) 3열 기능 카드 (calc_grid 1×3) — 슬랩 삭제분 흡수, 각주 위 0.12"까지 확장 ----
    y_cards = flow.cursor
    card_h = fn_top - int(Inches(0.12)) - y_cards
    grid = calc_grid(1, 3, area=(L, y_cards, W, card_h), gap=Inches(0.22))

    for cell, spec in zip(grid[0], CARDS):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                      cell.left, cell.top, cell.width, cell.height)
        card.adjustments[0] = 0.05
        card.fill.solid()
        card.fill.fore_color.rgb = CANVAS
        card.line.color.rgb = FOG
        card.line.width = Pt(1.0)
        card.shadow.inherit = False

        inner_x = cell.left + Inches(0.16)
        inner_w = cell.width - Inches(0.32)
        vf = VFlow(x=inner_x, w=inner_w,
                   y_start=cell.top + Inches(0.16),
                   y_max=cell.top + cell.height, gap=Inches(0.04))

        # 번호 01/02/03 — Pretendard Bold 20pt PRIMARY
        vf.textbox(slide, spec["num"], font_size=20, h=Inches(0.32),
                   font_name=FONT_R, bold=True, color=PRIMARY, gap=Inches(0.04))
        # 카드 헤더 — SemiBold 12.5pt INK
        vf.textbox(slide, spec["header"], font_size=12.5, h=Inches(0.26),
                   font_name=FONT_SB, color=INK, gap=Inches(0.10))
        # 기능 흐름 요약 — 10.5pt CHARCOAL (높이 동적 계산)
        flow_h = estimate_container_height(spec["flow"], 10.5, max_width=inner_w,
                                           padding_top=Inches(0.02),
                                           padding_bottom=Inches(0.02))
        vf.textbox(slide, spec["flow"], font_size=10.5, h=flow_h,
                   font_name=FONT_R, color=CHARCOAL)

        # 하단 현황 소표 [항목|현황|목표] — 카드 바닥 기준, 흐름 텍스트와 겹침 없음
        tbl_bottom = cell.top + cell.height - int(Inches(0.16))
        tbl_top = int(vf.cursor) + int(Inches(0.12))
        tbl_h = tbl_bottom - tbl_top
        max_tbl_h = int(Inches(3.0))
        if tbl_h > max_tbl_h:
            tbl_h = max_tbl_h
            tbl_top = tbl_bottom - tbl_h
        _status_table(slide, inner_x, tbl_top, inner_w, tbl_h, spec["rows"])

    # ---- 4) 각주 3줄 — 항목당 1줄 (RAG / MISRA / MCP), CONTENT_SAFE 하단 밀착 ----
    fn_box = slide.shapes.add_textbox(L, fn_top, W, fn_h)
    fn_tf = fn_box.text_frame
    fn_tf.word_wrap = True
    set_text_inset(fn_box, left=0, top=Inches(0.01), right=0, bottom=Inches(0.01))
    fn_p = fn_tf.paragraphs[0]
    fn_p.alignment = PP_ALIGN.LEFT
    _set_runs(fn_p, [{"text": fn_lines[0], "size": 9,
                      "font": FONT_R, "color": GRAPHITE}])
    for line in fn_lines[1:]:
        add_para(fn_tf, line, font_name=FONT_R, font_size=9, color=GRAPHITE)

    # ---- 5) 폰트 일괄 강제 — 맑은 고딕 (design_system.md 2절) ----
    force_font(slide, "맑은 고딕")
