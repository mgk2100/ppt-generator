"""slide_02 — 통합 플랫폼 ① : 코딩 지원 · 취약점 분석 (rev6 — 웹 캡처 복원 · 예상 표기 삭제).

spec:   input/smart-factory-briefing/slides/slide_02.spec.yaml
원본:   input/sl-sw-agent-v11/slides/slide_02.code.py (사용자 확정 구조) —
        좌측 라벨 열(프로젝트/시스템 구성/개발 내용/개발 항목(상세))
        + SL SW Agent 통합 밴드 + 2열(코딩 지원 | 취약점 분석) 레이아웃·문안 유지.
design: input/smart-factory-briefing/design_system.md — 맑은 고딕 단일 폰트,
        제목 20pt bold + 리드(0.72"/0.35" 15pt bold, 키워드 PRIMARY) 표준,
        각주 항목당 1줄, 마지막 force_font(slide, "맑은 고딕").

이식 변경 (rev5 → rev6):
  1) 제목 = "통합 플랫폼 ① — 코딩 지원 · 취약점 분석" (20pt bold INK)
  2) 표준 리드 추가 → 본문 T0 0.84 → 1.18, 행 높이 압축:
     banner 0.36→0.34, 헤더 0.36→0.34, 구성 1.57→1.48,
     내용 2.63→2.14, 상세 1.24→1.00 — Y_BOT 6.48, 이하 각주 3줄 영역
  3) 코딩 지원 열 '개발 내용' = ca_web_assist_conv.png (3200×2000) 웹 캡처 복원
     + ▲ 캡션 "실제 동작 화면 (라이브)" — 취약점 열과 동일 배치 문법
     (열 폭 내 최대 · 세로 중앙) (rev6)
  4) 취약점 분석 열 '개발 내용' = ca_code_analysis.png (3200×2000, /code/analysis 결과 실캡처) + ▲ 캡션
  5) '예상' 표기 전부 삭제 (rev6) — 스네이크 예상안 캡션 · "(예상 UI)" ·
     상세 "자료 자동 작성 (예상)" → "자료 자동 작성"
  6) 각주 3줄 분리 (RAG / MISRA / MCP — 항목당 1줄 + 약어 풀네임 병기,
     8.5pt GRAPHITE 유지)

텍스트 폭 검증 (estimate_text_size 실측 — 스모크 테스트에서 재확인):
  플로우 8.5pt 행별 필요폭 ≤ 배분 가능폭 (잔여 균등 가산, v11 동일 데이터)
  상세 소표 9pt 열 필요폭 합 ≤ COL_W (긴 셀 명시 \n 2행)
  리드 15pt·배너 13pt·캡션 8.5pt·각주 8.5pt(풀네임) 전부 1줄 실측 통과
"""

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR

from ppt_utils import (
    BASE_DIR,
    set_title, clear_placeholders,
    add_accent_bar, add_picture, add_grid_table,
    add_arrowhead, add_routed_connector,
    set_body_anchor, set_text_inset,
    estimate_text_size, force_font,
)
from template_contract import CONTENT_SAFE


# ---------- 디자인 토큰 (design_system.md — 이 목록 밖 색 사용 금지) ----------
PRIMARY      = RGBColor(0x02, 0x4A, 0xD8)   # Electric Blue — 유일한 신호색
PRIMARY_DEEP = RGBColor(0x0E, 0x31, 0x91)   # 보조 강조 (banner·LLM 블록 텍스트)
PRIMARY_SOFT = RGBColor(0xC9, 0xE0, 0xFC)   # 옅은 파랑 표면 (banner·LLM 블록)
INK          = RGBColor(0x1A, 0x1A, 0x1A)
CHARCOAL     = RGBColor(0x3D, 0x3D, 0x3D)
GRAPHITE     = RGBColor(0x63, 0x63, 0x63)
CANVAS       = RGBColor(0xFF, 0xFF, 0xFF)
CLOUD        = RGBColor(0xF7, 0xF7, 0xF7)
FOG          = RGBColor(0xE8, 0xE8, 0xE8)
STEEL        = RGBColor(0xC2, 0xC2, 0xC2)

FONT = "맑은 고딕"                           # 단일 폰트 — weight 는 bold 플래그로

# ---------- 지오메트리 — 좌측 라벨 열 + 2열, 5행 contiguous ----------
LEFT      = int(CONTENT_SAFE.left)
RIGHT     = int(CONTENT_SAFE.right)
FULL_W    = int(CONTENT_SAFE.width)
LABEL_W   = int(Inches(0.75))
LABEL_GAP = int(Inches(0.08))
COL_GAP   = int(Inches(0.14))
COL1_X    = LEFT + LABEL_W + LABEL_GAP
RIGHT_W   = RIGHT - COL1_X                  # 라벨 열 제외 우측 전체 폭 ≈9.45in
COL_W     = (RIGHT_W - COL_GAP) // 2        # ≈4.655in
COL2_X    = COL1_X + COL_W + COL_GAP
COL_XS    = (COL1_X, COL2_X)

T0    = int(Inches(1.18))       # 프레임 top — 리드(0.72~1.07) 아래 표준 시작선
H_BAN = int(Inches(0.34))       # agent 통합 밴드 (v11 0.36 -0.02)
H_HDR = int(Inches(0.34))       # 프로젝트 헤더 (v11 0.36 -0.02)
H_SYS = int(Inches(1.48))       # 시스템 구성 (v11 1.57 -0.09 — 스네이크 1.46 수용)
H_CNT = int(Inches(2.14))       # 개발 내용 (v11 2.63 -0.49 — img_h 축소)
H_DTL = int(Inches(1.00))       # 개발 항목(상세) (v11 1.24 -0.24)
Y_HDR = T0 + H_BAN              # 1.52
Y_SYS = Y_HDR + H_HDR           # 1.86
Y_CNT = Y_SYS + H_SYS           # 3.34
Y_DTL = Y_CNT + H_CNT           # 5.48
Y_BOT = Y_DTL + H_DTL           # 6.48 — 이하 각주 3줄 (bottom 7.00 ≤ 7.02)

HAIR = int(Inches(0.01))        # hairline 두께
PAD  = int(Inches(0.03))        # 행 내부 배지/밴드 상하 여백

# ---------- 콘텐츠 (v11 원문 유지) ----------
LEAD_SEGS = [                   # 표준 리드 — 15pt bold, 키워드만 PRIMARY (rev8 문구 교체)
    ("질문 · 코드에서 ", INK),
    ("답변", PRIMARY),
    ("과 ", INK),
    ("취약점 리포트", PRIMARY),
    ("를 자동 생성", INK),
]

BANNER_TEXT = ("SL SW Agent — 코딩 지원 · 취약점 분석 · SW 설계 문서 생성을 "
               "모두 제공하는 통합 agent")

LABELS = [                      # (텍스트, y, h) — 프로젝트는 banner+헤더 2행 걸침
    ("프로젝트", T0, H_BAN + H_HDR),
    ("시스템\n구성", Y_SYS, H_SYS),
    ("개발 내용", Y_CNT, H_CNT),
    ("개발 항목\n(상세)", Y_DTL, H_DTL),
]

HEADERS = ["코딩 지원", "취약점 분석"]

# 코딩 지원 — 좌→우 4단계 1행: 입력 → 검색 → LLM 활용 → 결과물.
FLOW_COL1 = [
    ["질문 · 코드", "입력"],
    ["사내 문서 ·", "코드 검색", "(RAG · 코드 저장소)"],
    ["답변 작성", "(LLM)"],
    ["답변 · 리뷰 ·", "수정안"],
]

# 취약점 분석 — 6단계 2행 스네이크 (다언어 전제, C/C++ 한정 표현 없음).
FLOW_COL2_ROWS = [
    [["코드베이스 업로드"],
     ["코드 파싱", "(tree-sitter)"],
     ["데이터 적재", "(PostgreSQL · MongoDB)"]],
    [["핵심 코드 분류"],
     ["취약점 진단 · 수정안", "(LLM)"],
     ["취약점 분석 자료 작성"]],
]

FLOW_FONT_PT = 8.5

# 블록 스타일 (fill, border, bold, text_color) — LLM만 SOFT+PRIMARY 강조
_ST_PLAIN = (CANVAS, FOG, False, INK)
_ST_LLM   = (PRIMARY_SOFT, PRIMARY, True, PRIMARY_DEEP)
_ST_OUT   = (CLOUD, None, True, INK)
STYLES_COL1 = [_ST_PLAIN, _ST_PLAIN, _ST_LLM, _ST_OUT]
STYLES_COL2 = [
    [_ST_PLAIN, _ST_PLAIN, _ST_PLAIN],
    [_ST_PLAIN, _ST_LLM, _ST_OUT],
]

# 개발 내용 — 두 열 모두 실물 캡처 + ▲캡션, 동일 배치 문법 (rev6 웹 캡처 복원)
WEB_PATH  = str(BASE_DIR / "sources" / "sl-sw-agent" / "assets"
                / "ca_web_assist_conv.png")           # 3200×2000 (1.6:1)
MOCK_PATH = str(BASE_DIR / "sources" / "sl-sw-agent" / "assets" / "ca_code_analysis.png")
WEB_CAPTION  = "▲ 웹 코드 질의응답 — 실제 동작 화면 (라이브)"
MOCK_CAPTION = "▲ 코드 분석 결과 — 실제 동작 화면 (라이브)"
CONTENT_PICS = [(WEB_PATH, WEB_CAPTION), (MOCK_PATH, MOCK_CAPTION)]

DETAILS = [
    [["코드 질의응답", "웹 라이브 동작", "사내 스킬 연계 고도화"],
     ["Claude Code 플러그인", "MCP 연동 구축", "skill 자동 실행 확대"]],
    [["정적 검증", "컴파일 · MISRA 검증 구축\n(cppcheck)", "tree-sitter 로\n다언어 확장"],
     ["분석 자료 작성", "AI 진단 · 수정안 구현", "자료 자동 작성"]],
]

FOOTNOTES = [                   # 각주 — 항목당 1줄 + 풀네임 병기 (design_system 2-2)
    "RAG(Retrieval-Augmented Generation): 사내 문서·코드를 검색해 답변 근거로 활용",
    "MISRA(Motor Industry Software Reliability Association): 차량 SW C 코딩 안전 규칙",
    "MCP(Model Context Protocol): AI 도구가 사내 시스템 기능을 호출하는 표준 프로토콜",
]


def _band(slide, x, y, w, h, fill):
    """섹션 밴드 — 사각형, 보더 없음 (design_system 표면 규칙)."""
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, int(x), int(y), int(w), int(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def _fill_lines(shp, lines, size, color, bold=False, align=PP_ALIGN.CENTER,
                line_spacing=None):
    """도형 안에 명시 라인만 채운다 (줄바꿈은 여기서 지정한 것만 발생)."""
    tf = shp.text_frame
    tf.word_wrap = True
    for j, line in enumerate(lines):
        p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
        p.alignment = align
        if line_spacing is not None:
            p.line_spacing = line_spacing
        run = p.add_run()
        run.text = line
        run.font.name = FONT
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color


def _label_badge(slide, x, y, w, h, text):
    """좌측 라벨 배지 — CLOUD 배경 + GRAPHITE 10pt bold 세로 중앙."""
    shp = _band(slide, x, y, w, h, CLOUD)
    set_text_inset(shp, left=int(Inches(0.02)), right=int(Inches(0.02)),
                   top=int(Inches(0.02)), bottom=int(Inches(0.02)))
    set_body_anchor(shp, "ctr")
    _fill_lines(shp, text.split("\n"), 10, GRAPHITE, bold=True)
    return shp


def _node(slide, x, y, w, h, lines, size, fill, line_color,
          text_color=INK, bold=False):
    """플로우 블록 — ROUNDED_RECT 소 radius. line_color=None이면 보더 없음."""
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 int(x), int(y), int(w), int(h))
    shp.adjustments[0] = 0.08
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line_color is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line_color
        shp.line.width = Pt(1)
    shp.shadow.inherit = False
    set_text_inset(shp, left=int(Inches(0.04)), right=int(Inches(0.04)),
                   top=int(Inches(0.02)), bottom=int(Inches(0.02)))
    set_body_anchor(shp, "ctr")
    _fill_lines(shp, lines, size, text_color, bold=bold)
    return shp


def _arrow(slide, x1, y1, x2, y2):
    """주 흐름 화살표 — PRIMARY 1.5pt STRAIGHT + arrowhead (design_system 문법)."""
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                      int(x1), int(y1), int(x2), int(y2))
    conn.line.color.rgb = PRIMARY
    conn.line.width = Pt(1.5)
    add_arrowhead(conn)
    return conn


def _bare_textbox(slide, x, y, w, h):
    """margin 0 텍스트박스."""
    tb = slide.shapes.add_textbox(int(x), int(y), int(w), int(h))
    tf = tb.text_frame
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    return tb


# ---------- 플로우 블록 폭 — 행별 실측 배분 (v11 로직 유지) ----------
_FLOW_FX_OFF = int(Inches(0.03))            # 열 내부 좌우 여백
_FLOW_GAP    = int(Inches(0.16))            # 블록 사이 화살표 구간
_FLOW_INSET  = int(Inches(0.10))            # 블록 텍스트 여백(0.04×2) + 슬랙(0.02)

_SNAKE_BLK_H = int(Inches(0.50))            # 취약점 열 블록 높이 (2행)
_SNAKE_TOP   = int(Inches(0.15))            # Y_SYS → 1행 top (캡션 삭제 후 세로 중앙)
_SNAKE_VGAP  = int(Inches(0.18))            # 행 사이 — 행 전환 화살표 구간
# 세로 합: 0.15+0.50+0.18+0.50 = 1.33, 하단 여백 0.15 — H_SYS(1.48) 내 세로 중앙 ✓


def _row_widths(blocks, avail):
    """한 행 블록 폭 실측 배분 — 유출·자동 줄바꿈이 발생하지 않음을 코드로 보장."""
    needed = []
    for lines in blocks:
        max_w = max(int(estimate_text_size(t, FLOW_FONT_PT).width) for t in lines)
        needed.append(max_w + _FLOW_INSET)
    leftover = avail - sum(needed)
    n = len(needed)
    if leftover >= 0:
        add = leftover // n
        widths = [nd + add for nd in needed]
        widths[-1] += leftover - add * n
    else:  # 방어적 축소 — 현 데이터로는 도달하지 않음
        scale = avail / float(sum(needed))
        widths = [int(nd * scale) for nd in needed]
    return widths


def _flow_row(slide, x, y, blocks, widths, h_blk, styles):
    """한 행 좌→우 블록 배치 + 순방향 수평 화살표. 블록 도형 리스트 반환."""
    shapes = []
    bx = int(x)
    for lines, w, (fill, border, bold, tcolor) in zip(blocks, widths, styles):
        shapes.append(_node(slide, bx, y, w, h_blk, lines, FLOW_FONT_PT,
                            fill, border, text_color=tcolor, bold=bold))
        bx += w + _FLOW_GAP
    cy = int(y) + h_blk // 2
    for a, b in zip(shapes, shapes[1:]):
        _arrow(slide, a.left + a.width + HAIR, cy, b.left - HAIR, cy)
    return shapes


def _seq_flow(slide, ci, cx, y):
    """시스템 구성 — 열별 분기 (v11 유지: 순방향 화살표만, LLM 수렴 구도 금지).

    ci=0 코딩 지원:  [입력]→[검색]→[LLM 활용]→[결과물] 1행 4블록 (h 0.88, 세로 중앙).
    ci=1 취약점 분석: 6단계 2행 스네이크 + 행 전환 꺾임 화살표.
    """
    fw = COL_W - 2 * _FLOW_FX_OFF
    fx = cx + _FLOW_FX_OFF

    if ci == 0:
        widths = _row_widths(FLOW_COL1, fw - (len(FLOW_COL1) - 1) * _FLOW_GAP)
        h_blk = int(Inches(0.88))
        by = y + (H_SYS - h_blk) // 2
        _flow_row(slide, fx, by, FLOW_COL1, widths, h_blk, STYLES_COL1)
        return

    # --- 취약점 분석 — 2행 스네이크 ---
    fy1 = y + _SNAKE_TOP
    fy2 = fy1 + _SNAKE_BLK_H + _SNAKE_VGAP
    rows = []
    for ri, (blocks, ry) in enumerate(zip(FLOW_COL2_ROWS, (fy1, fy2))):
        widths = _row_widths(blocks, fw - (len(blocks) - 1) * _FLOW_GAP)
        rows.append(_flow_row(slide, fx, ry, blocks, widths,
                              _SNAKE_BLK_H, STYLES_COL2[ri]))

    # 행 전환 화살표 — 1행 마지막(적재) 하단 → 꺾임 → 2행 첫(분류) 상단 (순방향)
    a, b = rows[0][-1], rows[1][0]
    ax = a.left + a.width // 2
    bx = b.left + b.width // 2
    mid_y = fy1 + _SNAKE_BLK_H + _SNAKE_VGAP // 2
    add_routed_connector(slide, [
        (ax, a.top + a.height + HAIR),
        (ax, mid_y),
        (bx, mid_y),
        (bx, fy2 - HAIR),
    ], arrow=True, color=PRIMARY, width_pt=1.5)


def _detail_table(slide, x, y, w, h, rows):
    """개발 항목(상세) 소표 — [항목|현황|목표] 헤더 + 2행, 9pt (v11 유지).

    열 폭 = 각 열 estimate_text_size 실측 최대폭(명시 \\n 라인 단위) + 셀 여백,
    잔여 균등 분배. 수평 hairline 위주, 수직선 없음, zebra 금지.
    """
    header = ["항목", "개발 현황", "목표"]
    mar_l, mar_r = int(Inches(0.05)), int(Inches(0.04))
    margins = (mar_l, int(Inches(0.02)), mar_r, int(Inches(0.02)))

    needed = []
    for c in range(3):
        texts = [header[c]]
        for r in rows:
            texts.extend(r[c].split("\n"))
        max_w = max(int(estimate_text_size(t, 9).width) for t in texts)
        needed.append(max_w + mar_l + mar_r)
    leftover = int(w) - sum(needed)
    if leftover >= 0:
        widths = [n + leftover // 3 for n in needed]
        widths[-1] += leftover - (leftover // 3) * 3
    else:  # 방어적 축소 — 현 데이터로는 도달하지 않음
        scale = int(w) / float(sum(needed))
        widths = [int(n * scale) for n in needed]

    cells = {}
    for c, htxt in enumerate(header):
        cells[(0, c)] = dict(text=htxt, bold=True, font_size=9,
                             font_color=INK, fill=CLOUD, align="l", anchor="ctr",
                             border_edges="b", border_color=STEEL,
                             border_width_pt=0.75, margins=margins)
    for r, row in enumerate(rows, start=1):
        for c, txt in enumerate(row):
            cells[(r, c)] = dict(text=txt, bold=(c == 0), font_size=9,
                                 font_color=INK if c == 0 else CHARCOAL,
                                 align="l", anchor="ctr",
                                 border_edges="b", border_color=FOG,
                                 border_width_pt=0.75, margins=margins)
    add_grid_table(
        slide, int(x), int(y), int(w), int(h),
        nrows=1 + len(rows), ncols=3, cells=cells,
        col_widths=widths, row_heights=[22, 37, 37],
        font_name=FONT, default_font_size=9, gridlines=False)


def build_slide_02(slide):
    # ---- 제목 — 표준 규격 (20pt bold INK) ----
    set_title(slide, "통합 플랫폼 ① — 코딩 지원 · 취약점 분석",
              font_size=20, color=INK, bold=True)
    clear_placeholders(slide, keep=[0])

    # ---- 리드 — 표준 규격 (y=0.72", h=0.35") 15pt bold, 키워드 PRIMARY ----
    lead = _bare_textbox(slide, LEFT, Inches(0.72), FULL_W, Inches(0.35))
    lead_tf = lead.text_frame
    lead_tf.word_wrap = True
    lead_p = lead_tf.paragraphs[0]
    lead_p.alignment = PP_ALIGN.LEFT
    for text, color in LEAD_SEGS:
        run = lead_p.add_run()
        run.text = text
        run.font.name = FONT
        run.font.size = Pt(15)
        run.font.bold = True
        run.font.color.rgb = color

    # ---- 행 경계 hairline (폭 전체 5줄 + banner/헤더 경계는 우측 영역만) ----
    for hy in (T0, Y_SYS, Y_CNT, Y_DTL, Y_BOT):
        add_accent_bar(slide, LEFT, hy, FULL_W, HAIR, FOG)
    add_accent_bar(slide, COL1_X, Y_HDR, RIGHT_W, HAIR, FOG)

    # ---- 2열 사이 세로 hairline (banner 아래부터) ----
    div_x = COL1_X + COL_W + (COL_GAP - HAIR) // 2
    add_accent_bar(slide, div_x, Y_HDR, HAIR, Y_BOT - Y_HDR, FOG)

    # ---- 좌측 라벨 배지 4개 (프로젝트는 banner+헤더 2행 걸침) ----
    for text, ry, rh in LABELS:
        _label_badge(slide, LEFT, ry + PAD, LABEL_W, rh - 2 * PAD, text)

    # ---- agent 통합 밴드 — 라벨 열 제외 우측 전체 폭, PRIMARY_SOFT 1행 ----
    banner = _band(slide, COL1_X, T0 + PAD, RIGHT_W, H_BAN - 2 * PAD, PRIMARY_SOFT)
    set_text_inset(banner, left=int(Inches(0.10)), right=int(Inches(0.10)),
                   top=int(Inches(0.01)), bottom=int(Inches(0.01)))
    set_body_anchor(banner, "ctr")
    _fill_lines(banner, [BANNER_TEXT], 12.5, PRIMARY_DEEP, bold=True)

    # ---- 개발 내용 지오메트리 — 행 높이 우선, 이미지 비율 유지 ----
    img_y = Y_CNT + int(Inches(0.05))
    img_h = int(Inches(1.82))               # web 3200×2000 → w≈2.91 · mock 2400×1240
    cap_y = img_y + img_h + int(Inches(0.04))   # 캡션 bottom 5.43 ≤ Y_DTL(5.48)
    tbl_y = Y_DTL + int(Inches(0.02))
    tbl_h = int(Inches(0.96))               # bottom 6.46 ≤ Y_BOT(6.48)

    # ---- 2열 — 행별 콘텐츠 ----
    for ci, cx in enumerate(COL_XS):
        # 1) 프로젝트 헤더 — CLOUD 밴드 + INK bold 11.5pt 중앙
        hdr = _band(slide, cx, Y_HDR + PAD, COL_W, H_HDR - 2 * PAD, CLOUD)
        set_text_inset(hdr, left=int(Inches(0.08)), right=int(Inches(0.08)),
                       top=int(Inches(0.01)), bottom=int(Inches(0.01)))
        set_body_anchor(hdr, "ctr")
        _fill_lines(hdr, [HEADERS[ci]], 11.5, INK, bold=True)

        # 2) 시스템 구성 — 열별 분기 (코딩 지원 1행 | 취약점 분석 2행 스네이크)
        _seq_flow(slide, ci, cx, Y_SYS)

        # 3) 개발 내용 — 실물 캡처 (비율 유지, hairline 보더, 열 중앙) + ▲캡션.
        #    두 열 동일 배치 문법: 행 높이 기준 최대, 폭 초과 시 폭 기준 + 세로 중앙
        pic_path, pic_caption = CONTENT_PICS[ci]
        pic = add_picture(slide, pic_path, cx, img_y, h=img_h,
                          line_color=FOG, line_width_pt=1.0)
        if int(pic.width) > COL_W:              # 방어 — 폭 초과 시 폭 기준 재조정
            slide.shapes._spTree.remove(pic._element)
            pic = add_picture(slide, pic_path, cx, img_y, w=COL_W,
                              line_color=FOG, line_width_pt=1.0)
            pic.top = img_y + (img_h - int(pic.height)) // 2
        pic.left = cx + (COL_W - int(pic.width)) // 2
        cap = _bare_textbox(slide, cx, cap_y, COL_W, int(Inches(0.18)))
        cap.text_frame.word_wrap = True
        _fill_lines(cap, [pic_caption], 8.5, GRAPHITE)

        # 4) 개발 항목(상세) — add_grid_table 소표
        _detail_table(slide, cx, tbl_y, COL_W, tbl_h, DETAILS[ci])

    # ---- 하단 각주 — 항목당 1줄 (design_system 2-2, 8.5pt GRAPHITE) ----
    fn = _bare_textbox(slide, LEFT, Inches(6.53), FULL_W, Inches(0.47))
    fn.text_frame.word_wrap = False
    _fill_lines(fn, [f"※ {t}" for t in FOOTNOTES], 8.5, GRAPHITE,
                align=PP_ALIGN.LEFT, line_spacing=Pt(11))

    # ---- 폰트 일괄 강제 — 맑은 고딕 (design_system.md 2절) ----
    force_font(slide, "맑은 고딕")


# 하네스 호환 별칭 (build_slide_{spec.idx})
build_slide_2 = build_slide_02
