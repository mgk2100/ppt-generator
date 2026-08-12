"""slide_03 — 통합 플랫폼 ② — SW 설계 문서 생성. (rev6 — 예상 표기 삭제 + 각주 풀네임)

이식 원본: input/sl-sw-agent-v11/slides/slide_03.code.py (사용자 확정 구조)
  좌측 라벨 열(0.75in CLOUD 배지 4개: 프로젝트/시스템 구성/개발 내용/개발 항목(상세))
  + SL SW Agent 통합 밴드(PRIMARY_SOFT) + SW 설계 문서 생성 와이드 1열:
  시스템 구성 = 좌→우 5단계 **순차** 파이프라인('LLM 활용' 4번째 강조, 크기 동일)
  → 개발 내용 = 큰 이미지 2장 병치(목업 0.48 | 실제 산출물 0.52, 셀 내 fit+중심 정렬)
  → 개발 항목(상세) = 와이드 add_grid_table 소표 [항목|개발 현황|목표].

이식 변경 (smart-factory-briefing design_system.md 표준 — spec rev5):
  1) 제목 = "통합 플랫폼 ② — SW 설계 문서 생성", set_title(font_size=20, bold=True)
  2) 표준 리드 추가 (y=0.72", h=0.35", 15pt bold INK, 'SAD · SDD'만 PRIMARY)
     → 본문 T0=1.18"로 하향, 행 높이 소폭 압축:
       banner 0.30 / 헤더 0.30 / 구성 1.10 / 내용 = 잔여 flex / 상세 1.10
  3) 이미지 2종: mock_docgen.png(2360×1560) "▲ 문서 생성 진행 화면"
     + docgen_dynamic_behavior_sample.png(1619×796)
       "▲ 자동 추출 시퀀스 다이어그램 — 조건 분기 구간 확대 (실제 산출물)"
  4) 각주 = 항목당 1줄 + 약어 풀네임 병기 (SAD/SDD · tree-sitter 분리,
     9pt GRAPHITE, 하단 밀착 — 1줄 fit 실측 ✓)
  5) 마지막 force_font(slide, "맑은 고딕") — SemiBold run 은 bold=True 자동 변환

행 배분 (contiguous, T0=1.18 → Y_BOT = 각주 상단 −0.05, 내용 행이 잔여 흡수):
  banner 0.30 — PRIMARY_SOFT 밴드 + PRIMARY_DEEP 12pt 중앙 (1줄 실측 ✓)
  헤더   0.30 — CLOUD 밴드 + INK 11.5pt 중앙
  구성   1.10 — 블록 y=+0.10 h=0.64, 캡션 y=+0.78 h=0.30 (2줄 허용, bottom +1.08)
              tree-sitter 캡션은 폭 실측 초과(1.93>1.87in) → 자연 분리점 2줄
  내용   ≈2.6 — zone pad 0.08: 이미지 fit + 캡션 0.20 + 간격 0.04, 셀별 중심
  상세   1.10 — 소표 y=+0.08 h=0.94 (헤더+3행, 9.5pt) + 열 폭 실측 재배분 방어
"""

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR

from ppt_utils import (
    BASE_DIR,
    set_title, clear_placeholders,
    add_para, add_accent_bar, add_picture, add_grid_table,
    add_arrowhead, set_body_anchor, set_text_inset,
    estimate_text_size, estimate_container_height,
    force_font,
)
from template_contract import CONTENT_SAFE


# ---------- 디자인 토큰 (design_system.md — 이 목록 밖 색 사용 금지) ----------
PRIMARY      = RGBColor(0x02, 0x4A, 0xD8)   # Electric Blue — 유일한 신호색
PRIMARY_DEEP = RGBColor(0x0E, 0x31, 0x91)   # 보조 강조 (banner 텍스트)
PRIMARY_SOFT = RGBColor(0xC9, 0xE0, 0xFC)   # 옅은 파랑 표면 (banner·LLM 활용 블록)
INK          = RGBColor(0x1A, 0x1A, 0x1A)
CHARCOAL     = RGBColor(0x3D, 0x3D, 0x3D)
GRAPHITE     = RGBColor(0x63, 0x63, 0x63)
CANVAS       = RGBColor(0xFF, 0xFF, 0xFF)
CLOUD        = RGBColor(0xF7, 0xF7, 0xF7)
FOG          = RGBColor(0xE8, 0xE8, 0xE8)
STEEL        = RGBColor(0xC2, 0xC2, 0xC2)

# 빌드 중 폰트 — 마지막 force_font("맑은 고딕")가 라틴+EA+CS 일괄 교체,
# 'SemiBold' 포함 run 은 bold=True 로 변환되어 weight 위계 보존.
F_REG  = "Pretendard"
F_SEMI = "Pretendard SemiBold"

# ---------- 콘텐츠 (v11 문안 유지 + spec rev5 변경분) ----------
TITLE = "통합 플랫폼 ② — SW 설계 문서 생성"

LEAD_SEGMENTS = [                # 15pt bold — 'SAD · SDD'만 PRIMARY
    ("소스코드에서 ", INK),
    ("SAD · SDD", PRIMARY),
    (" 설계 문서를 자동 생성", INK),
]

BANNER_TEXT = ("SL SW Agent — 코딩 지원 · 취약점 분석 · SW 설계 문서 생성을 "
               "모두 제공하는 통합 agent")

HEADER_TEXT = "SW 설계 문서 생성"

# 좌→우 5단계 순차 파이프라인 — LLM 활용은 4번째 단계 (수렴 구도 금지)
PIPELINE = [                    # (라벨 9.5pt 1줄, 보조 캡션 8pt 1줄, 강조 여부)
    ("소스코드 업로드", "압축파일 저장", False),
    ("코드 파싱", "함수 · 호출 관계\n(tree-sitter)", False),   # 폭 1.93>1.87 → 2줄 분리
    ("컴포넌트 구성", "APP · BSP · Driver 분류", False),
    ("섹션별 문서 생성", "LLM 활용 — 본문 · 다이어그램", True),
    ("문서 조립 · 변환", "SAD · SDD DOCX 출력", False),
]

PICTURES = [                    # (경로, px, 캡션, 폭 비율)
    (str(BASE_DIR / "sources" / "sl-sw-agent" / "mockups" / "mock_docgen.png"),
     (2360, 1560), "▲ 문서 생성 진행 화면", 0.48),
    (str(BASE_DIR / "sources" / "sl-sw-agent" / "assets"
         / "docgen_dynamic_behavior_sample.png"),
     (1619, 796),
     "▲ 자동 추출 시퀀스 다이어그램 — 조건 분기 구간 확대 (실제 산출물)", 0.52),
]

DETAILS = [                     # [항목 | 개발 현황 | 목표]
    ["SAD 생성", "자동 생성 파이프라인 검증", "생성 품질 고도화"],
    ["SDD 생성", "생성 파이프라인 구축", "산출물 정식 적용"],
    ["다이어그램 · 인터페이스 표", "코드에서 자동 추출", "적용 범위 확대"],
]

FOOTNOTES = [                   # 항목당 1줄 + 약어 풀네임 병기 (design_system 2-2)
    "SAD(Software Architecture Design Specification) / SDD(Software Detailed Design Specification): "
    "SW 아키텍처 설계서 / 상세 설계서",
]

# ---------- 지오메트리 — 좌측 라벨 열 + 와이드 1열, 5행 contiguous ----------
LEFT      = int(CONTENT_SAFE.left)
RIGHT     = int(CONTENT_SAFE.right)
FULL_W    = int(CONTENT_SAFE.width)
BOTTOM    = int(CONTENT_SAFE.bottom)
LABEL_W   = int(Inches(0.75))
LABEL_GAP = int(Inches(0.08))
COL1_X    = LEFT + LABEL_W + LABEL_GAP
RIGHT_W   = RIGHT - COL1_X                  # 라벨 열 제외 우측 전체 폭 ≈9.45in

# 각주 2줄 높이 → 프레임 하단 Y_BOT 역산 (하단 밀착)
FN_LINES = ["※ " + t for t in FOOTNOTES]
FN_H  = int(estimate_container_height("\n".join(FN_LINES), 9, max_width=FULL_W,
                                      padding_top=Inches(0.02),
                                      padding_bottom=Inches(0.02)))
FN_TOP = BOTTOM - FN_H

T0    = int(Inches(1.18))       # 프레임 top — 리드(0.72~1.07) 아래 표준 시작선
H_BAN = int(Inches(0.30))       # agent 통합 밴드
H_HDR = int(Inches(0.30))       # 프로젝트 헤더
H_SYS = int(Inches(1.10))       # 시스템 구성 (좌→우 5단계 순차 파이프라인)
H_DTL = int(Inches(1.10))       # 개발 항목(상세) — 소표 0.94 + 상단 0.08
Y_HDR = T0 + H_BAN              # 1.48
Y_SYS = Y_HDR + H_HDR           # 1.78
Y_CNT = Y_SYS + H_SYS           # 2.88
Y_BOT = FN_TOP - int(Inches(0.05))
Y_DTL = Y_BOT - H_DTL
H_CNT = Y_DTL - Y_CNT           # 개발 내용 — 잔여 flex (≈2.6in)
assert H_CNT >= int(Inches(2.2)), (
    f"개발 내용 행 높이 부족: {H_CNT / 914400:.2f}in")

HAIR = int(Inches(0.01))        # hairline 두께
PAD  = int(Inches(0.03))        # 행 내부 배지/밴드 상하 여백

LABELS = [                      # (텍스트, y, h) — 프로젝트는 banner+헤더 2행 걸침
    ("프로젝트", T0, H_BAN + H_HDR),
    ("시스템\n구성", Y_SYS, H_SYS),
    ("개발 내용", Y_CNT, H_CNT),
    ("개발 항목\n(상세)", Y_DTL, H_DTL),
]


def _band(slide, x, y, w, h, fill):
    """섹션 밴드 — 사각형, 보더 없음 (design_system 표면 규칙)."""
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, int(x), int(y), int(w), int(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    return shp


def _fill_lines(shp, lines, font, size, color, align=PP_ALIGN.CENTER):
    """도형 안에 명시 라인만 채운다 (줄바꿈은 여기서 지정한 것만 발생)."""
    tf = shp.text_frame
    tf.word_wrap = True
    for j, line in enumerate(lines):
        p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        run.font.name = font
        run.font.size = Pt(size)
        run.font.bold = False
        run.font.color.rgb = color


def _label_badge(slide, x, y, w, h, text):
    """좌측 라벨 배지 — CLOUD 배경 + GRAPHITE 10pt 세로 중앙."""
    shp = _band(slide, x, y, w, h, CLOUD)
    set_text_inset(shp, left=int(Inches(0.02)), right=int(Inches(0.02)),
                   top=int(Inches(0.02)), bottom=int(Inches(0.02)))
    set_body_anchor(shp, "ctr")
    _fill_lines(shp, text.split("\n"), F_SEMI, 10, GRAPHITE)
    return shp


def _node(slide, x, y, w, h, lines, font, size, fill, line_color):
    """파이프라인 블록 — ROUNDED_RECT 소 radius. line_color=None이면 보더 없음."""
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
    set_text_inset(shp, left=int(Inches(0.04)), right=int(Inches(0.04)),
                   top=int(Inches(0.02)), bottom=int(Inches(0.02)))
    set_body_anchor(shp, "ctr")
    _fill_lines(shp, lines, font, size, INK)
    return shp


def _arrow(slide, x1, y1, x2, y2):
    """주 흐름 화살표 — PRIMARY 1.5pt STRAIGHT + arrowhead (design_system §5)."""
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                      int(x1), int(y1), int(x2), int(y2))
    conn.line.color.rgb = PRIMARY
    conn.line.width = Pt(1.5)
    add_arrowhead(conn)
    return conn


def _pipeline_row(slide):
    """시스템 구성 — 좌→우 5블록 순차 파이프라인 + 블록 아래 8pt 보조 캡션.

    수렴 구도 금지: 화살표는 i→i+1 순방향 4개뿐. 'LLM 활용'(4번째)만
    PRIMARY_SOFT + PRIMARY 보더 강조, 크기는 나머지와 동일.
    """
    margin = int(Inches(0.05))
    gap    = int(Inches(0.25))
    fx     = COL1_X + margin
    fw     = RIGHT_W - 2 * margin
    blk_w  = (fw - 4 * gap) // 5            # ≈1.67in, usable 1.59in
    blk_h  = int(Inches(0.64))
    blk_y  = Y_SYS + int(Inches(0.10))
    cap_y  = blk_y + blk_h + int(Inches(0.04))
    cap_h  = int(Inches(0.30))              # 2줄 허용 — bottom Y_SYS+1.08 ≤ +1.10
    ext    = int(Inches(0.10))              # 캡션 좌우 확장 (인접 간 0.05 유지)
    mid_y  = blk_y + blk_h // 2

    blocks = []
    for i, (label, caption, emphasis) in enumerate(PIPELINE):
        bx = fx + i * (blk_w + gap)
        if emphasis:    # 'LLM 활용' — 강조 블록 (다이어그램 핵심 노드 전용)
            shp = _node(slide, bx, blk_y, blk_w, blk_h, [label],
                        F_SEMI, 9.5, PRIMARY_SOFT, PRIMARY)
        elif i == len(PIPELINE) - 1:
            # 결과물 블록 — 현황① 의 '답변·리뷰·수정안'/'취약점 분석 자료 작성'과
            # 동일 스타일 (CLOUD 배경 · 무보더, rev9 통일)
            shp = _node(slide, bx, blk_y, blk_w, blk_h, [label],
                        F_SEMI, 9.5, CLOUD, None)
        else:
            shp = _node(slide, bx, blk_y, blk_w, blk_h, [label],
                        F_SEMI, 9.5, CANVAS, FOG)
        blocks.append(shp)

        # 블록 아래 GRAPHITE 8pt 보조 캡션 1줄 (우측 끝은 CONTENT_SAFE 클램프)
        cx0 = bx - ext
        cx1 = min(bx + blk_w + ext, RIGHT)
        cap = slide.shapes.add_textbox(cx0, cap_y, cx1 - cx0, cap_h)
        cap.text_frame.word_wrap = True
        cap.text_frame.margin_left = 0
        cap.text_frame.margin_right = 0
        cap.text_frame.margin_top = 0
        cap.text_frame.margin_bottom = 0
        _fill_lines(cap, caption.split("\n"), F_REG, 8, GRAPHITE)

    # 순방향 화살표 4개 — i 블록 우변 → i+1 블록 좌변 (동일 y → 완전 수평)
    tip = int(Inches(0.03))
    for a, b in zip(blocks, blocks[1:]):
        _arrow(slide, a.left + a.width + tip, mid_y, b.left - tip, mid_y)


def _picture_row(slide):
    """개발 내용 — 큰 이미지 2장 병치 (0.48 | 0.52), 셀 내 fit + 세로 중심 정렬."""
    gap      = int(Inches(0.25))
    avail    = RIGHT_W - gap
    cell1_w  = int(avail * PICTURES[0][3])
    cell_ws  = [cell1_w, avail - cell1_w]
    cell_xs  = [COL1_X, COL1_X + cell1_w + gap]
    zone_top = Y_CNT + int(Inches(0.08))
    zone_h   = H_CNT - int(Inches(0.16))
    cap_h    = int(Inches(0.20))
    cap_gap  = int(Inches(0.04))
    img_max_h = zone_h - cap_h - cap_gap

    for (path, (px_w, px_h), caption, _ratio), cx, cw in zip(
            PICTURES, cell_xs, cell_ws):
        ratio = px_w / px_h
        # 셀 폭·최대 높이 중 먼저 닿는 축으로 fit (px 비율 → 원본 파일 비율 유지)
        if int(cw / ratio) <= img_max_h:
            pic = add_picture(slide, path, cx, zone_top, w=cw,
                              line_color=FOG, line_width_pt=1.0)
        else:
            pic = add_picture(slide, path, cx, zone_top, h=img_max_h,
                              line_color=FOG, line_width_pt=1.0)
        # 이미지+캡션 유닛을 행 높이에 맞춰 세로 중심 정렬, 셀 내 가로 중앙
        unit_h = int(pic.height) + cap_gap + cap_h
        pic.top = zone_top + max(0, (zone_h - unit_h) // 2)
        pic.left = cx + (cw - int(pic.width)) // 2

        cap = slide.shapes.add_textbox(
            cx, int(pic.top) + int(pic.height) + cap_gap, cw, cap_h)
        cap.text_frame.word_wrap = True
        cap.text_frame.margin_left = 0
        cap.text_frame.margin_right = 0
        cap.text_frame.margin_top = 0
        cap.text_frame.margin_bottom = 0
        _fill_lines(cap, [caption], F_REG, 8.5, GRAPHITE)


def _detail_table(slide, x, y, w, h, rows):
    """개발 항목(상세) 소표 — [항목|개발 현황|목표] 헤더 + 3행, 9.5pt.

    열 폭 비율 [1.6, 2.2, 2.2] + estimate_text_size 실측 검증.
    비율 폭이 실측 필요폭보다 작으면 실측 기반 재배분 → 유출·세로 줄바꿈 원천 차단.
    수평 hairline 위주 (헤더 아래 STEEL, 본문 아래 FOG), 수직선 없음, zebra 금지.
    """
    header = ["항목", "개발 현황", "목표"]
    mar_l, mar_r = int(Inches(0.08)), int(Inches(0.04))
    margins = (mar_l, int(Inches(0.02)), mar_r, int(Inches(0.02)))

    # 비율 폭
    ratios = [1.6, 2.2, 2.2]
    tot = sum(ratios)
    widths = [int(int(w) * r / tot) for r in ratios]
    widths[-1] = int(w) - sum(widths[:-1])

    # 실측 필요폭 검증 (현 데이터로는 비율 폭이 모두 여유 있음)
    needed = []
    for c in range(3):
        texts = [header[c]] + [r[c] for r in rows]
        max_w = max(int(estimate_text_size(t, 9.5).width) for t in texts)
        needed.append(max_w + mar_l + mar_r)
    if any(n > wd for n, wd in zip(needed, widths)):
        leftover = int(w) - sum(needed)
        if leftover >= 0:       # 실측 + 잔여 균등 분배
            widths = [n + leftover // 3 for n in needed]
            widths[-1] += int(w) - sum(widths)
        else:                   # 방어적 축소 — 현 데이터로는 도달하지 않음
            scale = int(w) / float(sum(needed))
            widths = [int(n * scale) for n in needed]

    cells = {}
    for c, htxt in enumerate(header):
        cells[(0, c)] = dict(text=htxt, font_name=F_SEMI, font_size=9.5,
                             font_color=INK, fill=CLOUD, align="l", anchor="ctr",
                             border_edges="b", border_color=STEEL,
                             border_width_pt=0.75, margins=margins)
    for r, row in enumerate(rows, start=1):
        for c, txt in enumerate(row):
            cells[(r, c)] = dict(text=txt,
                                 font_name=F_SEMI if c == 0 else F_REG,
                                 font_size=9.5,
                                 font_color=INK if c == 0 else CHARCOAL,
                                 align="l", anchor="ctr",
                                 border_edges="b", border_color=FOG,
                                 border_width_pt=0.75, margins=margins)
    add_grid_table(
        slide, int(x), int(y), int(w), int(h),
        nrows=1 + len(rows), ncols=3, cells=cells,
        col_widths=widths, row_heights=[28, 30, 30, 30],
        font_name=F_REG, default_font_size=9.5, gridlines=False)


def build_slide_03(slide):
    # ---- 제목 — 표준 규격 (design_system 2-1) ----
    set_title(slide, TITLE, font_size=20, color=INK, bold=True)
    clear_placeholders(slide, keep=[0])

    # ---- 리드 — 표준 규격 (y=0.72", h=0.35") 15pt bold, 'SAD · SDD' PRIMARY ----
    lead = slide.shapes.add_textbox(LEFT, int(Inches(0.72)),
                                    FULL_W, int(Inches(0.35)))
    lead_tf = lead.text_frame
    lead_tf.word_wrap = True
    set_text_inset(lead, left=0, top=int(Inches(0.02)),
                   right=0, bottom=int(Inches(0.02)))
    lead_p = lead_tf.paragraphs[0]
    lead_p.alignment = PP_ALIGN.LEFT
    for text, color in LEAD_SEGMENTS:
        run = lead_p.add_run()
        run.text = text
        run.font.name = F_SEMI
        run.font.size = Pt(15)
        run.font.bold = True
        run.font.color.rgb = color

    # ---- 행 경계 hairline (폭 전체 5줄 + banner/헤더 경계는 우측 영역만) ----
    for hy in (T0, Y_SYS, Y_CNT, Y_DTL, Y_BOT):
        add_accent_bar(slide, LEFT, hy, FULL_W, HAIR, FOG)
    add_accent_bar(slide, COL1_X, Y_HDR, RIGHT_W, HAIR, FOG)

    # ---- 좌측 라벨 배지 4개 (프로젝트는 banner+헤더 2행 걸침) ----
    for text, ry, rh in LABELS:
        _label_badge(slide, LEFT, ry + PAD, LABEL_W, rh - 2 * PAD, text)

    # ---- agent 통합 밴드 — 라벨 열 제외 우측 전체 폭, PRIMARY_SOFT 1행 ----
    banner = _band(slide, COL1_X, T0 + PAD, RIGHT_W, H_BAN - 2 * PAD, PRIMARY_SOFT)
    set_text_inset(banner, left=int(Inches(0.10)), right=int(Inches(0.10)),
                   top=int(Inches(0.01)), bottom=int(Inches(0.01)))
    set_body_anchor(banner, "ctr")
    _fill_lines(banner, [BANNER_TEXT], F_SEMI, 12.5, PRIMARY_DEEP)

    # ---- 프로젝트 헤더 — CLOUD 밴드 + INK 11.5pt 중앙 (와이드 1열) ----
    hdr = _band(slide, COL1_X, Y_HDR + PAD, RIGHT_W, H_HDR - 2 * PAD, CLOUD)
    set_text_inset(hdr, left=int(Inches(0.08)), right=int(Inches(0.08)),
                   top=int(Inches(0.01)), bottom=int(Inches(0.01)))
    set_body_anchor(hdr, "ctr")
    _fill_lines(hdr, [HEADER_TEXT], F_SEMI, 11.5, INK)

    # ---- 시스템 구성 — 좌→우 5단계 순차 파이프라인 ----
    _pipeline_row(slide)

    # ---- 개발 내용 — 큰 이미지 2장 병치 (목업 + 실제 산출물 다이어그램) ----
    _picture_row(slide)

    # ---- 개발 항목(상세) — 와이드 소표 헤더+3행 ----
    _detail_table(slide, COL1_X, Y_DTL + int(Inches(0.08)),
                  RIGHT_W, int(Inches(0.94)), DETAILS)

    # ---- 각주 2줄 — 항목당 1줄 (SAD/SDD · tree-sitter), 하단 밀착 ----
    fn_box = slide.shapes.add_textbox(LEFT, FN_TOP, FULL_W, FN_H)
    fn_tf = fn_box.text_frame
    fn_tf.word_wrap = True
    set_text_inset(fn_box, left=0, top=int(Inches(0.01)),
                   right=0, bottom=int(Inches(0.01)))
    fn_p = fn_tf.paragraphs[0]
    fn_p.alignment = PP_ALIGN.LEFT
    run = fn_p.add_run()
    run.text = FN_LINES[0]
    run.font.name = F_REG
    run.font.size = Pt(9)
    run.font.color.rgb = GRAPHITE
    for line in FN_LINES[1:]:
        add_para(fn_tf, line, font_name=F_REG, font_size=9, color=GRAPHITE)

    # ---- 폰트 일괄 강제 — 맑은 고딕 (design_system.md 2절) ----
    force_font(slide, "맑은 고딕")


# 하네스 호환 별칭 (build_slide_{spec.idx})
build_slide_3 = build_slide_03
