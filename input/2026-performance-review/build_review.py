#!/usr/bin/env python3
"""2026 개발 성과리뷰 - PPT 생성 스크립트 (팀장 보고, 표지 + 1장)

실행: cd ~/Share/ppt-generator && python3 input/2026-performance-review/build_review.py
산출: output/2026-성과리뷰.pptx · 사실관계 근거는 memory performance-review-2026-deck
"""
import sys
sys.path.insert(0, "/home/user/Share/ppt-generator")

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

from ppt_utils import (
    load_template, get_layout, clear_placeholders, ensure_fonts,
    add_textbox, add_rich_text, add_footnote, set_title, setup_cover,
    add_grid_table, force_font, CONTENT_SAFE, OUTPUT_DIR,
)

# smart-factory-briefing 디자인 토큰 (단일 파랑 + 무채색)
PRIMARY = RGBColor(0x02, 0x4A, 0xD8)
PRIMARY_DEEP = RGBColor(0x0E, 0x31, 0x91)
PRIMARY_SOFT = RGBColor(0xC9, 0xE0, 0xFC)
INK = RGBColor(0x1A, 0x1A, 0x1A)
CHARCOAL = RGBColor(0x3D, 0x3D, 0x3D)
GRAPHITE = RGBColor(0x63, 0x63, 0x63)
CANVAS = RGBColor(0xFF, 0xFF, 0xFF)
CLOUD = RGBColor(0xF7, 0xF7, 0xF7)
FOG = RGBColor(0xE8, 0xE8, 0xE8)
STEEL = RGBColor(0xC2, 0xC2, 0xC2)

FONT = "맑은 고딕"
L, W = CONTENT_SAFE.left, CONTENT_SAFE.width

ensure_fonts()
prs = load_template()


def lead(slide, segments):
    tb = add_textbox(slide, L, Inches(0.72), W, Inches(0.35), "", font_size=15)
    tf = tb.text_frame
    tf.paragraphs[0]._p.getparent().remove(tf.paragraphs[0]._p)
    add_rich_text(tf, [s if isinstance(s, dict) else {"text": s, "font_size": 15, "bold": True, "color": INK}
                       for s in segments])


def kw(t, size=15):
    return {"text": t, "font_size": size, "bold": True, "color": PRIMARY}


def lines_cell(lines, size=10.5, color=CHARCOAL, bold_first=False):
    """셀 안에 여러 줄을 각 단락으로 넣기 위한 style_cell 후처리용 텍스트."""
    return "\n".join(lines)


# ── 표지 ──
s = prs.slides.add_slide(get_layout(prs, "제목 슬라이드"))
setup_cover(s, "2026년 개발 성과 리뷰", purpose="보고", date="2026.10.16")

# ── 1페이지: 과제별 성과 ──
s = prs.slides.add_slide(get_layout(prs, "제목 및 내용"))
set_title(s, "2026년 개발 성과", font_size=20, bold=True)
clear_placeholders(s, keep=[0])
lead(s, ["올해 진행 상황"])

rows = [
    ("01", "SAD/SDD\n설계 문서 생성",
     ["SAD/SDD 생성 파이프라인 설계",
      "BLTN SAD/SDD 문서 생성 대응",
      "통합 플랫폼 문서 생성 기능으로 플랫폼화"],
     [("SDD A-SPICE 심사 대응 완료", ["컴포넌트 11개 · 유닛 124개 · 유닛 요구사항 2,017개"]),
      ("SAD 생성 플랫폼화", ["슈어소프트 SAD 인터페이스 재현율 0.742"])],
     ("보류", ["취약점 분석 종료 후 프로젝트별 표준 플랫폼화"])),
    ("02", "SW 설계용\nAI Agent 플랫폼",
     ["통합 플랫폼(웹) - 코딩 지원·정적 분석·문서 생성",
      "VS Code 플러그인 - 코딩·정적 분석 지원",
      "코드 분석 워크스페이스 - 정적 분석·변경 비교·코드 편집",
      "문서 생성 하네스 구축 - PPT·Excel·Word 생성",
      "Codebeamer 연동 - A-SPICE 추적성 검사",
      "사내 지식 베이스 구축·RAG 연동"],
     [("통합 플랫폼 기능 구축 완료", ["테스터 30명 기반 사내 데이터 축적",
                                   "개인·팀별 에이전트 설계·구현 진행 중"])],
     ("품질 측정·최적화", ["답변 품질 측정 기준 수립",
                         "개인·팀별 에이전트 최적화 마무리"])),
    ("03", "빌트인캠 SW\n취약점 분석",
     ["2.6세대 · 3세대 2개 프로젝트 분석",
      "분석 보고서(HTML·Word·PPT) 작성·사내 웹 게시"],
     [("분석 진행 중", ["2.6세대 4개 기능 함수 3,628개 판정 · 결함 332건",
                     "3세대 ADAS-PRK 함수 1,114개 판정 (검증 전)"])],
     ("최우선 과제", ["2.6·3세대 결함 분석 확대",
                    "통합 플랫폼 취약점 분석 기능화"])),
]

HEAD, BODY = "H", "B"   # 결과·계획 칸의 단락 종류 표시


def block(groups):
    """[(머리, [줄...])...] → 셀 텍스트와 단락별 종류 목록"""
    lines, kinds = [], []
    for h, items in groups:
        lines.append(h); kinds.append(HEAD)
        for it in items:
            lines.append(it); kinds.append(BODY)
    return "\n".join(lines), kinds


cells, kinds_map = {}, {}
for c, t in enumerate(["번호", "과제", "수행 내역", "결과"]):
    cells[(0, c)] = {"text": t, "bold": True, "font_size": 11, "font_color": INK,
                     "fill": CLOUD, "align": "c" if c == 0 else "l", "border_edges": "b",
                     "border_color": INK, "border_width_pt": 1}
for r, (no, name, work, result, plan) in enumerate(rows, start=1):
    edge = {"border_edges": "b", "border_color": FOG, "border_width_pt": 0.75}
    cells[(r, 0)] = {"text": no, "font_size": 13, "bold": True, "font_color": INK, "align": "c", **edge}
    cells[(r, 1)] = {"text": name, "font_size": 12, "bold": True, "font_color": INK, "align": "l", **edge}
    cells[(r, 2)] = {"text": "\n".join(work), "font_size": 10, "font_color": CHARCOAL, "align": "l", **edge}
    kinds_map[(r, 2)] = [BODY] * len(work)
    txt, kinds_map[(r, 3)] = block(result)
    cells[(r, 3)] = {"text": txt, "font_size": 10, "font_color": INK, "align": "l", **edge}

shape, table = add_grid_table(s, L, Inches(1.12), W, Inches(3.75), 4, 4, cells=cells,
                              col_widths=[0.75, 1.55, 3.95, 4.03], row_heights=[0.32, 1.1, 1.5, 0.83],
                              gridlines=False)
# 셀 안 단락별 서식 (python-pptx 원시 API: style_cell 은 단락별 서식을 받지 않음)
for (r, c), kinds in kinds_map.items():
    paras = table.cell(r, c).text_frame.paragraphs
    for i, (p, k) in enumerate(zip(paras, kinds)):
        p.space_after = Pt(2)
        if not p.runs:
            continue
        if k == HEAD:
            p.runs[0].font.bold = True
            p.runs[0].font.size = Pt(11)
            p.runs[0].font.color.rgb = PRIMARY
            if i > 0:
                p.space_before = Pt(5)
        else:
            p.runs[0].text = "▸ " + p.runs[0].text

# 월별 흐름 1~12월 (11~12월 = 계획, 점선 테두리)
add_textbox(s, L, Inches(5.1), W, Inches(0.28), "월별 진행 · 12월까지 계획", font_size=12, bold=True, color=INK)
g = {(0, 0): {"text": ""}}
for i in range(12):
    g[(0, i + 1)] = {"text": f"{i + 1}월", "font_size": 9, "bold": True,
                     "font_color": PRIMARY if i >= 10 else GRAPHITE, "align": "c",
                     "border_edges": "b", "border_color": INK, "border_width_pt": 1}
tracks = [
    ("설계 문서 생성", [(1, 8, "SAD/SDD 설계 · BLTN 문서 생성 대응", "done"), (9, 10, "통합 플랫폼 연동", "soft"), (11, 12, "보류", "hold")]),
    ("AI Agent 플랫폼", [(6, 7, "설계 · 상용 LLM 전환", "soft"), (8, 10, "정식 개통 · 테스터 운영·데이터 축적", "done"),
                       (11, 12, "품질 측정 · 최적화", "plan")]),
    ("취약점 분석", [(10, 10, "분석", "done"), (11, 12, "최우선 · 분석 확대", "plan")]),
]
STY = {
    "done": {"fill": PRIMARY, "font_color": CANVAS, "bold": True},
    "soft": {"fill": PRIMARY_SOFT, "font_color": INK, "bold": False},
    "plan": {"fill": CANVAS, "font_color": PRIMARY, "bold": True, "border_edges": "lrtb",
             "border_color": PRIMARY, "border_width_pt": 1, "border_dash": "dash"},
    "hold": {"fill": CLOUD, "font_color": GRAPHITE, "bold": False},
}
for r, (name, bars) in enumerate(tracks, start=1):
    g[(r, 0)] = {"text": name, "font_size": 10, "bold": True, "font_color": INK, "align": "l"}
    for (a, b, label, kind) in bars:
        g[(r, a)] = {"text": label, "font_size": 8.5, "align": "c", "span": (1, b - a + 1), **STY[kind]}
add_grid_table(s, L, Inches(5.42), W, Inches(1.22), 4, 13, cells=g,
               col_widths=[1.48] + [0.733] * 12, row_heights=[0.25, 0.32, 0.32, 0.32], gridlines=False)
force_font(s, FONT)

# 지정하지 않은 셀 변은 이웃 칸의 맞닿은 선을 따르고, 없으면 「선 없음」으로 명시
# (python-pptx 원시 API: set_cell_border 는 noFill 을 못 씀)
import copy
from pptx.oxml.ns import qn
TAGS = ("a:lnL", "a:lnR", "a:lnT", "a:lnB")
OPP = {"a:lnL": ("a:lnR", 0, -1), "a:lnR": ("a:lnL", 0, 1), "a:lnT": ("a:lnB", -1, 0), "a:lnB": ("a:lnT", 1, 0)}
for sl in list(prs.slides)[1:]:
    for shp in sl.shapes:
        if not shp.has_table:
            continue
        grid = [[tc for tc in tr.tc_lst] for tr in shp.table._tbl.tr_lst]
        orig = [[{t: tc.get_or_add_tcPr().find(qn(t)) for t in TAGS} for tc in row] for row in grid]
        for r, row in enumerate(grid):
            for c, tc in enumerate(row):
                tcPr = tc.get_or_add_tcPr()
                for t in TAGS:
                    if orig[r][c][t] is not None:
                        continue
                    o, dr, dc = OPP[t]
                    rr, cc = r + dr, c + dc
                    nb = orig[rr][cc][o] if 0 <= rr < len(grid) and 0 <= cc < len(row) else None
                    if nb is not None:
                        ln = copy.deepcopy(nb); ln.tag = qn(t)
                    else:
                        ln = tcPr.makeelement(qn(t), {"w": "0"})
                        ln.append(ln.makeelement(qn("a:noFill"), {}))
                    tcPr.append(ln)
                lns = [tcPr.find(qn(t)) for t in TAGS]
                for el in lns:
                    tcPr.remove(el)
                for i, el in enumerate(lns):
                    tcPr.insert(i, el)

out = f"{OUTPUT_DIR}/2026-성과리뷰.pptx"
prs.save(out)
print(out)
