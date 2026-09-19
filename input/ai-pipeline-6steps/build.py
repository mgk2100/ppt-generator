#!/usr/bin/env python3
"""LLM · Vision · 로봇 공통 6단계 — 본 덱 1장 + 단계별 첨부 6장.

본 덱 각 박스 우상단의 PPT 아이콘을 더블클릭하면 그 단계의 첨부가 열린다
(연결이 아니라 파일에 포함 — 본 덱 하나만 보내도 첨부 6개가 함께 간다).
모든 요소는 네이티브 도형 · 텍스트다. 그림 자리는 점선 자리표시자로만 잡아 두었다.
"""
import os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(os.path.dirname(HERE), "sw-dev-ai-llm")
sys.path.insert(0, SRC)

from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

from deck_common import (new_slide, header, shape, tb, add_shadow, embed_pptx, pic, REPO,
                         NAVY, BLUE, BLUE_T, BLUE_B, GOLD_L, GREY, WHITE, INK, TBL_RULE)
from steps import STEPS, PARTS, LOOP_LABEL
from PIL import Image

IMG = lambda f: os.path.join(HERE, "img", f)


def place(slide, path, cx, cy, max_w, max_h):
    """원본 비율을 지키며 (cx, cy) 중앙에 넣는다."""
    w0, h0 = Image.open(path).size
    r = min(max_w / w0, max_h / h0)
    w, h = w0 * r, h0 * r
    return pic(slide, path, cx - w / 2, cy - h / 2, w=w, h=h)

MASTER  = os.path.join(SRC, "base_master.pptx")
OBJICON = os.path.join(SRC, "assets", "attach1_objicon.png")
ICON    = lambda n: os.path.join(HERE, "png", "step%s.png" % n)
OUT     = os.path.join(REPO, "output")
MAIN    = os.path.join(OUT, "ai-pipeline-6steps.pptx")
ATTACH  = lambda n: os.path.join(OUT, "ai-pipeline-6steps-%s.pptx" % n)

SLATE   = RGBColor(0x64, 0x74, 0x8B)   # Vision
SLATE_T = RGBColor(0xF7, 0xF8, 0xFA); SLATE_B = RGBColor(0xD7, 0xDD, 0xE6)
ROBOT   = RGBColor(0x9A, 0xA3, 0xB2)   # 로봇 (타 팀 작성 예정)
ROBOT_T = RGBColor(0xFA, 0xFA, 0xFB)
ROBOT_B = RGBColor(0xE2, 0xE5, 0xEA)
HEAD_S  = RGBColor(0xE2, 0xE6, 0xEC); HEAD_S_L = RGBColor(0xD8, 0xE6, 0xF7)
DASH_C  = RGBColor(0xB4, 0xC0, 0xD2)   # 그림 자리 점선
ART_TX  = RGBColor(0x76, 0x82, 0x93)

X0, Y0, W, H = 0.35, 1.24, 12.63, 6.02
RAD = 0.06
def adj(h): return [min(RAD / max(h, 0.01), 0.5)]


# =====================================================================
# 첨부 N — 한 단계를 LLM · Vision · 로봇 3열로
#   각 열 = 흐름 도식(2~3단) + 설명 3줄(수행 주체 · 작업 내용 · 난이도 · 공수)
# =====================================================================
COL_GAP = 0.12
COL_W = (W - COL_GAP * 2) / 3
COL_X = [X0 + i * (COL_W + COL_GAP) for i in range(3)]
PH_H, ART_SIZE = 0.46, 3.78
DESC_H = H - PH_H - 0.12 - ART_SIZE - 0.14


def desc_block(slide, x, y, w, h, text, border):
    """작업 내용 + 어려운 점을 한 블록에 적는다."""
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h,
          fill=WHITE, line=border, line_w=0.75, adj=adj(h))
    tb(slide, x + 0.16, y, w - 0.32, h, text, size=10.5, color=INK,
       anchor=MSO_ANCHOR.MIDDLE, spacing=1.45)


def attach_slide(prs, st):
    s = new_slide(prs)
    header(s, st["n"] + "단계", st["title"], [(st["sub"], {"color": BLUE})],
           sub_w=11.6, sub_size=15)

    parts = [
        dict(name=PARTS[0], main=BLUE,  bord=BLUE_B,  data=st["llm"]),
        dict(name=PARTS[1], main=SLATE, bord=SLATE_B, data=st["vis"]),
        dict(name=PARTS[2], main=ROBOT, bord=ROBOT_B, data=None),
    ]

    fy = Y0 + PH_H + 0.12
    dy = fy + ART_SIZE + 0.14
    for i, p in enumerate(parts):
        x = COL_X[i]
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, Y0, COL_W, PH_H, fill=p["main"], adj=adj(PH_H))
        tb(s, x + 0.15, Y0, COL_W - 0.30, PH_H, p["name"],
           size=13, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)

        d = p["data"]
        if d is None:                                    # 로봇 — 타 팀 작성 예정
            shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, fy, COL_W, ART_SIZE,
                  fill=ROBOT_T, line=DASH_C, line_w=1.0,
                  dash=MSO_LINE_DASH_STYLE.DASH, adj=adj(ART_SIZE))
            shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, dy, COL_W, DESC_H,
                  fill=WHITE, line=ROBOT_B, line_w=0.75, adj=adj(DESC_H))
            continue

        art = IMG("step%s_%s.png" % (st["n"], "llm" if i == 0 else "vis"))
        if os.path.exists(art):
            place(s, art, x + COL_W / 2, fy + ART_SIZE / 2, COL_W - 0.16, ART_SIZE)
        desc_block(s, x, dy, COL_W, DESC_H, d["desc"], p["bord"])


# =====================================================================
# 본 덱 1장 — 6단계 카테고리
# =====================================================================
BCOL_GAP, BROW_GAP = 0.15, 0.16
MX0 = 0.84                                     # 본 덱 박스 좌측 — 순환 화살표가 지나갈 여백을 둔다
MW = W - (MX0 - X0)
LOOP_H = 0.52                                  # 하단 순환 화살표 영역
BW = (MW - BCOL_GAP * 2) / 3
BCOL_X = [MX0 + i * (BW + BCOL_GAP) for i in range(3)]
ICON_W = 1.30                                  # 단계 아이콘 (정사각)
LOOP_LW = 0.035                                # 순환 화살표 선 두께


def loop_arrow(slide, x_from, y_bottom, y_mid, x_to):
    """6단계 하단 → 좌측 → 1단계 좌측으로 되돌아가는 순환 화살표."""
    lx = MX0 - 0.42                            # 좌측 세로 구간
    ly = y_bottom + 0.24                       # 하단 가로 구간
    h = LOOP_LW
    shape(slide, MSO_SHAPE.RECTANGLE, x_from - h / 2, y_bottom, h, ly - y_bottom, fill=GOLD_L)
    shape(slide, MSO_SHAPE.RECTANGLE, lx, ly, x_from - lx, h, fill=GOLD_L)
    shape(slide, MSO_SHAPE.RECTANGLE, lx, y_mid, h, ly - y_mid + h, fill=GOLD_L)
    shape(slide, MSO_SHAPE.RECTANGLE, lx, y_mid, x_to - 0.16 - lx, h, fill=GOLD_L)
    head = shape(slide, MSO_SHAPE.ISOSCELES_TRIANGLE, x_to - 0.20, y_mid - 0.09, 0.20, 0.19, fill=GOLD_L)
    head.rotation = 90
    tb(slide, lx + 0.16, ly + 0.05, 5.4, 0.24, LOOP_LABEL,
       size=9.5, bold=True, color=GOLD_L, anchor=MSO_ANCHOR.MIDDLE)


def main_slide(prs):
    """6단계만 크게 — 아이콘 + 제목 + 한 줄, 하단에 1단계로 되돌아가는 순환."""
    s = new_slide(prs)
    header(s, "단계", "LLM · Vision · 로봇 개발 프로세스 6단계",
           [("LLM · Vision · 로봇 별도 운영 · 개발 필요", {"color": BLUE})],
           sub_w=11.6, sub_size=15)

    bh = (H - LOOP_H - BROW_GAP) / 2
    for i, st in enumerate(STEPS):
        x = BCOL_X[i % 3]
        y = Y0 + (i // 3) * (bh + BROW_GAP)
        box = shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, BW, bh, fill=WHITE,
                    line=TBL_RULE, line_w=0.75, adj=adj(bh))
        add_shadow(box, blur=9, dist=3, alpha=10)

        shape(s, MSO_SHAPE.OVAL, x + 0.24, y + 0.22, 0.36, 0.36, fill=NAVY)
        tb(s, x + 0.24, y + 0.22, 0.36, 0.36, st["n"], size=13, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        ip = ICON(st["n"])
        if os.path.exists(ip):
            pic(s, ip, x + (BW - ICON_W) / 2, y + 0.66, w=ICON_W, h=ICON_W)

        tb(s, x + 0.18, y + bh - 0.78, BW - 0.36, 0.40, st["title"],
           size=17, bold=True, color=INK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        tb(s, x + 0.18, y + bh - 0.40, BW - 0.36, 0.30, st["sub"],
           size=10.5, color=GREY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        ap = ATTACH(st["n"])
        if os.path.exists(ap) and os.path.exists(OBJICON):
            embed_pptx(s, ap, OBJICON, x + BW - 0.62, y + 0.22, 0.36, 0.36,
                       name="첨부 %s" % st["n"], idx=int(st["n"]), show_as_icon=True)

    row2_bottom = Y0 + bh + BROW_GAP + bh
    loop_arrow(s, BCOL_X[2] + BW / 2, row2_bottom, Y0 + bh / 2, MX0)


def main():
    os.makedirs(OUT, exist_ok=True)
    for st in STEPS:                       # 첨부 먼저 (본 덱이 임베드해야 하므로)
        prs = Presentation(MASTER)
        attach_slide(prs, st)
        prs.save(ATTACH(st["n"]))
        print(ATTACH(st["n"]))
    prs = Presentation(MASTER)
    main_slide(prs)
    prs.save(MAIN)
    print(MAIN)


if __name__ == "__main__":
    main()
