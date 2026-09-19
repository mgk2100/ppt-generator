#!/usr/bin/env python3
"""6단계 아이콘 6종 — 플랫 벡터 SVG → PNG (덱 팔레트에 맞춘 선화)

톤은 input/sw-dev-ai-llm/claude_icons 와 같다: 네이비 외곽선 + 파랑 채움 + 금색 포인트 1곳.
python3 make_icons.py  →  svg/*.svg, png/*.png
"""
import os, subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
SVG = os.path.join(HERE, "svg"); PNG = os.path.join(HERE, "png")
NAVY = "#1B2E53"; BLUE = "#0E5DB7"; LIGHT = "#A6CAEC"; PALE = "#F2F6FC"
GOLD = "#D9A93E"; GREY = "#93A1B5"; WHITE = "#FFFFFF"
SW = 5.5                                    # 기본 선 굵기 — 6종 공통

def svg(body, vb=200):
    return ('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 %d %d" width="%d" height="%d" '
            'fill="none" stroke-linecap="round" stroke-linejoin="round">%s</svg>' % (vb, vb, vb, vb, body))

def rr(x, y, w, h, r=8, fill=PALE, stroke=NAVY, sw=SW):
    return ('<rect x="%g" y="%g" width="%g" height="%g" rx="%g" fill="%s" stroke="%s" stroke-width="%g"/>'
            % (x, y, w, h, r, fill, stroke, sw))

def line(x1, y1, x2, y2, stroke=NAVY, sw=SW):
    return '<line x1="%g" y1="%g" x2="%g" y2="%g" stroke="%s" stroke-width="%g"/>' % (x1, y1, x2, y2, stroke, sw)

def circ(cx, cy, r, fill=BLUE, stroke="none", sw=0):
    return '<circle cx="%g" cy="%g" r="%g" fill="%s" stroke="%s" stroke-width="%g"/>' % (cx, cy, r, fill, stroke, sw)

def path(d, fill="none", stroke=NAVY, sw=SW):
    return '<path d="%s" fill="%s" stroke="%s" stroke-width="%g"/>' % (d, fill, stroke, sw)

def arrow(x1, y1, x2, y2, stroke=BLUE, sw=SW, head=11):
    """직선 화살표 — 끝점에 머리"""
    import math
    a = math.atan2(y2 - y1, x2 - x1)
    p1 = (x2 - head * math.cos(a - 0.45), y2 - head * math.sin(a - 0.45))
    p2 = (x2 - head * math.cos(a + 0.45), y2 - head * math.sin(a + 0.45))
    return (line(x1, y1, x2, y2, stroke, sw) +
            path("M%g %g L%g %g L%g %g" % (p1[0], p1[1], x2, y2, p2[0], p2[1]), stroke=stroke, sw=sw))

def cylinder(x, y, w, h, fill=WHITE, top=LIGHT):
    ry = w * 0.17
    s = path("M%g %g V%g A%g %g 0 0 0 %g %g V%g" % (x, y + ry, y + h - ry, w / 2, ry, x + w, y + h - ry, y + ry),
             fill=fill)
    s += ('<ellipse cx="%g" cy="%g" rx="%g" ry="%g" fill="%s" stroke="%s" stroke-width="%g"/>'
          % (x + w / 2, y + ry, w / 2, ry, top, NAVY, SW))
    return s

def doc(x, y, w=44, h=56, fold=13, fill=WHITE, lines=3):
    s = path("M%g %g H%g L%g %g V%g H%g Z" % (x, y, x + w - fold, x + w, y + fold, y + h, x), fill=fill)
    s += path("M%g %g V%g H%g" % (x + w - fold, y, y + fold, x + w), sw=SW * 0.8)
    for i in range(lines):
        s += line(x + 9, y + 24 + i * 11, x + w - 11 - (10 if i == lines - 1 else 0), y + 24 + i * 11,
                  stroke=LIGHT, sw=SW * 0.9)
    return s

def gear(cx, cy, r=27, teeth=8, fill=PALE):
    import math
    d = ""
    for i in range(teeth):
        a = 2 * math.pi * i / teeth
        x1, y1 = cx + (r + 1) * math.cos(a), cy + (r + 1) * math.sin(a)
        x2, y2 = cx + (r + 11) * math.cos(a), cy + (r + 11) * math.sin(a)
        d += line(x1, y1, x2, y2, stroke=NAVY, sw=SW * 1.5)
    return d + circ(cx, cy, r, fill=fill, stroke=NAVY, sw=SW) + circ(cx, cy, 9, fill=BLUE)


# ── 1. 데이터 취득 — 세 갈래 자료가 한 저장소로 모인다 ─────────
def i1():
    s = doc(18, 20, 40, 50, lines=3)
    s += rr(80, 20, 40, 50, r=7, fill=WHITE)
    s += path("M92 40 L84 46 L92 52", stroke=BLUE, sw=SW * 0.9) + path("M108 40 L116 46 L108 52", stroke=BLUE, sw=SW * 0.9)
    s += path("M142 24 H182 V58 H160 L150 70 V58 H142 Z", fill=WHITE)
    s += line(152, 36, 172, 36, stroke=LIGHT, sw=SW * 0.9) + line(152, 47, 166, 47, stroke=LIGHT, sw=SW * 0.9)
    s += arrow(38, 78, 86, 112) + arrow(100, 78, 100, 110) + arrow(162, 78, 114, 112)
    s += cylinder(62, 118, 76, 62)
    s += circ(100, 152, 7, fill=GOLD)
    return s

# ── 2. 데이터 가공 — 한 덩어리를 작은 조각으로 나눈다 ───────────
def i2():
    s = rr(16, 62, 58, 76, r=9, fill=WHITE)
    for i in range(4):
        s += line(28, 78 + i * 16, 62 - (12 if i == 3 else 0), 78 + i * 16, stroke=LIGHT, sw=SW * 0.9)
    s += arrow(84, 100, 116, 100)
    for r_ in range(2):
        for c in range(2):
            x, y = 128 + c * 34, 66 + r_ * 38
            s += rr(x, y, 28, 30, r=6, fill=PALE if (r_ + c) % 2 else WHITE)
    s += rr(128, 142, 62, 30, r=6, fill=WHITE)
    s += circ(159, 157, 7, fill=GOLD)
    return s

# ── 3. AI 모델 확보 — 사 오는 경로와 만드는 경로 ───────────────
def i3():
    s = path("M32 58 H82 V88 H32 Z", fill=WHITE)          # 구매 : 포장 상자(몸통)
    s += path("M26 40 H88 V58 H26 Z", fill=PALE)         # 뚜껑
    s += line(57, 40, 57, 58, sw=SW * 0.9)               # 뚜껑 테이프
    s += gear(156, 63, r=21, teeth=8)                      # 제작 : 톱니
    s += arrow(57, 96, 84, 122) + arrow(156, 96, 124, 122)
    s += rr(70, 128, 62, 54, r=10, fill=PALE)              # 확보된 모델
    for i, x in enumerate((70, 132)):
        for y in (142, 158):
            s += line(x - 12 if i == 0 else x, y, x if i == 0 else x + 12, y, stroke=NAVY, sw=SW * 0.9)
    s += circ(101, 155, 12, fill=BLUE) + circ(101, 155, 5, fill=GOLD)
    return s

# ── 4. 서빙 환경 구성 — 모델을 서버와 현장 장비에 올린다 ────────
def i4():
    s = rr(72, 14, 56, 44, r=9, fill=PALE)                 # 모델
    s += circ(100, 36, 11, fill=BLUE) + circ(100, 36, 4.5, fill=GOLD)
    s += arrow(86, 66, 62, 96) + arrow(114, 66, 138, 96)
    s += rr(16, 102, 74, 78, r=9, fill=WHITE)              # 사내 서버
    for i in range(3):
        s += line(30, 120 + i * 20, 68, 120 + i * 20, stroke=LIGHT, sw=SW * 0.9)
        s += circ(78, 120 + i * 20, 4.5, fill=BLUE)
    s += rr(110, 110, 74, 52, r=9, fill=WHITE)             # 현장 장비
    s += rr(124, 124, 46, 24, r=5, fill=LIGHT, sw=0)
    s += line(132, 162, 162, 162) + line(147, 162, 147, 176)
    return s

# ── 5. 성능 평가 · 개선 — 측정하고 되돌아가 고친다 ──────────────
def i5():
    s = path("M40 128 A60 60 0 0 1 160 128", stroke=LIGHT, sw=SW * 2.2)
    s += path("M40 128 A60 60 0 0 1 72 76", stroke=BLUE, sw=SW * 2.2)
    s += line(100, 128, 132, 92, stroke=NAVY, sw=SW * 1.2) + circ(100, 128, 9, fill=NAVY)
    s += circ(160, 128, 6, fill=GOLD)
    s += path("M158 150 A62 62 0 0 1 48 154", stroke=BLUE, sw=SW * 1.2)
    s += path("M60 142 L48 155 L62 166", stroke=BLUE, sw=SW * 1.2)
    return s

# ── 6. 배포 · 운영 — 현장으로 퍼지고 자료가 되돌아온다 ──────────
def i6():
    s = path("M100 26 A74 74 0 1 1 46 52", stroke=LIGHT, sw=SW * 1.3)
    s += path("M88 14 L100 26 L88 38", stroke=LIGHT, sw=SW * 1.3)
    s += rr(74, 78, 52, 44, r=9, fill=PALE)
    s += circ(100, 100, 10, fill=BLUE) + circ(100, 100, 4, fill=GOLD)
    s += arrow(80, 126, 52, 152) + arrow(100, 126, 100, 154) + arrow(120, 126, 148, 152)
    for cx in (40, 100, 160):
        s += rr(cx - 20, 156, 40, 28, r=6, fill=WHITE)
        s += line(cx - 10, 170, cx + 10, 170, stroke=LIGHT, sw=SW * 0.9)
    return s


ICONS = [("step1", i1), ("step2", i2), ("step3", i3), ("step4", i4), ("step5", i5), ("step6", i6)]


def main():
    os.makedirs(SVG, exist_ok=True); os.makedirs(PNG, exist_ok=True)
    for name, fn in ICONS:
        sp = os.path.join(SVG, name + ".svg")
        open(sp, "w", encoding="utf-8").write(svg(fn()))
        subprocess.run(["rsvg-convert", "-w", "600", "-h", "600", sp,
                        "-o", os.path.join(PNG, name + ".png")], check=True)
        print(os.path.join(PNG, name + ".png"))


if __name__ == "__main__":
    main()
