"""slide_06 — AI 서버 인프라: Dell PowerEdge R770 × 2대 (스펙 표 + 다크 슬랩).

spec: input/smart-factory-briefing/slides/slide_06.spec.yaml
design: input/smart-factory-briefing/design_system.md (Electric Blue + Ink, 그림자/이모지/zebra 금지)
"""

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from ppt_utils import (
    add_footnote,
    set_title, clear_placeholders,
    add_para, add_accent_bar,
    add_grid_table, estimate_container_height,
    force_font,
)
from template_contract import CONTENT_SAFE

# ---- 디자인 토큰 (design_system.md — 이 목록 밖 색 사용 금지) ----
PRIMARY        = RGBColor(0x02, 0x4A, 0xD8)
PRIMARY_BRIGHT = RGBColor(0x29, 0x6E, 0xF9)
PRIMARY_SOFT   = RGBColor(0xC9, 0xE0, 0xFC)
INK            = RGBColor(0x1A, 0x1A, 0x1A)
CHARCOAL       = RGBColor(0x3D, 0x3D, 0x3D)
GRAPHITE       = RGBColor(0x63, 0x63, 0x63)
CANVAS         = RGBColor(0xFF, 0xFF, 0xFF)
CLOUD          = RGBColor(0xF7, 0xF7, 0xF7)
FOG            = RGBColor(0xE8, 0xE8, 0xE8)
STEEL          = RGBColor(0xC2, 0xC2, 0xC2)
ON_INK         = RGBColor(0xFF, 0xFF, 0xFF)

# ---- 콘텐츠 데이터 (spec.content_blocks — 임의 추가 금지) ----
HEADER = ["구성 항목", "R770 (2-way GPU)", "R770 (1-way GPU)"]
BODY = [
    ("용도", "vLLM 서빙 (Gemma-4-31B · Qwen3-Embedding-8B)", "학습 · 실험"),
    ("GPU", "RTX PRO 6000 Blackwell 96GB × 2", "RTX PRO 6000 Blackwell 96GB × 1"),
    ("CPU", "Xeon 6 6515P 16C/32T × 2", "Xeon 6 6520P 24C/48T × 2"),
    ("메모리", "32GB DDR5-6400 × 8 = 256GB", "32GB DDR5-6400 × 16 = 512GB"),
    ("스토리지", "3.84TB SATA SSD × 2", "7.68TB NVMe SSD × 2"),
    ("네트워크", "10GbE Base-T × 4 (OCP 3.0)", "10GbE Base-T × 4 (OCP 3.0)"),
    ("전원", "3200W Titanium 이중화 (1+1)", "3200W Titanium 이중화 (1+1)"),
]
LEAD_SEGS = [
    ("RTX PRO 6000 Blackwell 96GB ", INK),
    ("총 3장", PRIMARY),
    (" — 서빙 · 학습 ", INK),
    ("분리 2대", PRIMARY),
    (" 구성", INK),
]
METRICS = [
    ("도입처", "(주)아이웍스", ON_INK),
    ("GPU 총량", "3장 · VRAM 총 288GB", ON_INK),
    ("합계 금액", "2대 · 약 8,000만원", PRIMARY_BRIGHT),  # 금액만 PRIMARY_BRIGHT
]

# 표 셀 내부 여백 (marL, marT, marR, marB)
_CELL_MARGINS = (Inches(0.10), Inches(0.03), Inches(0.08), Inches(0.03))
_CELL_PAD_W = int(Inches(0.20))   # 좌우 여백 합 — 줄바꿈 폭 계산용
_HEADER_PT = 10.5
_BODY_PT = 10
_MONO_PT = 9.5


def _cell_min_h(text, font_pt, cell_w, floor_in):
    """셀 텍스트 기반 최소 행 높이 (넘침 금지) — 바닥값과 비교해 큰 쪽."""
    need = estimate_container_height(
        text, font_pt,
        max_width=max(int(cell_w) - _CELL_PAD_W, int(Inches(0.5))),
        padding_top=Inches(0.075), padding_bottom=Inches(0.075),
    )
    return max(int(Inches(floor_in)), int(need))


def _gpu_segments(qty):
    """GPU 행 값 — 모델명 SemiBold INK 강조 + 수량만 JetBrains Mono (과용 금지)."""
    return [
        {"text": "RTX PRO 6000 Blackwell 96GB", "font_name": "Pretendard SemiBold",
         "color": INK, "font_size": _BODY_PT, "bold": False},
        {"text": "  × " + qty, "font_name": "JetBrains Mono",
         "color": INK, "font_size": _MONO_PT, "bold": True},
    ]


def build_slide_6(slide):
    set_title(slide, "AI 서버 인프라 — Dell PowerEdge R770 × 2대",
              font_size=20, bold=True)
    clear_placeholders(slide, keep=[0])

    left = int(CONTENT_SAFE.left)
    width = int(CONTENT_SAFE.width)

    # ---- 하단 고정 요소 역산: 다크 슬랩(안전영역 하단 앵커) → 표 가용 높이 ----
    # rev3: 각주 삭제 — 슬랩을 CONTENT_SAFE 하단에 붙이고, 남는 세로 공간은
    # 표 높이(table_h)로 흡수 (add_grid_table 이 row_heights 비율로 행 높이 재배분)
    slab_h = int(Inches(0.58))
    fn_gap = int(Inches(0.32))          # 하단 vLLM 각주 1줄 공간 (rev11 — 가림 해소 여유 확대)
    slab_top = int(CONTENT_SAFE.bottom) - slab_h - fn_gap

    # ---- 리드 (표준 규격: y=0.72" h=0.35", 15pt bold INK, 키워드만 PRIMARY) ----
    lead_top = int(Inches(0.72))
    lead_h = int(Inches(0.35))
    lead_box = slide.shapes.add_textbox(left, lead_top, width, lead_h)
    tf = lead_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    for text, color in LEAD_SEGS:
        run = p.add_run()
        run.text = text
        run.font.size = Pt(15)
        run.font.bold = True
        run.font.color.rgb = color

    # ---- 스펙 표 add_grid_table 3열 8행 (헤더 + 7행) ----
    table_top = lead_top + int(lead_h) + int(Inches(0.15))
    table_h = slab_top - int(Inches(0.16)) - table_top

    label_w = int(Inches(1.45))
    data_w = (width - label_w) // 2

    # 행 높이 — 텍스트 기반 산출 (병합 행은 2열 폭 기준)
    row_min = [max(
        _cell_min_h(HEADER[0], _HEADER_PT, label_w, 0.40),
        _cell_min_h(HEADER[1], _HEADER_PT, data_w, 0.40),
        _cell_min_h(HEADER[2], _HEADER_PT, data_w, 0.40),
    )]
    for label, v1, v2 in BODY:
        val_w = data_w * 2 if v1 == v2 else data_w
        row_min.append(max(
            _cell_min_h(label, _BODY_PT, label_w, 0.42),
            _cell_min_h(v1, _BODY_PT, val_w, 0.42),
            _cell_min_h(v2, _BODY_PT, val_w, 0.42),
        ))
    table_h = max(int(table_h), sum(row_min))  # 넘침 방지 가드 (정상 시 budget이 더 큼)

    cells = {}
    for c, text in enumerate(HEADER):
        cells[(0, c)] = dict(
            text=text, fill=CLOUD,
            font_name="Pretendard SemiBold", font_color=INK,
            font_size=_HEADER_PT, bold=False, align="l", anchor="ctr",
            border_edges="b", border_color=STEEL, border_width_pt=1.0,
            margins=_CELL_MARGINS,
        )
    for i, (label, v1, v2) in enumerate(BODY):
        r = i + 1
        row_fill = PRIMARY_SOFT if label == "용도" else CANVAS   # '용도' 행만 PRIMARY_SOFT
        val_color = INK if label == "용도" else CHARCOAL
        cells[(r, 0)] = dict(
            text=label, fill=row_fill,
            font_name="Pretendard SemiBold", font_color=INK,
            font_size=_BODY_PT, bold=False, align="l", anchor="ctr",
            border_edges="b", border_color=FOG, border_width_pt=0.75,  # 수평 hairline
            margins=_CELL_MARGINS,
        )
        base = dict(
            fill=row_fill, font_name="Pretendard", font_color=val_color,
            font_size=_BODY_PT, bold=False, align="l", anchor="ctr",
            border_edges="b", border_color=FOG, border_width_pt=0.75,
            margins=_CELL_MARGINS,
        )
        if label == "GPU":
            cells[(r, 1)] = dict(base, segments=_gpu_segments("2"))
            cells[(r, 2)] = dict(base, segments=_gpu_segments("1"))
        elif v1 == v2:
            cells[(r, 1)] = dict(base, text=v1, span=(1, 2))     # 동일 값 → 좌우 병합
        else:
            cells[(r, 1)] = dict(base, text=v1)
            cells[(r, 2)] = dict(base, text=v2)

    add_grid_table(
        slide, left, table_top, width, table_h,
        nrows=1 + len(BODY), ncols=3, cells=cells,
        col_widths=[label_w, data_w, data_w], row_heights=row_min,
        gridlines=False,                # 수직선 없음 — 수평 hairline만 셀별 지정
        default_font_size=_BODY_PT,
    )

    # ---- 하단 다크 슬랩 (INK 배경 · 마감 밴드 1곳) ----
    slab = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, slab_top, width, slab_h)
    slab.fill.solid()
    slab.fill.fore_color.rgb = INK
    slab.line.fill.background()
    slab.shadow.inherit = False         # 그림자 금지

    col_w = width // 3
    for i, (label, value, val_color) in enumerate(METRICS):
        box = slide.shapes.add_textbox(left + i * col_w, slab_top, col_w, slab_h)
        mtf = box.text_frame
        mtf.word_wrap = True
        mtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        mtf.margin_top = 0
        mtf.margin_bottom = 0
        p0 = mtf.paragraphs[0]
        p0.alignment = PP_ALIGN.CENTER   # 중앙 정렬은 다크 슬랩 문구만 허용
        run = p0.add_run()
        run.text = label
        run.font.name = "Pretendard"
        run.font.size = Pt(9)
        run.font.bold = False
        run.font.color.rgb = FOG
        add_para(mtf, value, font_name="Pretendard SemiBold", font_size=13,
                 color=val_color, bold=False, align=PP_ALIGN.CENTER,
                 space_before=Pt(2))
    for i in (1, 2):                     # 지표 사이 수직 hairline
        add_accent_bar(slide, left + i * col_w, slab_top + int(Inches(0.10)),
                       int(Pt(1)), slab_h - int(Inches(0.20)), GRAPHITE)

    # ---- 폰트 일괄 통일 (design_system.md 2절 — 맑은 고딕 단일) ----
    add_footnote(slide, "vLLM: 오픈소스 LLM 서빙 엔진 — 사내 GPU 서버에서 모델을 API 로 제공",
                 font_size=9, color=GRAPHITE)
    force_font(slide, "맑은 고딕")
