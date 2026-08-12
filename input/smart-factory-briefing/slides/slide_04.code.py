"""slide_04 — 전체 아키텍처 (2026-08-11 실코드 분석 기반 신규 구도, rev5 레이어 박스화).

레이어(위→아래): 사용자 컨테이너(2블록) → SL SW Agent 컨테이너(서비스 3블록 +
내부 커넥터 2) → 지식·저장소 컨테이너(4블록) → LLM 컨테이너(2블록, 상용 강조)
→ 각주 1줄. 4개 레이어 전부 동일 컨테이너 문법(CLOUD 배경 + 좌상단 bold 라벨).
커넥터 7개 — 실선 PRIMARY 1.5pt + 점선 STEEL 1pt 2종만.
rev5: '사용자'·'지식 · 저장소'·'LLM' 레이어를 SL SW Agent 와 동일한 컨테이너
박스로 통일(기존 텍스트 라벨은 박스 라벨로 흡수), coding-assistant →
ca-analysis-pipeline 라벨 '코드 분석 요청' 1줄화, 파일 저장소·MongoDB 부제 교체.
rev6: 컨테이너 라벨 좌측 accent 바 삭제, '문서 생성 요청' 커넥터 라벨 1행 표기.
design_system.md 토큰 강제 — 이모지·그림자·CHEVRON 금지, 마지막 force_font.
"""

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.ns import qn

from ppt_utils import (
    set_title, clear_placeholders,
    add_textbox, add_para, add_arrowhead, add_footnote,
    set_body_anchor, set_text_inset, calc_grid, align_shapes,
    force_font,
)
from template_contract import CONTENT_SAFE

# ---- 디자인 토큰 (design_system.md — 이 목록 밖 색 사용 금지) ----
PRIMARY      = RGBColor(0x02, 0x4A, 0xD8)
PRIMARY_SOFT = RGBColor(0xC9, 0xE0, 0xFC)
INK          = RGBColor(0x1A, 0x1A, 0x1A)
CHARCOAL     = RGBColor(0x3D, 0x3D, 0x3D)
GRAPHITE     = RGBColor(0x63, 0x63, 0x63)
CANVAS       = RGBColor(0xFF, 0xFF, 0xFF)
CLOUD        = RGBColor(0xF7, 0xF7, 0xF7)
FOG          = RGBColor(0xE8, 0xE8, 0xE8)
STEEL        = RGBColor(0xC2, 0xC2, 0xC2)

# ---- 수직 레이아웃 (CONTENT_SAFE y=0.68"~7.02") ----
# rev5: 4개 레이어 전부 컨테이너 박스화 — 라벨 존 0.32" + 내부 블록 + 하단 inset
Y_UBOX  = Inches(1.10); H_UBOX  = Inches(0.82)   # 사용자 (bottom 1.92)
Y_USERS = Inches(1.42); H_USERS = Inches(0.42)
Y_AGENT = Inches(2.20); H_AGENT = Inches(1.22)   # SL SW Agent (bottom 3.42)
Y_SVC   = Inches(2.52); H_SVC   = Inches(0.80)
Y_SBOX  = Inches(3.70); H_SBOX  = Inches(1.30)   # 지식 · 저장소 (bottom 5.00)
Y_STORE = Inches(4.02); H_STORE = Inches(0.88)
Y_LBOX  = Inches(5.28); H_LBOX  = Inches(1.26)   # LLM (bottom 6.54 — 각주 2줄 공간)
Y_LLM   = Inches(5.58); H_LLM   = Inches(0.88)

SERVICES = [  # 좌→우 (coding-assistant 를 허브로 중앙 배치) — rev3 부제 간략화
    ("doc-generator", "SW 설계 문서 생성"),
    ("coding-assistant", "질의 · 코딩 지원"),
    ("ca-analysis-pipeline", "코드 분석 엔진"),
]

STORES = [  # 실제 저장 내용 — 코드베이스 확정 사실 (임의 변경 금지) — rev5 문구
    ("PostgreSQL (pgvector)", "RAG 벡터 검색 · 대화 이력\n· 사용자 계정"),
    ("파일 저장소 (JSON · md)", "함수 · 변수 목록 · 호출 관계도\n· 모듈별 역할 요약"),
    ("MongoDB", "파싱된 소스 코드\n· 섹션별 중간 문서 (JSON)"),
    ("MinIO", "업로드 zip · 다이어그램 이미지\n· 최종 문서 파일 (DOCX)"),
]

LLMS = [  # rev4: 역할 표현 명확화
    ("상용 LLM — Claude · GPT",
     "사내 스킬 · 지식 연계 → SL SW 최적화 답변", True),
    ("사내 vLLM 서빙 (GPU 서버)",
     "Gemma-4-31B — 코드 분석 리포트 문구 작성\nQwen3-Embedding-8B — 검색용 임베딩 생성", False),
]

FOOTNOTES = [
    "pgvector: PostgreSQL 벡터 검색 확장 — RAG 임베딩 저장 · 유사도 검색",
    "지식증류: 고성능 AI(teacher)의 결과물로 소형 AI(student)를 학습시켜 성능을 근접시키는 기법",
]


def _block(slide, x, y, w, h, name, sub, name_size=11.5, sub_size=9.5,
           emphasized=False):
    """다이어그램 블록 — 이름 bold 1줄 + 부제 2줄. hairline 보더 (그림자 금지)."""
    box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, int(x), int(y), int(w), int(h))
    box.adjustments[0] = 0.07
    box.fill.solid()
    box.fill.fore_color.rgb = PRIMARY_SOFT if emphasized else CANVAS
    box.line.color.rgb = PRIMARY if emphasized else FOG
    box.line.width = Pt(1.0)
    box.shadow.inherit = False
    tf = box.text_frame
    tf.word_wrap = True
    set_text_inset(box, left=Inches(0.12), top=Inches(0.04),
                   right=Inches(0.12), bottom=Inches(0.04))
    set_body_anchor(box, "ctr")
    p0 = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.LEFT
    run = p0.add_run()
    run.text = name
    run.font.size = Pt(name_size)
    run.font.bold = True
    run.font.color.rgb = INK
    if sub:
        p = add_para(tf, sub, font_size=sub_size, color=CHARCOAL,
                     space_before=Pt(2))
        p.line_spacing = Pt(sub_size * 1.35)
    return box


def _conn(slide, x1, y1, x2, y2, color, width_pt, dash=None, arrow=True,
          both_heads=False):
    """커넥터 2종만 — 실선 PRIMARY 1.5pt / 점선 STEEL 1pt."""
    conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, int(x1), int(y1), int(x2), int(y2))
    conn.line.color.rgb = color
    conn.line.width = Pt(width_pt)
    ln = conn.line._get_or_add_ln()
    if dash:  # prstDash 는 headEnd/tailEnd 앞에 위치해야 함 — arrowhead 전에 삽입
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": dash}))
    if arrow:
        add_arrowhead(conn)
    if both_heads:
        he = ln.makeelement(
            qn("a:headEnd"), {"type": "triangle", "w": "med", "len": "med"})
        tail = ln.find(qn("a:tailEnd"))
        if tail is not None:
            tail.addprevious(he)
        else:
            ln.append(he)
    return conn


def _label(slide, x, y, w, h, text, size=8, align=PP_ALIGN.CENTER,
           color=GRAPHITE, bold=False):
    """커넥터/밴드 라벨 — inset 0 소형 텍스트박스."""
    box = add_textbox(slide, int(x), int(y), int(w), int(h), text,
                      font_size=size, color=color, bold=bold, align=align)
    set_text_inset(box, 0, 0, 0, 0)
    return box


def _container(slide, y, h, label):
    """레이어 컨테이너 박스 — CLOUD 배경 + 좌상단 bold 라벨.

    rev5: SL SW Agent 문법을 전 레이어(사용자/지식·저장소/LLM)에 통일 적용.
    rev6: 라벨 좌측 파란 accent 바 삭제 — 라벨 x 를 바 자리로 당겨 정렬.
    """
    box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, int(CONTENT_SAFE.left), int(y),
        int(CONTENT_SAFE.width), int(h))
    box.adjustments[0] = 0.03
    box.fill.solid()
    box.fill.fore_color.rgb = CLOUD
    box.line.fill.background()
    box.shadow.inherit = False
    _label(slide, int(CONTENT_SAFE.left) + int(Inches(0.14)),
           int(y) + int(Inches(0.02)), Inches(3.0), Inches(0.26), label,
           size=11.5, align=PP_ALIGN.LEFT, color=INK, bold=True)
    return box


def build_slide_04(slide):
    set_title(slide, "전체 아키텍처", font_size=20, color=INK, bold=True)
    clear_placeholders(slide, keep=[0])

    L = int(CONTENT_SAFE.left)
    W = int(CONTENT_SAFE.width)
    R = int(CONTENT_SAFE.right)

    # ---- 1) 리드 — 표준 규격 (y=0.72", h=0.35") 15pt bold INK, 키워드 PRIMARY ----
    lead = slide.shapes.add_textbox(L, Inches(0.72), W, Inches(0.35))
    lead_tf = lead.text_frame
    lead_tf.word_wrap = True
    set_text_inset(lead, left=0, top=Inches(0.02), right=0, bottom=Inches(0.02))
    lead_p = lead_tf.paragraphs[0]
    lead_p.alignment = PP_ALIGN.LEFT
    for text, color in [
        ("3개 서비스가 하나의 ", INK),
        ("SL SW Agent", PRIMARY),
        (" — 저장소 · LLM 을 ", INK),
        ("역할별", PRIMARY),
        ("로 분담", INK),
    ]:
        run = lead_p.add_run()
        run.text = text
        run.font.size = Pt(15)
        run.font.bold = True
        run.font.color.rgb = color

    # ---- 2) 사용자 컨테이너 — 박스 라벨 '사용자' + 내부 블록 2개 ----
    _container(slide, Y_UBOX, H_UBOX, "사용자")
    user_w = Inches(2.70)
    u_web = _block(slide, Inches(2.07), Y_USERS, user_w, H_USERS,
                   "웹 UI", None, name_size=11)
    u_ide = _block(slide, Inches(6.07), Y_USERS, user_w, H_USERS,
                   "IDE 플러그인 (Claude Code)", None, name_size=11)
    align_shapes(u_web, u_ide, axis='h')

    # ---- 3) SL SW Agent 컨테이너 — 동일 문법 + 내부 서비스 3블록 ----
    _container(slide, Y_AGENT, H_AGENT, "SL SW Agent")

    # 내부 서비스 블록 3개 (허브 = 중앙 coding-assistant)
    svc_area = (L + Inches(0.18), Y_SVC, W - Inches(0.36), H_SVC)
    svc_grid = calc_grid(1, 3, area=svc_area, gap=Inches(0.84))
    svc = []
    for cell, (name, sub) in zip(svc_grid[0], SERVICES):
        svc.append(_block(slide, cell.left, cell.top, cell.width, cell.height,
                          name, sub, name_size=12, sub_size=9.5))
    svc_dg, svc_ca, svc_cap = svc
    align_shapes(svc_dg, svc_ca, svc_cap, axis='h')

    # 내부 커넥터 2 — coding-assistant → 양쪽 (실선 PRIMARY 1.5pt)
    cy = svc_ca.top + svc_ca.height // 2
    _conn(slide, svc_ca.left, cy, svc_dg.left + svc_dg.width, cy,
          PRIMARY, 1.5)
    _conn(slide, svc_ca.left + svc_ca.width, cy, svc_cap.left, cy,
          PRIMARY, 1.5)
    gap_a_x = svc_dg.left + svc_dg.width
    gap_a_w = svc_ca.left - gap_a_x
    gap_b_x = svc_ca.left + svc_ca.width
    gap_b_w = svc_cap.left - gap_b_x
    # rev6: '문서 생성 요청' 1행 표기 — 폭을 갭 좌우로 넓혀 줄바꿈 방지
    _label(slide, gap_a_x - Inches(0.10), cy - Inches(0.26),
           gap_a_w + Inches(0.20), Inches(0.20), "문서 생성 요청")
    _label(slide, gap_b_x, cy - Inches(0.26), gap_b_w, Inches(0.20),
           "코드 분석 요청")

    # ---- 4) 사용자 → 컨테이너 화살표 2 (웹 'HTTPS' / IDE 'MCP' 라벨) ----
    u_bottom = Y_UBOX + H_UBOX
    for u in (u_web, u_ide):
        ux = u.left + u.width // 2
        _conn(slide, ux, u_bottom, ux, Y_AGENT, PRIMARY, 1.5)
    _label(slide, u_web.left + u_web.width // 2 + Inches(0.08), Inches(1.96),
           Inches(0.55), Inches(0.18), "HTTPS", size=8.5, align=PP_ALIGN.LEFT)
    _label(slide, u_ide.left + u_ide.width // 2 + Inches(0.08), Inches(1.96),
           Inches(0.55), Inches(0.18), "MCP", size=8.5, align=PP_ALIGN.LEFT)

    # ---- 5) 지식 · 저장소 컨테이너 — 박스 라벨 흡수 + 내부 블록 4개 ----
    _container(slide, Y_SBOX, H_SBOX, "지식 · 저장소")
    store_area = (L + Inches(0.18), Y_STORE, W - Inches(0.36), H_STORE)
    store_grid = calc_grid(1, 4, area=store_area, gap=Inches(0.16))
    stores = []
    for cell, (name, sub) in zip(store_grid[0], STORES):
        stores.append(_block(slide, cell.left, cell.top, cell.width,
                             cell.height, name, sub, name_size=11, sub_size=9))
    align_shapes(*stores, axis='h')

    # 컨테이너 ↕ 저장소 컨테이너 양방향 1 — 라벨 '검색 · 저장'
    mid_x = (L + R) // 2
    _conn(slide, mid_x, Y_AGENT + H_AGENT, mid_x, Y_SBOX,
          PRIMARY, 1.5, both_heads=True)
    _label(slide, mid_x + Inches(0.10), Inches(3.46), Inches(1.00),
           Inches(0.18), "검색 · 저장", size=8.5, align=PP_ALIGN.LEFT)

    # ---- 6) LLM 컨테이너 — 동일 문법 + 블록 2개 (상용 강조) ----
    _container(slide, Y_LBOX, H_LBOX, "LLM")
    llm_x1 = L + int(Inches(0.18))
    llm_w1 = Inches(4.30)
    llm_gap = Inches(1.10)
    llm_x2 = llm_x1 + int(llm_w1) + int(llm_gap)
    llm_w2 = R - int(Inches(0.18)) - llm_x2
    name1, sub1, emp1 = LLMS[0]
    name2, sub2, emp2 = LLMS[1]
    llm_com = _block(slide, llm_x1, Y_LLM, llm_w1, H_LLM, name1, sub1,
                     name_size=11.5, sub_size=9.5, emphasized=emp1)
    llm_int = _block(slide, llm_x2, Y_LLM, llm_w2, H_LLM, name2, sub2,
                     name_size=11.5, sub_size=9.5, emphasized=emp2)
    align_shapes(llm_com, llm_int, axis='h')

    # 저장소 컨테이너 → 상용 LLM (요청 · 데이터 흐름) 실선 1
    com_cx = llm_com.left + llm_com.width // 2
    _conn(slide, com_cx, Y_SBOX + H_SBOX, com_cx, Y_LBOX, PRIMARY, 1.5)

    # 상용 ⇢ 사내 vLLM 점선 STEEL 1pt — 라벨 '지식증류 데이터' (rev11 용어 통일)
    llm_cy = llm_com.top + llm_com.height // 2
    _conn(slide, llm_com.left + llm_com.width, llm_cy, llm_int.left, llm_cy,
          STEEL, 1.0, dash="dash")
    _label(slide, llm_com.left + llm_com.width, llm_cy - Inches(0.28),
           llm_gap, Inches(0.22), "지식증류 데이터")

    # ---- 7) 각주 1줄 (rev3: 도구 캡션 삭제, pgvector 각주만 유지) ----
    # 각주 2줄 — 항목당 1줄 (design_system 2-2), CONTENT_SAFE 하단 밀착
    from ppt_utils import CONTENT_SAFE as _CS
    from pptx.util import Inches as _In
    _fn_h = _In(0.38)
    _fb = slide.shapes.add_textbox(_CS.left, _CS.bottom - _fn_h, _CS.width, _fn_h)
    _ftf = _fb.text_frame; _ftf.word_wrap = True
    for _j, _fn in enumerate(FOOTNOTES):
        _p = _ftf.paragraphs[0] if _j == 0 else _ftf.add_paragraph()
        _r = _p.add_run(); _r.text = "※ " + _fn
        _r.font.size = Pt(9); _r.font.color.rgb = GRAPHITE

    # ---- 8) 폰트 일괄 강제 — 맑은 고딕 (design_system.md 2절) ----
    force_font(slide, "맑은 고딕")
