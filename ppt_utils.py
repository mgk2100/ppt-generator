"""
ppt_utils.py — 비시각적 유틸리티
디자인을 제약하는 코드 없음. 인프라 헬퍼만 포함.
"""

import os
import shutil
import subprocess
import platform
import unicodedata
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR_TYPE
from pptx.oxml.ns import qn


BASE_DIR = Path(__file__).parent
REF_DIR = BASE_DIR / "ref"
OUTPUT_DIR = BASE_DIR / "output"
DEFAULT_TEMPLATE = REF_DIR / "표지.pptx"
FONTS_DIR = REF_DIR / "fonts"

from collections import namedtuple

_SafeZone = namedtuple("_SafeZone", ["left", "top", "width", "height", "right", "bottom"])
CONTENT_SAFE = _SafeZone(
    left=Inches(0.28),
    top=Inches(0.68),
    width=Inches(10.28),
    height=Inches(6.34),
    right=Inches(10.56),
    bottom=Inches(7.02),
)


def ensure_fonts():
    """ref/fonts/ 의 .ttf 폰트를 시스템에 설치한다. 이미 설치됐으면 스킵."""
    if not FONTS_DIR.exists():
        return False

    font_files = list(FONTS_DIR.glob("*.ttf"))
    if not font_files:
        return False

    system = platform.system()

    if system == "Linux":
        dest_dir = Path.home() / ".local" / "share" / "fonts"
    elif system == "Darwin":
        dest_dir = Path.home() / "Library" / "Fonts"
    elif system == "Windows":
        dest_dir = Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "Windows" / "Fonts"
    else:
        return False

    dest_dir.mkdir(parents=True, exist_ok=True)
    installed = False

    for f in font_files:
        dest = dest_dir / f.name
        if not dest.exists():
            shutil.copy2(f, dest)
            installed = True

    if installed and system == "Linux":
        try:
            subprocess.run(["fc-cache", "-fv"], capture_output=True, check=True)
        except (subprocess.CalledProcessError, FileNotFoundError):
            pass

    return True


def load_template(page_numbers=True):
    """표지.pptx를 로드하고 샘플 슬라이드를 제거하여 빈 Presentation을 반환한다."""
    prs = Presentation(str(DEFAULT_TEMPLATE))

    # 샘플 슬라이드 제거
    slide_ids = [slide.slide_id for slide in prs.slides]
    for sid in slide_ids:
        idx = next(
            (i for i, s in enumerate(prs.slides._sldIdLst) if s.id == sid), -1
        )
        if idx >= 0:
            rId = prs.slides._sldIdLst[idx].rId
            prs.part.drop_rel(rId)
            del prs.slides._sldIdLst[idx]

    return prs


def get_layout(prs, name):
    """이름으로 슬라이드 레이아웃을 찾는다. 없으면 ValueError."""
    for layout in prs.slide_masters[0].slide_layouts:
        if layout.name == name:
            return layout
    raise ValueError(f"레이아웃 '{name}'을 찾을 수 없습니다.")


def clear_placeholders(slide, keep=None):
    """마스터 슬라이드에서 상속된 유령 플레이스홀더/텍스트를 제거한다.

    Args:
        slide: 슬라이드 객체
        keep: 유지할 플레이스홀더 idx 리스트
    """
    if keep is None:
        keep = []

    ghost_texts = [
        "마스터 텍스트 스타일 편집",
        "마스터 텍스트 스타일을 편집합니다",
        "마스터 제목 스타일 편집",
        "제목을 추가하려면 클릭하십시오",
        "제목을 입력하십시오",
        "부제목을 입력하십시오",
        "텍스트를 입력하십시오",
        "내용을 입력하십시오",
        "텍스트를 추가하려면 클릭하십시오",
        "Click to edit Master text styles",
        "Click to edit Master title style",
        "Click to add title",
        "Click to add text",
        "Click to add subtitle",
    ]

    to_remove = []

    for ph in list(slide.placeholders):
        if ph.placeholder_format.idx in keep:
            continue
        if ph.has_text_frame:
            text = ph.text_frame.text.strip().rstrip(".")
            if not text or any(g in text or text in g for g in ghost_texts):
                to_remove.append(ph)

    for shape in slide.shapes:
        if shape in to_remove:
            continue
        if shape.is_placeholder:
            continue  # 플레이스홀더는 첫 번째 루프에서 이미 처리됨
        if shape.has_text_frame:
            text = shape.text_frame.text.strip().rstrip(".")
            if any(g in text or text in g for g in ghost_texts):
                to_remove.append(shape)

    for shape in to_remove:
        shape._element.getparent().remove(shape._element)


def set_title(slide, text, font_name=None, font_size=None, color=None, bold=None):
    """TITLE 플레이스홀더(idx=0)에 텍스트를 설정한다.

    None인 파라미터는 테마/마스터에서 상속.
    """
    ph = slide.placeholders[0]
    ph.text = text
    if font_name is not None or font_size is not None or color is not None or bold is not None:
        for run in ph.text_frame.paragraphs[0].runs:
            if font_name is not None:
                run.font.name = font_name
            if font_size is not None:
                run.font.size = Pt(font_size)
            if color is not None:
                run.font.color.rgb = color
            if bold is not None:
                run.font.bold = bold
    return ph


def set_cell_anchor(cell, anchor="ctr"):
    """테이블 셀 세로정렬 XML 워크어라운드.

    Args:
        cell: python-pptx 테이블 셀
        anchor: 't' (위), 'ctr' (가운데), 'b' (아래)
    """
    from lxml import etree

    a_ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    tc = cell._tc

    # tcPr 에 anchor 설정
    tcPr = next((c for c in tc if "tcPr" in c.tag), None)
    if tcPr is None:
        tcPr = etree.Element(f"{a_ns}tcPr")
        tc.insert(0, tcPr)
    tcPr.set("anchor", anchor)

    # txBody > bodyPr 에도 anchor 설정
    txBody = next((c for c in tc if "txBody" in c.tag), None)
    if txBody is None:
        _ = cell.text_frame  # txBody 생성
        txBody = next((c for c in tc if "txBody" in c.tag), None)

    if txBody is not None:
        bodyPr = next((c for c in txBody if "bodyPr" in c.tag), None)
        if bodyPr is None:
            bodyPr = etree.Element(f"{a_ns}bodyPr")
            txBody.insert(0, bodyPr)
        bodyPr.set("anchor", anchor)


def add_arrowhead(connector):
    """커넥터에 화살표 머리를 추가한다 (python-pptx에 네이티브 API 없음)."""
    connector.line._ln.append(
        connector.line._ln.makeelement(
            qn("a:tailEnd"),
            {"type": "triangle", "w": "med", "len": "med"},
        )
    )


# ---------------------------------------------------------------------------
# 신규 유틸리티 — XML 배관 코드 캡슐화 (시각적 의견 없음)
# ---------------------------------------------------------------------------


def add_shadow(shape, blur_pt=4, dist_pt=3, direction=2700000,
               opacity_pct=40, color=None):
    """도형에 outerShadow를 추가한다.

    Args:
        shape: python-pptx 도형 객체
        blur_pt: 블러 반경 (포인트 단위)
        dist_pt: 그림자 거리 (포인트 단위)
        direction: 그림자 방향 (EMU 각도, 기본 270° = 아래)
        opacity_pct: 그림자 불투명도 (0-100)
        color: RGBColor 또는 (r,g,b) 튜플. None이면 검정
    """
    if color is None:
        r, g, b = 0, 0, 0
    elif isinstance(color, RGBColor):
        r, g, b = color[0], color[1], color[2]
    else:
        r, g, b = color

    alpha_val = int(opacity_pct * 1000)  # 40% → 40000
    blur_emu = str(Pt(blur_pt))
    dist_emu = str(Pt(dist_pt))
    hex_color = f"{r:02X}{g:02X}{b:02X}"

    spPr = shape._element.spPr if hasattr(shape._element, 'spPr') else None
    if spPr is None:
        spPr = shape._element.find(qn("p:spPr"))
    if spPr is None:
        return

    # effectLst 찾기/생성
    effectLst = spPr.find(qn("a:effectLst"))
    if effectLst is None:
        effectLst = spPr.makeelement(qn("a:effectLst"), {})
        spPr.append(effectLst)

    # 기존 outerShdw 제거
    for old in effectLst.findall(qn("a:outerShdw")):
        effectLst.remove(old)

    outerShdw = effectLst.makeelement(qn("a:outerShdw"), {
        "blurRad": blur_emu,
        "dist": dist_emu,
        "dir": str(direction),
        "rotWithShape": "0",
    })
    srgbClr = outerShdw.makeelement(qn("a:srgbClr"), {"val": hex_color})
    alphaElem = srgbClr.makeelement(qn("a:alpha"), {"val": str(alpha_val)})
    srgbClr.append(alphaElem)
    outerShdw.append(srgbClr)
    effectLst.append(outerShdw)


def set_shape_opacity(shape, opacity_pct):
    """도형의 채우기 투명도를 설정한다.

    Args:
        shape: python-pptx 도형 객체 (solidFill이 이미 적용된 상태여야 함)
        opacity_pct: 불투명도 (0=완전투명, 100=불투명)
    """
    alpha_val = str(int(opacity_pct * 1000))  # 50% → 50000

    spPr = shape._element.spPr if hasattr(shape._element, 'spPr') else None
    if spPr is None:
        spPr = shape._element.find(qn("p:spPr"))
    if spPr is None:
        return

    solidFill = spPr.find(qn("a:solidFill"))
    if solidFill is None:
        return

    # srgbClr 또는 schemeClr 찾기
    color_elem = solidFill.find(qn("a:srgbClr"))
    if color_elem is None:
        color_elem = solidFill.find(qn("a:schemeClr"))
    if color_elem is None:
        return

    # 기존 alpha 제거 후 새로 추가
    for old in color_elem.findall(qn("a:alpha")):
        color_elem.remove(old)
    alpha_elem = color_elem.makeelement(qn("a:alpha"), {"val": alpha_val})
    color_elem.append(alpha_elem)


def add_gradient_stop(shape, position, r, g, b):
    """그라디언트 fill에 추가 stop을 삽입한다.

    shape.fill.gradient() 로 기본 2-stop 설정 후,
    3번째 이상의 stop을 추가할 때 사용한다.

    Args:
        shape: python-pptx 도형 (이미 gradient fill 적용된 상태)
        position: 0.0~1.0 (0=시작, 1=끝)
        r, g, b: 정수 0-255
    """
    spPr = shape._element.spPr if hasattr(shape._element, 'spPr') else None
    if spPr is None:
        spPr = shape._element.find(qn("p:spPr"))
    if spPr is None:
        return

    gradFill = spPr.find(qn("a:gradFill"))
    if gradFill is None:
        return

    gsLst = gradFill.find(qn("a:gsLst"))
    if gsLst is None:
        gsLst = gradFill.makeelement(qn("a:gsLst"), {})
        gradFill.insert(0, gsLst)

    pos_val = str(int(position * 100000))  # 0.5 → 50000
    hex_color = f"{r:02X}{g:02X}{b:02X}"

    gs = gsLst.makeelement(qn("a:gs"), {"pos": pos_val})
    srgbClr = gs.makeelement(qn("a:srgbClr"), {"val": hex_color})
    gs.append(srgbClr)
    gsLst.append(gs)


def make_icon_circle(slide, x, y, size, fill_color, text="",
                     font_size=10, font_color=None):
    """원형 아이콘/배지를 생성한다 (OVAL + fill + 중앙정렬 텍스트).

    번호 배지, 상태 표시, 아이콘 대체 등에 사용.
    색상/크기는 파라미터이므로 시각적 의견 없음.

    Args:
        slide: 슬라이드 객체
        x, y: 위치 (Inches/Emu)
        size: 원 지름 (Inches/Emu)
        fill_color: RGBColor
        text: 원 안에 들어갈 텍스트
        font_size: 포인트 단위
        font_color: RGBColor. None이면 brightness_check로 자동 결정

    Returns:
        생성된 도형 객체
    """
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, size, size)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()  # 테두리 없음

    if text:
        tf = shape.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = str(text)
        run.font.size = Pt(font_size)
        run.font.bold = True

        if font_color is None:
            is_bright = brightness_check(fill_color[0], fill_color[1], fill_color[2])
            run.font.color.rgb = RGBColor(0x33, 0x33, 0x33) if is_bright else RGBColor(0xFF, 0xFF, 0xFF)
        else:
            run.font.color.rgb = font_color

        # 세로 중앙정렬
        tf_body = shape.text_frame._txBody
        bodyPr = tf_body.find(qn("a:bodyPr"))
        if bodyPr is not None:
            bodyPr.set("anchor", "ctr")

    return shape


def brightness_check(r, g, b):
    """배경 색상의 밝기를 판단한다.

    Args:
        r, g, b: 정수 0-255

    Returns:
        True = 밝은 배경 (어두운 텍스트 사용)
        False = 어두운 배경 (흰색 텍스트 사용)
    """
    return (r * 0.299 + g * 0.587 + b * 0.114) > 160


# ---------------------------------------------------------------------------
# 텍스트 편의 함수 — 반복 보일러플레이트 제거 (시각적 의견 없음)
# ---------------------------------------------------------------------------


def add_textbox(slide, x, y, w, h, text, font_name=None, font_size=12,
                color=None, bold=False, align=PP_ALIGN.LEFT, word_wrap=True):
    """텍스트박스를 추가하고 단일 단락을 설정한다.

    Args:
        slide: 슬라이드 객체
        x, y, w, h: 위치/크기 (Inches/Emu)
        text: 텍스트 내용
        font_name: 폰트 이름 (None이면 기본)
        font_size: 포인트 단위
        color: RGBColor (None이면 기본)
        bold: 볼드 여부
        align: PP_ALIGN 정렬
        word_wrap: 자동 줄바꿈

    Returns:
        생성된 텍스트박스 도형 객체
    """
    txBox = slide.shapes.add_textbox(x, y, w, h)
    tf = txBox.text_frame
    tf.word_wrap = word_wrap

    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = str(text)
    run.font.size = Pt(font_size)
    run.font.bold = bold
    if font_name:
        run.font.name = font_name
    if color:
        run.font.color.rgb = color

    return txBox


def add_para(text_frame, text, font_name=None, font_size=12, color=None,
             bold=False, align=PP_ALIGN.LEFT, space_before=None,
             space_after=None):
    """기존 text_frame에 새 단락을 추가한다.

    Args:
        text_frame: python-pptx TextFrame 객체
        text: 텍스트 내용
        font_name: 폰트 이름 (None이면 기본)
        font_size: 포인트 단위
        color: RGBColor (None이면 기본)
        bold: 볼드 여부
        align: PP_ALIGN 정렬
        space_before: 단락 전 간격 (Pt 단위 값)
        space_after: 단락 후 간격 (Pt 단위 값)

    Returns:
        생성된 단락 객체
    """
    p = text_frame.add_paragraph()
    p.alignment = align
    if space_before is not None:
        p.space_before = space_before
    if space_after is not None:
        p.space_after = space_after

    run = p.add_run()
    run.text = str(text)
    run.font.size = Pt(font_size)
    run.font.bold = bold
    if font_name:
        run.font.name = font_name
    if color:
        run.font.color.rgb = color

    return p


def set_body_anchor(shape, anchor="ctr"):
    """도형 텍스트 프레임의 세로정렬을 설정한다.

    set_cell_anchor의 형제 함수. 테이블 셀이 아닌 일반 도형용.

    Args:
        shape: python-pptx 도형 객체 (text_frame을 가진)
        anchor: 't' (위), 'ctr' (가운데), 'b' (아래)
    """
    if not shape.has_text_frame:
        return
    bodyPr = shape.text_frame._txBody.find(qn("a:bodyPr"))
    if bodyPr is not None:
        bodyPr.set("anchor", anchor)


def setup_cover(slide, title, purpose="정보공유", author="강민규 선임",
                department="미래융합설계센터 알고리즘개발팀", date=None):
    """표지 슬라이드를 표준 포맷으로 설정한다.

    - 제목(PH idx=0): 텍스트만 설정, 폰트는 레이아웃 lstStyle에서 상속
      (48pt bold white 현대하모니 M)
    - 날짜(PH idx=1): 부제목 PH를 날짜 표시용으로 활용, 폰트 상속
      (24pt white 현대하모니 L)
    - 체크박스: 템플릿과 동일한 위치에 텍스트박스 추가
    - 저자: 템플릿과 동일한 위치에 텍스트박스 추가
    """
    from datetime import date as date_cls

    if date is None:
        date = date_cls.today().strftime("%Y.%m.%d")

    # 1. 제목 (PH idx=0) — 텍스트만 설정, 폰트 상속
    ph0 = slide.placeholders[0]
    ph0.text = title
    # 폰트를 명시적으로 설정하지 않음 → lstStyle의 48pt bold white 현대하모니 M 상속

    # 2. 날짜 (PH idx=1) — 부제목 PH를 날짜용으로 활용
    try:
        ph1 = slide.placeholders[1]
        ph1.text = date
        # 폰트를 명시적으로 설정하지 않음 → lstStyle의 24pt white 현대하모니 L 상속
    except (KeyError, IndexError):
        pass

    # 3. 우측 상단: 체크박스 — 템플릿 EMU 위치
    purposes = ["의사결정", "보고", "정보공유"]
    parts = [f"{'☑' if p == purpose else '☐'} {p}" for p in purposes]
    tb_purpose = add_textbox(slide,
        x=Emu(6531701), y=Emu(360000), w=Emu(3097103), h=Emu(272758),
        text="  ".join(parts), font_name="맑은 고딕", font_size=13,
        color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.RIGHT)
    set_text_inset(tb_purpose,
        left=Emu(108000), right=Emu(108000),
        top=Emu(36000), bottom=Emu(36000))

    # 4. 하단: 부서 + 이름 — 흰색 배경 위에 테마 기본색(어두운색)
    tb_author = add_textbox(slide,
        x=Emu(1609310), y=Emu(5985284), w=Emu(6681380), h=Emu(442035),
        text=f"{department} {author}", font_name="맑은 고딕", font_size=24,
        bold=True, align=PP_ALIGN.CENTER)
    set_text_inset(tb_author,
        left=Emu(108000), right=Emu(108000),
        top=Emu(36000), bottom=Emu(36000))

    return slide


# ---------------------------------------------------------------------------
# 레이아웃 계산 — 좌표/크기 수학만, 시각적 의견 없음
# ---------------------------------------------------------------------------

_TextSize = namedtuple("_TextSize", ["width", "height"])
_CellRect = namedtuple("_CellRect", ["left", "top", "width", "height"])
_ConnectorPoints = namedtuple("_ConnectorPoints", [
    "begin_x", "begin_y", "end_x", "end_y", "begin_cxn_idx", "end_cxn_idx"
])


class _Grid:
    """calc_grid 반환 객체. grid[r][c] 및 grid.flat 지원."""

    def __init__(self, cells):
        self._cells = cells  # list of list of _CellRect

    def __getitem__(self, row):
        return self._cells[row]

    @property
    def flat(self):
        """모든 셀을 row-major 순서로 iterate."""
        for row in self._cells:
            yield from row

    @property
    def rows(self):
        return len(self._cells)

    @property
    def cols(self):
        return len(self._cells[0]) if self._cells else 0


def estimate_text_size(text, font_size_pt, font_name=None, max_width=None):
    """텍스트가 차지할 예상 크기를 Emu 단위로 반환한다.

    unicodedata.east_asian_width로 한글/라틴 구분하여 폭 추정.
    max_width 지정 시 줄바꿈 계산하여 필요한 높이 반환.

    Args:
        text: 텍스트 문자열
        font_size_pt: 폰트 크기 (포인트)
        font_name: 미사용 (향후 폰트별 메트릭 확장용)
        max_width: 최대 너비 (Emu). None이면 한 줄 기준

    Returns:
        _TextSize(width, height) — Emu 단위
    """
    font_emu = int(Pt(font_size_pt))
    line_height = int(font_emu * 1.3)

    # 문자별 폭 계산
    char_widths = []
    for ch in text:
        if ch == '\n':
            char_widths.append(0)  # 줄바꿈은 폭 0
            continue
        eaw = unicodedata.east_asian_width(ch)
        if eaw in ('W', 'F'):  # Wide, Fullwidth (CJK)
            char_widths.append(int(font_emu * 1.0))
        else:  # Narrow, Half, Neutral, Ambiguous
            char_widths.append(int(font_emu * 0.55))

    if max_width is None:
        # 명시적 줄바꿈 처리
        lines = text.split('\n') if '\n' in text else [text]
        max_line_w = 0
        for line in lines:
            line_w = sum(
                int(font_emu * 1.0) if unicodedata.east_asian_width(ch) in ('W', 'F')
                else int(font_emu * 0.55)
                for ch in line
            )
            max_line_w = max(max_line_w, line_w)
        return _TextSize(width=max_line_w, height=line_height * len(lines))

    # 줄바꿈 계산
    lines = 1
    current_width = 0
    max_line_width = 0

    for i, ch in enumerate(text):
        if ch == '\n':
            max_line_width = max(max_line_width, current_width)
            current_width = 0
            lines += 1
            continue
        cw = char_widths[i]
        if current_width + cw > max_width and current_width > 0:
            max_line_width = max(max_line_width, current_width)
            current_width = cw
            lines += 1
        else:
            current_width += cw

    max_line_width = max(max_line_width, current_width)
    return _TextSize(width=min(max_line_width, max_width), height=line_height * lines)


def set_text_inset(shape, left=None, top=None, right=None, bottom=None):
    """도형 텍스트 프레임의 내부 여백(inset)을 설정한다.

    기본값은 한글 친화적으로 python-pptx 기본보다 ~20% 넓음.

    Args:
        shape: python-pptx 도형 객체 (text_frame을 가진)
        left: 좌측 여백 (Emu). None이면 Inches(0.12)
        top: 상단 여백 (Emu). None이면 Inches(0.06)
        right: 우측 여백 (Emu). None이면 Inches(0.12)
        bottom: 하단 여백 (Emu). None이면 Inches(0.06)
    """
    if not shape.has_text_frame:
        return

    tf = shape.text_frame
    tf.margin_left = left if left is not None else Inches(0.12)
    tf.margin_top = top if top is not None else Inches(0.06)
    tf.margin_right = right if right is not None else Inches(0.12)
    tf.margin_bottom = bottom if bottom is not None else Inches(0.06)


def calc_grid(rows, cols, area=None, gap=None, col_widths=None, row_heights=None):
    """영역을 rows×cols 그리드로 분할하여 각 셀의 좌표/크기를 반환한다.

    Args:
        rows: 행 수
        cols: 열 수
        area: _CellRect 또는 (left, top, width, height) 튜플. None이면 CONTENT_SAFE
        gap: 셀 간격 (Emu). None이면 자동 계산 (셀 너비의 ~10%, 0.1"~0.4")
        col_widths: 열 너비 비율 리스트 (예: [1, 2, 1]). None이면 균등
        row_heights: 행 높이 비율 리스트 (예: [1, 3]). None이면 균등

    Returns:
        _Grid 객체. grid[r][c] → _CellRect(left, top, width, height)
    """
    if area is None:
        a_left = CONTENT_SAFE.left
        a_top = CONTENT_SAFE.top
        a_width = CONTENT_SAFE.width
        a_height = CONTENT_SAFE.height
    elif isinstance(area, tuple) and len(area) == 4:
        a_left, a_top, a_width, a_height = area
    else:
        a_left = area.left
        a_top = area.top
        a_width = area.width
        a_height = area.height

    # gap 자동 계산
    if gap is None:
        approx_cell_w = int(a_width) // max(cols, 1)
        gap = max(Inches(0.1), min(Inches(0.4), int(approx_cell_w * 0.10)))
    gap = int(gap)

    # 가용 영역 (gap 제외)
    total_h_gap = gap * (cols - 1) if cols > 1 else 0
    total_v_gap = gap * (rows - 1) if rows > 1 else 0
    avail_w = int(a_width) - total_h_gap
    avail_h = int(a_height) - total_v_gap

    # 열 너비 계산
    if col_widths is None:
        col_widths = [1] * cols
    total_cw = sum(col_widths)
    col_sizes = []
    remainder_w = avail_w
    for i, cw in enumerate(col_widths):
        if i == len(col_widths) - 1:
            col_sizes.append(remainder_w)
        else:
            size = avail_w * cw // total_cw
            col_sizes.append(size)
            remainder_w -= size

    # 행 높이 계산
    if row_heights is None:
        row_heights = [1] * rows
    total_rh = sum(row_heights)
    row_sizes = []
    remainder_h = avail_h
    for i, rh in enumerate(row_heights):
        if i == len(row_heights) - 1:
            row_sizes.append(remainder_h)
        else:
            size = avail_h * rh // total_rh
            row_sizes.append(size)
            remainder_h -= size

    # 셀 좌표 계산
    cells = []
    y = int(a_top)
    for r in range(rows):
        row = []
        x = int(a_left)
        for c in range(cols):
            row.append(_CellRect(left=x, top=y, width=col_sizes[c], height=row_sizes[r]))
            x += col_sizes[c] + gap
        cells.append(row)
        y += row_sizes[r] + gap

    return _Grid(cells)


def calc_connector(shape_a, shape_b, direction=None):
    """두 도형 가장자리 중점 좌표와 connection point 인덱스를 계산한다.

    Connection point 인덱스:
        0 = top-center, 1 = left-center, 2 = bottom-center, 3 = right-center

    Args:
        shape_a: 시작 도형
        shape_b: 끝 도형
        direction: 'LR', 'RL', 'TB', 'BT' 또는 None (자동 감지)

    Returns:
        _ConnectorPoints(begin_x, begin_y, end_x, end_y, begin_cxn_idx, end_cxn_idx)
    """
    # 도형 중심 좌표
    ax = shape_a.left + shape_a.width // 2
    ay = shape_a.top + shape_a.height // 2
    bx = shape_b.left + shape_b.width // 2
    by = shape_b.top + shape_b.height // 2

    # 방향 자동 감지
    if direction is None:
        dx = abs(bx - ax)
        dy = abs(by - ay)
        if dx >= dy:
            direction = 'LR' if bx >= ax else 'RL'
        else:
            direction = 'TB' if by >= ay else 'BT'

    direction = direction.upper()

    # 가장자리 중점 + cxn 인덱스 계산
    if direction == 'LR':
        # A의 오른쪽 → B의 왼쪽
        bx_pt = shape_a.left + shape_a.width
        by_pt = ay
        ex_pt = shape_b.left
        ey_pt = by
        begin_cxn, end_cxn = 3, 1
    elif direction == 'RL':
        # A의 왼쪽 → B의 오른쪽
        bx_pt = shape_a.left
        by_pt = ay
        ex_pt = shape_b.left + shape_b.width
        ey_pt = by
        begin_cxn, end_cxn = 1, 3
    elif direction == 'TB':
        # A의 아래 → B의 위
        bx_pt = ax
        by_pt = shape_a.top + shape_a.height
        ex_pt = bx
        ey_pt = shape_b.top
        begin_cxn, end_cxn = 2, 0
    else:  # BT
        # A의 위 → B의 아래
        bx_pt = ax
        by_pt = shape_a.top
        ex_pt = bx
        ey_pt = shape_b.top + shape_b.height
        begin_cxn, end_cxn = 0, 2

    return _ConnectorPoints(
        begin_x=bx_pt, begin_y=by_pt,
        end_x=ex_pt, end_y=ey_pt,
        begin_cxn_idx=begin_cxn, end_cxn_idx=end_cxn,
    )


def add_smart_connector(slide, shape_a, shape_b, connector_type=None,
                        direction=None, arrow=True):
    """두 도형을 커넥터로 연결한다.

    calc_connector로 좌표 계산 후 커넥터 생성 + 옵션 화살표.

    Args:
        slide: 슬라이드 객체
        shape_a: 시작 도형
        shape_b: 끝 도형
        connector_type: MSO_CONNECTOR_TYPE 값 또는 None (자동 선택)
        direction: 'LR', 'RL', 'TB', 'BT' 또는 None (자동 감지)
        arrow: True이면 끝점에 화살표 추가

    Returns:
        생성된 커넥터 객체
    """
    pts = calc_connector(shape_a, shape_b, direction=direction)

    # 커넥터 타입 자동 선택
    if connector_type is None:
        # 정렬 여부 판단: 시작/끝 좌표가 한 축에서 충분히 가까우면 STRAIGHT
        dx = abs(pts.begin_x - pts.end_x)
        dy = abs(pts.begin_y - pts.end_y)
        threshold = Inches(0.3)
        if dx < threshold or dy < threshold:
            connector_type = MSO_CONNECTOR_TYPE.STRAIGHT
        else:
            connector_type = MSO_CONNECTOR_TYPE.ELBOW

    connector = slide.shapes.add_connector(
        connector_type, pts.begin_x, pts.begin_y, pts.end_x, pts.end_y
    )

    # begin_connect / end_connect 시도 (비직사각형 도형에서 실패할 수 있음)
    try:
        connector.begin_connect(shape_a, pts.begin_cxn_idx)
    except (IndexError, ValueError, AttributeError):
        pass
    try:
        connector.end_connect(shape_b, pts.end_cxn_idx)
    except (IndexError, ValueError, AttributeError):
        pass

    if arrow:
        # 커넥터의 ln 요소가 없을 수 있으므로 보장
        cxnSp = connector._element
        spPr = cxnSp.find(qn("p:spPr"))
        if spPr is None:
            spPr = cxnSp.find(qn("p:cxnSp/p:spPr"))
        if spPr is not None:
            ln = spPr.find(qn("a:ln"))
            if ln is None:
                ln = spPr.makeelement(qn("a:ln"), {})
                spPr.append(ln)
            # tailEnd 추가
            for old in ln.findall(qn("a:tailEnd")):
                ln.remove(old)
            ln.append(ln.makeelement(
                qn("a:tailEnd"),
                {"type": "triangle", "w": "med", "len": "med"},
            ))

    return connector
