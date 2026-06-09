"""
ppt_utils.py — 비시각적 유틸리티
디자인을 제약하는 코드 없음. 인프라 헬퍼만 포함.
"""

import os
import re
import math
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
ASSETS_DIR = BASE_DIR / "assets"        # 이미지/스크린샷/차트 PNG 등 시각 자산 (1급 채널)
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


def load_template():
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


def set_cell_fill(cell, color):
    """테이블 셀 배경색을 안전하게 설정한다.

    기존 fill 요소를 제거하고 올바른 위치에 solidFill을 삽입한다.
    OOXML 스키마: solidFill은 lnBlToTr 뒤, cell3D 앞에 위치해야 한다.

    Args:
        cell: 테이블 셀 객체
        color: RGBColor 또는 (r, g, b) 튜플
    """
    tcPr = cell._tc.get_or_add_tcPr()
    # 기존 fill 제거 (solidFill, gradFill, blipFill, pattFill, grpFill, noFill)
    for fill_tag in ("solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "noFill"):
        for old in tcPr.findall(qn(f"a:{fill_tag}")):
            tcPr.remove(old)
    # 새 solidFill 생성
    solidFill = tcPr.makeelement(qn("a:solidFill"), {})
    srgbClr = solidFill.makeelement(qn("a:srgbClr"), {
        "val": f"{color[0]:02X}{color[1]:02X}{color[2]:02X}"
    })
    solidFill.append(srgbClr)
    # 올바른 위치에 삽입: 모든 ln* 요소 뒤
    ln_tags = {"lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr"}
    insert_idx = 0
    for i, child in enumerate(tcPr):
        if child.tag.split('}')[-1] in ln_tags:
            insert_idx = i + 1
    tcPr.insert(insert_idx, solidFill)


def add_arrowhead(connector):
    """커넥터에 화살표 머리를 추가한다 (python-pptx에 네이티브 API 없음)."""
    cxnSp = connector._element
    spPr = cxnSp.find(qn("p:spPr"))
    if spPr is None:
        return
    ln = spPr.find(qn("a:ln"))
    if ln is None:
        ln = spPr.makeelement(qn("a:ln"), {})
        spPr.append(ln)
    for old in ln.findall(qn("a:tailEnd")):
        ln.remove(old)
    ln.append(ln.makeelement(
        qn("a:tailEnd"),
        {"type": "triangle", "w": "med", "len": "med"},
    ))


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
        direction: 그림자 방향 (OOXML 1/60000도 단위. 2700000=아래, 5400000=오른쪽아래)
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


def add_accent_bar(slide, x, y, w, h, color):
    """얇은 색상 바를 추가한다 (RECTANGLE, 테두리 없음).

    카드 좌측/상단 accent, Key Message Bar 좌측 바, 칼럼 구분선 등에 사용.
    용도에 따라 w/h 비율을 조절:
    - 세로 바: w=Inches(0.05), h=카드높이
    - 가로 바: w=카드폭, h=Inches(0.04)
    - 칼럼 구분선: w=Inches(0.01), h=영역높이

    Args:
        slide: 슬라이드 객체
        x, y: 위치 (Inches/Emu)
        w, h: 크기 (Inches/Emu)
        color: RGBColor 채우기 색상

    Returns:
        생성된 도형 객체
    """
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


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


def make_icon_badge(slide, x, y, w, h, text, fill_color,
                    font_size=14, font_color=None, corner_radius=0.15):
    """사각형 아이콘/배지를 생성한다 (ROUNDED_RECTANGLE + fill + 중앙정렬 텍스트).

    make_icon_circle()의 사각형 변형. 이모지, 텍스트 라벨, 번호 등에 사용.
    색상/크기는 파라미터이므로 시각적 의견 없음.

    Args:
        slide: 슬라이드 객체
        x, y: 위치 (Inches/Emu)
        w, h: 크기 (Inches/Emu). 정사각형이면 w=h
        text: 배지 안에 들어갈 텍스트 (이모지, 번호, 약어 등)
        fill_color: RGBColor 배경색
        font_size: 포인트 단위
        font_color: RGBColor. None이면 brightness_check로 자동 결정
        corner_radius: 모서리 반경 비율 (0.0~0.5). 0이면 직각

    Returns:
        생성된 도형 객체
    """
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()  # 테두리 없음

    if corner_radius is not None:
        shape.adjustments[0] = corner_radius

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

        set_body_anchor(shape, 'ctr')
        set_text_inset(shape)

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


def add_rich_text(text_frame, segments, align=PP_ALIGN.LEFT,
                  space_before=None, space_after=None, line_spacing=None):
    """단일 단락 내 혼합 서식 텍스트를 추가한다.

    핵심 수치/용어만 accent 색상으로 강조하거나 bold 대비를 줄 때 사용.

    Args:
        text_frame: python-pptx TextFrame 객체
        segments: 세그먼트 리스트. 각 항목은:
            - str → 기본 서식 run
            - dict → 커스텀 서식 run. 키:
                text (필수), font_size, font_name, color (RGBColor),
                bold, italic
        align: PP_ALIGN 정렬
        space_before: 단락 전 간격 (Pt 단위 값)
        space_after: 단락 후 간격 (Pt 단위 값)
        line_spacing: 줄 간격 (Pt 단위 값)

    Returns:
        생성된 단락 객체
    """
    p = text_frame.add_paragraph()
    p.alignment = align
    if space_before is not None:
        p.space_before = space_before
    if space_after is not None:
        p.space_after = space_after
    if line_spacing is not None:
        p.line_spacing = line_spacing

    for seg in segments:
        run = p.add_run()
        if isinstance(seg, str):
            run.text = seg
        elif isinstance(seg, dict):
            run.text = str(seg.get("text", ""))
            if "font_size" in seg:
                run.font.size = Pt(seg["font_size"])
            if "font_name" in seg:
                run.font.name = seg["font_name"]
            if "color" in seg:
                run.font.color.rgb = seg["color"]
            if "bold" in seg:
                run.font.bold = seg["bold"]
            if "italic" in seg:
                run.font.italic = seg["italic"]

    return p


def add_bullet_list(slide, x, y, w, items, font_size=12,
                    font_name=None, color=None, bullet_char="\u2022",
                    indent_emu=None, item_spacing=None, h=None):
    """단일 text_frame에 multi-paragraph 리스트를 생성한다.

    개별 textbox 대신 단일 text_frame을 사용하여 도형 수를 줄이고
    간격 일관성을 보장한다.

    Args:
        slide: 슬라이드 객체
        x, y, w: 위치/너비 (Inches/Emu)
        items: 항목 리스트. 각 항목은:
            - str → 기본 서식 리스트 항목
            - dict → 커스텀 서식. 키:
                text (필수), level (0-2, 기본 0), bold, color
        font_size: 기본 폰트 크기 (포인트)
        font_name: 폰트 이름 (None이면 기본)
        color: 기본 텍스트 RGBColor (None이면 기본)
        bullet_char: 글머리 기호 문자 (기본 "•")
        indent_emu: 레벨당 들여쓰기 (Emu). None이면 Inches(0.25)
        item_spacing: 항목 간격 (Pt 값). None이면 Pt(font_size * 0.3)
        h: 텍스트박스 높이. None이면 자동 계산

    Returns:
        생성된 텍스트박스 도형 객체
    """
    if indent_emu is None:
        indent_emu = Inches(0.25)
    if item_spacing is None:
        item_spacing = Pt(max(3, int(font_size * 0.3)))

    # 높이 자동 계산
    if h is None:
        total_text = "\n".join(
            (it if isinstance(it, str) else it.get("text", "")) for it in items
        )
        h = estimate_container_height(total_text, font_size, max_width=int(w))
        # 항목 간격 보정
        h += int(item_spacing) * max(0, len(items) - 1)

    txBox = slide.shapes.add_textbox(int(x), int(y), int(w), int(h))
    tf = txBox.text_frame
    tf.word_wrap = True

    for i, item in enumerate(items):
        if isinstance(item, str):
            text = item
            level = 0
            item_bold = False
            item_color = color
        else:
            text = str(item.get("text", ""))
            level = item.get("level", 0)
            item_bold = item.get("bold", False)
            item_color = item.get("color", color)

        # 글머리 기호 + 들여쓰기 적용
        indent = int(indent_emu) * level
        display_text = f"{bullet_char} {text}" if bullet_char else text

        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()

        p.alignment = PP_ALIGN.LEFT
        p.space_before = item_spacing if i > 0 else Pt(0)
        p.space_after = Pt(0)
        p.line_spacing = Pt(font_size * 1.2)

        # 들여쓰기 설정
        if indent > 0:
            pPr = p._pPr
            if pPr is None:
                pPr = p._p.get_or_add_pPr()
            pPr.set("marL", str(indent))

        run = p.add_run()
        run.text = display_text
        run.font.size = Pt(font_size)
        run.font.bold = item_bold
        if font_name:
            run.font.name = font_name
        if item_color:
            run.font.color.rgb = item_color

    return txBox


def add_footnote(slide, text, font_size=9, color=None, align=PP_ALIGN.LEFT,
                 prefix="※ "):
    """슬라이드 하단에 각주 텍스트를 추가한다.

    CONTENT_SAFE 하단에 자동 배치. 보조 정보, 출처, 참고사항에 사용.

    Args:
        slide: 슬라이드 객체
        text: 각주 텍스트 (여러 줄이면 \\n으로 구분)
        font_size: 포인트 단위 (기본 9)
        color: RGBColor (기본 #999999)
        align: PP_ALIGN 정렬
        prefix: 텍스트 앞 접두어 (기본 "※ "). 빈 문자열이면 접두어 없음

    Returns:
        생성된 텍스트박스 도형 객체
    """
    if color is None:
        color = RGBColor(0x99, 0x99, 0x99)

    display_text = f"{prefix}{text}" if prefix else text
    footnote_h = estimate_container_height(
        display_text, font_size, max_width=int(CONTENT_SAFE.width)
    )
    footnote_y = int(CONTENT_SAFE.bottom) - int(footnote_h)

    return add_textbox(
        slide, CONTENT_SAFE.left, footnote_y, CONTENT_SAFE.width, footnote_h,
        display_text, font_size=font_size, color=color, align=align,
    )


def auto_shrink_text(shape):
    """도형 텍스트가 넘칠 때 자동으로 폰트 크기를 축소하도록 설정한다."""
    if not shape.has_text_frame:
        return
    bodyPr = shape.text_frame._txBody.find(qn("a:bodyPr"))
    if bodyPr is not None:
        # wrap 속성을 제거하지 않음 (줄바꿈 유지)
        # 기존 noAutofit / spAutoFit 제거 (normAutofit과 충돌 방지)
        for old in bodyPr.findall(qn("a:noAutofit")):
            bodyPr.remove(old)
        for old in bodyPr.findall(qn("a:spAutoFit")):
            bodyPr.remove(old)
        autofit = bodyPr.find(qn("a:normAutofit"))
        if autofit is None:
            autofit = bodyPr.makeelement(qn("a:normAutofit"), {})
            bodyPr.append(autofit)
        # fontScale을 명시하지 않음 — PowerPoint가 자동 계산
        # (이전 버그: min_font_pt*1000=8000 → 8% 축소로 폰트가 1pt로 표시됨)
        autofit.attrib.pop("fontScale", None)


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


def setup_cover(slide, title, purpose="정보공유", author="강민규 책임",
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
    parts = [f"{'■' if p == purpose else '☐'} {p}" for p in purposes]
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


def add_section_divider(prs, section_title, subtitle="",
                        accent_color=None, keywords=None):
    """표준 섹션 구분 슬라이드를 생성한다.

    제목 중복 표시를 방지하고, 일관된 구조로 섹션을 구분한다.

    Args:
        prs: Presentation 객체
        section_title: 섹션 제목 (set_title로 1회만 설정)
        subtitle: 부제목 텍스트 (1줄, 12-14pt)
        accent_color: 강조 색상 RGBColor (기본 #2D5B8A). 수평 바와 키워드 pills에 사용
        keywords: 키워드 리스트 (선택). pill 형태로 표시

    Returns:
        생성된 슬라이드 객체
    """
    if accent_color is None:
        accent_color = RGBColor(0x2D, 0x5B, 0x8A)

    layout = get_layout(prs, "제목 및 내용 (페이지 번호 삭제)")
    slide = prs.slides.add_slide(layout)
    set_title(slide, section_title)
    clear_placeholders(slide, keep=[0])

    # 중앙 영역 계산
    center_y = CONTENT_SAFE.top + CONTENT_SAFE.height // 3

    # accent 수평 바
    bar_w = Inches(2.0)
    bar_h = Inches(0.06)
    bar_x = CONTENT_SAFE.left + (CONTENT_SAFE.width - int(bar_w)) // 2
    bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        bar_x, center_y, bar_w, bar_h,
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = accent_color
    bar.line.fill.background()

    # 부제목
    if subtitle:
        sub_y = center_y + int(bar_h) + Inches(0.3)
        sub_w = Inches(8.0)
        sub_x = CONTENT_SAFE.left + (CONTENT_SAFE.width - int(sub_w)) // 2
        add_textbox(slide, sub_x, sub_y, sub_w, Inches(0.5),
                    subtitle, font_size=13, align=PP_ALIGN.CENTER,
                    color=RGBColor(0x59, 0x59, 0x59))

    # 키워드 pills
    if keywords:
        pill_y = center_y + int(bar_h) + Inches(0.9)
        if subtitle:
            pill_y += Inches(0.3)

        # 간단한 수평 배치
        total_pill_w = sum(
            int(estimate_text_size(kw, 10).width) + Inches(0.4)
            for kw in keywords
        )
        pill_gap = Inches(0.12)
        total_w = total_pill_w + int(pill_gap) * (len(keywords) - 1)
        pill_x = CONTENT_SAFE.left + (CONTENT_SAFE.width - total_w) // 2

        for kw in keywords:
            kw_w = int(estimate_text_size(kw, 10).width) + Inches(0.4)
            pill = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                pill_x, pill_y, kw_w, Inches(0.35),
            )
            pill.fill.solid()
            pill.fill.fore_color.rgb = accent_color
            set_shape_opacity(pill, 12)
            pill.line.color.rgb = accent_color
            pill.line.width = Pt(0.5)

            tf = pill.text_frame
            tf.word_wrap = False
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = kw
            run.font.size = Pt(10)
            run.font.color.rgb = accent_color
            run.font.bold = True
            set_body_anchor(pill, 'ctr')
            set_text_inset(pill, left=Inches(0.08), right=Inches(0.08),
                           top=Inches(0.02), bottom=Inches(0.02))

            pill_x += kw_w + int(pill_gap)

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
            char_widths.append(int(font_emu * 1.1))
        else:  # Narrow, Half, Neutral, Ambiguous
            char_widths.append(int(font_emu * 0.6))

    if max_width is None:
        # 명시적 줄바꿈 처리
        lines = text.split('\n') if '\n' in text else [text]
        max_line_w = 0
        for line in lines:
            line_w = sum(
                int(font_emu * 1.1) if unicodedata.east_asian_width(ch) in ('W', 'F')
                else int(font_emu * 0.6)
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


def estimate_container_height(text, font_size_pt, max_width,
                               padding_top=None, padding_bottom=None):
    """텍스트를 담는 컨테이너의 필요 높이를 계산한다.

    estimate_text_size()로 텍스트 높이를 구한 뒤 상하 여백을 더한다.

    Args:
        text: 텍스트 문자열
        font_size_pt: 폰트 크기 (포인트)
        max_width: 컨테이너 너비 (Emu/Inches)
        padding_top: 상단 여백 (Emu). None이면 Inches(0.06)
        padding_bottom: 하단 여백 (Emu). None이면 Inches(0.06)

    Returns:
        int — 컨테이너 필요 높이 (Emu)
    """
    if padding_top is None:
        padding_top = Inches(0.06)
    if padding_bottom is None:
        padding_bottom = Inches(0.06)
    text_size = estimate_text_size(text, font_size_pt, max_width=int(max_width))
    return int(text_size.height) + int(padding_top) + int(padding_bottom)


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


class VFlow:
    """수직 요소 배치 트래커. y좌표를 자동 관리하여 겹침을 방지한다.

    레이아웃 엔진이 아님 — y좌표만 추적. 렌더링은 기존 함수(add_textbox 등)가 담당.
    """
    __slots__ = ('_x', '_w', '_y', '_y_max', '_gap', '_cursor')

    def __init__(self, x=None, w=None, y_start=None, y_max=None, gap=None):
        self._x = int(x if x is not None else CONTENT_SAFE.left)
        self._w = int(w if w is not None else CONTENT_SAFE.width)
        self._y = int(y_start if y_start is not None else CONTENT_SAFE.top)
        self._y_max = int(y_max if y_max is not None else CONTENT_SAFE.bottom)
        self._gap = int(gap if gap is not None else Inches(0.15))
        self._cursor = self._y

    @property
    def cursor(self): return self._cursor          # 현재 y (Emu)
    @property
    def remaining(self): return max(0, self._y_max - self._cursor)  # 남은 공간
    @property
    def x(self): return self._x
    @property
    def w(self): return self._w

    def advance(self, height, gap=None):
        """공간을 소비하고 배치 y좌표를 반환한다."""
        y = self._cursor
        g = int(gap) if gap is not None else self._gap
        self._cursor = y + int(height) + g
        return y

    def textbox(self, slide, text, font_size=12, h=None, gap=None, **kwargs):
        """텍스트박스를 cursor 위치에 배치. h=None이면 자동 계산."""
        if h is None:
            h = estimate_container_height(text, font_size, max_width=self._w)
        y = self.advance(h, gap=gap)
        return add_textbox(slide, self._x, y, self._w, h, text,
                           font_size=font_size, **kwargs)

    def reserve(self, height, gap=None):
        """높이를 예약하고 (left, top, width, height) 좌표를 반환한다. 도형/테이블용."""
        y = self.advance(height, gap=gap)
        return _CellRect(left=self._x, top=y, width=self._w, height=int(height))

    def skip(self, amount=None):
        """요소 없이 커서만 전진."""
        self._cursor += int(amount) if amount is not None else self._gap


class HFlow:
    """수평 요소 배치 트래커. x좌표를 자동 관리하여 겹침을 방지한다.

    VFlow의 수평 버전. 같은 인터페이스 패턴.
    """
    __slots__ = ('_y', '_h', '_x', '_x_max', '_gap', '_cursor')

    def __init__(self, y=None, h=None, x_start=None, x_max=None, gap=None):
        self._y = int(y if y is not None else CONTENT_SAFE.top)
        self._h = int(h if h is not None else CONTENT_SAFE.height)
        self._x = int(x_start if x_start is not None else CONTENT_SAFE.left)
        self._x_max = int(x_max if x_max is not None else CONTENT_SAFE.right)
        self._gap = int(gap if gap is not None else Inches(0.15))
        self._cursor = self._x

    @property
    def cursor(self): return self._cursor          # 현재 x (Emu)
    @property
    def remaining(self): return max(0, self._x_max - self._cursor)  # 남은 공간
    @property
    def y(self): return self._y
    @property
    def h(self): return self._h

    def advance(self, width, gap=None):
        """공간을 소비하고 배치 x좌표를 반환한다."""
        x = self._cursor
        g = int(gap) if gap is not None else self._gap
        self._cursor = x + int(width) + g
        return x

    def textbox(self, slide, text, font_size=12, w=None, gap=None, **kwargs):
        """텍스트박스를 cursor 위치에 배치. w=None이면 자동 계산."""
        if w is None:
            size = estimate_text_size(text, font_size)
            w = int(size.width) + Inches(0.24)  # 좌우 여백
        x = self.advance(w, gap=gap)
        return add_textbox(slide, x, self._y, int(w), self._h, text,
                           font_size=font_size, **kwargs)

    def reserve(self, width, gap=None):
        """너비를 예약하고 (left, top, width, height) 좌표를 반환한다."""
        x = self.advance(width, gap=gap)
        return _CellRect(left=x, top=self._y, width=int(width), height=self._h)

    def skip(self, amount=None):
        """요소 없이 커서만 전진."""
        self._cursor += int(amount) if amount is not None else self._gap


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

    # 커넥터 타입 자동 선택 — 방향별 교차축 정렬 기준
    if connector_type is None:
        dx = abs(pts.begin_x - pts.end_x)
        dy = abs(pts.begin_y - pts.end_y)
        threshold = Inches(0.3)

        # 유효 방향 결정 (direction=None이면 calc_connector와 동일 로직)
        if direction is not None:
            eff_dir = direction.upper()
        else:
            _ax = shape_a.left + shape_a.width // 2
            _ay = shape_a.top + shape_a.height // 2
            _bx = shape_b.left + shape_b.width // 2
            _by = shape_b.top + shape_b.height // 2
            eff_dir = 'LR' if abs(_bx - _ax) >= abs(_by - _ay) else 'TB'

        # 교차축 정렬 여부로 STRAIGHT/ELBOW 결정
        # TB/BT: 주축=Y → 교차축=X(dx)가 작아야 STRAIGHT
        # LR/RL: 주축=X → 교차축=Y(dy)가 작아야 STRAIGHT
        cross = dx if eff_dir in ('TB', 'BT') else dy
        if cross < threshold:
            connector_type = MSO_CONNECTOR_TYPE.STRAIGHT
        else:
            connector_type = MSO_CONNECTOR_TYPE.ELBOW

    connector = slide.shapes.add_connector(
        connector_type, pts.begin_x, pts.begin_y, pts.end_x, pts.end_y
    )

    # begin_connect / end_connect 시도
    # 비직사각형 도형(CHEVRON, CLOUD 등)은 연결점 인덱스가 직사각형과 다르므로
    # calc_connector가 반환한 cxn_idx가 유효하지 않을 수 있다.
    # 이 경우 커넥터는 좌표 기반으로만 배치되며 도형에 스냅되지 않는다 (의도된 폴백).
    try:
        connector.begin_connect(shape_a, pts.begin_cxn_idx)
    except (IndexError, ValueError, AttributeError):
        pass
    try:
        connector.end_connect(shape_b, pts.end_cxn_idx)
    except (IndexError, ValueError, AttributeError):
        pass

    # 기본 선 스타일 보장 (테마에 의존하지 않음)
    cxnSp = connector._element
    spPr = cxnSp.find(qn("p:spPr"))
    if spPr is None:
        spPr = cxnSp.find(qn("p:cxnSp/p:spPr"))
    if spPr is not None:
        ln = spPr.find(qn("a:ln"))
        if ln is None:
            ln = spPr.makeelement(qn("a:ln"), {"w": str(int(Pt(1)))})
            spPr.append(ln)
        # 선 색상이 없으면 기본 회색 설정
        if ln.find(qn("a:solidFill")) is None and ln.find(qn("a:noFill")) is None:
            solidFill = ln.makeelement(qn("a:solidFill"), {})
            srgbClr = solidFill.makeelement(qn("a:srgbClr"), {"val": "666666"})
            solidFill.append(srgbClr)
            ln.insert(0, solidFill)
        # 선 두께가 없으면 기본 1pt
        if "w" not in ln.attrib:
            ln.set("w", str(int(Pt(1))))

        if arrow:
            for old in ln.findall(qn("a:tailEnd")):
                ln.remove(old)
            ln.append(ln.makeelement(
                qn("a:tailEnd"),
                {"type": "triangle", "w": "med", "len": "med"},
            ))

    return connector


def align_shapes(*shapes, axis='h'):
    """같은 행/열의 도형을 정렬 (크기 통일 + 중심 맞춤).

    axis='h': 높이→max, center_y→평균  (수평 커넥터 평행 보장)
    axis='v': 너비→max, center_x→평균  (수직 커넥터 평행 보장)

    Args:
        *shapes: 정렬할 도형 객체들 (2개 이상)
        axis: 'h' (수평 행 정렬) 또는 'v' (수직 열 정렬)

    Returns:
        list — 정렬된 도형 리스트
    """
    if len(shapes) < 2:
        return list(shapes)
    if axis == 'h':
        max_h = max(s.height for s in shapes)
        avg_cy = sum(s.top + s.height // 2 for s in shapes) // len(shapes)
        for s in shapes:
            s.height = max_h
            s.top = avg_cy - max_h // 2
    elif axis == 'v':
        max_w = max(s.width for s in shapes)
        avg_cx = sum(s.left + s.width // 2 for s in shapes) // len(shapes)
        for s in shapes:
            s.width = max_w
            s.left = avg_cx - max_w // 2
    return list(shapes)


# ---------------------------------------------------------------------------
# Mermaid 파싱 — 순수 텍스트 처리, 외부 의존성 없음
# ---------------------------------------------------------------------------


def parse_mermaid_metadata(mermaid_text):
    """Mermaid 텍스트에서 구조 메타데이터를 추출한다 (렌더링 없음).

    간단한 regex 기반 파싱으로 content_inventory 자동 생성에 사용.

    Args:
        mermaid_text: Mermaid 다이어그램 텍스트 (```mermaid 펜스 없이)

    Returns:
        dict: {
            type: str,              # "graph_TD", "graph_LR", "sequenceDiagram", etc.
            participants: list,     # participant/actor 이름 리스트
            subgraphs: list,        # subgraph 라벨 리스트
            nodes: int,             # 고유 노드 수 (graph 타입)
            edges: int,             # 화살표 연결 수
            has_loop: bool,         # loop 블록 존재 여부
            has_alt: bool,          # alt/else 블록 존재 여부
            suggested_strategy: str # PPT 변환 전략 제안
        }
    """
    lines = mermaid_text.strip().splitlines()
    first_line = lines[0].strip().lower() if lines else ""

    # 다이어그램 타입 감지
    diagram_type = "unknown"
    if first_line.startswith("graph td") or first_line.startswith("graph tb"):
        diagram_type = "graph_TD"
    elif first_line.startswith("graph lr"):
        diagram_type = "graph_LR"
    elif first_line.startswith("graph rl"):
        diagram_type = "graph_RL"
    elif first_line.startswith("sequencediagram"):
        diagram_type = "sequenceDiagram"
    elif first_line.startswith("statediagram"):
        diagram_type = "stateDiagram"
    elif first_line.startswith("flowchart"):
        direction = first_line.replace("flowchart", "").strip().upper()
        diagram_type = f"flowchart_{direction}" if direction else "flowchart_TD"

    # participant/actor 추출 (sequence diagram)
    participants = []
    for line in lines:
        m = re.match(r'\s*(?:participant|actor)\s+(\S+)', line, re.IGNORECASE)
        if m:
            participants.append(m.group(1))

    # subgraph 추출
    subgraphs = []
    for line in lines:
        m = re.match(r'\s*subgraph\s+(.*)', line, re.IGNORECASE)
        if m:
            label = m.group(1).strip().strip('"').strip("'")
            if label:
                subgraphs.append(label)

    # 노드 추출 (graph/flowchart 타입)
    nodes = set()
    # 화살표 패턴: 체인 엣지(A --> B --> C)를 지원하기 위해 화살표 기준 split
    arrow_re = re.compile(
        r'(?:---->|--->|-->|==>|-\.->|-.->|~~>|--\s)'
        r'(?:\|[^|]*\|)?'
    )
    node_re = re.compile(r'(\w+)')
    edge_count = 0
    for line in lines:
        stripped = line.strip()
        if stripped.startswith(('graph', 'flowchart', 'subgraph', 'end', 'style',
                                'classDef', 'class ', '%%', 'direction')):
            continue
        # 화살표로 분할하여 모든 노드 추출
        parts = arrow_re.split(stripped)
        seg_nodes = []
        for p in parts:
            p = p.strip()
            # 노드 ID 추출 (데코레이터 [label], {label} 등 앞의 식별자)
            m = node_re.match(p)
            if m:
                seg_nodes.append(m.group(1))
        if len(seg_nodes) >= 2:
            for n in seg_nodes:
                nodes.add(n)
            edge_count += len(seg_nodes) - 1
        # 노드 선언만 있는 라인 (A["라벨"])
        node_decl = re.match(r'\s*(\w+)[\[\(\{]', stripped)
        if node_decl:
            nodes.add(node_decl.group(1))

    # sequence diagram 메시지 수 (-->>, ->>, ->, -->> 등)
    seq_messages = 0
    if diagram_type == "sequenceDiagram":
        for line in lines:
            if re.search(r'->>|->|-->>|-->', line) and not line.strip().startswith(('participant', 'actor', 'Note', 'loop', 'alt', 'else', 'end', '%%')):
                seq_messages += 1
        edge_count = seq_messages

    # loop/alt 감지
    has_loop = any(re.match(r'\s*loop\b', l, re.IGNORECASE) for l in lines)
    has_alt = any(re.match(r'\s*alt\b', l, re.IGNORECASE) for l in lines)

    # PPT 변환 전략 제안
    if diagram_type == "graph_TD":
        strategy = "layer_diagram"
    elif diagram_type in ("graph_LR", "graph_RL"):
        if any(kw in mermaid_text.lower() for kw in ['상태', 'state', 'status']):
            strategy = "state_machine"
        else:
            strategy = "horizontal_flow"
    elif diagram_type == "sequenceDiagram":
        if has_loop or has_alt:
            strategy = "flowchart"
        elif seq_messages <= 5:
            strategy = "horizontal_flow"
        else:
            strategy = "summary_diagram_with_table"
    elif diagram_type == "stateDiagram":
        strategy = "state_machine"
    else:
        strategy = "horizontal_flow"

    return {
        "type": diagram_type,
        "participants": participants,
        "subgraphs": subgraphs,
        "nodes": len(nodes),
        "edges": edge_count,
        "has_loop": has_loop,
        "has_alt": has_alt,
        "suggested_strategy": strategy,
    }


# ---------------------------------------------------------------------------
# 코드 블록 — 어두운 배경 + 고정폭 폰트 + 라인 하이라이트
# ---------------------------------------------------------------------------


def add_code_block(slide, x, y, w, h, code_text, font_name="Consolas",
                   font_size=14, bg_color=None, text_color=None,
                   highlight_lines=None, highlight_color=None):
    """코드 스니펫 블록을 슬라이드에 추가한다.

    CLAUDE.md "코드 스니펫 규칙"을 함수화:
    - 어두운 배경(#1E1E1E) + 고정폭 폰트
    - 선택적 라인 하이라이트

    Args:
        slide: 슬라이드 객체
        x, y, w, h: 위치/크기 (Inches/Emu)
        code_text: 코드 텍스트
        font_name: 고정폭 폰트 (기본 Consolas)
        font_size: 포인트 단위 (기본 11)
        bg_color: 배경 RGBColor (기본 #1E1E1E)
        text_color: 텍스트 RGBColor (기본 #D4D4D4)
        highlight_lines: 하이라이트할 라인 번호 리스트 (1-indexed). None이면 없음
        highlight_color: 하이라이트 배경색 RGBColor (기본 #264F78)

    Returns:
        생성된 배경 도형 + 텍스트박스 튜플 (bg_shape, text_shape)
    """
    if bg_color is None:
        bg_color = RGBColor(0x1E, 0x1E, 0x1E)
    if text_color is None:
        text_color = RGBColor(0xD4, 0xD4, 0xD4)
    if highlight_color is None:
        highlight_color = RGBColor(0x26, 0x4F, 0x78)
    if highlight_lines is None:
        highlight_lines = []

    # 배경 도형 (둥근 모서리)
    bg_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = bg_color
    bg_shape.line.fill.background()
    # 모서리 반경 줄이기
    bg_shape.adjustments[0] = 0.02

    # 코드 라인 분할
    code_lines = code_text.split('\n')
    line_height_emu = int(Pt(font_size) * 1.5)
    inset_left = Inches(0.15)
    inset_top = Inches(0.1)

    # 하이라이트 배경 (해당 라인에만)
    for ln_num in highlight_lines:
        if 1 <= ln_num <= len(code_lines):
            hl_y = int(y) + int(inset_top) + (ln_num - 1) * line_height_emu
            hl_shape = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                int(x) + int(Inches(0.05)), hl_y,
                int(w) - int(Inches(0.1)), line_height_emu,
            )
            hl_shape.fill.solid()
            hl_shape.fill.fore_color.rgb = highlight_color
            hl_shape.line.fill.background()

    # 텍스트박스 (코드 전체)
    text_shape = slide.shapes.add_textbox(x, y, w, h)
    tf = text_shape.text_frame
    tf.word_wrap = True
    set_text_inset(text_shape,
                   left=inset_left, top=inset_top,
                   right=Inches(0.15), bottom=Inches(0.1))

    for i, line in enumerate(code_lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        p.line_spacing = Pt(font_size * 1.5)
        run = p.add_run()
        run.text = line
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.font.color.rgb = text_color

    return bg_shape, text_shape


def add_styled_table(slide, x, y, w, rows, cols, data,
                     header_color=None, header_font_color=None,
                     font_size=10, header_font_size=11,
                     font_name=None, zebra=True, zebra_color=None,
                     col_widths=None, row_height=None,
                     border_color=None, border_width_pt=0.5):
    """헤더+zebra+border 스타일 테이블을 추가한다.

    Args:
        slide: 슬라이드 객체
        x, y, w: 위치/너비 (Inches/Emu)
        rows, cols: 행/열 수 (헤더 행 포함)
        data: 2D 리스트 [[h1,h2,...],[r1c1,r1c2,...],...]
        header_color: 헤더 배경 RGBColor (기본 #2D5B8A)
        header_font_color: 헤더 텍스트 RGBColor (None이면 brightness_check로 자동)
        font_size: 본문 폰트 크기 (pt)
        header_font_size: 헤더 폰트 크기 (pt)
        font_name: 폰트 이름 (None이면 기본)
        zebra: 짝수 행 배경색 적용 여부
        zebra_color: 짝수 행 배경색 (기본 #F5F5F5)
        col_widths: 열 너비 비율 리스트 (예: [1, 2, 1]). None이면 균등
        row_height: 행 높이 (Emu). None이면 Inches(0.35)
        border_color: 테두리 색상 RGBColor (기본 #D0D0D0)
        border_width_pt: 테두리 두께 (pt)

    Returns:
        생성된 테이블 도형 객체
    """
    if header_color is None:
        header_color = RGBColor(0x2D, 0x5B, 0x8A)
    if zebra_color is None:
        zebra_color = RGBColor(0xF5, 0xF5, 0xF5)
    if border_color is None:
        border_color = RGBColor(0xD0, 0xD0, 0xD0)
    if row_height is None:
        row_height = Inches(0.35)

    table_h = int(row_height) * rows
    table_shape = slide.shapes.add_table(rows, cols, int(x), int(y), int(w), table_h)
    table = table_shape.table

    # 열 너비 설정
    if col_widths is not None:
        total_ratio = sum(col_widths)
        avail_w = int(w)
        for ci, cw in enumerate(col_widths):
            if ci < cols:
                table.columns[ci].width = avail_w * cw // total_ratio
    else:
        col_w = int(w) // cols
        for ci in range(cols):
            table.columns[ci].width = col_w

    # 헤더 텍스트 색상 자동 결정
    if header_font_color is None:
        is_bright = brightness_check(header_color[0], header_color[1], header_color[2])
        header_font_color = RGBColor(0x33, 0x33, 0x33) if is_bright else RGBColor(0xFF, 0xFF, 0xFF)

    # 셀 채우기
    border_w = Pt(border_width_pt)
    for ri in range(rows):
        for ci in range(cols):
            cell = table.cell(ri, ci)

            # 텍스트 설정
            if ri < len(data) and ci < len(data[ri]):
                cell.text = str(data[ri][ci])

            # 서식 적용
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(header_font_size if ri == 0 else font_size)
                    run.font.bold = (ri == 0)
                    if font_name:
                        run.font.name = font_name
                    if ri == 0:
                        run.font.color.rgb = header_font_color
                    else:
                        run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

            # 테두리 설정 (OOXML 순서: ln* → solidFill)
            tcPr = cell._tc.get_or_add_tcPr()
            border_hex = f"{border_color[0]:02X}{border_color[1]:02X}{border_color[2]:02X}"
            for edge in ("lnL", "lnR", "lnT", "lnB"):
                ln = tcPr.makeelement(qn(f"a:{edge}"), {"w": str(int(border_w))})
                sf = ln.makeelement(qn("a:solidFill"), {})
                srgbClr = sf.makeelement(qn("a:srgbClr"), {"val": border_hex})
                sf.append(srgbClr)
                ln.append(sf)
                for old in tcPr.findall(qn(f"a:{edge}")):
                    tcPr.remove(old)
                tcPr.append(ln)

            # 배경색 (테두리 뒤에 설정 — OOXML 순서 준수)
            if ri == 0:
                set_cell_fill(cell, header_color)
            elif zebra and ri % 2 == 0:
                set_cell_fill(cell, zebra_color)

            set_cell_anchor(cell, 'ctr')

    return table_shape


# ---------------------------------------------------------------------------
# 셀 단위 서식 테이블 — 전면 격자 레이아웃 (고품질 레퍼런스 덱 관용구)
#
# 독립 도형을 좌표로 흩뿌려 '가짜 테이블'을 만드는 대신, 슬라이드 영역을
# 하나의 네이티브 a:tbl 로 격자 분할하고 셀별 배경·테두리·정렬·병합·
# rich-text 를 정밀 제어한다. 행·열 정렬이 OOXML 차원에서 보장되므로
# 빈틈없는 인상과 높은 정보 밀도를 얻는다.
# ---------------------------------------------------------------------------

_CELL_ALIGN = {
    "l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER,
    "r": PP_ALIGN.RIGHT, "j": PP_ALIGN.JUSTIFY,
}

# CT_TableCellProperties 자식 순서 (OOXML 스키마). 위반 시 파일 손상.
_TCPR_CHILD_ORDER = [
    "lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr", "cell3D",
    "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill",
    "headers", "extLst",
]


def _tcpr_insert(tcPr, el):
    """CT_TableCellProperties 자식 순서를 지키며 요소를 삽입한다."""
    tag = el.tag.split("}")[-1]
    order = _TCPR_CHILD_ORDER.index(tag) if tag in _TCPR_CHILD_ORDER else len(_TCPR_CHILD_ORDER)
    idx = len(tcPr)
    for i, child in enumerate(tcPr):
        ctag = child.tag.split("}")[-1]
        corder = _TCPR_CHILD_ORDER.index(ctag) if ctag in _TCPR_CHILD_ORDER else len(_TCPR_CHILD_ORDER)
        if corder > order:
            idx = i
            break
    tcPr.insert(idx, el)


def set_cell_border(cell, edge, color=None, width_pt=0.5, dash=None):
    """셀의 한 변(또는 여러 변)에 테두리를 설정한다. dash 패턴 지원.

    Args:
        cell: 테이블 셀
        edge: 'l' | 'r' | 't' | 'b' (또는 'lrtb' 같은 조합 문자열)
        color: RGBColor 또는 (r,g,b). None이면 #D0D0D0
        width_pt: 두께 (pt)
        dash: None | 'solid' | 'dash' | 'dot' | 'dashDot' | 'lgDash' | 'sysDash'
    """
    if len(edge) > 1:
        for e in edge:
            set_cell_border(cell, e, color=color, width_pt=width_pt, dash=dash)
        return
    if color is None:
        color = RGBColor(0xD0, 0xD0, 0xD0)
    edge_tag = {"l": "lnL", "r": "lnR", "t": "lnT", "b": "lnB"}[edge]
    hexc = f"{color[0]:02X}{color[1]:02X}{color[2]:02X}"
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn(f"a:{edge_tag}")):
        tcPr.remove(old)
    ln = tcPr.makeelement(qn(f"a:{edge_tag}"),
                          {"w": str(int(Pt(width_pt))), "cap": "flat"})
    sf = ln.makeelement(qn("a:solidFill"), {})
    sf.append(sf.makeelement(qn("a:srgbClr"), {"val": hexc}))
    ln.append(sf)
    if dash:
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": dash}))
    _tcpr_insert(tcPr, ln)


def style_cell(cell, text=None, segments=None, fill=None,
               font_color=None, font_size=None, bold=None, italic=None,
               align=None, anchor="ctr", font_name=None, wrap=True,
               border_edges=None, border_color=None, border_width_pt=0.5,
               border_dash=None, margins=None):
    """테이블 셀 하나에 종합 서식을 적용한다 (전면 격자 레이아웃용).

    Args:
        text: 평문 텍스트 (segments 와 택일)
        segments: add_rich_text 형식 세그먼트 (셀 내 혼합 서식)
        fill: 배경 RGBColor
        font_color/font_size/bold/italic/font_name: 폰트 서식
        align: 'l'|'c'|'r'|'j' 또는 PP_ALIGN
        anchor: 't'|'ctr'|'b' 세로정렬
        border_edges: 'lrtb' 등 테두리 변. None이면 변경 없음
        border_color/border_width_pt/border_dash: 테두리 스타일
        margins: (left, top, right, bottom) Emu 셀 내부 여백
    Returns:
        cell
    """
    tf = cell.text_frame
    tf.word_wrap = wrap

    if segments is not None:
        p = tf.paragraphs[0]
        for r in list(p.runs):
            r._r.getparent().remove(r._r)
        for seg in segments:
            run = p.add_run()
            if isinstance(seg, str):
                run.text = seg
                if font_color is not None:
                    run.font.color.rgb = font_color
                if font_size is not None:
                    run.font.size = Pt(font_size)
                if bold is not None:
                    run.font.bold = bold
                if font_name:
                    run.font.name = font_name
            else:
                run.text = str(seg.get("text", ""))
                if "color" in seg:
                    run.font.color.rgb = seg["color"]
                elif font_color is not None:
                    run.font.color.rgb = font_color
                run.font.size = Pt(seg.get("font_size", font_size or 11))
                run.font.bold = seg.get("bold", bool(bold))
                if seg.get("italic"):
                    run.font.italic = True
                fn = seg.get("font_name", font_name)
                if fn:
                    run.font.name = fn
    elif text is not None:
        cell.text = str(text)

    plain = segments is None
    for p in tf.paragraphs:
        if align is not None:
            p.alignment = _CELL_ALIGN.get(align, align)
        for run in p.runs:
            if plain and font_size is not None:
                run.font.size = Pt(font_size)
            if plain and bold is not None:
                run.font.bold = bold
            if italic is not None:
                run.font.italic = italic
            if plain and font_name:
                run.font.name = font_name
            if plain and font_color is not None:
                run.font.color.rgb = font_color

    if fill is not None:
        set_cell_fill(cell, fill)
    if border_edges:
        set_cell_border(cell, border_edges, color=border_color,
                        width_pt=border_width_pt, dash=border_dash)
    if anchor is not None:
        set_cell_anchor(cell, anchor)
    if margins is not None:
        l, t, r, b = margins
        tcPr = cell._tc.get_or_add_tcPr()
        for k, v in (("marL", l), ("marT", t), ("marR", r), ("marB", b)):
            if v is not None:
                tcPr.set(k, str(int(v)))
    return cell


def merge_cells(table, r0, c0, r1, c1):
    """(r0,c0)~(r1,c1) 직사각형 범위를 병합하고 origin 셀을 반환한다."""
    table.cell(r0, c0).merge(table.cell(r1, c1))
    return table.cell(r0, c0)


def add_grid_table(slide, x, y, w, h, nrows, ncols, cells=None,
                   col_widths=None, row_heights=None, font_name=None,
                   default_fill=None, default_font_size=11,
                   gridlines=True, gridline_color=None, gridline_width_pt=0.5):
    """전면 격자 레이아웃용 셀 단위 서식 테이블.

    고품질 레퍼런스 덱의 핵심 관용구. 슬라이드 영역을 하나의 네이티브 표로
    격자 분할하고 셀별 배경·테두리·정렬·병합·rich-text 를 정밀 제어한다.

    Args:
        x, y, w, h: 표 전체 bbox (Emu/Inches)
        nrows, ncols: 행/열 수
        cells: dict[(row, col)] -> style_cell kwargs + 선택적 'span': (rowspan, colspan)
               예: {(0,0): {"text":"항목", "bold":True, "fill":HEADER,
                            "font_color":WHITE, "span":(1,2)}}
        col_widths: 열 너비 비율 리스트 또는 None(균등)
        row_heights: 행 높이 비율 리스트 또는 None(균등)
        default_fill: 모든 셀 기본 배경 (None이면 채우지 않음)
        gridlines: 모든 셀에 기본 테두리 적용 여부
        gridline_color: 기본 테두리 색 (None이면 #D9D9D9)
    Returns:
        (table_shape, table)
    """
    if gridline_color is None:
        gridline_color = RGBColor(0xD9, 0xD9, 0xD9)
    table_shape = slide.shapes.add_table(nrows, ncols, int(x), int(y), int(w), int(h))
    table = table_shape.table
    table.first_row = False
    table.horz_banding = False
    # 기본 테마 스타일(파란 헤더/밴딩) 제거 → 셀 서식을 완전히 명시 제어.
    # "No Style, No Grid" → 채우지 않은 셀은 투명(흰 배경), 테두리는 우리가 명시.
    _tblPr = table._tbl.find(qn("a:tblPr"))
    if _tblPr is not None:
        for _sid in _tblPr.findall(qn("a:tableStyleId")):
            _sid.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"

    total_w = int(w)
    if col_widths:
        tot = sum(col_widths)
        for ci in range(ncols):
            cw = col_widths[ci] if ci < len(col_widths) else 1
            table.columns[ci].width = int(total_w * cw / tot)
    else:
        cw = total_w // ncols
        for ci in range(ncols):
            table.columns[ci].width = cw

    total_h = int(h)
    if row_heights:
        tot = sum(row_heights)
        for ri in range(nrows):
            rh = row_heights[ri] if ri < len(row_heights) else 1
            table.rows[ri].height = int(total_h * rh / tot)
    else:
        rh = total_h // nrows
        for ri in range(nrows):
            table.rows[ri].height = rh

    # 병합 먼저 (이후 서식은 origin 셀만 대상)
    if cells:
        for (r, c), st in cells.items():
            span = st.get("span")
            if span and (span[0] > 1 or span[1] > 1):
                r1 = min(r + span[0] - 1, nrows - 1)
                c1 = min(c + span[1] - 1, ncols - 1)
                merge_cells(table, r, c, r1, c1)

    # 기본 그리드/배경
    for r in range(nrows):
        for c in range(ncols):
            cell = table.cell(r, c)
            if cell.is_spanned:
                continue
            if default_fill is not None:
                set_cell_fill(cell, default_fill)
            if gridlines:
                set_cell_border(cell, "lrtb", color=gridline_color,
                                width_pt=gridline_width_pt)

    # 셀별 명시 서식
    if cells:
        for (r, c), st in cells.items():
            cell = table.cell(r, c)
            if cell.is_spanned:
                continue
            kwargs = {k: v for k, v in st.items() if k != "span"}
            kwargs.setdefault("font_name", font_name)
            kwargs.setdefault("font_size", default_font_size)
            style_cell(cell, **kwargs)

    return table_shape, table


# ---------------------------------------------------------------------------
# 시각 자산 채널 — 이미지/스크린샷/차트 (1급 시민)
#
# python-pptx 도형만으로 모든 것을 '그리는' 대신, 실제 UI 캡처·matplotlib
# 차트 PNG·로고·일러스트를 임베드하고, 수치는 네이티브 차트로 표현한다.
# ---------------------------------------------------------------------------


def add_picture(slide, image_path, x, y, w=None, h=None,
                shadow=False, line_color=None, line_width_pt=1.0):
    """이미지/스크린샷/차트 PNG 를 슬라이드에 삽입한다.

    Args:
        image_path: 파일 경로 (예: assets/screenshot.png, sources/{name}/assets/...)
        x, y: 좌상단 위치 (Emu/Inches)
        w, h: 너비/높이. 하나만 주면 비율 유지, 둘 다 None이면 원본 크기
        shadow: 그림자 추가
        line_color: 테두리 색 (RGBColor)
        line_width_pt: 테두리 두께
    Returns:
        Picture shape
    """
    p = Path(image_path)
    if not p.exists():
        raise FileNotFoundError(f"이미지 파일 없음: {image_path}")
    kw = {}
    if w is not None:
        kw["width"] = int(w)
    if h is not None:
        kw["height"] = int(h)
    pic = slide.shapes.add_picture(str(p), int(x), int(y), **kw)
    if line_color is not None:
        pic.line.color.rgb = line_color
        pic.line.width = Pt(line_width_pt)
    if shadow:
        add_shadow(pic, blur_pt=6, dist_pt=3, opacity_pct=35)
    return pic


def add_chart(slide, x, y, w, h, chart_type, categories, series,
              legend=True, legend_position="BOTTOM", title=None,
              number_format=None, data_labels=False, colors=None,
              font_size=10):
    """카테고리 차트(막대/선/원/도넛 등)를 추가한다.

    3개 이상의 수치는 텍스트 나열 대신 차트로 표현 (CLAUDE.md 규칙).

    Args:
        chart_type: XL_CHART_TYPE enum 또는 문자열
            ("COLUMN_CLUSTERED", "BAR_CLUSTERED", "LINE_MARKERS",
             "PIE", "DOUGHNUT", "COLUMN_STACKED", ...)
        categories: 카테고리 라벨 리스트
        series: [(name, [values...]), ...]
        legend_position: "BOTTOM"|"RIGHT"|"TOP"|"LEFT" 또는 XL_LEGEND_POSITION
        colors: 시리즈(또는 PIE/DOUGHNUT 포인트) 색상 RGBColor 리스트
    Returns:
        chart 객체
    """
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

    cd = CategoryChartData()
    cd.categories = list(categories)
    for name, vals in series:
        cd.add_series(str(name), tuple(vals))

    ct = chart_type
    if isinstance(ct, str):
        ct = getattr(XL_CHART_TYPE, chart_type)
    gframe = slide.shapes.add_chart(ct, int(x), int(y), int(w), int(h), cd)
    chart = gframe.chart

    chart.has_legend = bool(legend)
    if legend and legend_position is not None:
        pos = legend_position
        if isinstance(pos, str):
            pos = getattr(XL_LEGEND_POSITION, legend_position)
        chart.legend.position = pos
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(font_size)

    if title:
        chart.has_title = True
        chart.chart_title.text_frame.text = str(title)
    else:
        chart.has_title = False

    plot = chart.plots[0]
    if data_labels:
        plot.has_data_labels = True
        if number_format:
            plot.data_labels.number_format = number_format
            plot.data_labels.number_format_is_linked = False
        plot.data_labels.font.size = Pt(font_size)

    try:
        chart.category_axis.tick_labels.font.size = Pt(font_size)
        chart.value_axis.tick_labels.font.size = Pt(font_size)
        if number_format:
            chart.value_axis.tick_labels.number_format = number_format
            chart.value_axis.tick_labels.number_format_is_linked = False
    except Exception:
        pass  # PIE/DOUGHNUT 은 축 없음

    if colors:
        if ct in (XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT):
            ser = plot.series[0]
            for i, pt in enumerate(ser.points):
                if i < len(colors):
                    pt.format.fill.solid()
                    pt.format.fill.fore_color.rgb = colors[i]
        else:
            for i, s in enumerate(plot.series):
                if i < len(colors):
                    s.format.fill.solid()
                    s.format.fill.fore_color.rgb = colors[i]
    return chart


# ---------------------------------------------------------------------------
# 상태 머신 다이어그램 — Mermaid 상태 전이의 표준 PPT 변환
# ---------------------------------------------------------------------------


def add_state_machine(slide, states, transitions, area=None,
                      state_color=None, active_color=None,
                      font_size=10, arrow=True):
    """상태 머신 다이어그램을 슬라이드에 추가한다.

    Mermaid 상태 전이 다이어그램의 표준 PPT 변환 대상.

    Args:
        slide: 슬라이드 객체
        states: 상태 이름 리스트 (예: ["대기", "진행중", "완료", "실패"])
        transitions: (from, to, label) 튜플 리스트
            (예: [("대기", "진행중", "시작"), ("진행중", "완료", "완료")])
        area: _CellRect 또는 (left, top, width, height). None이면 CONTENT_SAFE
        state_color: 상태 도형 채우기 RGBColor (기본 #E3F2FD)
        active_color: 첫 번째/마지막 상태 강조색 (기본 #1565C0)
        font_size: 상태 라벨 폰트 크기
        arrow: 커넥터에 화살표 추가 여부

    Returns:
        dict: {shapes: {state_name: shape}, connectors: [connector]}
    """
    if state_color is None:
        state_color = RGBColor(0xE3, 0xF2, 0xFD)
    if active_color is None:
        active_color = RGBColor(0x15, 0x65, 0xC0)

    n = len(states)
    if n == 0:
        return {"shapes": {}, "connectors": []}

    # 그리드 레이아웃 결정
    if n <= 5:
        rows, cols = 1, n
    else:
        rows = 2
        cols = math.ceil(n / 2)

    grid = calc_grid(rows, cols, area=area, gap=Inches(0.3))

    # 상태 도형 생성
    shapes = {}
    for i, state_name in enumerate(states):
        r = i // cols
        c = i % cols
        if r >= grid.rows or c >= grid.cols:
            break
        cell = grid[r][c]

        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            cell.left, cell.top, cell.width, cell.height,
        )
        # 첫/마지막 상태 강조
        if i == 0 or i == n - 1:
            shape.fill.solid()
            shape.fill.fore_color.rgb = active_color
            font_color = RGBColor(0xFF, 0xFF, 0xFF)
        else:
            shape.fill.solid()
            shape.fill.fore_color.rgb = state_color
            font_color = RGBColor(0x21, 0x21, 0x21)

        shape.line.color.rgb = RGBColor(0x90, 0xCA, 0xF9)
        shape.line.width = Pt(1)

        tf = shape.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = state_name
        run.font.size = Pt(font_size)
        run.font.bold = True
        run.font.color.rgb = font_color
        set_body_anchor(shape, 'ctr')
        set_text_inset(shape)

        shapes[state_name] = shape

    # 같은 행 도형 정렬
    for r in range(rows):
        row_shapes = []
        for c in range(cols):
            idx = r * cols + c
            if idx < n:
                row_shapes.append(shapes[states[idx]])
        if len(row_shapes) >= 2:
            align_shapes(*row_shapes, axis='h')

    # 전이 커넥터
    connectors = []
    for from_state, to_state, label in transitions:
        if from_state not in shapes or to_state not in shapes:
            continue
        conn = add_smart_connector(slide, shapes[from_state], shapes[to_state],
                                    arrow=arrow)
        connectors.append(conn)

        # 라벨 텍스트박스 (커넥터 중간에 작은 텍스트)
        if label:
            s_a = shapes[from_state]
            s_b = shapes[to_state]
            mid_x = (s_a.left + s_a.width // 2 + s_b.left + s_b.width // 2) // 2
            mid_y = (s_a.top + s_a.height // 2 + s_b.top + s_b.height // 2) // 2
            label_w = Inches(0.8)
            label_h = Inches(0.25)
            add_textbox(slide,
                        mid_x - int(label_w) // 2,
                        mid_y - int(label_h) // 2,
                        label_w, label_h,
                        label, font_size=8,
                        align=PP_ALIGN.CENTER,
                        color=RGBColor(0x66, 0x66, 0x66))

    return {"shapes": shapes, "connectors": connectors}


def add_routed_connector(slide, waypoints, arrow=True, color=None, width_pt=1.5):
    """다중 경유점을 거치는 라우팅 커넥터 (freeform 기반).

    직선/ELBOW 커넥터로는 중간 도형을 우회할 수 없을 때 사용.
    각 경유점 사이는 직선으로 연결 (직각 경로 권장).

    Args:
        slide: 슬라이드 객체
        waypoints: [(x, y), ...] 경유점 리스트 (최소 2개, Emu 단위)
        arrow: True이면 끝점에 화살표 추가
        color: RGBColor 또는 None (기본 #666666)
        width_pt: 선 두께 (pt)

    Returns:
        생성된 freeform 도형
    """
    if len(waypoints) < 2:
        raise ValueError("waypoints는 최소 2개 필요")

    x0, y0 = int(waypoints[0][0]), int(waypoints[0][1])
    builder = slide.shapes.build_freeform(x0, y0)
    builder.add_line_segments(
        [(int(x), int(y)) for x, y in waypoints[1:]],
        close=False,
    )
    shape = builder.convert_to_shape()

    # 채우기 없음, 선만
    shape.fill.background()
    shape.line.color.rgb = color or RGBColor(0x66, 0x66, 0x66)
    shape.line.width = Pt(width_pt)

    # 화살표 (끝점)
    if arrow:
        spPr = shape._element.find(qn("p:spPr"))
        if spPr is not None:
            ln = spPr.find(qn("a:ln"))
            if ln is None:
                ln = spPr.makeelement(qn("a:ln"), {})
                spPr.append(ln)
            for old in ln.findall(qn("a:tailEnd")):
                ln.remove(old)
            ln.append(ln.makeelement(
                qn("a:tailEnd"),
                {"type": "triangle", "w": "med", "len": "med"},
            ))


# ---------------------------------------------------------------------------
# 동적 페이지 수 적용 — "‹#› / 00" → "‹#› / <total>"
# ---------------------------------------------------------------------------


def apply_page_total(prs, total=None):
    """마스터의 페이지 번호 텍스트박스에서 "/ <N>" 부분만 실제 총 슬라이드 수로 치환.

    템플릿(ref/표지.pptx)의 master shape '슬라이드 번호 개체 틀 5' 는
    `<a:fld>‹#›</a:fld> <a:r>/ 00</a:r>` 형태로 구성됨 ("‹#›"는 PowerPoint 필드,
    " / 00" 은 별도 text run). 이 함수는 **" / 00" run 만** 찾아서
    " / <total>" 로 교체한다. 필드 구조는 그대로 유지.

    Args:
        prs: Presentation 객체
        total: 치환할 총 페이지 수. None이면 len(prs.slides) 사용.

    Returns:
        치환 성공 여부 (bool). 대상 run 을 찾지 못하면 False.

    Note:
        이 함수는 master XML을 수정한다. TemplateGuard 안에서 호출하지 말고
        Assembler 가 guard 종료 후 save 직전에 호출해야 한다.
        master_xml_hash 는 "/ <N>" 패턴을 "/ 00" 으로 정규화하므로 hash 비교에 영향 없음.
    """
    import re

    if total is None:
        total = len(prs.slides)

    changed = False
    for master in prs.slide_masters:
        for shape in master.shapes:
            if not shape.has_text_frame:
                continue
            tf = shape.text_frame
            # 전체 text 에 "‹#›" 가 있어야 페이지 번호 textbox 로 간주
            if "‹#›" not in tf.text:
                continue
            # run 단위 순회 — "/ <digits>" 패턴만 찾아 교체
            for para in tf.paragraphs:
                for run in para.runs:
                    m = re.match(r"^(\s*/\s*)\d+(\s*)$", run.text)
                    if m:
                        run.text = f"{m.group(1)}{total}{m.group(2)}"
                        changed = True
    return changed

    return shape
