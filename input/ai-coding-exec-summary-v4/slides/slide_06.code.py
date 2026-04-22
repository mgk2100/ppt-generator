from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from ppt_utils import (
    set_title, clear_placeholders,
    add_textbox, add_accent_bar,
    add_styled_table, set_body_anchor,
)
from template_contract import CONTENT_SAFE


# ---------- 색상 ----------
C_NAVY   = RGBColor(0x1F, 0x49, 0x7D)
C_INK    = RGBColor(0x1A, 0x1F, 0x2E)
C_LGRAY  = RGBColor(0xE0, 0xE4, 0xEA)
C_ZEBRA  = RGBColor(0xF5, 0xF7, 0xFA)
C_BORDER = RGBColor(0xD0, 0xD4, 0xDA)
C_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)


def build_slide_6(slide):
    set_title(slide, "부록  ·  전문용어 해설")
    clear_placeholders(slide, keep=[0])

    # ================================================================
    # 레이아웃 계획 (CONTENT_SAFE: left=0.28, top=0.68, w=10.28, h=6.34)
    # - Key Message Bar: top 0.72, height 0.50
    # - Glossary Table:  top 1.32, height ~5.62 (11 rows x ~0.48")
    # ================================================================

    safe_left  = CONTENT_SAFE.left
    safe_top   = CONTENT_SAFE.top
    safe_w     = CONTENT_SAFE.width
    safe_bottom = CONTENT_SAFE.bottom

    # ---------------- 1) Key Message Bar ----------------
    kmb_x = safe_left
    kmb_y = safe_top + Inches(0.04)
    kmb_w = safe_w
    kmb_h = Inches(0.50)

    # 좌측 액센트 바
    accent_bar_w = Inches(0.10)
    add_accent_bar(slide, kmb_x, kmb_y, accent_bar_w, kmb_h, C_NAVY)

    # 배경 박스 (연한 회색)
    kmb_bg_x = kmb_x + accent_bar_w
    kmb_bg_w = kmb_w - accent_bar_w
    kmb_bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, kmb_bg_x, kmb_y, kmb_bg_w, kmb_h
    )
    kmb_bg.fill.solid()
    kmb_bg.fill.fore_color.rgb = C_LGRAY
    kmb_bg.line.fill.background()

    # Key Message 텍스트 (박스 안)
    kmb_text = add_textbox(
        slide,
        kmb_bg_x + Inches(0.15), kmb_y,
        kmb_bg_w - Inches(0.30), kmb_h,
        "Key Message  |  본문에 등장한 자동차 SW · AI 안전 관련 표준 · 약어 해설",
        font_size=13, bold=True, color=C_INK, align=PP_ALIGN.LEFT,
    )
    set_body_anchor(kmb_text, 'ctr')

    # ---------------- 2) Glossary Table ----------------
    tbl_x = safe_left
    tbl_y = kmb_y + kmb_h + Inches(0.12)
    tbl_w = safe_w

    # 테이블 높이는 rows * row_height. 하단 안전 여유 0.06" 확보.
    available_h = safe_bottom - tbl_y - Inches(0.06)
    rows = 11  # 헤더 + 10행
    # 행 높이 계산: available_h / rows. 최대 0.49, 최소 0.42
    raw_rh = int(available_h) // rows
    row_h_min = int(Inches(0.42))
    row_h_max = int(Inches(0.49))
    if raw_rh > row_h_max:
        row_h = row_h_max
    elif raw_rh < row_h_min:
        row_h = row_h_min
    else:
        row_h = raw_rh

    data = [
        ["용어 / 약어", "분류", "설명"],
        [
            "ASPICE CL2 / CL3",
            "프로세스 품질",
            "자동차 SW 개발 프로세스 품질 표준. OEM이 Tier-1 에 거의 예외 없이 요구하는 인증 수준. 모든 변경에 '요구→설계→코드→테스트' 추적성 필요.",
        ],
        [
            "ISO 26262 (ASIL A~D)",
            "기능안전",
            "차량 기능안전 표준. ASIL(Automotive Safety Integrity Level)은 안전등급 (A 낮음 → D 최고). 안전 관련 항목은 검증 · 증거 자료 추가 요구.",
        ],
        [
            "ISO 21448 (SOTIF)",
            "의도된 기능 안전",
            "Safety Of The Intended Functionality. ADAS · 자율주행처럼 '오작동 아니지만 의도 벗어남' 리스크에 적용.",
        ],
        [
            "ISO/PAS 8800:2024",
            "AI 차량 안전",
            "AI 시스템의 차량 안전 표준 (2024.12 신규 발행). AI 생성 코드 검증 프로세스의 참조 기준. Geely 가 2025.07 세계 최초 인증.",
        ],
        [
            "MISRA C:2025",
            "코딩 가이드",
            "자동차 임베디드 C 언어 코딩 가이드라인. 2025년 개정에서 'AI 생성 코드도 수기 코드와 동일 규칙 준수' 명문화.",
        ],
        [
            "AUTOSAR / BSW / MCAL",
            "SW 아키텍처",
            "차량 제어 기반 SW 아키텍처. BSW = Basic Software, MCAL = Microcontroller Abstraction Layer. Tier-1 벤더 라이선스 코드라 공개 OSS 거의 없음.",
        ],
        [
            "SPACE 프레임워크",
            "개발 생산성 측정",
            "Satisfaction · Performance · Activity · Communication · Efficiency 5축. GitHub · ACM이 제안. BMW 사내 측정에 사용.",
        ],
        [
            "EU AI Act",
            "법규",
            "EU 인공지능 규제법. 2024.08 발효, 2026.08 일반 high-risk AI 적용, 2027.08 자동차 등 embedded high-risk AI 완전 준수 D-Day.",
        ],
        [
            "SDV",
            "차량 아키텍처",
            "Software Defined Vehicle. 차량 기능을 SW 업데이트로 확장 · 변경하는 차세대 아키텍처. 현대모비스 Mobis Development Studio 대상.",
        ],
        [
            "RAG / 파인튜닝",
            "AI 학습 기법",
            "RAG = Retrieval-Augmented Generation (외부 데이터 기반 응답 증강). 파인튜닝 = 특정 도메인 데이터로 모델 재학습. 사내 BSW 코드로 적용 시 효과 추가 가능.",
        ],
    ]

    add_styled_table(
        slide,
        tbl_x, tbl_y, tbl_w,
        rows=rows, cols=3,
        data=data,
        header_color=C_NAVY,
        header_font_color=C_WHITE,
        font_size=10,
        header_font_size=11,
        zebra=True,
        zebra_color=C_ZEBRA,
        col_widths=[3, 2, 7],
        row_height=row_h,
        border_color=C_BORDER,
        border_width_pt=0.5,
    )
