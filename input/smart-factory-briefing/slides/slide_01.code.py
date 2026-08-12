"""slide_01 — 표지: AI 플랫폼 개발 현황 및 AI 서버 운영 경험 공유.

spec: input/smart-factory-briefing/slides/slide_01.spec.yaml
레이아웃 "제목 슬라이드" — setup_cover 만 호출한다.
추가 도형·배경 Picture·마스터 요소 변경 금지 (마스터 고정).
"""

from ppt_utils import setup_cover


def build_slide_01(slide):
    setup_cover(
        slide,
        "AI 플랫폼 개발 현황 및 AI 서버 운영 경험 공유",
        purpose="정보공유",
        author="강민규 책임",
        department="SDP센터 알고리즘개발팀",
        date="2026.08.13",
    )


# 하네스 호환 별칭 (build_slide_{spec.idx})
build_slide_1 = build_slide_01
