"""Slide 01 — 표지 (Cover).

ref/표지.pptx 의 "제목 슬라이드" 레이아웃을 그대로 사용.
setup_cover() 가 placeholder 텍스트만 채우고 배경 이미지 · 로고는 템플릿 그대로 상속.
"""

from ppt_utils import setup_cover


def build_slide_1(slide):
    setup_cover(
        slide,
        title="AI 코딩 어시스턴트\n도입 영향 분석",
        # purpose/author/department 는 setup_cover() 기본값 사용
        # (정보공유 / 강민규 책임 / 미래융합설계센터 알고리즘개발팀)
        date="2026.04.21",
    )
