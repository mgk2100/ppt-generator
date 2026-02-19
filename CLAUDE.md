# PPT Generator

## 목적
입력 자료를 분석하여 python-pptx 기반 PPT를 자동 생성한다.

## 작동 방식
Claude가 python-pptx 코드를 직접 작성 → 실행 → PPT 생성.
고정 렌더러 없음. 매번 콘텐츠에 맞는 코드를 새로 작성한다.

## 지원 입력
하나의 /generate-ppt 명령으로 다양한 입력을 자동 판단하여 처리한다.
- 기존 분석 결과 폴더 (project-analyzer 결과물)
- 소스 코드 프로젝트
- 텍스트/문서 파일 (.md, .txt, .pdf 등)
- 기업명, 주제 등 텍스트 입력

## 핵심 원칙
- 있는 데이터만 슬라이드로 만든다. 데이터 없으면 해당 슬라이드 생성 안 함.
- 카드 개수 = 실제 데이터 항목 수. 빈 카드 금지.
- 슬라이드 타입은 데이터 성격으로 결정한다.

## 적응형 슬라이드 구조

고정 구조 금지. 전체 구조도 데이터가 결정한다.

### 구조 패턴 선택

데이터의 성격과 볼륨으로 전체 구조 패턴을 결정한다:

| 조건 | 패턴 | 구조 | 예시 |
|------|------|------|------|
| 2-3개 독립 그룹, 각 moderate | **섹션 분할형** | 표지→목차→[섹션구분→콘텐츠]×N→마무리 | 프로젝트 소개, 분석 보고서 |
| 5개+ 연속적 주제, 하나의 깊은 흐름 | **내러티브형** | 표지→연속 슬라이드→마무리 | 기술 딥다이브, 설계 문서 |
| 단일 주제, 소량 데이터 | **컴팩트형** | 표지→콘텐츠(3-5장)→마무리 | 짧은 제안, 상태 보고 |
| 다층 구조 + 부록 필요 | **본문+부록형** | 표지→본문→부록 구분→부록 | 온보딩, 기술 레퍼런스 |

### 목차/섹션 구분 판단

- **목차 슬라이드**: 독립적인 그룹이 3개 이상이고 청중이 전체 맥을 먼저 파악해야 할 때만 생성. 2개 이하이거나 순차 내러티브이면 생략.
- **섹션 구분 슬라이드**: 주제가 명확히 전환될 때만 사용. 같은 맥락이 계속되면 생략. 섹션 수 = 실제 데이터 그룹 수 (2개면 2개, 5개면 5개).

### 슬라이드 병합/분할

- **병합**: 관련 데이터가 각각 항목 2-3개 이하이고 한 슬라이드에 시각적으로 수용 가능하면 병합
- **분할**: 단일 주제라도 항목이 7개+이거나 다이어그램이 복잡하면 여러 슬라이드로 분할
- **기준**: 슬라이드 하나에 시각적 요소 3-5개가 적정. 6개+이면 분할 고려

### 내러티브 흐름

콘텐츠 목적에 따라 슬라이드 순서를 결정한다:
- **기술 소개**: 왜 필요한가 → 전체 구조 → 핵심 기능 → 기술 상세 → 운영/현황
- **분석 보고서**: 개요/규모 → 아키텍처 → 데이터 흐름 → 기술 스택 → 리스크/이슈
- **딥다이브**: 문제 정의 → 전체 조감 → [메커니즘 A → B → C...] → 리스크 → 운영
- **제안/설득**: 문제 → 해결책 → 근거/데이터 → 기대효과 → 다음 단계

## 금지 사항
- **이모지를 시각 지표로 사용 금지** (🔴🟠🟡 등). 색상 도형(`OVAL` 또는 `make_icon_circle()`) 사용
- **가짜 그림자(오프셋 사각형) 금지**. `add_shadow()` 사용
- **모든 다이어그램에서 사각형만 사용 금지**. 의미 도형 활용 (CAN, CLOUD, CUBE, GEAR_6 등)
- **통계 데이터를 텍스트 숫자만으로 나열 금지**. 3개+ 수치 → 차트 활용
- **ppt_utils에 있는 함수를 직접 재구현 금지**. import하여 사용
- **"반드시 N섹션" 같은 고정 구조 강제 금지**. 전체 구조는 데이터 볼륨과 성격이 결정

## 실행 방법
- 생성 스크립트는 `/tmp/`에 작성하고, 실행 후 삭제한다
- output/ 디렉토리에는 .pptx 파일만 남긴다 (.py 파일 금지)
- 테스트용 파일(test_*.pptx 등)은 검증 완료 후 즉시 삭제한다

## 폰트
- 주제와 분위기에 맞는 폰트를 자유롭게 선택한다
- 사용 전 fc-list 명령으로 시스템에 설치된 폰트를 확인한다
- 설치된 폰트 중에서만 선택한다
- 커스텀 폰트 경로: ref/fonts/

---

## 생성 스크립트 관례

### 스크립트 구조
```python
#!/usr/bin/env python3
"""[프레젠테이션 제목] - PPT 생성 스크립트"""

import sys
sys.path.insert(0, "/home/ubuntu/Share/ppt-generator")

from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.ns import qn

from ppt_utils import (
    load_template, get_layout, clear_placeholders,
    ensure_fonts, set_cell_anchor, add_arrowhead,
    add_shadow, set_shape_opacity, add_gradient_stop,
    make_icon_circle, brightness_check,
    add_textbox, add_para, set_body_anchor,
    OUTPUT_DIR,
)

ensure_fonts()
prs = load_template()

# 슬라이드 크기 변수
SLIDE_W = prs.slide_width
SLIDE_H = prs.slide_height

# 색상 팔레트 — 이 프레젠테이션 전용
PRIMARY = RGBColor(...)

# --- 슬라이드 1: ... ---
# (각 슬라이드 섹션에 설계 의도 주석)

# 저장
output_path = OUTPUT_DIR / "파일명.pptx"
prs.save(str(output_path))
print(f"생성 완료: {output_path}")
```

### 필수 관례
- 슬라이드 크기: `SLIDE_W`, `SLIDE_H` 변수명 통일
- 도형에 텍스트 넣을 때 `tf.word_wrap = True` 필수
- 한글 문자 폭 ≈ 라틴 1.5배 — 박스 너비 계산 시 반영
- 레이아웃: `get_layout(prs, "제목 슬라이드")`, `get_layout(prs, "제목 및 내용")`, `get_layout(prs, "제목 및 내용 (페이지 번호 삭제)")`
- 슬라이드 추가 후 `clear_placeholders(slide, keep=[0])` 으로 유령 텍스트 제거

---

## python-pptx 능력 레퍼런스

### ppt_utils 함수 목록
| 함수 | 용도 |
|------|------|
| `load_template()` | 표지.pptx 로드, 빈 프레젠테이션 반환 |
| `get_layout(prs, name)` | 이름으로 슬라이드 레이아웃 검색 |
| `clear_placeholders(slide, keep=[])` | 유령 플레이스홀더 제거 |
| `ensure_fonts()` | ref/fonts/ 폰트 시스템 설치 |
| `add_textbox(slide, x, y, w, h, text, ...)` | 텍스트박스 추가 (font_name, font_size, color, bold, align) |
| `add_para(text_frame, text, ...)` | 기존 text_frame에 단락 추가 (font_name, font_size, color, bold, align, space_before, space_after) |
| `set_cell_anchor(cell, 'ctr')` | 테이블 셀 세로정렬 |
| `set_body_anchor(shape, 'ctr')` | 도형 텍스트 세로정렬 |
| `add_arrowhead(connector)` | 커넥터에 화살표 머리 추가 |
| `add_shadow(shape, blur_pt, dist_pt, direction, opacity_pct, color)` | 도형에 그림자 추가 |
| `set_shape_opacity(shape, opacity_pct)` | 도형 채우기 투명도 |
| `add_gradient_stop(shape, position, r, g, b)` | 그라디언트 3번째+ stop 추가 |
| `make_icon_circle(slide, x, y, size, fill_color, text, font_size)` | 원형 아이콘/배지 |
| `brightness_check(r, g, b)` | 밝기 판단 (True=밝음→어두운 텍스트) |

### 그림자 (Shadow)
```python
add_shadow(card, blur_pt=6, dist_pt=3, direction=2700000, opacity_pct=35)
```
- direction: 2700000=아래, 5400000=오른쪽아래
- blur_pt 4~8, dist_pt 2~4, opacity_pct 30~50이 자연스러움

### 그라디언트 (Gradient Fill)
```python
shape.fill.gradient()
shape.fill.gradient_stops[0].color.rgb = RGBColor(0x1A, 0x1A, 0x2E)
shape.fill.gradient_stops[0].position = 0.0
shape.fill.gradient_stops[1].color.rgb = RGBColor(0x16, 0x21, 0x3E)
shape.fill.gradient_stops[1].position = 1.0
shape.fill.gradient_angle = 270.0  # 위→아래
add_gradient_stop(shape, position=0.5, r=0x20, g=0x30, b=0x50)  # 3-stop
```
- **gradient_angle 단위: 도(degrees)**. 0=좌→우, 90=하→상, 180=우→좌, 270=상→하
- **절대 1/60000도 단위(16200000 등)를 쓰지 말 것** — 파일 손상됨

### 투명도 (Opacity)
```python
overlay.fill.solid()
overlay.fill.fore_color.rgb = RGBColor(0x00, 0x00, 0x00)
set_shape_opacity(overlay, opacity_pct=30)
overlay.line.fill.background()
```

### 차트 (Charts)
`slide.shapes.add_chart(chart_type, x, y, w, h, chart_data)` 사용.
- 비율/비중 → `DOUGHNUT`, `PIE`
- 비교 → `COLUMN_CLUSTERED`, `BAR_CLUSTERED`
- 추세 → `LINE`, `LINE_MARKERS`
- 분포 → `SCATTER`, `BUBBLE`
- 누적 → `COLUMN_STACKED`, `BAR_STACKED`

### 도형 종류
`MSO_SHAPE` enum으로 193종 도형 사용 가능. 사각형만 쓰지 말 것.

| 용도 | 도형 |
|------|------|
| 다이어그램 | `HEXAGON`, `CHEVRON`, `PENTAGON`, `DIAMOND` |
| 플로차트 | `FLOWCHART_PROCESS`, `FLOWCHART_DECISION`, `FLOWCHART_DATA`, `FLOWCHART_TERMINATOR` |
| 인프라 | `CUBE`(서버), `CAN`(DB), `CLOUD`(클라우드), `GEAR_6`(서비스), `FUNNEL`(깔때기) |
| 화살표 | `RIGHT_ARROW`, `CHEVRON`, `NOTCHED_RIGHT_ARROW`, `CURVED_RIGHT_ARROW` |
| 사각형 변형 | `ROUNDED_RECTANGLE`, `SNIP_1_RECTANGLE`, `ROUND_1_RECTANGLE` |
| 콜아웃 | `RECTANGULAR_CALLOUT`, `ROUNDED_RECTANGULAR_CALLOUT`, `CLOUD_CALLOUT` |

### 커넥터 유형
```python
connector = slide.shapes.add_connector(MSO_CONNECTOR.ELBOW, x1, y1, x2, y2)
add_arrowhead(connector)
```
- `STRAIGHT`: 직선, `ELBOW`: 꺾인선 (아키텍처용), `CURVE`: 곡선

### 기타 기능
- **그룹 도형**: `slide.shapes.add_group_shape()` — 관련 컴포넌트 묶기
- **프리폼**: `slide.shapes.build_freeform(x, y)` — 커스텀 도형
- **이미지**: `slide.shapes.add_picture(path, x, y, w, h)`
- **회전**: `shape.rotation = 45.0`

---

## 디자인 원칙

### 타이포그래피
- 슬라이드 제목: 20-28pt, bold
- 섹션 헤더: 14-16pt, bold
- 본문: 11-13pt
- 캡션: 9-10pt (최소)
- 줄간격: 1.2-1.5

### 색상
- 4-6색 팔레트를 콘텐츠 분위기에 맞게 선택
- 밝기 체크: `brightness_check(r, g, b)` 사용하여 텍스트 색상 결정
- 인접 요소에 같은 accent 색상 반복 금지

### 레이아웃
- 여백: 0.3"-0.5"
- 카드 간격: 카드 너비의 8-15%
- 콘텐츠가 컨테이너 가장자리에 닿지 않게
- 항상 같은 그리드 아닌, 콘텐츠에 맞는 레이아웃 선택
- 비대칭, 엇갈림, 흐름형 등 다양한 패턴 활용

### 시각적 인코딩 원칙
- **숫자/통계** → 차트. 3개+ 수치 → 반드시 차트 고려
- **우선순위/심각도** → 색상 채운 작은 원(`OVAL`). 이모지 금지
- **프로세스 흐름** → 플로차트 도형
- **아키텍처** → 의미 도형: DB=`CAN`, 클라우드=`CLOUD`, 서버=`CUBE`, 서비스=`GEAR_6`
- **파이프라인** → `CHEVRON` 도형 연결
- **계층/단계** → 크기와 위치로 중요도 표현

### 다이어그램 패턴
- **서비스 토폴로지**: 계층 배치 + `MSO_CONNECTOR.ELBOW` 커넥터
- **데이터 흐름**: 좌→우 `CHEVRON` 파이프라인
- **의사결정 트리**: `FLOWCHART_DECISION` 분기점
- **그룹**: `add_group_shape()`로 관련 컴포넌트 묶기

## 언어
한국어로 작성하면 한국어로 응답할 것.
