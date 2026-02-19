대상: $ARGUMENTS
출력 경로: /home/ubuntu/Share/ppt-generator/output/

## Step 0: 입력 유형 판단

대상 경로/내용을 확인하여 유형을 판단한다:

A) 분석 결과 폴더 (아래 파일 중 하나 이상 존재)
   - phase0-raw-facts.md
   - phase1-architecture.md
   - phase2-data-flow.md
   - phase3-tech-stack.md
   - phase4-maintenance-guide.md
   - project-overview.md
   → 존재하는 파일들을 모두 읽어서 바로 Step 2로
B) 소스 코드 프로젝트 (*.py, package.json 등 존재)
   → 코드를 직접 분석하여 Step 1 수행
C) 텍스트/문서 파일 (.md, .txt, .pdf 등)
   → 문서 내용을 분석하여 Step 1 수행
D) 대상이 경로가 아닌 텍스트 (기업명, 주제 등)
   → 주어진 정보와 웹 검색으로 Step 1 수행

대상에서 프로젝트 식별자 `{name}`을 결정한다 (예: `sl-coding-assistant`, `company-intro`).
판단 결과를 사용자에게 안내한다:
"입력 유형: [A/B/C/D]로 판단했습니다. 프로젝트명: `{name}`. 진행할까요?"

## Step 1: 데이터 수집 및 분석
(유형 A는 이 단계 생략)

대상에서 PPT에 담을 데이터를 수집한다.
- fc-list로 시스템 폰트 목록을 확인하고, 주제에 어울리는 폰트를 선택한다
- 실제로 존재하는 데이터 항목만 나열 (없는 건 절대 포함 금지)
- 각 항목의 데이터 성격 판단:
  - 특징/가치 나열 → cards
  - 숫자/비교 → table
  - 순차 흐름 → flowchart
  - 계층 구조 → architecture
  - 디렉토리 → tree
  - 대비 → comparison
  - 설명 → content_boxed
  - 비율/비중 수치 → chart (pie/doughnut)
  - 비교 수치 → chart (bar/column)
  - 추세 수치 → chart (line)
  - 서비스 연결 관계 → architecture_diagram (의미 도형 사용)
  - 파이프라인/순차 단계 → pipeline (chevron)
  - 의사결정 분기 → decision_flow (flowchart shapes)
- 각 항목의 풍부도 (sparse/moderate/rich)
- 각 항목의 실제 개수 기록

저장: `input/{name}_analysis.yaml`

## Step 2: 슬라이드 설계

analysis.yaml (또는 기존 분석 파일들)을 읽고 **시각적 설계**를 한다.

각 슬라이드별로 계획:
- 시각적 컨셉 (레이아웃 패턴, 시각적 비유)
- 색상 팔레트 (이 프레젠테이션 전용 RGB 값)
- 공간 배치 (대략적 비율)
- 타이포그래피 선택
- 인접 슬라이드와 시각적으로 구분되는 요소
- 사용할 MSO_SHAPE 도형 종류 (의미 도형: CAN, CLOUD, CUBE, GEAR_6 등)
- 차트 슬라이드: 차트 타입(DOUGHNUT/COLUMN_CLUSTERED/LINE 등) + 데이터→시리즈 매핑
- 다이어그램 슬라이드: 노드 도형 종류 + 연결 토폴로지 (ELBOW/CURVE/STRAIGHT)
- 그림자/그라디언트 적용 대상 요소

설계 규칙:
- CLAUDE.md "적응형 슬라이드 구조"에 따라 전체 구조 패턴(섹션 분할형/내러티브형/컴팩트형/본문+부록형)을 먼저 결정
- 목차/섹션 구분 슬라이드의 필요성을 데이터 그룹 수와 내러티브 성격으로 판단 (불필요하면 생략)
- 카드 수 = 실제 데이터 항목 수
- 풍부한 데이터(5개+)는 별도 슬라이드로 분리
- 데이터 없는 슬라이드는 설계에 포함 금지

색상: 섹션마다 다른 accent 색상 사용, 단일 색상 반복 금지

저장: `input/{name}_slide-plan.md` (Markdown — 설계는 창의적 문서)

**사용자에게 슬라이드 설계를 보여주고 확인을 받는다.** 승인 후 Step 3 진행.

## Step 3: python-pptx 스크립트 작성

slide-plan.md를 읽고 완전한 python-pptx 스크립트를 작성한다.

스크립트 요구사항:
1. ppt_utils 임포트 (ensure_fonts, load_template, clear_placeholders, add_shadow, set_shape_opacity, add_gradient_stop, make_icon_circle, brightness_check, add_textbox, add_para, set_body_anchor 등)
2. ref/표지.pptx 템플릿 사용
3. 각 슬라이드 섹션에 설계 의도 주석
4. 독립 실행 가능 (python script.py)
5. output/{name}.pptx에 저장
6. `add_shadow()` 사용 — 오프셋 사각형으로 가짜 그림자 금지
7. 그라디언트 적극 활용 (최소 표지/섹션 구분 슬라이드)
8. 수치 데이터 → `add_chart()` API 사용
9. 다이어그램에 의미 도형 사용 (CAN, CLOUD, CUBE, GEAR_6, CHEVRON 등)
10. 이모지를 시각 지표로 사용 금지 — 색상 도형/`make_icon_circle()` 사용
11. ppt_utils의 헬퍼 함수 활용 (add_textbox, add_para, set_body_anchor 등) — 직접 재구현 금지
12. `set_title(slide, "...")` 으로 제목 설정 — add_textbox()로 제목 금지
13. 콘텐츠는 CONTENT_SAFE 영역 안에 배치
14. 콘텐츠 슬라이드에서 전면 배경으로 마스터 요소 덮지 말 것

저장: /tmp/{name}_generate.py

## Step 4: 실행 및 검증

python /tmp/{name}_generate.py

검증:
- [ ] 에러 없이 실행되는가?
- [ ] 빈 슬라이드 없는가?
- [ ] 카드 수 = 실제 데이터 항목 수인가?
- [ ] 섹션별 시각적 변화가 있는가?
- [ ] 이모지가 시각 지표로 사용되지 않았는가?
- [ ] 수치 데이터에 차트가 활용되었는가?
- [ ] 다이어그램에 의미 도형이 사용되었는가?
- [ ] 가짜 그림자(오프셋 사각형) 없는가?
- [ ] 그라디언트가 적절히 활용되었는가?
- [ ] ppt_utils 헬퍼를 직접 재구현하지 않았는가?
- [ ] 제목이 set_title()로 플레이스홀더를 사용하는가?
- [ ] 콘텐츠가 안전 영역(0.68"~7.02") 안에 있는가?
- [ ] 콘텐츠 슬라이드에서 마스터 요소가 가려지지 않는가?

실패 시 스크립트 수정 후 재실행.
