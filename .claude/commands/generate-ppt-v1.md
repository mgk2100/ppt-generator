대상: $ARGUMENTS
출력 경로: /home/ubuntu/Share/ppt-generator/output/

## Step 0: 입력 유형 판단

대상 경로/내용을 확인하여 유형을 판단한다:

A) 분석 결과 폴더 — 3단계 폴백 감지:
   → 입력 스키마: CLAUDE.md "Type A 입력 인터페이스 계약" 참조
   1. `.analysis-meta.json` 존재 확인 (최우선)
      → `project_name` 필드로 `{name}` 결정
      → `last_analyzed_date`를 PPT 표지 "분석 기준일"로 사용
   2. 한글 패턴 매칭 (폴백):
      - `프로젝트 종합 분석.md`
      - `Phase 0 - *.md` ~ `Phase 5 - *.md`
      - `심화 - *.md`
   3. 영문 레거시 (최종 폴백):
      - `phase0-raw-facts.md` ~ `phase4-maintenance-guide.md`
      - `project-overview.md`
   → 감지된 파일들을 모두 읽어서 바로 Step 1로 (content_inventory 생성)
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

대상에서 PPT에 담을 데이터를 수집한다.

### 1a. 기본 수집
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

### 1b. 콘텐츠 인벤토리 (유형 A 필수, B/C 권장)

소스 파일별 구조화된 콘텐츠를 카탈로깅한다. analysis.yaml에 다음 블록을 포함:

```yaml
project:
  name: "{name}"
  summary: "프로젝트 한줄 설명"
  analyzed_date: "2026-03-05"  # .analysis-meta.json 또는 오늘 날짜

sources:  # 각 소스 파일 메타데이터
  - file: "Phase 0 - 전체 구조와 팩트.md"
    size_bytes: 16504
    headings: {H1: 1, H2: 10, H3: 5}
    tables: 8
    mermaid_diagrams: 0
    code_blocks: 2

content_inventory:  # 구조화된 콘텐츠 카탈로그
  mermaid_diagrams:
    - id: "p1_arch_overview"
      source: "Phase 1 - 아키텍처.md"
      type: "graph_TD"           # graph_TD, graph_LR, sequenceDiagram, stateDiagram 등
      nodes: 13
      subgraphs: 4
      edges: 12
      ppt_strategy: "layer_diagram"  # CLAUDE.md Mermaid 변환 매핑 참조

  code_snippets:
    - id: "celery_routing"
      source: "Phase 5 - 학습갭분석.md"
      language: "python"
      lines: 5
      topic: "Celery 태스크 라우팅"

  tables:
    - id: "tech_overview"
      source: "Phase 3 - 기술 스택.md"
      rows: 16
      columns: 4

  state_machines:
    - id: "document_status"
      source: "Phase 2 - 데이터 흐름.md"
      states: 9
      transitions: 8

cross_references:  # Phase 간 연결 관계 (있을 때만)
  - from: "Phase 0 > DB 스키마"
    to: "Phase 2 > 데이터 저장소 관계"
    relationship: "schema_implements_storage"
```

Mermaid 다이어그램은 `parse_mermaid_metadata()` 함수로 자동 분석:
```python
from ppt_utils import parse_mermaid_metadata
meta = parse_mermaid_metadata(mermaid_text)
# → {type, participants, subgraphs, nodes, edges, has_loop, has_alt, suggested_strategy}
```

### 1c. 콘텐츠 밀도 스코어링 (H2 섹션 단위)

각 H2 섹션의 밀도를 계산하여 슬라이드 할당 기준으로 사용:

```
density_score = (
    tables * 2.0 +
    mermaid_diagrams * 3.0 +
    code_snippets * 2.5 +
    state_machines * 3.0 +
    ascii_trees * 1.5 +
    call_chains * 2.0 +
    bullet_items * 0.3 +
    text_paragraphs * 0.5
)
```

density_score → 슬라이드 할당:

| density_score | 할당 | 설명 |
|---------------|------|------|
| 0~2 | 인접 섹션과 병합 | 단독 슬라이드 불가 |
| 2~5 | 1장 | 보통 콘텐츠 |
| 5~10 | 1~2장 | 풍부한 콘텐츠 |
| 10~20 | 2~3장 | 매우 풍부, 분할 필요 |
| 20+ | 3장+ | 반드시 분할 |

각 섹션의 density_score를 analysis.yaml sections에 기록:
```yaml
sections:
  - name: "아키텍처 개요"
    type: "architecture"
    richness: "rich"
    density_score: 12.5
    source_ref: "Phase 1 > 아키텍처 다이어그램"
```

저장: `input/{name}_analysis.yaml`

## Step 2: 슬라이드 설계

analysis.yaml (또는 기존 분석 파일들)을 읽고 **시각적 설계**를 한다.

### 2a. 슬라이드 총량 할당

density_score 기반으로 슬라이드를 할당한다:

```
target_content = total_target - overhead(표지+목차+섹션구분+마무리)
각 Phase 기본 할당 = round(phase_bytes / total_bytes * target_content)
밀도 보정: density 상위 20% → +1장, 하위 20% → -1장(최소 1)
Mermaid 보정: 다이어그램 3개+ Phase → ceil(diagrams/2) 이상 확보
최종 클램핑: 20~25장 범위
```

### 2b. 슬라이드별 설계

각 슬라이드별로 계획:
- 시각적 컨셉 (레이아웃 패턴, 시각적 비유)
- 색상 팔레트 (이 프레젠테이션 전용 RGB 값)
- 공간 배치 (대략적 비율)
- 타이포그래피 선택
- 인접 슬라이드와 시각적으로 구분되는 요소
- 타이프 스케일 계획 (역할별 pt — 타이포그래피 스케일 참조)
- 카드 구조 (accent 바 위치, 내부 요소 — 카드 컴포넌트 표준 참조)
- 비교 패턴 선택 (Before/After, Pros/Cons, 기술 선택 — 비교 슬라이드 패턴 참조)
- 사용할 MSO_SHAPE 도형 종류 (의미 도형: CAN, CLOUD, CUBE, GEAR_6 등)
- 차트 슬라이드: 차트 타입(DOUGHNUT/COLUMN_CLUSTERED/LINE 등) + 데이터→시리즈 매핑
- 다이어그램 슬라이드: 노드 도형 종류 + 연결 토폴로지 (ELBOW/CURVE/STRAIGHT)
- Mermaid 변환: content_inventory의 ppt_strategy 참조 → CLAUDE.md "Mermaid→PPT 변환 전략" 매핑 적용
- 그림자/그라디언트 적용 대상 요소

설계 규칙:
- CLAUDE.md "적응형 슬라이드 구조"에 따라 전체 구조 패턴(섹션 분할형/내러티브형/컴팩트형/본문+부록형)을 먼저 결정
- 목차/섹션 구분 슬라이드의 필요성을 데이터 그룹 수와 내러티브 성격으로 판단 (불필요하면 생략)
- 카드 수 = 실제 데이터 항목 수
- 풍부한 데이터(5개+)는 별도 슬라이드로 분리
- 데이터 없는 슬라이드는 설계에 포함 금지
- 오버헤드 비율 체크: 비콘텐츠(표지+목차+섹션구분+마무리)가 전체의 30% 이하인지 확인
- 총 15장 이하 프레젠테이션에서 섹션 구분 슬라이드는 지양 → 색상 전환으로 대체

### 2c. 커버리지 매핑 테이블 (필수)

소스의 모든 H2/H3 섹션 및 구조화된 콘텐츠가 슬라이드에 매핑되는지 검증:

```
| 소스 파일 | H2 섹션 | 주요 요소 | 슬라이드 | 생략 사유 |
|-----------|---------|----------|---------|----------|
| Phase 0 | 디렉토리 구조 | ASCII tree | S4 | |
| Phase 1 | 아키텍처 개요 | Mermaid graph_TD | S5-S6 | |
| Phase 4 | 로컬 실행 | bash commands | 생략 | 운영 상세 |
```

### 2d. 시각 패턴 시퀀스 테이블 (필수)

슬라이드 설계 완료 후 시각적 다양성을 검증:

```
| S# | 시각 패턴 | accent 색상 | 비텍스트 요소 |
|----|---------|------------|-------------|
| S1 | cover | - | 배경 이미지 |
| S2 | cards | C_TEAL | 5 메트릭 카드 |
| S3 | pipeline+chart | C_BLUE | CHEVRON + DOUGHNUT |
```

검증 규칙:
1. 연속 3장 동일 시각 패턴 금지 → 위반 시 중간 슬라이드 패턴 변경
2. 인접 슬라이드 간 최소 2개 시각적 차이 (패턴/색상/도형/방향 중 2+)
3. 전체 비율: 다이어그램/플로차트 25%+, 코드 2장+, 테이블 3장 이하, 텍스트 전용 0장
4. content_inventory의 모든 mermaid_diagrams 매핑 필수 (생략 시 사유 명시)

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
10. 이모지는 배지/라벨/마커로 사용 가능 — CLAUDE.md "이모지 활용 가이드" 준수
11. ppt_utils의 헬퍼 함수 활용 (add_textbox, add_para, set_body_anchor 등) — 직접 재구현 금지
12. `set_title(slide, "...")` 으로 제목 설정 — add_textbox()로 제목 금지
13. 콘텐츠는 CONTENT_SAFE 영역 안에 배치
14. 콘텐츠 슬라이드에서 전면 배경으로 마스터 요소 덮지 말 것
15. 표지는 `setup_cover(slide, title)` 사용 — 수동 placeholder 텍스트 설정 금지

저장: /tmp/{name}_generate.py

## Step 4: 실행 및 검증

python /tmp/{name}_generate.py

검증:
- [ ] 에러 없이 실행되는가?
- [ ] 빈 슬라이드 없는가?
- [ ] 카드 수 = 실제 데이터 항목 수인가?
- [ ] 섹션별 시각적 변화가 있는가?
- [ ] 이모지 사용 시 슬라이드당 5종 이하이며 배지/라벨/마커 위치에만 있는가?
- [ ] 수치 데이터에 차트가 활용되었는가?
- [ ] 다이어그램에 의미 도형이 사용되었는가?
- [ ] 가짜 그림자(오프셋 사각형) 없는가?
- [ ] 그라디언트가 적절히 활용되었는가?
- [ ] ppt_utils 헬퍼를 직접 재구현하지 않았는가?
- [ ] 제목이 set_title()로 플레이스홀더를 사용하는가?
- [ ] 콘텐츠가 안전 영역(0.68"~7.02") 안에 있는가?
- [ ] 콘텐츠 슬라이드에서 마스터 요소가 가려지지 않는가?
- [ ] 표지가 setup_cover()로 생성되었는가? (체크박스, 날짜, 저자 포함)
- [ ] 오버헤드 비율이 30% 이하인가? (비콘텐츠 ÷ 전체)
- [ ] 각 콘텐츠 슬라이드에 최소 1개 비텍스트 시각 요소가 있는가?
- [ ] 소스 마크다운의 모든 주요 섹션이 슬라이드에 반영되었는가? (커버리지 매핑 테이블 대조)
- [ ] content_inventory의 모든 mermaid_diagrams가 PPT 도형으로 매핑되었는가?
- [ ] 슬라이드 총수가 20~25장 범위 내인가?
- [ ] density_score 상위 Phase에 충분한 슬라이드가 할당되었는가?
- [ ] 시각 패턴 시퀀스에서 연속 3장 동일 패턴이 없는가?
- [ ] 다이어그램/플로차트 슬라이드가 전체의 25% 이상인가?

### 가독성·디자인 품질 체크

- [ ] 본문 텍스트 모두 11pt 이상인가? (8-9pt는 캡션/라벨에만)
- [ ] 슬라이드당 폰트 크기 종류 4개 이하인가?
- [ ] 섹션 구분 슬라이드에 제목 중복 표시 없는가?
- [ ] 리스트 콘텐츠가 단일 text_frame multi-paragraph인가? (개별 textbox 남발 아닌가)
- [ ] 핵심 메시지에 최소 1개 강조 기법 적용되었는가?
- [ ] 테이블에 헤더 스타일(배경+bold) + zebra 적용되었는가?
- [ ] 흰색 배경 위 텍스트가 #999999보다 어두운가?
- [ ] 색상 팔레트가 역할별로 정의되어 있는가? (PRIMARY~DANGER)

실패 시 스크립트 수정 후 재실행.
