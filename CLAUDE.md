# PPT Generator

## 목적
입력 자료를 분석하여 python-pptx 기반 PPT를 자동 생성한다.

## 작동 방식
Claude가 python-pptx 코드를 직접 작성 → 실행 → PPT 생성.
고정 렌더러 없음. 매번 콘텐츠에 맞는 코드를 새로 작성한다.

## 지원 입력
하나의 /generate-ppt 명령으로 다양한 입력을 자동 판단하여 처리한다.
- 기존 분석 결과 폴더 (obsidian-notes/03_학습노트/ 내 분석 결과)
- 소스 코드 프로젝트
- 텍스트/문서 파일 (.md, .txt, .pdf 등)
- 기업명, 주제 등 텍스트 입력

## Type A 입력 인터페이스 계약

obsidian-notes 분석 결과(Type A)를 PPT로 변환할 때의 입력 요구사항.
분석기 측 출력 스키마는 `obsidian-notes/_analyzer/schemas/phase-output.schema.md` 참조.

### 필수 파일 구조

| 파일 | PPT 활용 |
|------|---------|
| `.analysis-meta.json` | 표지 (`project_name` → 대주제, `last_analyzed_date` → 분석 기준일) |
| `프로젝트 종합 분석.md` | 도입 메트릭 카드 (핵심 수치), 마무리 슬라이드 (리스크/조치) |
| `Phase 0 - 전체 구조와 팩트.md` | 프로젝트 구조 슬라이드 (ASCII tree, 핵심 수치) |
| `Phase 1 - 아키텍처.md` | 아키텍처 다이어그램 슬라이드 (Mermaid graph → PPT 도형) |
| `Phase 2 - 데이터 흐름.md` | 데이터 흐름 슬라이드 (sequence → 플로차트), 상태 머신 |
| `Phase 3 - 기술 스택.md` | 기술 스택 슬라이드 (매핑 테이블, 의존 관계) |
| `Phase 4 - 운영 가이드.md` | 아키텍처 보충 정보 (PPT 직접 활용도 낮음) |
| `Phase 5 - 학습갭분석.md` | 코드 스니펫 슬라이드, 학습 포인트, 기술 갭 |
| `심화 - *.md` (선택) | 딥다이브 슬라이드 (존재 시에만) |

### Phase별 PPT 변환 필수 데이터

| Phase | 필수 데이터 요소 | PPT 변환 결과 | 누락 시 영향 |
|-------|-----------------|-------------|------------|
| 종합 | 핵심 수치 테이블, 리스크 테이블 | 도입 메트릭 카드, 마무리 슬라이드 | 도입/마무리 빈약 |
| 0 | ASCII tree, 핵심 수치 | 프로젝트 구조 슬라이드 | 개요 축약 |
| 1 | **Mermaid `graph` 1+**, 컴포넌트 역할 테이블 | 아키텍처 다이어그램 슬라이드 | 아키텍처가 텍스트 전용 |
| 2 | **Mermaid `sequenceDiagram` 2+**, 상태 전이 | 데이터 흐름 슬라이드, 상태 머신 | 흐름 슬라이드 생성 불가 |
| 3 | 기술 매핑 테이블 (10행+), 의존 관계 다이어그램 | 기술 스택 슬라이드 | 스택 축약 |
| 4 | (PPT 직접 활용도 낮음) | 아키텍처 보충 정보 | 영향 미미 |
| 5 | **코드 스니펫 3+**, 기술 갭 테이블, 학습 포인트 | 코드 슬라이드, 학습 포인트 | 코드 슬라이드 부족 |

### Mermaid PPT 호환 규칙

`phase-output.schema.md`의 렌더링 제약(참여자 ≤5, 노드 ≤20, 레이블 ≤40자 등)을 따르면 PPT 변환도 호환된다.
PPT 전용 추가 규칙:

- 노드 레이블에 기술명 원문 유지 → PPT 도형 텍스트로 직접 사용 (번역 금지)
- subgraph 이름 → PPT 배경 ROUNDED_RECT 좌상단 라벨 텍스트박스
- 화살표 라벨 → PPT 커넥터 옆 8-9pt 텍스트박스 (**30자 이내**)
- sequence 참여자 별칭 → PPT 도형 내부 텍스트 (한글 2-4자 또는 영문 약어)
- 변환 전략 상세: "Mermaid→PPT 변환 전략" 섹션 참조

### 코드 스니펫 PPT 호환 규칙

Phase 5 코드 레퍼런스가 PPT 슬라이드에 직접 매핑된다. 작성 시:

- **15줄 이하** — PPT 슬라이드당 코드 제한
- 언어 표시 필수 (` ```python `, ` ```yaml ` 등)
- 파일 경로 주석 포함 (`# src/tasks/worker.py`)
- 학습 포인트 텍스트 동반 필수 — 코드만 단독 배치 금지
- 행 길이 70자 이내 (`phase-output.schema.md` 제약과 동일)

### density_score 가중치

콘텐츠 밀도 스코어가 슬라이드 할당량을 결정한다. 시각 요소가 풍부한 Phase에 더 많은 슬라이드가 배정된다.

| 요소 | 가중치 | 비고 |
|------|--------|------|
| Mermaid 다이어그램 | 3.0 | 아키텍처/흐름 슬라이드 생성 |
| 상태 머신 | 3.0 | `add_state_machine()` 슬라이드 |
| 코드 스니펫 | 2.5 | 코드 블록 슬라이드 |
| 테이블 | 2.0 | 비교/매핑 슬라이드 |
| ASCII tree | 1.5 | 구조 슬라이드 |
| 불릿 항목 | 0.3 | 텍스트 보조 |

공식 및 할당 기준 상세: `generate-ppt.md` Step 1c 참조

### PPT 호환성 체크리스트

분석 완료 후 PPT 변환 품질을 위해 검증:

- [ ] Phase 1에 Mermaid `graph` 다이어그램이 1개 이상 존재하는가
- [ ] Phase 2에 Mermaid `sequenceDiagram`이 2개 이상 존재하는가
- [ ] Phase 5에 코드 스니펫이 3개 이상이며 각각 15줄 이하인가
- [ ] 모든 Mermaid 노드 레이블이 기술명 원문을 유지하는가
- [ ] 모든 화살표 라벨이 30자 이내인가
- [ ] 코드 스니펫마다 학습 포인트 텍스트가 동반되는가
- [ ] Phase 3 기술 매핑 테이블이 10행 이상인가
- [ ] 종합 분석에 핵심 수치 테이블과 리스크 테이블이 존재하는가

---

## 핵심 원칙
- 있는 데이터만 슬라이드로 만든다. 데이터 없으면 해당 슬라이드 생성 안 함.
- 카드 개수 = 실제 데이터 항목 수. 빈 카드 금지.
- 슬라이드 타입은 데이터 성격으로 결정한다.

---

## PPT 생성 설정

### 발표 정보
- 유형: 기술 세미나 (엔지니어링 딥다이브)
- 청중: 개발자/기술팀 (AI/ML 기본 이해도 있음)
- 목적: 프로젝트 개발 과정에서 학습한 기술 스택과 설계 패턴 공유
- 분량: 20-25장

### 템플릿 규칙
1. 회사 지정 PPT 템플릿을 반드시 사용하라
   - 템플릿 경로: /home/ubuntu/Share/ppt-generator/ref/
   - 표지 슬라이드: 템플릿의 표지 레이아웃을 유지하고 제목/발표자/날짜만 교체
   - 로고: 템플릿에 포함된 회사 로고 위치와 크기를 변경하지 말 것
   - 슬라이드 마스터: 템플릿의 마스터 레이아웃(폰트, 색상, 배경)을 그대로 따를 것
2. 새 슬라이드 추가 시 템플릿의 레이아웃 중 적절한 것을 선택하여 사용하라
3. 템플릿에 정의된 폰트, 색상 팔레트를 벗어나지 말 것
4. 마스터 텍스트 스타일의 글머리 기호는 무시하라. 글머리 기호를 따라하지 말 것

### 레이아웃 자유도
1. 고정 요소 (절대 변경 금지): 표지, 회사 로고, 마스터 레이아웃(폰트/색상/배경)
2. 자유 요소: 위 고정 요소를 제외한 모든 콘텐츠 배치, 도형, 텍스트 박스, 이미지 위치, 슬라이드 내부 구성은 자유롭게 생성 가능
3. 템플릿에 적합한 레이아웃이 없으면 빈 슬라이드(Blank)에 마스터 스타일을 적용하여 자유롭게 구성하라

### 코드 스니펫 규칙
1. 슬라이드당 코드는 최대 15줄 이내로 제한
2. 핵심 로직만 발췌하고 생략 부분은 `# ...` 으로 표시
3. 코드 폰트: 고정폭 폰트 (Consolas 또는 D2Coding), 14pt 이상
4. 코드 블록 배경은 어두운 색(#1E1E1E 등)으로 구분

### 코드-설명 연결 규칙
1. 코드 스니펫 슬라이드에는 반드시 핵심 포인트를 2-3줄로 함께 표시
2. 코드 내 주목할 라인은 하이라이트 색상 또는 화살표로 강조
3. 필요 시 코드의 입력/출력 예시를 함께 보여줄 것

### 다이어그램 규칙
1. 아키텍처 다이어그램은 PPT 도형(Shape)으로 직접 구성하라
2. 컴포넌트 간 화살표로 데이터 흐름 방향을 명시
3. 도형 색상은 템플릿 팔레트 내에서 역할별로 구분 (예: API=파란계열, DB=초록계열, 외부서비스=회색)
4. 도형 내 텍스트는 기술명만 간결하게 (예: "FastAPI", "PostgreSQL")

### 시각 요소 규칙
1. 다음 시각 요소를 콘텐츠에 맞게 자유롭게 활용하라:
   - 코드 스니펫: 핵심 구현 설명 시
   - 아키텍처 다이어그램: 시스템 구조 설명 시
   - 비교 테이블: 기술 선택지 비교, Before/After 시
   - 플로우차트: 처리 순서, 파이프라인 설명 시
   - 하이라이트 박스: 핵심 개념 정의, 용어 설명 시
2. 연속 3장 이상 같은 레이아웃이 반복되지 않도록 할 것
3. 텍스트만으로 구성된 슬라이드는 최소화하라

### 강조 전략

**Key Message Bar** — 제목 아래 첫 요소:
- CONTENT_SAFE 폭 80%+, ROUNDED_RECTANGLE, accent 배경 opacity 8-15%
- 좌측 세로 바 (accent 100%, 4px)
- 본문보다 1-2pt 크게 bold, 한 문장 요약

**본문 내 강조 (우선순위 순)**:
1. accent 색상 텍스트 (`add_rich_text()`)
2. bold 대비
3. 배경 하이라이트 ROUNDED_RECTANGLE
4. `make_icon_circle()` 배지

금지: 밑줄 강조, 전체 bold, 한 슬라이드 3개+ 강조 기법 혼용

### 비교 슬라이드 패턴

- **Before/After**: `calc_grid(1,2)`, 좌=회색/연빨강, 우=accent, 중앙 화살표
- **Pros/Cons**: 좌=녹색, 우=빨강, 상단에 아이콘
- **기술 선택**: `add_styled_table()` 또는 카드형 `calc_grid(1,N)`
- 공통: 비교축 명시, 대칭 구조, 우열은 색상 강도로 표현

### 언어 규칙
1. 슬라이드 제목, 설명: 한국어
2. 코드 스니펫, 기술 용어, 변수명: 영어 원문 유지
3. 기술 스택명은 번역하지 말 것 (예: "작업 큐" ✗ → "Celery" ✓)
4. 한국어로 작성하면 한국어로 응답할 것

### 구성 흐름
1. 프로젝트별로 "왜 이 기술을 선택했는가 → 어떻게 구현했는가 → 무엇을 배웠는가" 순서로 전개
2. 섹션 전환 시 구분 슬라이드(Section Divider)를 삽입하라
3. 각 슬라이드 상단 핵심 메시지는 앞뒤 슬라이드와 논리적으로 연결되어야 한다

### 슬라이드 구성
1. 슬라이드 총량: 20-25장
2. 구성 비율은 마크다운 내용 분량에 비례하여 자동 조절
3. 단, 도입과 정리는 각각 최소 2장 이상 포함할 것
4. 기본 구조:
   - 도입: 프로젝트 배경, 전체 아키텍처 개요
   - 본론: 기술 스택별 설명 + 설계 패턴 + 코드 스니펫
   - 정리: 기술적 교훈, 향후 학습 방향

### 다중 마크다운 처리
1. 마크다운 파일명 또는 내부 제목(H1) 기준으로 프로젝트를 식별
2. 프로젝트 간 순서: 마크다운 파일의 알파벳순 또는 파일 내 명시된 순서를 따르라
3. 프로젝트 간 공통 기술 스택은 도입부에서 한번만 설명하고 중복하지 말 것

### PPT 생성 규칙
1. 마크다운의 프로젝트 구조(tree), 아키텍처, 기술 스택, 핵심 코드 스니펫을 기반으로 슬라이드를 구성하라
2. 슬라이드 스타일:
   - 코드 스니펫은 [코드 스니펫 규칙], [코드-설명 연결 규칙]을 따를 것
   - 아키텍처는 [다이어그램 규칙]을 따를 것
   - 시각 요소는 [시각 요소 규칙]을 따를 것
   - 슬라이드 상단에 핵심 메시지 1줄
   - 정보 밀도 높되 시각적으로 정돈
3. 마크다운에 없는 내용을 임의로 추가하지 말 것
4. 경로 내 마크다운이 여러 개면 [다중 마크다운 처리] 규칙을 따를 것

### Mermaid→PPT 변환 전략

마크다운 내 Mermaid 다이어그램을 PPT 도형으로 변환할 때 아래 매핑을 따른다.

#### 변환 매핑 테이블

| Mermaid 타입 | PPT 변환 | 도형 | 레이아웃 |
|-------------|---------|------|---------|
| `graph TD` (subgraph 포함) | 계층 다이어그램 | subgraph→ROUNDED_RECT 배경, 노드→의미도형 | 상→하, calc_grid 레이어 분할 |
| `graph LR` | 수평 흐름 | 노드→의미도형, add_smart_connector(LR) | 좌→우, align_shapes(axis='h') |
| `graph LR` (상태 전이) | 상태 머신 | ROUNDED_RECT + add_state_machine() | 좌→우 또는 2행 |
| `sequenceDiagram` (5 이하) | 수평 흐름도 | ROUNDED_RECT + CHEVRON 파이프라인 | 좌→우 |
| `sequenceDiagram` (6+) | 축약 다이어그램 + 테이블 | 의미도형 + 단계 테이블 | 상하 분할 |
| `sequenceDiagram` (loop/alt) | 플로차트 | FLOWCHART_PROCESS + DECISION | VFlow 또는 calc_grid |

#### 변환 규칙
- participant → 도형 텍스트 그대로 사용
- subgraph → 배경 ROUNDED_RECTANGLE + 좌상단 라벨 텍스트박스
- 화살표 라벨 → 커넥터 옆 작은 텍스트박스 (8~9pt)
- 5개+ 메시지 sequence → CHEVRON 파이프라인 요약 + 상세 테이블 분리
- loop/alt → FLOWCHART_DECISION 분기
- Note → RECTANGULAR_CALLOUT
- Mermaid 메타데이터 추출: `parse_mermaid_metadata(mermaid_text)` 사용

### 생성 후 검증
PPT 생성 완료 후 다음을 스스로 검증하라:
1. 표지, 로고, 마스터 레이아웃이 원본 템플릿과 동일한가
2. 마크다운에 없는 내용이 추가되지 않았는가
3. 코드 스니펫이 15줄을 초과하지 않는가
4. 기술 용어가 한국어로 번역되지 않았는가
5. 연속 3장 이상 동일 레이아웃이 반복되지 않는가
6. 모든 다이어그램의 화살표 방향이 데이터 흐름과 일치하는가
7. 코드 스니펫 슬라이드에 핵심 포인트 설명이 포함되어 있는가
8. 텍스트만으로 구성된 슬라이드가 연속되지 않는가

---

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

- **목차 슬라이드**: 독립적인 그룹이 **4개 이상**이고 청중이 전체 맥을 먼저 파악해야 할 때만 생성. 3개 이하이거나 순차 내러티브이면 생략.
- **섹션 구분 슬라이드**: 콘텐츠 슬라이드가 **10장 이상**이고 주제가 명확히 전환될 때만 사용. 같은 맥락이 계속되면 생략. 섹션 수 = 실제 데이터 그룹 수.
- **오버헤드 비율 제한**: 비콘텐츠 슬라이드(표지+목차+섹션구분+마무리)가 전체의 **30% 이하**여야 한다. 초과 시 섹션 구분 제거 또는 내러티브형으로 전환.
- **소규모 프레젠테이션** (총 15장 이하): 섹션 구분 슬라이드 대신 콘텐츠 슬라이드 내 시각적 색상 전환으로 섹션을 구분한다. 제목 좌측 accent 색상 바, 슬라이드 배경 색조 변화 등 활용.

### 섹션 구분 슬라이드 스타일

`add_section_divider()` 사용 권장. 수동 생성 시 아래 규칙:

필수 구성:
1. `set_title()` — 플레이스홀더 1개만 (textbox 중복 금지)
2. 부제목 1줄 (12-14pt)
3. 시각 요소 1개+ (accent 수평 바, 키워드 pills, 미니 진행표시 중 택)

금지: 제목 2번 표시, 빈 콘텐츠 영역

### 슬라이드 병합/분할

- **병합**: 관련 데이터가 각각 항목 2-3개 이하이고 한 슬라이드에 시각적으로 수용 가능하면 병합
- **분할**: 단일 주제라도 항목이 7개+이거나 다이어그램이 복잡하면 여러 슬라이드로 분할
- **기준**: 슬라이드 하나에 시각적 요소 3-5개가 적정. 6개+이면 분할 고려

### 분할 슬라이드 연결 전략

단일 주제 분할 시 맥락 유지:
1. **제목 연속성**: 동일 접두어 ("데이터 흐름 — 입력" → "데이터 흐름 — 처리")
2. **미니맵 패턴**: 우상단 CHEVRON 파이프라인, 현재 위치 accent 강조
3. **요약 앵커**: 후속 슬라이드 최상단 1줄 이전 내용 요약 (9-10pt, opacity 60%)
4. **시각 연속성**: 동일 accent + 동일 레이아웃 유지

### 내러티브 흐름

콘텐츠 목적에 따라 슬라이드 순서를 결정한다:
- **기술 소개**: 왜 필요한가 → 전체 구조 → 핵심 기능 → 기술 상세 → 운영/현황
- **분석 보고서**: 개요/규모 → 아키텍처 → 데이터 흐름 → 기술 스택 → 리스크/이슈
- **딥다이브**: 문제 정의 → 전체 조감 → [메커니즘 A → B → C...] → 리스크 → 운영
- **제안/설득**: 문제 → 해결책 → 근거/데이터 → 기대효과 → 다음 단계

---

## 금지 사항
- **이모지를 시각 지표로 사용 금지** (🔴🟠🟡 등). 색상 도형(`OVAL` 또는 `make_icon_circle()`) 사용
- **가짜 그림자(오프셋 사각형) 금지**. `add_shadow()` 사용
- **모든 다이어그램에서 사각형만 사용 금지**. 의미 도형 활용 (CAN, CLOUD, CUBE, GEAR_6 등)
- **통계 데이터를 텍스트 숫자만으로 나열 금지**. 3개+ 수치 → 차트 활용
- **ppt_utils에 있는 함수를 직접 재구현 금지**. import하여 사용
- **"반드시 N섹션" 같은 고정 구조 강제 금지**. 전체 구조는 데이터 볼륨과 성격이 결정
- **콘텐츠 슬라이드에서 마스터 요소를 덮는 전면 배경 금지**. 로고, 구분선, 푸터가 가려짐 (표지 레이아웃은 예외)
- **제목에 `add_textbox()` 사용 금지**. `set_title(slide, "...")` 사용하여 TITLE 플레이스홀더 활용
- **안전 영역(y=0.68"~7.02") 밖에 콘텐츠 배치 금지**. `CONTENT_SAFE` 상수 참조
- **세로 배치 y좌표 하드코딩 금지**. 같은 열에 2개+ 요소 → `VFlow`로 순차 배치. 개별 요소 높이도 `estimate_container_height()`로 동적 계산
- **다른 높이의 도형 간 수평 커넥터 금지**. 같은 행 도형은 `align_shapes(*shapes, axis='h')`로 높이/center_y 통일 후 연결. 같은 열 도형은 `axis='v'`

---

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

## 슬라이드 마스터 템플릿

`ref/표지.pptx` 슬라이드 마스터의 구조를 따른다.

### 슬라이드 크기
10.833" × 7.500" (와이드스크린 변형)

### 마스터 요소 (자동 표시)

| 요소 | 위치 | 비고 |
|------|------|------|
| 로고 (이미지) | 좌상단, y ≈ 0.09" | 모든 슬라이드에 자동 표시 |
| 제목 구분선 | y ≈ 0.60", 슬라이드 폭 전체 | 제목 아래 수평선 |
| 푸터 구분선 | y ≈ 7.02", 슬라이드 폭 전체 | 푸터 위 수평선 |
| 페이지 번호 | 우하단 | "제목 및 내용" 레이아웃에서 자동 |
| 부서명 | 좌하단 | 푸터 영역 |
| 기밀 문구 | 우하단 | 푸터 영역 |

### 레이아웃별 구조

**"제목 슬라이드"** — 표지 전용
- 플레이스홀더: idx=0 (제목), idx=1 (부제목)
- 레이아웃 자체에 전면 배경 이미지 포함 → 마스터 요소를 덮는 것은 의도된 동작
- 사용 패턴:
```python
layout = get_layout(prs, "제목 슬라이드")
slide = prs.slides.add_slide(layout)
slide.placeholders[0].text = "프레젠테이션 제목"
slide.placeholders[1].text = "부제목"
```

**"제목 및 내용"** — 일반 콘텐츠 (페이지 번호 있음)
- 플레이스홀더: idx=0 (제목, y=0.13"~0.56"), idx=1 (본문)
- 마스터 요소(로고, 구분선, 푸터) 자동 표시
- 사용 패턴:
```python
layout = get_layout(prs, "제목 및 내용")
slide = prs.slides.add_slide(layout)
set_title(slide, "슬라이드 제목")
clear_placeholders(slide, keep=[0])
```

**"제목 및 내용 (페이지 번호 삭제)"** — 섹션 구분 등
- "제목 및 내용"과 동일하나 페이지 번호 없음
- 사용 패턴: 위와 동일, 레이아웃 이름만 변경

### 콘텐츠 안전 영역
마스터 요소(로고, 구분선, 푸터)를 침범하지 않는 영역:
- **x**: 0.28" ~ 10.56" (너비 10.28")
- **y**: 0.68" ~ 7.02" (높이 6.34")
- 코드에서 `CONTENT_SAFE` 상수 사용:
```python
from ppt_utils import CONTENT_SAFE
# CONTENT_SAFE.left, .top, .width, .height, .right, .bottom
```

### 테마 폰트
- 라틴: Trebuchet MS
- 동아시아(EA): 휴먼모음T
- `set_title()` 에서 font_name=None이면 테마 폰트가 자동 상속됨

### 표지 기본 포맷

`setup_cover()` 함수로 표준 표지를 생성한다:

```python
layout = get_layout(prs, "제목 슬라이드")
slide = prs.slides.add_slide(layout)
setup_cover(slide, "프레젠테이션 제목")

# 커스텀:
setup_cover(slide, "제목",
    purpose="보고",           # "의사결정" | "보고" | "정보공유"
    author="강민규 선임",
    department="미래융합설계센터 알고리즘개발팀",
    date="2026.02.19",
)
```

표지 구성:
- **우측 상단**: ☐ 의사결정  ☐ 보고  ☑ 정보공유 (기본: 정보공유)
- **중앙**: 대주제 (PH idx=0, 폰트 상속 — 48pt bold white 현대하모니 M)
- **중간**: 날짜 (PH idx=1 활용, 폰트 상속 — 24pt white 현대하모니 L)
- **하단 중앙**: 부서명 + 이름

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
    add_textbox, add_para, add_rich_text,
    add_bullet_list, set_body_anchor,
    set_title, setup_cover, add_section_divider,
    CONTENT_SAFE,
    estimate_text_size, set_text_inset, calc_grid,
    calc_connector, add_smart_connector,
    add_code_block, add_styled_table,
    add_state_machine,
    auto_shrink_text, add_routed_connector,
    VFlow, HFlow,
    OUTPUT_DIR,
)

ensure_fonts()
prs = load_template()

# 슬라이드 크기 변수
SLIDE_W = prs.slide_width
SLIDE_H = prs.slide_height

# 색상 팔레트 — 이 프레젠테이션 전용
PRIMARY = RGBColor(...)

# --- 슬라이드 1: 표지 ---
layout = get_layout(prs, "제목 슬라이드")
slide = prs.slides.add_slide(layout)
setup_cover(slide, "프레젠테이션 제목")

# --- 슬라이드 2: 콘텐츠 ---
layout = get_layout(prs, "제목 및 내용")
slide = prs.slides.add_slide(layout)
set_title(slide, "슬라이드 제목")
clear_placeholders(slide, keep=[0])
# (콘텐츠는 CONTENT_SAFE 영역 안에 배치)

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
- 콘텐츠 슬라이드 제목: `set_title(slide, "...")` 사용 — `add_textbox()`로 제목 금지
- 슬라이드 추가 후 `clear_placeholders(slide, keep=[0])` 으로 유령 텍스트 제거
- 콘텐츠 배치: `CONTENT_SAFE` 영역(y=0.68"~7.02") 안에 배치
- 마스터 요소(로고, 구분선, 푸터) 보호: 콘텐츠 슬라이드에서 전면 배경으로 덮지 말 것
- 텍스트 컨테이너 높이: `estimate_container_height(text, font_size, max_width=w)`로 동적 계산. 복합 카드 내부 요소 위치도 이전 요소의 실제 높이 기반으로 산출
  ```python
  # BAD — 높이 하드코딩
  add_textbox(slide, x, y, w, Inches(0.5), label, font_size=12)
  # GOOD — 동적 계산
  label_h = estimate_container_height(label, 12, max_width=w)
  add_textbox(slide, x, y, w, label_h, label, font_size=12)
  ```
- 같은 열에 2개+ 요소를 세로로 배치할 때 `VFlow`를 사용하여 y좌표를 관리할 것. 절대 y좌표 하드코딩 금지
  ```python
  # BAD — y좌표 하드코딩 → 요소 겹침 위험
  add_code_block(slide, x, Inches(1.2), w, Inches(2.3), code)
  add_table(slide, x, Inches(3.7), w, 6, 3, data)           # 겹침!
  add_textbox(slide, x, Inches(5.6), w, Inches(0.4), text)  # 겹침!

  # GOOD — VFlow가 y 관리
  flow = VFlow(x=x, w=w, y_start=Inches(1.2))
  rect = flow.reserve(Inches(2.3))
  add_code_block(slide, rect.left, rect.top, rect.width, rect.height, code)
  rect = flow.reserve(Inches(0.35) * 6)
  add_table(slide, rect.left, rect.top, rect.width, 6, 3, data)
  flow.textbox(slide, text, font_size=11)
  ```
- 커넥터 평행 보장: 수평 커넥터로 연결되는 도형은 같은 center_y, 수직 커넥터로 연결되는 도형은 같은 center_x를 가져야 한다. 도형 배치 후 커넥터 생성 전에 `align_shapes()` 호출
  ```python
  # BAD — 높이가 달라 대각선 커넥터
  a = add_shape_box(slide, x1, y, w, Inches(0.7), ...)
  b = add_shape_box(slide, x2, y, w, Inches(0.8), ...)
  add_arrow_connector(slide, a, b)  # 비스듬함
  # GOOD — 정렬 후 연결
  a = add_shape_box(slide, x1, y, w, Inches(0.7), ...)
  b = add_shape_box(slide, x2, y, w, Inches(0.8), ...)
  align_shapes(a, b, axis='h')     # 높이 통일 + center_y 정렬
  add_arrow_connector(slide, a, b)  # 완벽한 수평
  ```

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
| `add_rich_text(text_frame, segments, ...)` | 혼합 서식 단락 (accent 색상 키워드, bold 대비) |
| `add_bullet_list(slide, x, y, w, items, ...)` | 구조화된 글머리 리스트 (단일 text_frame) |
| `set_cell_anchor(cell, 'ctr')` | 테이블 셀 세로정렬 |
| `set_cell_fill(cell, color)` | 테이블 셀 배경색 안전 설정 (중복 fill 방지 + OOXML 순서 준수) |
| `set_body_anchor(shape, 'ctr')` | 도형 텍스트 세로정렬 |
| `add_arrowhead(connector)` | 커넥터에 화살표 머리 추가 |
| `add_shadow(shape, blur_pt, dist_pt, direction, opacity_pct, color)` | 도형에 그림자 추가 |
| `set_shape_opacity(shape, opacity_pct)` | 도형 채우기 투명도 |
| `add_gradient_stop(shape, position, r, g, b)` | 그라디언트 3번째+ stop 추가 |
| `make_icon_circle(slide, x, y, size, fill_color, text, font_size)` | 원형 아이콘/배지 |
| `brightness_check(r, g, b)` | 밝기 판단 (True=밝음→어두운 텍스트) |
| `set_title(slide, text, ...)` | TITLE 플레이스홀더에 텍스트 설정 (font_name, font_size, color, bold) |
| `setup_cover(slide, title, ...)` | 표지 표준 포맷 (purpose, author, department, date) |
| `add_section_divider(prs, section_title, ...)` | 섹션 구분 슬라이드 (제목 중복 방지) |
| `CONTENT_SAFE` | 콘텐츠 안전 영역 (.left, .top, .width, .height, .right, .bottom) |
| `estimate_text_size(text, font_size_pt, ...)` | 텍스트 크기 추정 (한글/라틴 혼합, 줄바꿈 고려) |
| `set_text_inset(shape, ...)` | 도형 텍스트 내부 여백 설정 (한글 친화적 기본값) |
| `calc_grid(rows, cols, ...)` | 그리드 셀 좌표 계산 (균등/비율 분할) |
| `calc_connector(shape_a, shape_b, ...)` | 두 도형 간 커넥터 좌표/cxn 인덱스 계산 |
| `add_smart_connector(slide, shape_a, shape_b, ...)` | 스마트 커넥터 생성 (방향/타입 자동) |
| `align_shapes(*shapes, axis)` | 같은 행/열 도형 정렬 (높이/너비 통일 + 중심 맞춤) |
| `estimate_container_height(text, font_size_pt, max_width, ...)` | 텍스트+여백 포함 컨테이너 높이 계산 |
| `VFlow(x, w, y_start, y_max, gap)` | 수직 요소 배치 트래커 — y좌표 자동 관리, 겹침 방지 |
| `HFlow(y, h, x_start, x_max, gap)` | 수평 요소 배치 트래커 — x좌표 자동 관리 |
| `parse_mermaid_metadata(mermaid_text)` | Mermaid 텍스트 구조 분석 → dict(type, participants, subgraphs, nodes, edges, has_loop, has_alt, suggested_strategy) |
| `add_code_block(slide, x, y, w, h, code_text, ...)` | 코드 스니펫 블록 (어두운 배경 + 고정폭 폰트 + 라인 하이라이트) |
| `add_styled_table(slide, x, y, w, rows, cols, data, ...)` | 헤더+zebra+border 스타일 테이블 |
| `add_state_machine(slide, states, transitions, ...)` | 상태 머신 다이어그램 (자동 그리드 + 커넥터 + 라벨) |
| `auto_shrink_text(shape)` | 도형 텍스트 자동 축소 (normAutofit 설정) |
| `add_routed_connector(slide, waypoints, ...)` | 다중 경유점 라우팅 커넥터 (freeform path, 중간 도형 우회) |
| `OUTPUT_DIR` | 출력 디렉토리 경로 (`output/`) |

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

#### 커넥터 스타일 가이드

| 용도 | 두께 | 색상 | 타입 |
|------|------|------|------|
| 주 데이터 흐름 | 1.5-2pt | PRIMARY/#444444 | ELBOW+화살표 |
| 보조 관계 | 1pt | #999999 | STRAIGHT+화살표 |
| 양방향 | 1pt | #666666 | STRAIGHT 화살표 없음 |
| 강조 흐름 | 2pt | accent | STRAIGHT+큰 화살표 |

다이어그램당 커넥터 스타일 최대 2종

### 레이아웃 계산 헬퍼

#### 텍스트 크기 추정
```python
size = estimate_text_size("한글 텍스트", font_size_pt=14)
# size.width, size.height (Emu)

# 줄바꿈 고려
size = estimate_text_size("긴 텍스트...", 12, max_width=Inches(3))
# → 3인치 안에서 줄바꿈된 높이 반환
```

#### 도형 텍스트 여백
```python
shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, ...)
set_text_inset(shape)  # 한글 친화적 기본값 (0.12"/0.06")
set_text_inset(shape, left=Inches(0.2), right=Inches(0.2))  # 커스텀
```

#### 그리드 레이아웃
```python
# CONTENT_SAFE를 2×3 균등 분할
grid = calc_grid(2, 3)
for cell in grid.flat:
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
        cell.left, cell.top, cell.width, cell.height)
    set_text_inset(shape)

# 비율 분할 + 커스텀 영역
grid = calc_grid(1, 3, col_widths=[1, 2, 1], gap=Inches(0.2))
sidebar = grid[0][0]   # 좁은 좌측
main = grid[0][1]      # 넓은 중앙
```

#### 스마트 커넥터
```python
# 방향/타입 자동 감지
add_smart_connector(slide, shape_a, shape_b)

# 명시적 방향 지정
add_smart_connector(slide, shape_a, shape_b, direction='TB')

# 화살표 없이
add_smart_connector(slide, shape_a, shape_b, arrow=False)

# 좌표만 계산 (커넥터 직접 생성 시)
pts = calc_connector(shape_a, shape_b, direction='LR')
# pts.begin_x, pts.begin_y, pts.end_x, pts.end_y, pts.begin_cxn_idx, pts.end_cxn_idx
```

#### 수직 흐름 (VFlow)
```python
# 슬라이드 우측 열에 코드+테이블+텍스트 순서 배치
flow = VFlow(x=Inches(5.5), w=Inches(5.0), y_start=Inches(1.2))
rect = flow.reserve(Inches(2.3))
add_code_block(slide, rect.left, rect.top, rect.width, rect.height, code)
rect = flow.reserve(Inches(0.35) * 5)
add_table(slide, rect.left, rect.top, rect.width, 5, 3, data)
flow.textbox(slide, "요약 텍스트", font_size=11, bold=True)
print(f"남은 공간: {flow.remaining}")  # 경계 확인

# 카드 내부 서브플로우 (복합 요소 내부 배치)
card_flow = VFlow(x=cx, w=cw, y_start=cy+Inches(0.1), y_max=cy+ch, gap=Inches(0.05))
card_flow.textbox(slide, "42", font_size=28, bold=True)
card_flow.textbox(slide, "라벨 텍스트", font_size=12)
```

#### 수평 흐름 (HFlow)
```python
# 같은 행에 키워드 pills 수평 배치
hflow = HFlow(y=Inches(2), h=Inches(0.4), x_start=Inches(1), x_max=Inches(10))
for kw in ["FastAPI", "Celery", "Redis"]:
    rect = hflow.reserve(Inches(1.5))
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
        rect.left, rect.top, rect.width, rect.height)
    # ...
print(f"남은 공간: {hflow.remaining}")
```

#### 도형 행/열 정렬
```python
# 같은 행 → 높이 통일 + center_y 정렬
a = add_shape_box(slide, Inches(0.3), y, w, Inches(0.7), ...)
b = add_shape_box(slide, Inches(2.6), y, w, Inches(0.8), ...)
align_shapes(a, b, axis='h')  # 둘 다 h=0.8", 같은 center_y

# 같은 열 → 너비 통일 + center_x 정렬
align_shapes(a, b, axis='v')

# 3개 이상도 가능
align_shapes(box_a, box_b, box_c, axis='h')
```

#### 텍스트 컨테이너 높이
```python
label_h = estimate_container_height("PG Tables\n+ Mongo Collections", 12, max_width=Inches(1.5))
add_textbox(slide, x, y, Inches(1.5), label_h, "PG Tables\n+ Mongo Collections", font_size=12)
```

### 기타 기능
- **그룹 도형**: `slide.shapes.add_group_shape()` — 관련 컴포넌트 묶기
- **프리폼**: `slide.shapes.build_freeform(x, y)` — 커스텀 도형
- **이미지**: `slide.shapes.add_picture(path, x, y, w, h)`
- **회전**: `shape.rotation = 45.0`

---

## 디자인 원칙

### 타이포그래피
- 폰트는 프레젠테이션 주제에 맞게 자유 선택
- 제목은 `set_title()` 사용 — 테마 폰트 상속 또는 직접 지정 모두 가능
- 한글 문자 폭 ≈ 라틴 1.5배 — 박스 너비 계산 시 반영

#### 타이포그래피 스케일

| 역할 | pt 범위 | 비고 |
|------|---------|------|
| 섹션 구분 제목 | 28-40pt | 슬라이드당 1개 |
| 소제목/카드 헤더 | 14-18pt | bold 권장 |
| 본문 텍스트 | 11-14pt | **최소 11pt** (투영 가독성) |
| 캡션/주석 | 9-10pt | 보조 정보만 |
| 다이어그램 라벨 | 8-10pt | 도형 내부 |
| 코드 | 11-14pt | 고정폭 |

- 금지: 10pt 미만 본문, 슬라이드당 5종+ 폰트 크기

#### 줄 간격·단락 간격

| 요소 | line_spacing | space_before | space_after |
|------|-------------|-------------|------------|
| 본문 단락 | Pt(font×1.4) | Pt(4) | Pt(4) |
| 카드 내부 | Pt(font×1.3) | Pt(2) | Pt(2) |
| 리스트 항목 | Pt(font×1.2) | Pt(3) | Pt(1) |
| 코드 블록 | Pt(font×1.5) | Pt(0) | Pt(0) |
| 소제목 | 기본 | Pt(8) | Pt(4) |

### 색상

#### 색상 팔레트 프레임워크
역할별 색상 정의를 필수화한다:
```python
PRIMARY   = RGBColor(...)   # 주제색
SECONDARY = RGBColor(...)   # 보조색
ACCENT_1  = RGBColor(...)   # 강조1 — 하이라이트, 수치
ACCENT_2  = RGBColor(...)   # 강조2 — 분류, 카테고리
NEUTRAL   = RGBColor(...)   # 배경, 보더, 비활성
DANGER    = RGBColor(...)   # 리스크, 에러 전용
```
- 섹션마다 PRIMARY/ACCENT 교대 사용
- DANGER는 리스크/에러 전용, 장식 금지
- `brightness_check(r, g, b)` 사용하여 텍스트 색상 결정
- 인접 요소에 같은 accent 색상 반복 금지

#### 텍스트 대비 요구사항
- 본문(11pt+): 배경 대비 4.5:1 이상
- 대형(18pt+ 또는 14pt+ bold): 3:1 이상
- 흰색 배경 → 텍스트 #595959 이하, **#999999+ 밝은 회색 본문 금지**
- 어두운 배경 → 텍스트 #C0C0C0 이상

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

### 콘텐츠 슬라이드 풍부함

마스터 요소를 보호하면서 시각적 풍부함을 유지하는 기법:

- **색상 영역 분할**: CONTENT_SAFE 내부를 ROUNDED_RECTANGLE로 분할하여 각 영역에 accent 색상 배경 (opacity 5-15%) 적용
- **카드 기반 디자인**: 정보를 카드에 담고 그림자 + 상단/좌측 accent 바 적용. 카드 배경은 WHITE 또는 연한 색상
- **시각적 앵커**: 각 슬라이드에 최소 1개 비텍스트 요소 (도형, 차트, 다이어그램, 아이콘) 포함
- **색상 전환**: 같은 섹션 내 슬라이드도 accent 색상을 단계적으로 변화시켜 단조로움 방지
- **데이터 밀도**: 슬라이드당 2-3개 시각 유형 혼합 (수치 카드 + 차트, 텍스트 + 다이어그램 등)

### 카드 컴포넌트 표준

```
┌──────────────────────┐
│ ■ accent 바 (좌 4px 또는 상 4px)
│ 헤더 (14-16pt bold)
│ ── 구분선 (선택) ──
│ 본문 (11-12pt)
└──────────────────────┘ + shadow
```
- 배경: WHITE/#F5F5F5 ROUNDED_RECTANGLE
- 그림자: `add_shadow(blur_pt=4, dist_pt=2, opacity_pct=25)`
- 내부: VFlow로 순서 관리
- 정렬: `calc_grid()`로 균등 분할

## WORKLOG 규칙
- LLM은 WORKLOG.md를 직접 수정하지 않는다
- SessionEnd hook만 WORKLOG.md에 기록한다
- 수동 엔트리, LLM 해석, 토픽 추출 모두 금지
