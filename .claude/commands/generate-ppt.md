대상: $ARGUMENTS
출력 경로: /home/user/Share/ppt-generator/output/

# Harness-Based PPT Generation (v2)

이 명령은 Anthropic 하네스 엔지니어링 원칙을 적용한 6-역할 파이프라인으로 PPT를 생성한다.

**격리의 실제 구현**: Generator / Evaluator / Refiner 는 **메인 세션이 아니라 `Agent` tool로 dispatch된 subagent**에서 실행된다. 각 subagent는 fresh context — 부모 대화 히스토리 없음. 명시적으로 prompt와 파일 경로만 받는다.

**고정** (절대 변경 금지):
- 첫 페이지 표지 레이아웃
- 페이지별 마스터 요소 (회사 로고, 페이지 번호, "SL Corporation Confidential", "연구개발본부", 상/하 divider)

**자유** (LLM 주도):
- CONTENT_SAFE 영역 내 모든 콘텐츠

## 역할별 실행 주체

| 역할 | 실행 주체 | 이유 |
|------|----------|------|
| Planner | 메인 Claude | Human gate (사용자 승인) 필요 |
| **Generator** | **Agent subagent** | Fresh context — 이전 슬라이드 코드 미노출 |
| Validator | Bash (deterministic) | LLM 아님 |
| **Evaluator** | **Agent subagent** | **코드 미노출 격리 — 자기 과대평가 편향 방지** |
| **Refiner** | **Agent subagent** | Fresh context — attempt 히스토리 배제 |
| Assembler | Bash (deterministic) | LLM 아님 |

## Step 0: 입력 유형 판단

대상 경로/텍스트를 확인해 유형 판단:
- A) 분석 결과 폴더 (.analysis-meta.json 등)
- B) 소스 코드 프로젝트
- C) 텍스트/문서 파일 (.md, .txt, .pdf, .html)
- D) 기업명 / 주제 등 텍스트

프로젝트 식별자 `{name}` 결정 (예: `ai-coding-exec-summary`).

사용자에게 보고:
"입력 유형: [A/B/C/D]로 판단했습니다. 프로젝트명: `{name}`. 진행할까요?"

## Step 1: 분석 및 인벤토리 수집

원본 자료를 읽고 `input/{name}/analysis.yaml` 생성 (v1과 동일 — 섹션·테이블·수치·다이어그램 인벤토리).

## Step 2: PLANNER (메인 Claude)

Planner는 사용자와 상호작용(Human gate)이 필요하므로 **메인 Claude가 직접 수행**.

### 2a. 덱 설계
- 총 슬라이드 수 결정 (표지 포함)
- 표지는 항상 `slide_01`, layout=`제목 슬라이드`
- 각 슬라이드 layout 결정:
  - `제목 슬라이드` — 표지만
  - `제목 및 내용 (페이지 번호 삭제)` — 섹션 구분 또는 5장 이하 콤팩트 덱
  - `제목 및 내용` — 기본

### 2b. plan.yaml 작성

`input/{name}/plan.yaml`:
```yaml
project_name: "{name}"
total_slides: N
sources: ["sources/{name}/..."]
narrative: "..."
slides: [1, 2, 3, ..., N]
style_guide:
  primary_accent: "#1F497D"
  secondary_accents: ["#4F81BD", "#9BBB59", "#C0504D", "#4BACC6"]
  # 레퍼런스 앵커 — Evaluator가 "결함 없는가"가 아니라 "이 수준에 가까운가"로 평가.
  # 기본값: 전면 격자표·색상코딩·셀 병합의 고밀도 기준 슬라이드. 덱 성격에 맞게 교체 가능.
  reference_anchor: "ref/anchors/goal_slide.png"
assets:                        # 시각 자산 인벤토리 (add_picture로 삽입할 캡처/차트/로고)
  - path: ""                   # 예: "sources/{name}/assets/ui_capture.png"
    use_on_slide: 0
```

### 2b-1. 자산 수집 (1급 단계)

분석 중 발견한 실제 UI 캡처·스크린샷·로고·외부 차트를 `assets/` 또는
`sources/{name}/assets/`에 모으고 위 `assets:` 인벤토리에 등록한다.
python-pptx 도형만으로 모든 것을 '그리지' 말고, 실물 자산은 `add_picture`로 삽입한다.

### 2c. slide_NN.spec.yaml 작성

각 슬라이드마다 `input/{name}/slides/slide_NN.spec.yaml`:
```yaml
idx: N
layout: "제목 및 내용"
title: "..."
pattern: "cards"    # cover / section_divider / cards / table / chart / chevron / diagram / code
accent_color: "#4F81BD"
key_message: "..."
content_blocks:
  - kind: card
    title: "..."
    body: "..."
data_refs: ["analysis.yaml > Section X"]
constraints:
  min_nontext_elements: 1
  # 시각 야심 — Evaluator의 visual_ambition/density 게이트 대비.
  # 정렬된 격자형 데이터는 add_grid_table, 수치 3+는 add_chart, 실물은 add_picture.
  # max_shapes 상한은 두지 않는다 (밀도 억제 금지). 빈 영역을 남기지 말 것.
  visual_primitive: "grid_table"   # grid_table | chart | picture | cards | diagram
```

### 2d. ▶ Human Gate

생성한 plan·specs을 사용자에게 요약해서 보여주고 승인 대기. 사용자 승인 없이 Step 3 진행 금지.

## Step 3: 슬라이드별 루프

### 3·0. 렌더 사전 점검 (필수 — blind 생성 방지)

루프 진입 전 LibreOffice 가용 여부를 확인한다. 없으면 시각 피드백 루프(3c/3d)가
작동하지 못하고 "눈 없이" 생성된다 — 반드시 먼저 설치한다.

```bash
Bash: which libreoffice soffice || echo "MISSING: sudo apt-get install -y libreoffice-impress 필요"
```

미설치 시 사용자에게 설치를 요청하고, 설치 완료까지 Evaluator 단계를 건너뛰지 말 것.

각 슬라이드 `N = 1..total_slides`에 대해:

### 3a. GENERATOR — **Agent subagent**

```
Agent(
  description="Generate slide NN code",
  subagent_type="general-purpose",
  prompt="""
당신은 PPT 슬라이드 Generator다. Fresh context — 이전 대화 없음.

## 읽어야 할 파일
1. /home/user/Share/ppt-generator/input/{name}/slides/slide_NN.spec.yaml
2. /home/user/Share/ppt-generator/harness/prompts/generator.md  (엄격한 제약)
3. /home/user/Share/ppt-generator/template_contract.py  (CONTENT_SAFE, LAYOUT 상수)
4. /home/user/Share/ppt-generator/ppt_utils.py  (사용 가능한 헬퍼 시그니처 — 읽고 참조)

## 작업
1. spec.yaml 읽기
2. generator.md 제약 준수하며 build_slide_NN(slide) 함수 작성:
   - import whitelist 안에서만
   - CONTENT_SAFE 영역 안에 shape 배치
   - set_title(slide, ...) + clear_placeholders(slide, keep=[0])
   - spec.content_blocks 각 항목 반영
3. Write tool 로 아래 경로에 저장:
   /home/user/Share/ppt-generator/input/{name}/slides/slide_NN.code.py
4. 저장 완료 후 "Saved: <경로>" 만 짧게 보고. 코드 요약·설명 불필요.
"""
)
```

### 3b. VALIDATOR — Bash (결정적)

```bash
Bash: python3 -m harness.loop {name} --action validate --slide N
```

내부적으로 `harness.validator.validate_slide_code()` 실행.
결과: `input/{name}/slides/slide_NN.validation.json`

판정: 메인 Claude가 JSON 읽고 `checks` 전부 true 인지 확인.

### 3c. PNG 렌더 (Validator 통과 && Evaluator 필요 시)

`spec.pattern ∈ {cover, section_divider}` 이면 Evaluator 생략 → 3d로 직행.

그 외:
```bash
Bash: python3 -c "
import sys; sys.path.insert(0, '/home/user/Share/ppt-generator')
from harness.render import render_single_slide
from harness.schemas import SlideSpec
from pathlib import Path
spec = SlideSpec.load(Path('/home/user/Share/ppt-generator/input/{name}/slides/slide_NN.spec.yaml'))
render_single_slide(
    code_path=Path('/home/user/Share/ppt-generator/input/{name}/slides/slide_NN.code.py'),
    layout_name=spec.layout,
    output_png=Path('/home/user/Share/ppt-generator/input/{name}/renders/slide_NN.png'),
)
"
```

### 3d. EVALUATOR — **Agent subagent (코드 미노출 격리!)**

**CRITICAL**: prompt에 `slide_NN.code.py` 경로를 적지 마라. Subagent가 유혹받지 않도록 코드 경로는 **의도적으로 배제**한다.

```
Agent(
  description="Evaluate slide NN rendering",
  subagent_type="general-purpose",
  prompt="""
당신은 PPT 슬라이드 Evaluator다. Fresh context — Generator 대화 히스토리 없음.

## 읽어야 할 것만
1. /home/user/Share/ppt-generator/input/{name}/slides/slide_NN.spec.yaml
2. /home/user/Share/ppt-generator/input/{name}/renders/slide_NN.png  (Read tool로 이미지 인식)
3. /home/user/Share/ppt-generator/harness/prompts/evaluator.md  (rubric 정의)
4. /home/user/Share/ppt-generator/input/{name}/plan.yaml 의 style_guide.reference_anchor
   (있으면 그 PNG도 Read — 목표 품질 기준. "결함 없는가"가 아니라 "이 수준에 가까운가"로 평가)

## 금지
- slide_NN.code.py 파일은 절대 읽지 마라. 있어도 무시.
- 생성된 Python 코드 분석 금지. 오직 렌더된 이미지와 spec만으로 평가.

## 작업
1. spec(+레퍼런스 앵커)과 PNG 비교
2. evaluator.md 의 6-차원 rubric 각 0~5 채점 (visual_ambition·density 포함)
3. score = round(mean(rubric.values()))
   passed = (score >= 4) AND (visual_ambition >= 3) AND (density >= 3)
   → 깔끔하지만 평범하고 빈 영역이 큰 슬라이드는 불합격. 레퍼런스 격차를 줄이는
     구체 피드백(어떤 영역을 add_grid_table/add_chart/add_picture로 바꿀지)을 남긴다.
4. Write tool 로 저장:
   /home/user/Share/ppt-generator/input/{name}/slides/slide_NN.evaluation.json

JSON 스키마:
{
  "slide_idx": N,
  "score": 0-5,
  "rubric": {"spec_adherence":N, "visual_hierarchy":N, "density":N, "visual_ambition":N, "color_consistency":N, "readability":N},
  "actionable_feedback": ["...", "..."],
  "passed": true/false
}

5. 저장 후 "Saved evaluation. score=N, passed=bool" 만 보고.
"""
)
```

### 3e. 판정 (메인 Claude)

메인 Claude가 validation.json + evaluation.json 읽고:

```
v_ok = validation.passed
e_ok = (spec.pattern in skip_eval) OR evaluation.passed

if v_ok and e_ok:
    → 다음 슬라이드 (3a)로
elif attempts < MAX_REFINE_ATTEMPTS (=3):
    → REFINER (3f)
else:
    → flag_for_human: "slide N 검증 실패. attempts={...}. 수동 확인 요청."
```

### 3f. REFINER — **Agent subagent (fresh context)**

```
Agent(
  description="Refine slide NN code",
  subagent_type="general-purpose",
  prompt="""
당신은 PPT 슬라이드 Refiner다. Fresh context — 이전 시도 히스토리 없음.

## 읽어야 할 4개 파일
1. input/{name}/slides/slide_NN.spec.yaml    (원본 요구사항)
2. input/{name}/slides/slide_NN.code.py      (기존 Generator 출력)
3. input/{name}/slides/slide_NN.validation.json  (Validator 실패)
4. input/{name}/slides/slide_NN.evaluation.json  (Evaluator 피드백 — 있으면)
5. /home/user/Share/ppt-generator/harness/prompts/refiner.md  (원칙)

## 작업
1. 네 파일 읽고 실패 원인 파악
2. refiner.md 원칙 준수:
   - 핀포인트 수정 (전체 재작성 금지)
   - 제약 유지 (import whitelist, CONTENT_SAFE, 함수 시그니처)
3. Write tool 로 기존 code.py 덮어쓰기:
   input/{name}/slides/slide_NN.code.py
4. 저장 후 "Refined. changed: [간단 bullet]" 보고.
"""
)
```

Refine attempt trace 저장 (메인 Claude가 기록):
```bash
Bash: python3 -c "
import json
from pathlib import Path
trace = {'attempt': ATTEMPT, 'spec_path': '...', 'validation': {...}, 'evaluation': {...}}
p = Path('/home/user/Share/ppt-generator/input/{name}/traces/slide_NN_attempt_MM.json')
p.parent.mkdir(parents=True, exist_ok=True)
p.write_text(json.dumps(trace, ensure_ascii=False, indent=2))
"
```

Refine 후 3b(Validator)부터 다시.

## Step 4: ASSEMBLE — Bash (결정적)

모든 슬라이드가 검증 통과하면:

```bash
Bash: python3 -m harness.loop {name} --action assemble
```

내부 동작:
1. `plan.yaml`의 `slides` 순서대로 `code.py`들 import·실행
2. `TemplateGuard`로 감싸 마스터 무결성 감사
3. `output/{name}.pptx` 저장
4. `validate_full_deck` 실행 (마스터 XML 해시 + 전 슬라이드 safe_zone)

실패 시 어느 슬라이드가 원인인지 추적해 3a로 돌아가 재생성.

## Step 5: 최종 검증 체크리스트

- [ ] `output/{name}.pptx` 생성됨
- [ ] 슬라이드 수 == `plan.total_slides`
- [ ] 마스터 XML 해시가 `ref/표지.pptx`와 동일 (validate_full_deck 통과)
- [ ] 모든 custom shape bbox ⊂ CONTENT_SAFE
- [ ] 표지(slide 1) — setup_cover 사용, 배경 Picture 변경 없음
- [ ] 각 content 슬라이드 비텍스트 시각 요소 1+
- [ ] `input/{name}/traces/`에 refine attempt 기록

## Step 6: 최종 보고

사용자에게:
```
✓ output/{name}.pptx  (N 슬라이드)
  · Generator Agent 호출: N회
  · Evaluator Agent 호출: M회 (cover/section_divider 제외)
  · Refiner Agent 호출: K회
  · Audit: passed
  · Refine 필요 슬라이드: [...]
```

---

## 격리 검증 체크리스트

v2 구현이 진짜로 격리되었는지 확인:

- [ ] Generator prompt에 **다른 슬라이드의 code.py 경로 미포함**
- [ ] Evaluator prompt에 **자기 슬라이드의 code.py 경로조차 미포함** (spec + PNG만)
- [ ] Refiner prompt에 **이전 attempt의 trace 미포함** (현재 시점의 4개 파일만)
- [ ] 각 Agent 호출은 `subagent_type="general-purpose"` — main과 별도 컨텍스트
- [ ] Agent 리턴 메시지는 파일 저장 확인 정도만. 코드 덤프 금지.

## Notes

- **Planner는 메인 Claude** (사용자 interaction 필요). Generator/Evaluator/Refiner만 subagent.
- **Validator/Assembler는 Bash** (deterministic, LLM 불필요).
- **Agent 호출 비용**: 슬라이드당 평균 1.5~3회 (Gen 1 + 선택 Eval 1 + 선택 Refine 0~2). 25장 덱 기준 40~75회.
- **부분 재생성**: 특정 slide_NN.code.py만 바꾸고 `--action assemble` 재실행하면 결정적.
- **v1 호환**: 기존 `sources/*/generate.py`는 완성된 .pptx를 `validate_full_deck`으로 감사하면 v2 체계에 통합 가능.
