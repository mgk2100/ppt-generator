# Harness-Based PPT Generation

Anthropic 하네스 엔지니어링 원칙을 적용한 6-역할 PPT 생성 파이프라인.

## 목적

- **고정**: 첫 페이지 표지 + 페이지별 마스터 요소 (로고, 페이지 번호, "SL Corporation Confidential", "연구개발본부", 상/하 divider)
- **자유**: CONTENT_SAFE 영역 내 모든 콘텐츠는 LLM이 자유롭게 생성
- **안전**: 런타임 검증으로 마스터 오염·safe_zone 위반 차단

## 5-Layer 아키텍처

```
L5 Human-in-loop     : plan 승인 게이트, refiner 상한 시 flag
L4 Command / CLI     : .claude/commands/generate-ppt.md (v2)
L3 Artifacts         : input/{name}/slides/*.{spec.yaml,code.py,validation.json,evaluation.json}
L2 Harness Roles     : Planner · Generator · Validator · Evaluator · Refiner · Assembler
L1 Template Contract : template_contract.py + ref/locked_registry.yaml + ref/표지.pptx
```

## 6개 역할

| 역할 | 실행 주체 | 방식 | 입력 | 출력 |
|------|---------|------|------|------|
| Planner | 메인 Claude | LLM + Human gate | analysis.yaml | plan.yaml + NN.spec.yaml |
| **Generator** | **Agent subagent** | LLM (fresh ctx) | NN.spec.yaml + 제약 문서 | NN.code.py |
| Validator | Bash | Deterministic | NN.code.py | NN.validation.json |
| **Evaluator** | **Agent subagent** | LLM (isolated, **코드 미노출**) | NN.spec.yaml + PNG | NN.evaluation.json |
| **Refiner** | **Agent subagent** | LLM (fresh ctx) | code + v + e | 수정된 code |
| Assembler | Bash | Deterministic | 통과 code 전체 | output/{name}.pptx |

### 격리는 어떻게 실제로 구현되나

Generator / Evaluator / Refiner 는 메인 세션이 아니라 Claude Code의 **`Agent` tool로 dispatch 된 subagent**에서 실행된다:

```
Agent(
  subagent_type="general-purpose",
  prompt="...역할 지시 + 읽을 파일 경로...",
)
```

각 subagent는:
- **Fresh context** — 부모 대화 히스토리 없음
- **명시적 파일만** — prompt에 적힌 경로만 Read
- **격리된 도구 실행** — 자기 ToolUse는 부모와 별개

특히 Evaluator 호출 시 prompt에 `slide_NN.code.py` 경로를 **의도적으로 포함하지 않는다** — subagent가 코드를 볼 수 없으므로 자기 과대평가 편향이 구조적으로 차단된다.

핵심 원칙:
- **Generator·Evaluator 격리** — Agent tool로 mechanism-level 분리
- **Context reset per slide** — 각 슬라이드가 새 subagent
- **File-based state handoffs** — subagent 간 통신은 디스크 파일로만

## 빠른 사용법

### 1) 슬래시 커맨드로 전체 자동 실행

```
/generate-ppt 대상경로_or_텍스트
```

Claude가 v2 워크플로 (`.claude/commands/generate-ppt.md`)를 따라:
1. 입력 유형 판단 → analysis.yaml
2. **Planner**: plan.yaml + spec.yaml (Human gate)
3. 슬라이드별 루프: Generator → Validator → Evaluator → Refiner (최대 3회)
4. Assembler: output/{name}.pptx

### 2) Validator만 단독 실행 (기존 PPT 감사)

```bash
python3 -m harness.validator output/some_deck.pptx
```

### 3) 수동 조립 (slide code들 준비되어 있을 때)

```bash
python3 -m harness.assembler \
    --slides input/myproj/slides/slide_01.code.py \
             input/myproj/slides/slide_02.code.py \
    --output output/myproj.pptx
```

### 4) 프로젝트 상태 확인

```bash
python3 -m harness.loop myproj --action summary
```

## 디렉토리 구조

```
ppt-generator/
├── template_contract.py          # L1: 불변 계약 + TemplateGuard
├── ref/
│   ├── 표지.pptx                  # 원본 템플릿 (read-only)
│   └── locked_registry.yaml      # 로크드 요소 레지스트리
├── harness/
│   ├── schemas.py                # 공용 dataclass (SlideSpec, DeckPlan, ...)
│   ├── planner.py                # I/O 헬퍼 (save_plan, save_specs)
│   ├── generator.py              # spec → code 프로토콜
│   ├── validator.py              # Deterministic 검증 (CLI)
│   ├── evaluator.py              # rubric 정의 + I/O
│   ├── refiner.py                # 패치 프로토콜
│   ├── assembler.py              # 최종 조립 (CLI)
│   ├── loop.py                   # 오케스트레이션 (CLI)
│   ├── render.py                 # PNG 렌더 (LibreOffice)
│   └── prompts/
│       ├── generator.md
│       ├── evaluator.md
│       └── refiner.md
├── input/{name}/
│   ├── analysis.yaml
│   ├── plan.yaml
│   ├── renders/slide_NN.png
│   ├── slides/
│   │   ├── slide_NN.spec.yaml
│   │   ├── slide_NN.code.py
│   │   ├── slide_NN.validation.json
│   │   └── slide_NN.evaluation.json
│   └── traces/slide_NN_attempt_MM.json
├── output/
│   └── {name}.pptx
└── .claude/commands/
    ├── generate-ppt.md           # v2 워크플로
    └── generate-ppt-v1.md        # 이전 워크플로 보관
```

## Template Contract (핵심)

`template_contract.py` 에서 import 해서 사용:

```python
from template_contract import (
    CONTENT_SAFE,         # SafeZone(left, top, width, height)
    LAYOUT_COVER,         # "제목 슬라이드"
    LAYOUT_CONTENT,       # "제목 및 내용"
    LAYOUT_NO_PAGENUM,    # "제목 및 내용 (페이지 번호 삭제)"
    TemplateGuard,        # 컨텍스트 매니저
    is_shape_in_safe_zone,
    master_xml_hash,
)
```

### TemplateGuard 사용 예

```python
from template_contract import TemplateGuard, LAYOUT_CONTENT, load_guarded_presentation

prs, guard = load_guarded_presentation()
with guard:
    guard.add_slide(LAYOUT_CONTENT, build_slide_1)
    guard.add_slide(LAYOUT_CONTENT, build_slide_2)
# __exit__ 시:
#   - 마스터 XML 해시 비교 → 변경 시 raise TemplateContractViolation
#   - 각 슬라이드 safe_zone 준수 확인
prs.save("output/deck.pptx")
```

## Backward Compatibility

기존 `sources/*/generate.py` 스크립트들은 변경 없이 계속 작동.
완성된 PPTX만 `validate_full_deck` 로 감사하면 Validator 체계에 통합.

```bash
# 기존 스크립트로 생성 후 감사
python3 sources/ai-coding-exec-summary/generate.py
python3 -m harness.validator output/ai-coding-exec-summary.pptx
```

## 적용된 하네스 원칙 매핑

| 원칙 (Anthropic 아티클) | 구현 |
|---|---|
| Task decomposition | 6개 역할 분리 |
| Specialized roles | planner/generator/evaluator/refiner/assembler/loop 모듈 |
| Generation ↔ Evaluation 격리 | **Agent tool subagent dispatch** + Evaluator prompt에 코드 경로 미포함 |
| File-based state handoffs | input/{name}/slides/*.yaml/*.py/*.json |
| Concrete test criteria | validation.json의 기계 검사 가능 체크 dict |
| Context reset per unit | 각 슬라이드마다 **새 Agent subagent** — mechanism-level fresh context |
| Generator-Evaluator loop | Validator + Evaluator → Refiner subagent, MAX_REFINE_ATTEMPTS=3 |
| Empirical validation | traces/ 디렉토리에 attempt 로그 |
| Model-adaptive | 간단 패턴(cover/section_divider)은 Evaluator skip |

## 참고

- 상세 설계: `/home/ubuntu/.claude/plans/ppt-kind-valiant.md`
- v1 워크플로: `.claude/commands/generate-ppt-v1.md`
