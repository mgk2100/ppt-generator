"""Evaluator — 렌더된 PNG + SlideSpec → EvaluationReport (score 0-5).

실제 채점은 Claude가 **isolated context** 에서 수행 (Generator 코드 미노출).
이 모듈은 I/O + rubric 정의만 담당.

Generator-Evaluator 격리 원칙 (Anthropic harness):
- Evaluator에게 spec + PNG만 주고 **code는 보여주지 않음**.
- Evaluator가 낮은 점수를 주면 Refiner가 code + feedback을 보고 수정.

rubric (각 0~5):
- spec_adherence: spec의 content_blocks를 얼마나 충실히 반영했나
- visual_hierarchy: 정보 계층이 명확한가 (제목 > 요약 > 본문)
- density: 콘텐츠 밀도가 적절한가 (빈 영역 없이 정보가 화면을 채우나)
- visual_ambition: 레퍼런스 수준의 시각적 야심이 있나 (격자표/차트/이미지 활용,
  '결함 없는 평범함'을 넘어서나) — 이 차원이 local-search 함정을 깬다
- color_consistency: 색상 팔레트 일관성 + accent 준수
- readability: 폰트 크기 11pt+ 본문, 대비 충분, 겹침 없음
"""

from __future__ import annotations

from pathlib import Path

from harness.schemas import EvaluationReport


RUBRIC_DIMENSIONS = [
    "spec_adherence",
    "visual_hierarchy",
    "density",
    "visual_ambition",   # 레퍼런스 대비 시각적 야심 — 평범함 자동통과 차단
    "color_consistency",
    "readability",
]

DEFAULT_THRESHOLD = 4   # 5점 만점 중 4 이상 pass
# 추가 게이트: '결함 없는 평범함'이 자동 통과하지 못하도록, 아래 두 조건도 요구.
#   passed = (score >= DEFAULT_THRESHOLD)
#            and (rubric["visual_ambition"] >= 3)
#            and (rubric["density"] >= 3)
MIN_AMBITION = 3
MIN_DENSITY = 3


EVALUATOR_PROMPT = """\
당신은 PPT 슬라이드 평가자다. **생성 코드는 보지 않고** 렌더된 이미지와 spec만 본다.

## 입력
- SlideSpec (YAML): 슬라이드가 담아야 할 내용
- PNG: 렌더된 결과
- 레퍼런스 앵커(있으면): 목표 품질 수준을 보여주는 참조 PNG 또는 기술서

## rubric (각 0~5)

| 차원 | 기준 |
|---|---|
| spec_adherence | spec.content_blocks 각 항목이 슬라이드에 반영됐나? 1:1 매칭 |
| visual_hierarchy | 제목 > Key Message > 본문 계층이 시각적으로 명확한가? |
| density | 정보가 화면을 채우나? **하단/측면에 큰 빈 영역이 있으면 ≤2** |
| visual_ambition | 레퍼런스 수준의 시각적 야심 — 격자표(add_grid_table)·차트·이미지를 활용했나? 텍스트 카드 나열에 그치면 ≤2. 레퍼런스와 동급이면 5 |
| color_consistency | spec.accent_color 사용 + 전체 팔레트 일관성 |
| readability | 본문 11pt+, 대비 충분, shape 겹침 없음 |

## 합격 규칙 (Goodhart 방지)

'결함이 없다'만으로는 통과시키지 마라. 다음을 **모두** 만족해야 passed=true:
- score(평균 반올림) >= 4
- visual_ambition >= 3
- density >= 3

즉 깔끔하지만 평범하고 빈 공간이 큰 슬라이드는 **불합격** 처리하고, 레퍼런스
격차를 줄이는 구체적 피드백(어떤 영역을 격자표/차트/이미지로 바꿀지)을 남긴다.

## 출력 (JSON)

```json
{
  "slide_idx": N,
  "score": AVG_INT,
  "rubric": {
    "spec_adherence": 4,
    "visual_hierarchy": 3,
    "density": 4,
    "visual_ambition": 3,
    "color_consistency": 5,
    "readability": 4
  },
  "actionable_feedback": [
    "위치(어디) · 무엇 · 어떻게 고칠지",
    "예: 하단 30% 공백 → 3개 카드를 add_grid_table 5×4 비교표로 전환해 밀도 확보"
  ],
  "passed": true
}
```

평가는 엄격히. 5점은 레퍼런스와 동급일 때만.
"""


def build_evaluation_stub(slide_idx: int) -> EvaluationReport:
    """LLM 없이 테스트용 통과 스텁."""
    return EvaluationReport(
        slide_idx=slide_idx,
        score=5,
        rubric={d: 5 for d in RUBRIC_DIMENSIONS},
        actionable_feedback=[],
        passed=True,
    )


def save_evaluation(report: EvaluationReport, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    report.save(path)
    return path


def load_evaluation(path: Path) -> EvaluationReport:
    return EvaluationReport.load(path)


def should_skip_eval(pattern: str) -> bool:
    """cover / section_divider 는 deterministic 이므로 평가 생략 가능."""
    return pattern in {"cover", "section_divider"}
