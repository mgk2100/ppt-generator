# Reference Anchors — 목표 품질 기준 슬라이드

Evaluator(`harness/prompts/evaluator.md`)가 슬라이드를 채점할 때 "결함이 없는가"가
아니라 **"이 기준에 얼마나 가까운가"** 로 평가하기 위한 참조 PNG.

`plan.yaml`의 `style_guide.reference_anchor` 가 이 디렉토리의 PNG를 가리킨다.
local-search 함정(결함 없는 평범함이 자동 통과)을 깨는 장치.

## 파일

| 파일 | 설명 |
|------|------|
| `goal_slide.png` | 전면 격자표 + 셀별 색상코딩 + 카테고리 셀 병합의 고밀도 기준. `add_grid_table` 관용구의 목표 수준 |

## 교체 방법

덱 성격에 맞는 다른 기준이 있으면 PNG를 추가하고 `plan.yaml`에서 경로만 바꾼다:

```yaml
style_guide:
  reference_anchor: "ref/anchors/<your_anchor>.png"
```
