"""Planner — 원본 자료 + analysis.yaml → plan.yaml + NN.spec.yaml.

이 모듈은 **인터페이스 + I/O**만 담당. 실제 플래닝은 generate-ppt.md v2 워크플로에서
Claude가 수행하고, 결과를 이 헬퍼로 저장한다.

사용:
    from harness.planner import save_plan, save_specs
    save_plan(plan, paths)
    save_specs(specs, paths)
"""

from __future__ import annotations

from pathlib import Path

from harness.schemas import DeckPlan, SlideSpec, ProjectPaths


def save_plan(plan: DeckPlan, paths: ProjectPaths) -> Path:
    paths.ensure_dirs()
    plan.save(paths.plan_path)
    return paths.plan_path


def save_specs(specs: list[SlideSpec], paths: ProjectPaths) -> list[Path]:
    paths.ensure_dirs()
    out = []
    for spec in specs:
        p = paths.spec_path(spec.idx)
        spec.save(p)
        out.append(p)
    return out


def load_plan(paths: ProjectPaths) -> DeckPlan:
    return DeckPlan.load(paths.plan_path)


def load_specs(paths: ProjectPaths) -> list[SlideSpec]:
    specs = []
    for p in sorted(paths.slides_dir.glob("slide_*.spec.yaml")):
        specs.append(SlideSpec.load(p))
    return specs


# ============================ 플래닝 프롬프트 참고용 ============================

PLANNER_PROMPT_HINT = """\
당신은 PPT 플래너다. analysis.yaml을 읽고 다음을 출력하라:

1. plan.yaml:
   - project_name, total_slides, sources, narrative, slides(idx 순서)

2. 각 슬라이드별 slide_NN.spec.yaml:
   - idx, layout, title, pattern, accent_color, content_blocks, data_refs, constraints, key_message

제약:
- 허용 layout: "제목 슬라이드" (표지만) / "제목 및 내용" (기본) / "제목 및 내용 (페이지 번호 삭제)" (섹션 구분)
- 표지는 slide idx=1로 고정
- Content blocks는 kind별로: text / card / table / chart / chevron / diagram
- constraints.max_shapes 기본 30

출력은 ProjectPaths의 경로에 yaml 파일로 저장.
"""
