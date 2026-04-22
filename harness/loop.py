"""Loop — 전체 오케스트레이션.

generate-ppt.md v2 워크플로를 프로그래밍 방식으로 표현.
LLM 호출은 외부(Claude) 에서 수행하므로 이 모듈은:
- 파일 기반 상태 전이
- Validator 결정적 실행
- 상한 제어 (MAX_REFINE_ATTEMPTS)
- trace 저장

CLI:
    python -m harness.loop <project_name> [--from-step planner|generator|...]

사용 시나리오:
    1. Claude가 analysis.yaml 작성
    2. Claude가 Planner 역할 수행 → plan.yaml + NN.spec.yaml 저장
    3. Claude가 Generator 역할 수행 → NN.code.py 저장
    4. 이 모듈의 run_validation() 호출 → NN.validation.json
    5. Claude가 Evaluator 역할 (필요시) → NN.evaluation.json
    6. 결과 기반 Refiner 필요성 판단
    7. 최종 Assembler 실행 → output/{name}.pptx
"""

from __future__ import annotations

import json
import sys
from dataclasses import dataclass
from pathlib import Path

_HARNESS_DIR = Path(__file__).parent
_PROJECT_ROOT = _HARNESS_DIR.parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

from harness.schemas import ProjectPaths, SlideSpec, EvaluationReport, DeckPlan
from harness.validator import validate_slide_code, validate_full_deck, ValidationReport
from harness.evaluator import should_skip_eval, DEFAULT_THRESHOLD
from harness.assembler import assemble
from harness.refiner import MAX_REFINE_ATTEMPTS


@dataclass
class SlideStatus:
    idx: int
    attempts: int
    validation_passed: bool
    evaluation_passed: bool
    needs_refine: bool
    flagged: bool
    failure_summary: str = ""


def run_validation(paths: ProjectPaths, idx: int) -> ValidationReport:
    """단일 슬라이드 Validator 실행.

    Reads NN.code.py + NN.spec.yaml, writes NN.validation.json.
    """
    report = validate_slide_code(paths.code_path(idx), spec_path=paths.spec_path(idx))
    paths.validation_path(idx).parent.mkdir(parents=True, exist_ok=True)
    paths.validation_path(idx).write_text(report.to_json(), encoding="utf-8")
    return report


def decide_slide_status(
    paths: ProjectPaths,
    idx: int,
    attempts: int,
    eval_threshold: int = DEFAULT_THRESHOLD,
) -> SlideStatus:
    """현재 상태 판정: 통과 / 재시도 / human flag."""
    # 1. validation
    val_path = paths.validation_path(idx)
    if not val_path.exists():
        return SlideStatus(idx, attempts, False, False, True, False,
                           "no validation.json yet")

    val_data = json.loads(val_path.read_text(encoding="utf-8"))
    v_ok = val_data.get("passed", False)

    # 2. evaluation (optional)
    spec = SlideSpec.load(paths.spec_path(idx))
    skip_eval = should_skip_eval(spec.pattern)

    e_ok = True
    if not skip_eval:
        eval_path = paths.evaluation_path(idx)
        if eval_path.exists():
            e_report = EvaluationReport.load(eval_path)
            e_ok = e_report.score >= eval_threshold
        else:
            # 평가 결과 없음 — 아직 평가 안 됨
            e_ok = False

    passed = v_ok and e_ok

    # 상한 검사
    if not passed and attempts >= MAX_REFINE_ATTEMPTS:
        return SlideStatus(idx, attempts, v_ok, e_ok, False, True,
                           "max attempts reached")

    return SlideStatus(
        idx=idx, attempts=attempts,
        validation_passed=v_ok, evaluation_passed=e_ok,
        needs_refine=not passed, flagged=False,
        failure_summary="" if passed else "needs refine",
    )


def save_trace(paths: ProjectPaths, idx: int, attempt: int, state: dict):
    """슬라이드별 attempt 단위 trace 저장 — 향후 튜닝용."""
    paths.traces_dir.mkdir(parents=True, exist_ok=True)
    paths.trace_path(idx, attempt).write_text(
        json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def run_assembly(paths: ProjectPaths) -> dict:
    """모든 검증 통과 NN.code.py 를 순서대로 조립해 output/{name}.pptx 저장."""
    plan = DeckPlan.load(paths.plan_path)
    code_paths = [paths.code_path(idx) for idx in plan.slides]

    # 존재 확인
    missing = [p for p in code_paths if not p.exists()]
    if missing:
        raise FileNotFoundError(f"Missing code files: {missing}")

    return assemble(code_paths, paths.output_path, use_guard=True, audit=True)


def summary_report(paths: ProjectPaths) -> dict:
    """전체 진행 상태 요약 — workflow 중간 점검용."""
    plan = DeckPlan.load(paths.plan_path) if paths.plan_path.exists() else None
    total = plan.total_slides if plan else 0
    slides = []
    for idx in (plan.slides if plan else []):
        s = {
            "idx": idx,
            "has_spec": paths.spec_path(idx).exists(),
            "has_code": paths.code_path(idx).exists(),
            "has_validation": paths.validation_path(idx).exists(),
            "has_evaluation": paths.evaluation_path(idx).exists(),
        }
        if s["has_validation"]:
            v = json.loads(paths.validation_path(idx).read_text(encoding="utf-8"))
            s["validation_passed"] = v.get("passed")
        slides.append(s)
    return {"project": paths.name, "total": total, "slides": slides}


def _cli():
    import argparse
    parser = argparse.ArgumentParser(description="Harness loop orchestrator.")
    parser.add_argument("project", help="Project name (input/{project}/ must exist)")
    parser.add_argument("--root", default="/home/ubuntu/Share/ppt-generator")
    parser.add_argument("--action",
                        choices=["validate", "summary", "assemble"],
                        default="summary")
    parser.add_argument("--slide", type=int, help="slide idx (for validate)")
    args = parser.parse_args()

    paths = ProjectPaths(name=args.project, project_root=Path(args.root))

    if args.action == "summary":
        print(json.dumps(summary_report(paths), ensure_ascii=False, indent=2))
    elif args.action == "validate":
        if args.slide is None:
            # 전체
            plan = DeckPlan.load(paths.plan_path)
            for idx in plan.slides:
                if paths.code_path(idx).exists():
                    report = run_validation(paths, idx)
                    mark = "✓" if report.passed else "✗"
                    print(f"  {mark} slide {idx}: passed={report.passed}")
        else:
            report = run_validation(paths, args.slide)
            print(report.to_json())
    elif args.action == "assemble":
        result = run_assembly(paths)
        print(f"✓ {result['output']} (slides={result['slides']}, "
              f"audit_passed={result['audit_passed']})")


if __name__ == "__main__":
    _cli()
