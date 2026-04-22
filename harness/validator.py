"""Validator — Layer 2 결정적(deterministic) 검증.

한 개의 `NN.code.py`에 대해 AST + 런타임 검사를 수행하고 `NN.validation.json`을 출력.

CLI 사용:
    python -m harness.validator path/to/slide_01.code.py [--spec path/to/slide_01.spec.yaml]

프로그래밍 사용:
    from harness.validator import validate_slide_code
    report = validate_slide_code(code_path, spec_path=None)
    if not report.passed:
        print(report.failures)
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any

import yaml
from pptx import Presentation

# 루트 경로를 sys.path에 추가 (template_contract / ppt_utils import)
_HARNESS_DIR = Path(__file__).parent
_PROJECT_ROOT = _HARNESS_DIR.parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

from template_contract import (  # noqa: E402
    REGISTRY,
    LAYOUT_COVER,
    LOCKED_PH_IDX,
    TemplateGuard,
    TemplateContractViolation,
    is_shape_in_safe_zone,
    slide_uses_allowed_layout,
    placeholder_idx_respected,
    master_xml_hash,
    layout_xml_hashes,
    only_imports_whitelisted,
    no_direct_master_edit,
)
from ppt_utils import load_template  # noqa: E402


# ============================ 데이터 모델 ============================

@dataclass
class ValidationReport:
    slide_idx: int | None
    code_path: str
    checks: dict[str, bool] = field(default_factory=dict)
    failures: list[dict[str, Any]] = field(default_factory=list)

    @property
    def passed(self) -> bool:
        return bool(self.checks) and all(self.checks.values())

    def to_json(self) -> str:
        d = asdict(self)
        d["passed"] = self.passed
        return json.dumps(d, ensure_ascii=False, indent=2)


# ============================ 내부 헬퍼 ============================

def _load_spec(spec_path: Path | None) -> dict:
    if spec_path is None or not spec_path.exists():
        return {}
    return yaml.safe_load(spec_path.read_text(encoding="utf-8")) or {}


def _import_code_module(code_path: Path):
    """NN.code.py를 모듈로 import. build_slide_N 함수를 반환."""
    mod_name = code_path.stem.replace(".", "_")
    spec = importlib.util.spec_from_file_location(mod_name, code_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Cannot load {code_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _find_build_function(module, slide_idx: int | None):
    """module에서 build_slide_N 또는 build_slide_<idx> 함수를 찾는다.

    엄격한 이름 고정이 부담스러우면 'build_slide' 또는 'main'도 허용.
    """
    candidates = []
    if slide_idx is not None:
        candidates.append(f"build_slide_{slide_idx}")
    candidates.extend(["build_slide", "main"])
    for name in candidates:
        fn = getattr(module, name, None)
        if callable(fn):
            return name, fn
    # 폴백: build_slide_로 시작하는 첫 함수
    for attr in dir(module):
        if attr.startswith("build_slide_"):
            fn = getattr(module, attr)
            if callable(fn):
                return attr, fn
    return None, None


# ============================ 검사 함수 ============================

def _check_imports(code_path: Path) -> tuple[bool, list[str]]:
    passed, viols = only_imports_whitelisted(code_path)
    return passed, viols


def _check_no_master_edit(code_path: Path) -> tuple[bool, list[str]]:
    return no_direct_master_edit(code_path)


def _check_function_signature(module, slide_idx: int | None) -> tuple[bool, str]:
    name, fn = _find_build_function(module, slide_idx)
    if fn is None:
        return False, "build_slide_N or build_slide function not found"
    return True, name


def _runtime_check(
    code_path: Path, slide_idx: int | None, spec: dict
) -> tuple[dict[str, bool], list[dict[str, Any]]]:
    """실제로 TemplateGuard 하에 슬라이드를 build한 후 결과 검사.

    Returns:
        (checks dict, failures list)
    """
    checks: dict[str, bool] = {}
    failures: list[dict[str, Any]] = []

    prs = load_template()
    before_master = master_xml_hash(prs)
    before_layouts = layout_xml_hashes(prs)

    try:
        module = _import_code_module(code_path)
    except Exception as e:
        checks["import_succeeds"] = False
        failures.append({"check": "import_succeeds", "error": repr(e)})
        return checks, failures
    checks["import_succeeds"] = True

    ok, name = _check_function_signature(module, slide_idx)
    checks["has_build_function"] = ok
    if not ok:
        failures.append({"check": "has_build_function", "error": name})
        return checks, failures

    fn = getattr(module, name)

    # spec에서 layout 결정 (없으면 CONTENT 기본)
    layout_name = spec.get("layout", "제목 및 내용")
    if layout_name not in REGISTRY.layouts:
        checks["layout_allowed"] = False
        failures.append(
            {"check": "layout_allowed", "error": f"unknown layout: {layout_name}"}
        )
        return checks, failures
    checks["layout_allowed"] = True

    # TemplateGuard로 빌드. build 함수가 (prs,)를 받거나 (slide,)를 받거나 둘 다 지원.
    guard = TemplateGuard(prs)

    def _wrapper(slide):
        # 사용자 build 함수가 prs를 받는 스타일일 수도 있으므로 heuristic 검사.
        import inspect
        sig = inspect.signature(fn)
        params = list(sig.parameters.keys())
        try:
            if len(params) == 1 and params[0] in ("slide",):
                fn(slide)
            else:
                # prs 스타일: 사용자가 자기 방식대로 add_slide 할 것이므로
                # TemplateGuard 통합이 어려움. 이 경우 별도 처리.
                raise RuntimeError(
                    "prs-style build function not supported in isolated validator. "
                    "Use build_slide_N(slide) signature."
                )
        except Exception as e:
            raise

    try:
        with guard:
            guard.add_slide(layout_name, _wrapper)
    except TemplateContractViolation as e:
        checks["master_and_safe_zone_intact"] = False
        failures.append({"check": "template_guard", "error": str(e)})
    except RuntimeError as e:
        # Fallback: prs 스타일 함수일 수 있음. 가드 밖에서 직접 호출해보고,
        # 호출 후 master hash만 비교하는 약한 검증.
        msg = str(e)
        if "prs-style build function not supported" in msg:
            try:
                fn(prs)
                after_master = master_xml_hash(prs)
                after_layouts = layout_xml_hashes(prs)
                checks["master_xml_unchanged"] = (after_master == before_master)
                if after_master != before_master:
                    failures.append({
                        "check": "master_xml_unchanged",
                        "expected": before_master[:12], "got": after_master[:12],
                    })
                # 각 slide의 safe_zone 검사
                safe_zone_ok = True
                for slide in prs.slides:
                    if slide.slide_layout.name == LAYOUT_COVER:
                        continue
                    for shape in slide.shapes:
                        if shape.is_placeholder:
                            continue
                        if not is_shape_in_safe_zone(shape):
                            safe_zone_ok = False
                            failures.append({
                                "check": "all_shapes_in_safe_zone",
                                "slide_layout": slide.slide_layout.name,
                                "shape": shape.name,
                                "bbox_emu": [shape.left, shape.top, shape.width, shape.height],
                            })
                checks["all_shapes_in_safe_zone"] = safe_zone_ok
                # placeholder idx 존중 검사
                ph_ok = True
                for slide in prs.slides:
                    expected_keep = LOCKED_PH_IDX.get(
                        slide.slide_layout.name, [0]
                    )
                    if not placeholder_idx_respected(slide, expected_keep):
                        ph_ok = False
                        failures.append({
                            "check": "placeholder_idx_respected",
                            "slide_layout": slide.slide_layout.name,
                            "expected_keep": expected_keep,
                        })
                checks["placeholder_idx_respected"] = ph_ok
                # layout이 허용 목록에 있는지
                for slide in prs.slides:
                    if not slide_uses_allowed_layout(slide):
                        failures.append({
                            "check": "layout_allowed",
                            "slide_layout": slide.slide_layout.name,
                        })
            except Exception as inner:
                checks["build_function_runs"] = False
                failures.append({"check": "build_function_runs", "error": repr(inner)})
        else:
            checks["build_function_runs"] = False
            failures.append({"check": "build_function_runs", "error": repr(e)})
    except Exception as e:
        checks["build_function_runs"] = False
        failures.append({"check": "build_function_runs", "error": repr(e)})
    else:
        # TemplateGuard 통과
        checks["master_and_safe_zone_intact"] = True

        # placeholder idx 검사 (guard 밖에서)
        placeholder_ok = True
        for slide in prs.slides:
            expected_keep = LOCKED_PH_IDX.get(slide.slide_layout.name, [0])
            if not placeholder_idx_respected(slide, expected_keep):
                placeholder_ok = False
                failures.append({
                    "check": "placeholder_idx_respected",
                    "slide_layout": slide.slide_layout.name,
                    "expected_keep": expected_keep,
                })
        checks["placeholder_idx_respected"] = placeholder_ok

    return checks, failures


# ============================ 공개 API ============================

def validate_slide_code(
    code_path: str | Path,
    spec_path: str | Path | None = None,
) -> ValidationReport:
    """단일 NN.code.py에 대한 검증 리포트 생성."""
    code_path = Path(code_path)
    spec_path = Path(spec_path) if spec_path else None
    spec = _load_spec(spec_path)
    slide_idx = spec.get("idx")

    report = ValidationReport(slide_idx=slide_idx, code_path=str(code_path))

    # 1. import whitelist
    ok, viols = _check_imports(code_path)
    report.checks["imports_whitelisted"] = ok
    if not ok:
        report.failures.append({"check": "imports_whitelisted", "violations": viols})

    # 2. AST: 마스터 직접 수정
    ok, viols = _check_no_master_edit(code_path)
    report.checks["no_direct_master_edit"] = ok
    if not ok:
        report.failures.append({"check": "no_direct_master_edit", "violations": viols})

    # 3. 런타임 검사: import + build + guard
    rt_checks, rt_failures = _runtime_check(code_path, slide_idx, spec)
    report.checks.update(rt_checks)
    report.failures.extend(rt_failures)

    return report


def validate_full_deck(pptx_path: str | Path) -> ValidationReport:
    """완성된 .pptx 파일에 대한 검증.

    마스터 XML 해시를 템플릿 원본과 비교하고, 모든 슬라이드의 shape가 CONTENT_SAFE 안인지 검사.
    """
    pptx_path = Path(pptx_path)
    report = ValidationReport(slide_idx=None, code_path=str(pptx_path))

    # 원본 템플릿 해시
    template_prs = load_template()
    expected_master = master_xml_hash(template_prs)
    expected_layouts = layout_xml_hashes(template_prs)

    deck_prs = Presentation(str(pptx_path))
    actual_master = master_xml_hash(deck_prs)
    actual_layouts = layout_xml_hashes(deck_prs)

    report.checks["master_xml_unchanged"] = (actual_master == expected_master)
    if actual_master != expected_master:
        report.failures.append({
            "check": "master_xml_unchanged",
            "expected": expected_master[:12], "got": actual_master[:12],
        })

    layouts_ok = True
    for name, exp in expected_layouts.items():
        if actual_layouts.get(name) != exp:
            layouts_ok = False
            report.failures.append({
                "check": "layout_xml_unchanged",
                "layout": name,
            })
    report.checks["layouts_xml_unchanged"] = layouts_ok

    safe_ok = True
    for i, slide in enumerate(deck_prs.slides):
        if slide.slide_layout.name == LAYOUT_COVER:
            continue
        for shape in slide.shapes:
            if shape.is_placeholder:
                continue
            if not is_shape_in_safe_zone(shape):
                safe_ok = False
                report.failures.append({
                    "check": "all_shapes_in_safe_zone",
                    "slide_index": i,
                    "slide_layout": slide.slide_layout.name,
                    "shape": shape.name,
                    "bbox_emu": [shape.left, shape.top, shape.width, shape.height],
                })
    report.checks["all_shapes_in_safe_zone"] = safe_ok

    return report


# ============================ CLI ============================

def _cli():
    parser = argparse.ArgumentParser(description="Validate PPT slide code or full deck.")
    parser.add_argument("target", help="Path to NN.code.py or .pptx file")
    parser.add_argument("--spec", help="Path to NN.spec.yaml (for code.py only)")
    parser.add_argument("--output", help="Write JSON report to this path")
    args = parser.parse_args()

    target = Path(args.target)
    if target.suffix == ".pptx":
        report = validate_full_deck(target)
    else:
        report = validate_slide_code(target, spec_path=args.spec)

    if args.output:
        Path(args.output).write_text(report.to_json(), encoding="utf-8")
    else:
        print(report.to_json())

    sys.exit(0 if report.passed else 1)


if __name__ == "__main__":
    _cli()
