"""Assembler — 검증 통과한 NN.code.py를 순서대로 빌드해 완성된 .pptx를 만든다.

두 가지 스타일 지원:
    A) NN.code.py 가 build_slide_N(slide) 를 정의 — TemplateGuard로 감쌈 (권장)
    B) NN.code.py 가 build_slide_N(prs) 를 정의 — 기존 스크립트 호환 (get_layout/add_slide 직접)

두 경우 모두 Assembler는 load_template()로 prs 생성 후 각 build 함수를 순서대로 호출.
완료 후 전체 deck에 대해 validate_full_deck 실행해 마스터 무결성 최종 감사.

CLI:
    python -m harness.assembler \
        --slides input/myproj/slides/*.code.py \
        --output output/myproj.pptx
"""

from __future__ import annotations

import argparse
import importlib.util
import inspect
import sys
from pathlib import Path
from typing import Callable

_HARNESS_DIR = Path(__file__).parent
_PROJECT_ROOT = _HARNESS_DIR.parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

from template_contract import (  # noqa: E402
    TemplateGuard,
    TemplateContractViolation,
    master_xml_hash,
    LAYOUT_CONTENT,
)
from ppt_utils import load_template  # noqa: E402
from harness.validator import validate_full_deck  # noqa: E402


# ============================ 내부 헬퍼 ============================

def _load_build_fn(code_path: Path) -> tuple[str, Callable]:
    """NN.code.py에서 build_slide_* 함수 하나를 찾아 반환.

    Returns:
        (function_name, callable)
    """
    mod_name = f"slide_{code_path.stem.replace('.', '_')}"
    spec = importlib.util.spec_from_file_location(mod_name, code_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)

    for attr in dir(module):
        if attr.startswith("build_slide_") or attr == "build_slide":
            fn = getattr(module, attr)
            if callable(fn):
                return attr, fn
    for attr in ("main",):
        fn = getattr(module, attr, None)
        if callable(fn):
            return attr, fn
    raise RuntimeError(f"No build_slide_* function in {code_path}")


def _fn_signature_style(fn: Callable) -> str:
    """함수 시그니처가 'slide' 스타일인지 'prs' 스타일인지 heuristic."""
    sig = inspect.signature(fn)
    params = list(sig.parameters.keys())
    if not params:
        return "none"
    first = params[0].lower()
    if first in ("slide", "sld"):
        return "slide"
    if first in ("prs", "presentation", "pr"):
        return "prs"
    # 이름 기반 폴백: 두 가지 모두 가능 — 'slide' 기본
    return "slide"


def _infer_layout_from_code(code_path: Path) -> str:
    """소스를 한 번 훑어 LAYOUT_COVER/CONTENT/NO_PAGENUM 중 어떤 걸 쓰는지 추론.

    판단 순서:
    1. setup_cover 호출 또는 LAYOUT_COVER/"제목 슬라이드" 언급 → 표지
    2. LAYOUT_NO_PAGENUM 또는 "페이지 번호 삭제" 언급 → 페이지번호 없음
    3. 그 외 → 기본 콘텐츠 레이아웃
    """
    src = code_path.read_text(encoding="utf-8")
    if "setup_cover" in src or "LAYOUT_COVER" in src or "제목 슬라이드" in src:
        return "제목 슬라이드"
    if "LAYOUT_NO_PAGENUM" in src or "페이지 번호 삭제" in src:
        return "제목 및 내용 (페이지 번호 삭제)"
    return LAYOUT_CONTENT


# ============================ 공개 API ============================

def assemble(
    slide_code_paths: list[Path],
    output_path: Path,
    use_guard: bool = True,
    audit: bool = True,
) -> dict:
    """슬라이드 코드 파일들을 순서대로 빌드해 .pptx 생성.

    Args:
        slide_code_paths: NN.code.py 경로 리스트 (순서 중요).
        output_path: 저장할 .pptx 경로.
        use_guard: slide 스타일 함수에 대해 TemplateGuard 사용 (권장).
        audit: 저장 후 validate_full_deck 실행.

    Returns:
        dict: {"output": path, "slides": n, "audit": ValidationReport or None}
    """
    prs = load_template()
    initial_hash = master_xml_hash(prs)

    guard = TemplateGuard(prs) if use_guard else None

    def _run_all():
        for code_path in slide_code_paths:
            name, fn = _load_build_fn(code_path)
            style = _fn_signature_style(fn)
            if style == "prs":
                # 레거시 스타일: 사용자가 직접 add_slide 하므로 guard 밖에서 호출
                fn(prs)
            else:
                # slide 스타일: guard로 슬라이드 추가 및 빌드
                layout_name = _infer_layout_from_code(code_path)
                if guard is not None:
                    guard.add_slide(layout_name, fn)
                else:
                    # use_guard=False → 수동 처리
                    from ppt_utils import get_layout, clear_placeholders
                    from template_contract import LOCKED_PH_IDX
                    layout = get_layout(prs, layout_name)
                    slide = prs.slides.add_slide(layout)
                    clear_placeholders(slide, keep=LOCKED_PH_IDX.get(layout_name, [0]))
                    fn(slide)

    if use_guard:
        with guard:
            _run_all()
    else:
        _run_all()

    # 마스터 해시 최종 확인 (guard 없이 호출된 경우를 위해)
    final_hash = master_xml_hash(prs)
    if final_hash != initial_hash:
        raise TemplateContractViolation(
            f"Master XML changed during assembly. "
            f"initial={initial_hash[:12]} final={final_hash[:12]}"
        )

    # 동적 페이지 번호 적용: "‹#› / 00" → "‹#› / <total>"
    # master_xml_hash 의 정규화가 "00" 부분을 무시하므로 hash 비교는 여전히 통과.
    from ppt_utils import apply_page_total
    apply_page_total(prs, total=len(prs.slides))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))

    audit_report = None
    if audit:
        audit_report = validate_full_deck(output_path)

    return {
        "output": str(output_path),
        "slides": len(prs.slides),
        "audit": audit_report.to_json() if audit_report else None,
        "audit_passed": audit_report.passed if audit_report else None,
    }


# ============================ CLI ============================

def _cli():
    parser = argparse.ArgumentParser(description="Assemble PPT from slide code files.")
    parser.add_argument("--slides", nargs="+", required=True,
                        help="NN.code.py files (order matters)")
    parser.add_argument("--output", required=True, help="Output .pptx path")
    parser.add_argument("--no-guard", action="store_true",
                        help="Disable TemplateGuard (legacy compat)")
    parser.add_argument("--no-audit", action="store_true",
                        help="Skip post-build validate_full_deck")
    args = parser.parse_args()

    slides = [Path(p) for p in args.slides]
    result = assemble(
        slide_code_paths=slides,
        output_path=Path(args.output),
        use_guard=not args.no_guard,
        audit=not args.no_audit,
    )
    print(f"✓ {result['output']}")
    print(f"  slides: {result['slides']}")
    if result["audit"] is not None:
        print(f"  audit_passed: {result['audit_passed']}")
        if not result["audit_passed"]:
            print(result["audit"])
            sys.exit(1)


if __name__ == "__main__":
    _cli()
