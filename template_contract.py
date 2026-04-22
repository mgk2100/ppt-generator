"""Template Contract — Layer 1 of the harness.

불변 계약(immutable contract): 표지 레이아웃과 마스터 요소들은 절대 건드리지 않는다.
이 모듈은 ref/locked_registry.yaml을 로드해 다음을 제공한다:

- 상수: LAYOUT_*, LOCKED_PH_IDX, CONTENT_SAFE
- 검증 primitives: is_shape_in_safe_zone, master_elements_intact, only_imports_whitelisted 등
- 보호 컨텍스트 매니저: TemplateGuard

사용 예:
    with TemplateGuard(prs) as guard:
        guard.add_slide(LAYOUT_CONTENT, lambda slide: build_slide_N(slide))
        ...
    # __exit__에서 마스터 XML 해시 검증. 변경 시 AssertionError.

의존성: pyyaml, python-pptx, lxml (이미 requirements.txt에 포함).
"""

from __future__ import annotations

import ast
import hashlib
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable

import yaml
from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Emu


# ============================ 경로 ============================

_BASE_DIR = Path(__file__).parent
_REGISTRY_PATH = _BASE_DIR / "ref" / "locked_registry.yaml"
_TEMPLATE_PATH = _BASE_DIR / "ref" / "표지.pptx"


# ============================ 레이아웃 상수 ============================

LAYOUT_COVER      = "제목 슬라이드"
LAYOUT_CONTENT    = "제목 및 내용"
LAYOUT_NO_PAGENUM = "제목 및 내용 (페이지 번호 삭제)"

LOCKED_PH_IDX: dict[str, list[int]] = {
    LAYOUT_COVER:      [0, 1],   # title + subtitle(date)
    LAYOUT_CONTENT:    [0],      # title only
    LAYOUT_NO_PAGENUM: [0],
}


# ============================ SafeZone ============================

@dataclass(frozen=True)
class SafeZone:
    """Content safe zone. 이 영역 밖으로 shape가 나가면 마스터 요소 가림."""
    left: int
    top: int
    width: int
    height: int
    tolerance_emu: int = 9525  # 1pt

    @property
    def right(self) -> int:
        return self.left + self.width

    @property
    def bottom(self) -> int:
        return self.top + self.height


# ============================ Registry Load ============================

@dataclass(frozen=True)
class LockedShapeSpec:
    id: str
    source_name: str
    layout: str
    kind: str
    bbox_in: tuple[float, float, float, float]
    cover_exception: bool = False
    hidden_in_layouts: tuple[str, ...] = field(default_factory=tuple)
    fixed_text_pattern: str | None = None
    fixed_text_prefix: str | None = None


@dataclass(frozen=True)
class TemplateRegistry:
    safe_zone: SafeZone
    layouts: tuple[str, ...]
    locked_shapes: tuple[LockedShapeSpec, ...]
    import_whitelist: frozenset[str]

    @classmethod
    def load(cls, path: Path = _REGISTRY_PATH) -> "TemplateRegistry":
        data = yaml.safe_load(path.read_text(encoding="utf-8"))

        sz = data["safe_zone"]
        safe = SafeZone(
            left=Inches(sz["left_in"]),
            top=Inches(sz["top_in"]),
            width=Inches(sz["width_in"]),
            height=Inches(sz["height_in"]),
            tolerance_emu=int(sz.get("tolerance_emu", 9525)),
        )

        layouts = tuple(layout["name"] for layout in data["layouts"])

        locked_shapes = tuple(
            LockedShapeSpec(
                id=s["id"],
                source_name=s["source_name"],
                layout=s["layout"],
                kind=s["kind"],
                bbox_in=tuple(s["bbox_in"]),
                cover_exception=bool(s.get("cover_exception", False)),
                hidden_in_layouts=tuple(s.get("hidden_in_layouts", ())),
                fixed_text_pattern=s.get("fixed_text_pattern"),
                fixed_text_prefix=s.get("fixed_text_prefix"),
            )
            for s in data["master_locked_shapes"]
        )

        imports = frozenset(data.get("import_whitelist", []))

        return cls(
            safe_zone=safe,
            layouts=layouts,
            locked_shapes=locked_shapes,
            import_whitelist=imports,
        )


# 모듈 로드 시 1회 읽어서 캐시
REGISTRY = TemplateRegistry.load()
CONTENT_SAFE = REGISTRY.safe_zone


# ============================ 검증 primitives ============================

def is_shape_in_safe_zone(
    shape, safe: SafeZone = CONTENT_SAFE, tol: int | None = None
) -> bool:
    """shape.bbox ⊂ safe_zone + tolerance.

    line(직선 연결선) 같이 width/height=0인 shape는 point 취급 → 완화된 검사.
    """
    tol = tol if tol is not None else safe.tolerance_emu
    try:
        l, t, w, h = shape.left, shape.top, shape.width, shape.height
    except Exception:
        return True  # bbox 없음 (placeholder만 있고 실제 도형 아닌 경우)
    if any(v is None for v in (l, t, w, h)):
        return True
    return (
        l >= safe.left - tol
        and t >= safe.top - tol
        and (l + w) <= safe.right + tol
        and (t + h) <= safe.bottom + tol
    )


def slide_uses_allowed_layout(slide, allowed: Iterable[str] = None) -> bool:
    """slide의 layout.name이 허용 목록 안에 있는지."""
    allowed = allowed or REGISTRY.layouts
    return slide.slide_layout.name in allowed


def placeholder_idx_respected(slide, expected_keep: list[int]) -> bool:
    """slide에 남아있는 placeholder의 idx가 expected_keep의 부분집합인지.

    ghost text만 있는 placeholder는 clear_placeholders()에서 제거됨. 여기서는
    '비어있지 않은(사용 중인) placeholder들만' 검사해서 expected_keep 안에 있는지 확인.
    """
    used_idx = set()
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        if ph.has_text_frame and ph.text_frame.text.strip():
            used_idx.add(idx)
    return used_idx.issubset(set(expected_keep))


def _xml_hash(element) -> str:
    """XML element를 canonical string으로 변환 후 SHA256 해시."""
    s = etree.tostring(element, method="c14n")
    return hashlib.sha256(s).hexdigest()


def _normalize_page_number(xml_bytes: bytes) -> bytes:
    """마스터 XML 중 페이지 번호 텍스트를 표준형으로 정규화.

    XML 구조: `<a:fld>...<a:t>‹#›</a:t></a:fld><a:r>...<a:t>/ 5</a:t></a:r>`
    "‹#›"는 필드(별도 `<a:t>`), "/ <N>"은 별도 run 의 `<a:t>` 안에 위치.

    따라서 `<a:t>` 태그 안의 "/ <숫자>" 패턴만 찾아 "/ 00" 으로 정규화.
    Assembler 의 apply_page_total 로 인한 동적 수정이 hash 비교에 영향 없도록 한다.
    """
    import re as _re
    text = xml_bytes.decode("utf-8")
    text = _re.sub(r"<a:t>(\s*/\s*)\d+(\s*)</a:t>", r"<a:t>\g<1>00\g<2></a:t>", text)
    return text.encode("utf-8")


def master_xml_hash(prs: Presentation) -> str:
    """Presentation의 slide_masters[0] XML 해시.

    페이지 번호 텍스트 ("‹#› / <N>") 는 정규화 후 해시 — Assembler 의
    apply_page_total 로 인한 수정을 허용한다.
    """
    raw = etree.tostring(prs.slide_masters[0]._element, method="c14n")
    return hashlib.sha256(_normalize_page_number(raw)).hexdigest()


def layout_xml_hashes(prs: Presentation) -> dict[str, str]:
    """각 레이아웃의 XML 해시."""
    return {
        layout.name: _xml_hash(layout._element)
        for layout in prs.slide_masters[0].slide_layouts
    }


def only_imports_whitelisted(
    code_path: Path, whitelist: Iterable[str] = None
) -> tuple[bool, list[str]]:
    """AST 분석: code_path의 import가 whitelist 안에 있는지.

    Returns:
        (passed, violating_imports)
    """
    whitelist = set(whitelist) if whitelist else set(REGISTRY.import_whitelist)
    src = Path(code_path).read_text(encoding="utf-8")
    try:
        tree = ast.parse(src)
    except SyntaxError as e:
        return False, [f"SyntaxError: {e}"]

    violations: list[str] = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                root = alias.name.split(".")[0]
                if alias.name not in whitelist and root not in whitelist:
                    violations.append(f"import {alias.name}")
        elif isinstance(node, ast.ImportFrom):
            if node.module is None:
                continue
            root = node.module.split(".")[0]
            if node.module not in whitelist and root not in whitelist:
                violations.append(f"from {node.module} import ...")
    return (len(violations) == 0, violations)


# AST 검사: forbidden patterns
_FORBIDDEN_AST_PATTERNS = [
    # prs.slide_masters[...]... 접근 후 shapes.add_* 호출
    # prs.slide_layouts[...]... 접근 후 shapes.add_* 호출
    # 간단 버전: 'slide_masters' 또는 'slide_layouts'와 'shapes.add_' 함께 출현
]


def no_direct_master_edit(code_path: Path) -> tuple[bool, list[str]]:
    """AST 분석: prs.slide_masters[...] 또는 prs.slide_layouts[...] 에 shape 추가 시도.

    Returns:
        (passed, violations)
    """
    src = Path(code_path).read_text(encoding="utf-8")
    try:
        tree = ast.parse(src)
    except SyntaxError as e:
        return False, [f"SyntaxError: {e}"]

    violations: list[str] = []

    class _MasterEditVisitor(ast.NodeVisitor):
        def visit_Attribute(self, node: ast.Attribute):
            # slide_master(s).shapes.add_* 또는 slide_layouts[n].shapes.add_*
            # chain을 unparse해서 슬쩍 검사
            try:
                src_text = ast.unparse(node)
            except Exception:
                src_text = ""
            if (
                ("slide_master" in src_text or "slide_layouts" in src_text)
                and ("shapes.add_" in src_text or ".shapes.remove" in src_text)
            ):
                violations.append(src_text)
            self.generic_visit(node)

    _MasterEditVisitor().visit(tree)
    return (len(violations) == 0, violations)


# ============================ TemplateGuard ============================

class TemplateGuard:
    """프리젠테이션 구축 중 마스터·레이아웃 무결성을 보장하는 컨텍스트 매니저.

    사용 예::

        from pptx import Presentation
        from template_contract import TemplateGuard, LAYOUT_CONTENT

        prs = Presentation("ref/표지.pptx")
        # 샘플 슬라이드 제거 (ppt_utils.load_template이 하는 일)
        ...

        with TemplateGuard(prs) as guard:
            guard.add_slide(LAYOUT_CONTENT, build_slide_1)
            guard.add_slide(LAYOUT_CONTENT, build_slide_2)
            ...
        prs.save("output/deck.pptx")

    __exit__ 시:
        - 마스터 XML 해시가 초기값과 동일한지 확인. 다르면 AssertionError.
        - 각 레이아웃 XML 해시도 확인.
        - add_slide 중에 safe_zone 위반이 있었다면 모아서 raise.
    """

    def __init__(self, prs: Presentation, strict_safe_zone: bool = True):
        self.prs = prs
        self.strict_safe_zone = strict_safe_zone
        self._master_hash: str | None = None
        self._layout_hashes: dict[str, str] = {}
        self._violations: list[str] = []

    def __enter__(self) -> "TemplateGuard":
        self._master_hash = master_xml_hash(self.prs)
        self._layout_hashes = layout_xml_hashes(self.prs)
        return self

    def add_slide(
        self,
        layout_name: str,
        on_build: Callable[[object], None],
    ):
        """layout_name 레이아웃 슬라이드 추가 후 on_build(slide) 실행.

        on_build는 slide를 받아 shape/text를 추가하는 함수.
        반환: 추가된 slide 객체.
        """
        # 늦은 import — 순환 회피
        from ppt_utils import get_layout, clear_placeholders

        if layout_name not in REGISTRY.layouts:
            raise ValueError(
                f"Layout '{layout_name}' not in registry. "
                f"Allowed: {list(REGISTRY.layouts)}"
            )

        layout = get_layout(self.prs, layout_name)
        slide = self.prs.slides.add_slide(layout)

        keep = LOCKED_PH_IDX.get(layout_name, [0])
        clear_placeholders(slide, keep=keep)

        # on_build 실행
        on_build(slide)

        # post-check: safe_zone 위반
        if self.strict_safe_zone and layout_name != LAYOUT_COVER:
            for shape in slide.shapes:
                # placeholder는 레이아웃에서 상속 — 검사 제외
                if shape.is_placeholder:
                    continue
                if not is_shape_in_safe_zone(shape):
                    self._violations.append(
                        f"Slide (layout={layout_name}): shape '{shape.name}' "
                        f"out of CONTENT_SAFE. "
                        f"bbox=({shape.left},{shape.top})+{shape.width}x{shape.height}"
                    )

        return slide

    def __exit__(self, exc_type, exc_val, exc_tb):
        # 예외 이미 발생한 경우에도 무결성 체크는 수행
        errors: list[str] = list(self._violations)

        now_master = master_xml_hash(self.prs)
        if now_master != self._master_hash:
            errors.append(
                f"Master XML changed! "
                f"expected={self._master_hash[:12]}... got={now_master[:12]}..."
            )

        now_layouts = layout_xml_hashes(self.prs)
        for name, before_hash in self._layout_hashes.items():
            now = now_layouts.get(name)
            if now != before_hash:
                errors.append(
                    f"Layout '{name}' XML changed! "
                    f"expected={before_hash[:12]}... got={now[:12] if now else 'MISSING'}..."
                )

        if errors and exc_type is None:
            raise TemplateContractViolation(
                f"{len(errors)} template contract violation(s):\n  - "
                + "\n  - ".join(errors)
            )
        # 이미 다른 예외 중이면 우리 오류는 append만 하고 원 예외를 그대로 올림
        return False


# ============================ 예외 ============================

class TemplateContractViolation(AssertionError):
    """Template contract 위반. 마스터 오염, safe_zone 위반 등."""


# ============================ 편의 API ============================

def load_guarded_presentation() -> tuple[Presentation, TemplateGuard]:
    """편의: ref/표지.pptx 로드 후 샘플 슬라이드 제거, TemplateGuard 래핑.

    반환: (prs, guard) — caller가 `with guard:` 블록 안에서 build.
    """
    from ppt_utils import load_template
    prs = load_template()
    guard = TemplateGuard(prs)
    return prs, guard


__all__ = [
    "LAYOUT_COVER",
    "LAYOUT_CONTENT",
    "LAYOUT_NO_PAGENUM",
    "LOCKED_PH_IDX",
    "SafeZone",
    "CONTENT_SAFE",
    "TemplateRegistry",
    "REGISTRY",
    "is_shape_in_safe_zone",
    "slide_uses_allowed_layout",
    "placeholder_idx_respected",
    "master_xml_hash",
    "layout_xml_hashes",
    "only_imports_whitelisted",
    "no_direct_master_edit",
    "TemplateGuard",
    "TemplateContractViolation",
    "load_guarded_presentation",
]
