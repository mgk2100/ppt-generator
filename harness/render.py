"""단일 슬라이드를 PNG로 렌더 — Evaluator 입력용.

LibreOffice headless → PDF → pdftoppm → PNG.
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

_HARNESS_DIR = Path(__file__).parent
_PROJECT_ROOT = _HARNESS_DIR.parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

from pptx import Presentation  # noqa: E402
from ppt_utils import load_template  # noqa: E402


def _have(cmd: str) -> bool:
    return shutil.which(cmd) is not None


def _soffice() -> str:
    """설치된 LibreOffice 실행 파일 경로를 반환 (libreoffice 또는 soffice)."""
    return shutil.which("libreoffice") or shutil.which("soffice")


def render_single_slide(
    code_path: Path,
    layout_name: str,
    output_png: Path,
    tmp_dir: Path | None = None,
    dpi: int = 120,
) -> Path:
    """단일 NN.code.py를 빌드 → PNG 렌더.

    Args:
        code_path: 빌드할 NN.code.py
        layout_name: 슬라이드 레이아웃 이름
        output_png: 저장 경로
        tmp_dir: 임시 작업 디렉토리 (없으면 /tmp/harness_render)
        dpi: PNG 해상도

    Returns:
        생성된 PNG 경로.
    """
    if not _have("libreoffice") and not _have("soffice"):
        raise RuntimeError("LibreOffice 필요 (apt install libreoffice)")
    if not _have("pdftoppm"):
        raise RuntimeError("pdftoppm 필요 (apt install poppler-utils)")

    tmp_dir = tmp_dir or Path("/tmp/harness_render")
    tmp_dir.mkdir(parents=True, exist_ok=True)

    from harness.assembler import assemble
    tmp_pptx = tmp_dir / f"{code_path.stem}.pptx"
    assemble([code_path], tmp_pptx, use_guard=True, audit=False)

    # PPTX → PDF
    subprocess.run(
        [_soffice(), "--headless", "--convert-to", "pdf",
         "--outdir", str(tmp_dir), str(tmp_pptx)],
        check=True, capture_output=True,
    )
    tmp_pdf = tmp_dir / f"{tmp_pptx.stem}.pdf"
    if not tmp_pdf.exists():
        raise RuntimeError(f"PDF 변환 실패: {tmp_pdf}")

    # PDF → PNG
    output_png.parent.mkdir(parents=True, exist_ok=True)
    stem = output_png.stem
    subprocess.run(
        ["pdftoppm", "-r", str(dpi), "-png",
         "-f", "1", "-l", "1",
         str(tmp_pdf), str(output_png.parent / stem)],
        check=True, capture_output=True,
    )
    # pdftoppm은 파일명 뒤에 -1 붙임 → 리네임
    produced = output_png.parent / f"{stem}-1.png"
    if produced.exists():
        produced.replace(output_png)
    return output_png


def render_full_deck(
    pptx_path: Path,
    output_dir: Path,
    dpi: int = 120,
) -> list[Path]:
    """완성된 .pptx의 모든 슬라이드를 PNG로 렌더.

    Returns:
        생성된 PNG 파일 리스트 (슬라이드 순서).
    """
    if not _have("libreoffice") and not _have("soffice"):
        raise RuntimeError("LibreOffice 필요")
    if not _have("pdftoppm"):
        raise RuntimeError("pdftoppm 필요")

    output_dir.mkdir(parents=True, exist_ok=True)
    tmp = output_dir / "_tmp.pdf"

    subprocess.run(
        [_soffice(), "--headless", "--convert-to", "pdf",
         "--outdir", str(output_dir), str(pptx_path)],
        check=True, capture_output=True,
    )
    pdf = output_dir / f"{pptx_path.stem}.pdf"
    if not pdf.exists():
        raise RuntimeError("PDF 변환 실패")

    subprocess.run(
        ["pdftoppm", "-r", str(dpi), "-png",
         str(pdf), str(output_dir / "slide")],
        check=True, capture_output=True,
    )

    return sorted(output_dir.glob("slide-*.png"))
