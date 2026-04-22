"""Refiner — (code + validation + evaluation) → 수정된 code.

실제 패치는 Claude가 fresh context에서 수행:
- 입력: 기존 code.py, validation.json, evaluation.json, spec.yaml
- 출력: 패치된 code.py

이 모듈은 I/O + 상한 제어만 담당.
"""

from __future__ import annotations

import json
from dataclasses import asdict
from pathlib import Path

from harness.schemas import SlideSpec, EvaluationReport


REFINER_PROMPT = """\
당신은 PPT 슬라이드 리파이너(수정자)다.

## 입력
- spec.yaml: 원본 요구사항
- 기존 code.py: 이전 Generator 출력
- validation.json: Deterministic validator 실패 목록
- evaluation.json: LLM Evaluator 피드백

## 원칙
1. **fresh context** — 이전 대화 히스토리 없음. 네 개의 파일만 참조.
2. **핀포인트 수정** — 전체 재작성 금지. 실패한 check/feedback만 정확히 고침.
3. **제약 유지**:
   - 함수 시그니처 build_slide_N(slide) 고정
   - import whitelist 준수
   - CONTENT_SAFE 경계 준수

## 실패 분류

| 실패 종류 | 처리 |
|---|---|
| all_shapes_in_safe_zone | bbox 수정 — left/top/width/height 재배치 |
| imports_whitelisted | 금지 import 제거 또는 허용 모듈로 대체 |
| no_direct_master_edit | prs.slide_masters/slide_layouts 접근 삭제 |
| placeholder_idx_respected | set_title(idx=0) 외 placeholder 사용 금지 |
| Evaluator score < 4 | actionable_feedback 항목별 반영 |

## 출력
수정된 code.py 전체 텍스트.
"""


MAX_REFINE_ATTEMPTS = 3


def save_patched_code(code_text: str, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(code_text, encoding="utf-8")
    return path


def build_refine_context(
    spec: SlideSpec,
    old_code: str,
    validation: dict,
    evaluation: EvaluationReport | None,
) -> dict:
    """Refiner LLM에 전달할 구조화된 입력.

    Claude가 generate-ppt.md v2 워크플로에서 이 dict를 prompt로 꾸며 호출.
    """
    return {
        "spec": asdict(spec),
        "old_code": old_code,
        "validation": validation,
        "evaluation": asdict(evaluation) if evaluation else None,
    }


def save_trace(trace: dict, path: Path) -> Path:
    """디버깅·튜닝용 trace 저장."""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(trace, ensure_ascii=False, indent=2), encoding="utf-8")
    return path
