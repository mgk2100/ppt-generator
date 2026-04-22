"""공용 데이터 스키마 — Planner/Generator/Evaluator/Refiner 간 오가는 구조.

YAML ↔ dataclass 변환을 단순화.
"""

from __future__ import annotations

from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any

import yaml


# ============================ SlideSpec (Planner → Generator) ============================

@dataclass
class SlideSpec:
    idx: int
    layout: str                  # "제목 슬라이드" / "제목 및 내용" / "제목 및 내용 (페이지 번호 삭제)"
    title: str
    pattern: str                 # "cover" / "section_divider" / "cards" / "table" / "chart" / ...
    accent_color: str = "#1F497D"
    content_blocks: list[dict] = field(default_factory=list)
    data_refs: list[str] = field(default_factory=list)
    constraints: dict = field(default_factory=dict)
    key_message: str = ""        # 상단 Key Message Bar 텍스트
    notes: str = ""              # 생성자(LLM)를 위한 자유 메모

    def to_yaml(self) -> str:
        return yaml.safe_dump(asdict(self), allow_unicode=True, sort_keys=False)

    @classmethod
    def from_yaml(cls, text: str) -> "SlideSpec":
        data = yaml.safe_load(text) or {}
        return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})

    @classmethod
    def load(cls, path: Path) -> "SlideSpec":
        return cls.from_yaml(Path(path).read_text(encoding="utf-8"))

    def save(self, path: Path):
        Path(path).write_text(self.to_yaml(), encoding="utf-8")


# ============================ DeckPlan (Planner 출력) ============================

@dataclass
class DeckPlan:
    project_name: str
    total_slides: int
    sources: list[str]           # 원본 자료 파일 경로들
    narrative: str               # 덱의 전체 흐름 요약 (1~2 문단)
    slides: list[int]            # SlideSpec idx 순서
    style_guide: dict = field(default_factory=dict)

    def to_yaml(self) -> str:
        return yaml.safe_dump(asdict(self), allow_unicode=True, sort_keys=False)

    @classmethod
    def load(cls, path: Path) -> "DeckPlan":
        data = yaml.safe_load(Path(path).read_text(encoding="utf-8")) or {}
        return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})

    def save(self, path: Path):
        Path(path).write_text(self.to_yaml(), encoding="utf-8")


# ============================ EvaluationReport (Evaluator → Refiner) ============================

@dataclass
class EvaluationReport:
    slide_idx: int
    score: int                   # 0~5
    rubric: dict[str, int]       # sub-scores
    actionable_feedback: list[str]
    passed: bool                 # score >= threshold?

    def to_json(self) -> str:
        import json
        return json.dumps(asdict(self), ensure_ascii=False, indent=2)

    @classmethod
    def from_json(cls, text: str) -> "EvaluationReport":
        import json
        return cls(**json.loads(text))

    @classmethod
    def load(cls, path: Path) -> "EvaluationReport":
        return cls.from_json(Path(path).read_text(encoding="utf-8"))

    def save(self, path: Path):
        Path(path).write_text(self.to_json(), encoding="utf-8")


# ============================ LoopArtifact 경로 도우미 ============================

@dataclass(frozen=True)
class ProjectPaths:
    name: str
    project_root: Path                           # /home/ubuntu/Share/ppt-generator

    @property
    def input_dir(self) -> Path:
        return self.project_root / "input" / self.name

    @property
    def slides_dir(self) -> Path:
        return self.input_dir / "slides"

    @property
    def traces_dir(self) -> Path:
        return self.input_dir / "traces"

    @property
    def plan_path(self) -> Path:
        return self.input_dir / "plan.yaml"

    @property
    def analysis_path(self) -> Path:
        return self.input_dir / "analysis.yaml"

    @property
    def output_path(self) -> Path:
        return self.project_root / "output" / f"{self.name}.pptx"

    def spec_path(self, idx: int) -> Path:
        return self.slides_dir / f"slide_{idx:02d}.spec.yaml"

    def code_path(self, idx: int) -> Path:
        return self.slides_dir / f"slide_{idx:02d}.code.py"

    def validation_path(self, idx: int) -> Path:
        return self.slides_dir / f"slide_{idx:02d}.validation.json"

    def evaluation_path(self, idx: int) -> Path:
        return self.slides_dir / f"slide_{idx:02d}.evaluation.json"

    def trace_path(self, idx: int, attempt: int) -> Path:
        return self.traces_dir / f"slide_{idx:02d}_attempt_{attempt:02d}.json"

    def png_path(self, idx: int) -> Path:
        return self.input_dir / "renders" / f"slide_{idx:02d}.png"

    def ensure_dirs(self):
        self.slides_dir.mkdir(parents=True, exist_ok=True)
        self.traces_dir.mkdir(parents=True, exist_ok=True)
        (self.input_dir / "renders").mkdir(parents=True, exist_ok=True)
        self.output_path.parent.mkdir(parents=True, exist_ok=True)
