"""harness — Layer 2 역할 분리 기반 PPT 생성 파이프라인.

역할:
    planner   — analysis.yaml → plan.yaml + NN.spec.yaml
    generator — NN.spec.yaml → NN.code.py (LLM, fresh context)
    validator — NN.code.py → NN.validation.json (deterministic)
    evaluator — rendered PNG → NN.evaluation.json (LLM, isolated context)
    refiner   — (code+v+e) → patched code (LLM, fresh context)
    assembler — 검증 통과 NN.code.py → output/{name}.pptx
    loop      — 오케스트레이터
"""
