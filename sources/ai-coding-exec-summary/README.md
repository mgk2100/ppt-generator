# AI 코딩 어시스턴트 도입 영향 분석 — 임원용 요약

> **작성일**: 2026-04-21
> **대상**: 팀장(우선) · 실장 · 센터장
> **스코프**: 자동차 OEM · Tier-1 벤더사 (일반 IT 제외)

## 관련 파일 위치

### 최종 산출물
| 파일 | 경로 | 용도 |
|---|---|---|
| **임원용 PPT (4페이지)** | `output/ai-coding-exec-summary.pptx` | 팀장/실장/센터장 보고용 |
| **슬라이드 설계서** | `input/ai-coding-exec-summary_slide-plan.md` | PPT 구조 · 용어 순화 가이드 |

### 원본 분석 자료 (본 폴더)
| 파일 | 용도 |
|---|---|
| `AI_Coding_Assistant_Automotive_Impact_Analysis.md` | 10개 섹션 완전 분석 리포트 (본문) |
| `AI_Coding_Assistant_Automotive_Impact_Analysis.html` | 위 리포트의 HTML 버전 (인쇄/공유용) |
| `AI_Coding_Assistant_Automotive_Impact_Analysis_2pager.html` | 2페이지 축약 HTML (실무자용) |
| `generate.py` | PPT 재생성 스크립트 |

## PPT 재생성 방법

```bash
cd /home/ubuntu/Share/ppt-generator
python3 sources/ai-coding-exec-summary/generate.py
# → output/ai-coding-exec-summary.pptx 갱신
```

## 자료 계층

```
상세도 ↓                                  대상
─────────────────────────────────────────────
10-section 풀 리포트 (.md/.html)  →  실무 담당자
2-페이지 요약 (.html)              →  중간 관리자
4-페이지 임원용 PPT (.pptx)        →  팀장 · 실장 · 센터장
```

## 슬라이드 구성 요약

| # | 슬라이드 | 핵심 메시지 |
|---|---|---|
| 1 | 핵심 요약 | 영역별 차등 적용 · 6개월 내 투자 회수 · 즉시 결정 4건 |
| 2 | 왜 지금인가 | 경쟁사 이미 도입 · 2027.08 EU AI Act D-Day |
| 3 | 영역별 효과 | 신호등 평가 · 100명 기준 연 15억 절감 |
| 4 | 실행 계획 | 3단계 로드맵 · 4개월 시범 사업 · 승인 요청 4건 |

## 용어 순화 적용

본 임원용 PPT는 원본 리포트의 전문용어를 다음 기준으로 순화:

| 원문 | PPT 표현 |
|---|---|
| V&V / 검증 엔지니어 | 품질검증 인력 |
| PR (Pull Request) | 코드 리뷰 요청 |
| 컴플라이언스 | 규정 준수 |
| ASPICE / MISRA C | 자동차 SW 품질 표준 |
| ISO 26262 / ASIL B+ | 차량 안전 등급 (높음↑) |
| BSW / MCAL / AUTOSAR | 차량 제어 기반 SW |
| RAG / 파인튜닝 | 사내 데이터 기반 AI 학습 |
| ROI / BEP | 투자 회수 효과 / 손익분기점 |
| SDV | SW 중심 차량 |
