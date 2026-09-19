#!/usr/bin/env python3
"""SW Agent vs 인지 · 로봇 파이프라인 대조 2장 — 네이티브 도형 · 텍스트로 생성.

이미지 한 장을 얹지 않는다. 카드 · 배지 · 밴드는 모두 python-pptx 도형이고
글자는 모두 텍스트 박스라, PowerPoint 에서 문구 · 색 · 위치를 그대로 고칠 수 있다.
좌표는 아래 LAYOUT 상수에서 계산으로만 만든다(손으로 찍지 않는다).
"""
import os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))

from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

from deck_common import (new_slide, header, shape, tb, REPO,
                         NAVY, BLUE, BLUE_T, BLUE_B, GOLD_L, GREY, WHITE, TBL_RULE)

# ── 이 자료에서만 쓰는 색 ──────────────────────────────────
SLATE      = RGBColor(0x64, 0x74, 0x8B)   # 인지 · 로봇 열 (= 우리 소관 아님)
SLATE_T    = RGBColor(0xF7, 0xF8, 0xFA)
SLATE_B    = RGBColor(0xD7, 0xDD, 0xE6)
LAST_BG    = RGBColor(0xFB, 0xFC, 0xFD)
LAST_BG_L  = RGBColor(0xE9, 0xF1, 0xFB)
HEAD_S_L   = RGBColor(0xD8, 0xE6, 0xF7)
HEAD_S_M   = RGBColor(0xE2, 0xE6, 0xEC)
BAND_TX    = RGBColor(0xE8, 0xED, 0xF5)
NOTE_M_BG  = RGBColor(0xEE, 0xF0, 0xF4);  NOTE_M_TX = RGBColor(0x3C, 0x46, 0x55)
NOTE_L_BG  = RGBColor(0xE4, 0xEE, 0xFA);  NOTE_L_TX = RGBColor(0x2B, 0x4A, 0x74)

# ── 본문 영역 (마스터의 SL 로고 · 금색 헤어라인 아래) ──────
X0, Y0, W, H = 0.35, 1.24, 12.63, 6.02
LAB_W, GAP_X = 1.04, 0.10                       # 행 라벨 열 · 열 간격
COL_W = (W - LAB_W - GAP_X * 3) / 3
COL_X = [X0 + LAB_W + GAP_X + i * (COL_W + GAP_X) for i in range(3)]

PT = 0.72                                        # HTML px → pt
RAD = 0.05                                       # 카드 모서리 반지름(인치)
def adj(h):                                      # roundRect 조정값 = 반지름 / 짧은 변
    return [min(RAD / max(h, 0.01), 0.5)]

F_COLT, F_COLS = 13, 9                           # 열 헤더 제목 · 부제
F_RL          = 10                               # 행 라벨
F_K,  F_D     = 11, 9.5                          # 카드 제목 · 설명
F_LEAD, F_EXP = 14.5, 10.5                       # 결론 밴드
F_FN          = 9                                # 각주
LSP = 1.32                                       # 줄 간격 배수


def card(slide, x, y, w, h, fill, line):
    return shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h,
                 fill=fill, line=line, line_w=0.75, adj=adj(h))


def para_box(slide, x, y, w, h, paras, pad_x=0.13, gap_pt=2,
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT):
    """문단 리스트를 한 텍스트 박스에 담는다. paras = [[(글, 스타일), ...], ...]"""
    box = tb(slide, x + pad_x, y, w - pad_x * 2, h, paras,
             align=align, anchor=anchor, spacing=LSP)
    for p in box.text_frame.paragraphs[1:]:
        p.space_before = Pt(gap_pt)
    return box


def band(slide, y, h, lead, exp, lead_w=2.62):
    """결론 밴드 — 좌측 금색 결론 + 우측 근거 2줄"""
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, X0, y, W, h, fill=NAVY, adj=adj(h))
    tb(slide, X0 + 0.28, y, lead_w, h, lead, size=F_LEAD, bold=True, color=GOLD_L,
       anchor=MSO_ANCHOR.MIDDLE, spacing=1.3)
    ex = X0 + 0.28 + lead_w + 0.38
    tb(slide, ex, y, X0 + W - 0.28 - ex, h, exp, size=F_EXP, color=BAND_TX,
       anchor=MSO_ANCHOR.MIDDLE, spacing=1.5)


def footnotes(slide, y, h, lines):
    tb(slide, X0, y, W, h, [[(t, {})] for t in lines],
       size=F_FN, color=GREY, anchor=MSO_ANCHOR.TOP, spacing=1.55)


# =====================================================================
# 1장 — 비교 항목 6개 × 3열
# =====================================================================
COLS = [("SW Agent (LLM)", "AI 활용 — 완성된 상용 모델 도입", BLUE,  HEAD_S_L),
        ("영상 · LiDAR 인지", "AI 개발 — 모델 직접 제작",     SLATE, HEAD_S_M),
        ("로봇 강화학습",     "AI 개발 — 모델 직접 제작",     SLATE, HEAD_S_M)]

ROWS = [
    ("원천 데이터", [
        ("업무 수행 중 자동 축적",       "설계자의 질의 · 답변, 업로드 코드 · 문서가 그대로 기록"),
        ("현장 취득 선행 필요",          "야간 · 우천 · 역광 등 필요 상황을 계획해 직접 주행 · 촬영"),
        ("시뮬레이터 구축 선행 필요",    "환경 · 물리 조건 구축 후 실기 시행착오로 보완")]),
    ("정답 부여", [
        ("사람의 정답 부여 없음",        "설계자의 채택 · 수정이 정답 신호를 대체"),
        ("한 장씩 사람이 annotation",    "업계 통상 수만 장 규모 — 전체 공수의 최대 구간"),
        ("정답 부재",                    "사람이 보상 함수를 설계하고 결과에 따라 반복 조정")]),
    ("모델", [
        ("학습 단계 없음",               "상용 모델을 그대로 호출 — 모델 제작 없음"),
        ("태스크별 네트워크 선정 후 학습", "가공 데이터로 학습 → 재학습 반복"),
        ("정책 네트워크 학습",           "시뮬레이터 내 반복 시행으로 획득")]),
    ("성능 개선", [
        ("모델이 아닌 데이터를 개선",    "지식 유닛 품질 · 검색 정확도 향상 작업"),
        ("모델 구조 직접 수정",          "레이어 추가 · 삭제, 손실 함수 · 하이퍼파라미터 조정"),
        ("보상 재설계",                  "시뮬레이터와 실기의 차이 보정 작업 수반")]),
    ("적용 · 배포", [
        ("사내 서버 탑재 후 호출",       "탑재 시 전 설계자가 동일 경로로 사용"),
        ("SOC 포팅 후 재검증",           "경량화 · 양자화 및 실시간 처리 제약 내 재측정"),
        ("실기 안전 검증",               "시뮬레이터 결과의 실기 재현 여부 확인")]),
    ("필요 인력\n· 장비", [
        ("설계 도메인 이해 SW 설계자\n+ 검색용 사내 서버",
         "학습용 GPU · annotation 인력 · 포팅 검증 설비 모두 불필요"),
        ("annotation 인력 + 학습용 GPU\n+ 임베디드 포팅 엔지니어",
         "촬영 장비 · 대용량 저장소 별도 필요"),
        ("시뮬레이션 개발자 + 학습용 GPU\n+ 실기 시험 설비",
         "안전 검증 절차 별도 구축 필요")]),
]

HEAD_H, BAND_H, FN_H, GAP_Y, ROW_GAP = 0.50, 0.70, 0.39, 0.09, 0.06
LAST_RATIO = 1.30


def slide_one(prs):
    s = new_slide(prs)
    header(s, "비교", "SW Agent 는 AI 활용, 인지 · 로봇은 AI 개발 업무",
           [("AI 라는 명칭은 동일하나, ", {}),
            ("데이터 · 학습 · 검증 · 인력이 단계마다 상이하여 동일 조직 수행 불가", {"color": BLUE})],
           sub_w=11.6, sub_size=15)

    # 열 헤더
    tb(s, X0, Y0, LAB_W, HEAD_H, "비교 항목", size=9, bold=True, color=SLATE,
       anchor=MSO_ANCHOR.BOTTOM)
    for i, (name, sub, fill, sub_c) in enumerate(COLS):
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, COL_X[i], Y0, COL_W, HEAD_H,
              fill=fill, adj=adj(HEAD_H))
        para_box(s, COL_X[i], Y0, COL_W, HEAD_H,
                 [[(name, {"size": F_COLT, "bold": True, "color": WHITE})],
                  [(sub,  {"size": F_COLS, "bold": True, "color": sub_c})]],
                 pad_x=0.14, gap_pt=1)

    # 본문 행
    rows_y = Y0 + HEAD_H + GAP_Y
    rows_h = H - HEAD_H - GAP_Y * 3 - BAND_H - FN_H
    unit = (rows_h - ROW_GAP * (len(ROWS) - 1)) / (len(ROWS) - 1 + LAST_RATIO)
    y = rows_y
    for ri, (label, cells) in enumerate(ROWS):
        last = ri == len(ROWS) - 1
        h = unit * (LAST_RATIO if last else 1)
        tb(s, X0, y, LAB_W, h, [[(t, {})] for t in label.split("\n")],
           size=F_RL, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE, spacing=1.25)
        for ci, (k, d) in enumerate(cells):
            llm = ci == 0
            fill = (LAST_BG_L if llm else LAST_BG) if last else (BLUE_T if llm else WHITE)
            card(s, COL_X[ci], y, COL_W, h, fill, BLUE_B if llm else TBL_RULE)
            kst = {"size": F_K if not last else F_K - 0.5, "bold": True,
                   "color": BLUE if llm else NAVY}
            paras = [[(line, kst)] for line in k.split("\n")]
            paras.append([(d, {"size": F_D, "color": GREY})])
            para_box(s, COL_X[ci], y, COL_W, h, paras, gap_pt=2)
        y += h + ROW_GAP

    band(s, rows_y + rows_h + GAP_Y, BAND_H,
         [[("AI 라는 명칭만 동일", {})], [("수행 업무는 상이", {})]],
         [[("SW Agent 는 ", {}),
           ("완성된 상용 모델을 도입해 사내 지식을 연결하는 업무", {"bold": True, "color": WHITE}),
           (", 인지 · 로봇은 ", {}),
           ("데이터에 정답을 부여해 모델을 직접 제작하는 업무", {"bold": True, "color": WHITE}),
           (".", {})],
          [("6개 항목 전부에서 데이터 · 학습 유무 · 인력 · 장비가 상이 — 한쪽 인력의 업무 이관 불가.", {})]])

    footnotes(s, Y0 + H - FN_H, FN_H, [
        "※ annotation(주석 작업) : 사진 · 점군 한 장마다 사람이 정답(대상과 위치)을 표시하는 작업",
        "※ 영상 · LiDAR 인지 · 로봇 강화학습 항목의 장수 · 반복 횟수는 업계 통상 범위이며 SL 실적치 아님"])


# =====================================================================
# 2장 — 앞단 데이터 확보 방식
# =====================================================================
P_LW, P_GAP = 4.00, 0.16
P_RX = X0 + P_LW + P_GAP
P_RW = W - P_LW - P_GAP
PH_H, ST_GAP, NOTE_H = 0.40, 0.08, 0.38
NUM_D = 0.22                                     # 번호 원 지름
TXT_X = 0.15 + NUM_D + 0.09                      # 카드 좌측 여백 + 번호 + 간격

LEFT_STEPS = [
    ("1", "현장 취득", 2,
     [[("야간 · 우천 · 역광 등 필요 상황을 계획해 직접 주행 · 촬영.", {})],
      [("촬영 차량 · 센서 · 대용량 저장소 선행 구비 필요", {})]]),
    ("2", "annotation — 한 장씩 정답 표시", 3,
     [[("취득한 사진 · 점군에 사람이 개별 정답을 부여해야 학습 가능.", {})],
      [("업계 통상 ", {}), ("수만 장", {"bold": True, "color": NAVY}),
       (" 규모로 전체 공수의 최대 구간", {})]]),
    ("3", "검수", 2,
     [[("부여한 정답의 정확성 재확인.", {})],
      [("오류 시 학습 결과에 그대로 반영되어 사후 정정 곤란", {})]]),
]

RIGHT_STEPS = [
    ("1", "자동 축적 — 별도 취득 작업 없음", 1,
     [[("설계자의 일상 질의 시 질문 · 입력 자료 · 답변이 그대로 기록. 업로드 코드 · 설계 문서도 함께 축적", {})]]),
    ("2", "정제 — 축적분 그대로 사용 불가", 1,
     [[("중복 제거 · 사내 보안 규칙 기준 민감 문자열 삭제 · 개인정보 삭제 · 길이 필터 · 플랫폼 내부 구현 정보 제외", {})]]),
    ("3", "유닛화 — 재사용 가능 단위로 분할", 1,
     [[("대화 단위 보관 시 검색 불가. 성격별 분할 후 7종 분류", {})]]),
    ("4", "선별 적재 — 신뢰 가능분만 유지", 1,
     [[("조각별 신뢰도 점수 부여 후 ", {}), ("기준 미만 제외", {"bold": True, "color": NAVY}),
       (", 잔여분만 검색 DB 적재하여 후속 질의에 인용", {})]]),
]
PILLS = ["질답", "결정", "패턴", "방법", "주의점", "참조", "기타"]
PILL_H, PILL_GAP = 0.20, 0.05
LINE_K, LINE_D, PILL_BLOCK = 0.20, 0.19, 0.26    # 제목 1줄 · 설명 1줄 · pills 한 줄 높이


def panel(slide, x, y, w, h, title, sub, steps, note, llm, pills_at=None):
    main   = BLUE if llm else SLATE
    bg     = BLUE_T if llm else SLATE_T
    border = BLUE_B if llm else SLATE_B
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, PH_H, fill=main, adj=adj(PH_H))
    tb(slide, x + 0.15, y, w - 0.30, PH_H,
       [(title, {"size": 12, "bold": True, "color": WHITE}),
        ("   " + sub, {"size": 9.5, "color": HEAD_S_L if llm else HEAD_S_M})],
       anchor=MSO_ANCHOR.MIDDLE)

    steps_y = y + PH_H + 0.08
    steps_h = h - PH_H - 0.08 - 0.09 - NOTE_H
    sh_each = (steps_h - ST_GAP * (len(steps) - 1)) / len(steps)
    sy = steps_y
    for si, (num, k, dlines, dparas) in enumerate(steps):
        shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, sy, w, sh_each,
              fill=bg, line=border, line_w=0.75, adj=adj(sh_each))
        body_h = LINE_K + 0.03 + LINE_D * dlines + (PILL_BLOCK if si == pills_at else 0)
        top = sy + (sh_each - body_h) / 2
        # 번호 원 + 제목
        circ = shape(slide, MSO_SHAPE.OVAL, x + 0.15, top + (LINE_K - NUM_D) / 2,
                     NUM_D, NUM_D, fill=main)
        tb(slide, x + 0.15, top + (LINE_K - NUM_D) / 2, NUM_D, NUM_D, num,
           size=8.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        tb(slide, x + TXT_X, top, w - TXT_X - 0.15, LINE_K, k,
           size=F_K, bold=True, color=main if llm else NAVY,
           anchor=MSO_ANCHOR.MIDDLE, spacing=1.28)
        tb(slide, x + TXT_X, top + LINE_K + 0.03, w - TXT_X - 0.15, LINE_D * dlines,
           dparas, size=F_D, color=GREY, anchor=MSO_ANCHOR.TOP, spacing=1.45)
        if si == pills_at:
            px = x + TXT_X
            py = top + LINE_K + 0.03 + LINE_D * dlines + 0.06
            for t in PILLS:
                pw = 0.18 + len(t) * 0.125
                shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, px, py, pw, PILL_H,
                      fill=WHITE, line=BLUE_B, line_w=0.6, adj=adj(PILL_H))
                tb(slide, px, py, pw, PILL_H, t, size=8.5, bold=True, color=BLUE,
                   align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
                px += pw + PILL_GAP
        sy += sh_each + ST_GAP

    ny = y + h - NOTE_H
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, ny, w, NOTE_H,
          fill=NOTE_L_BG if llm else NOTE_M_BG, adj=adj(NOTE_H))
    tb(slide, x + 0.14, ny, w - 0.28, NOTE_H, note, size=10,
       color=NOTE_L_TX if llm else NOTE_M_TX, anchor=MSO_ANCHOR.MIDDLE, spacing=1.42)


def slide_two(prs):
    s = new_slide(prs)
    header(s, "비교", "앞단 데이터 확보 방식 상이 — 인력 투입 vs 사용량 축적",
           [("인지 · 로봇은 사람의 정답 부여로 데이터 성립, ", {}),
            ("SW Agent 는 설계자 사용 중 자동 축적", {"color": BLUE})],
           sub_w=11.6, sub_size=15)

    panels_h = H - 0.11 * 2 - 0.74 - 0.20
    panel(s, X0, Y0, P_LW, panels_h,
          "영상 · LiDAR 인지", "사람의 정답 부여로 데이터 성립", LEFT_STEPS,
          [("앞단 확대 방법 = ", {}), ("annotation 인력 추가 투입", {"bold": True, "color": NAVY})],
          llm=False)
    panel(s, P_RX, Y0, P_RW, panels_h,
          "SW Agent (LLM)", "축적된 대화를 검색 가능한 단위로 분할 — 지식 유닛화", RIGHT_STEPS,
          [("앞단 확대 방법 = ", {}), ("설계자 사용량 증가", {"bold": True, "color": BLUE}),
           (" — 인력 투입으로 확대 불가", {})],
          llm=True, pills_at=2)

    band(s, Y0 + panels_h + 0.11, 0.74,
         [[("앞단 필요 인력 상이", {})]],
         [[("인지 · 로봇의 앞단은 ", {}), ("annotation 인력 투입량", {"bold": True, "color": WHITE}),
           ("에 비례 확대, SW Agent 의 앞단은 ", {}),
           ("설계 도메인 이해 설계자의 사용량", {"bold": True, "color": WHITE}), ("에 비례 확대.", {})],
          [("인력 전환으로 상대 앞단 대체 불가하며, 앞단 상이 시 후속 학습 · 검증 단계도 상이.", {})]],
         lead_w=2.45)

    footnotes(s, Y0 + H - 0.20, 0.20, [
        "※ 지식 유닛 7종 분류 및 선별 기준(신뢰도 0.6 미만 제외)은 SW Agent 지식 추출 기능의 현재 구현값"])


def main(out: str = ""):
    out = out or os.path.join(REPO, "output", "sw-dev-ai-llm-compare.pptx")
    prs = Presentation(os.path.join(os.path.dirname(HERE), "base_master.pptx"))
    slide_one(prs)
    slide_two(prs)
    prs.save(out)
    print(out)


if __name__ == "__main__":
    main(*sys.argv[1:])
