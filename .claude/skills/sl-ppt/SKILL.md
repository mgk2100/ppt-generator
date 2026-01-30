---
name: sl-ppt
description: SL 템플릿을 사용하여 PowerPoint 프레젠테이션을 생성합니다. PPT 생성, 아키텍처 문서화, 보고서 작성 시 사용합니다.
argument-hint: [주제 또는 프로젝트 경로]
allowed-tools: Read, Write, Bash, Grep, Glob, Task
---

# SL-PPT Generator 스킬

이 스킬은 전문적인 PowerPoint 프레젠테이션을 생성합니다.

## 사용 방법

사용자가 "SL템플릿으로 PPT 만들어줘" 또는 유사한 요청을 하면 이 스킬을 사용합니다.

## 작업 흐름

### 1. 요청 분석
- 사용자의 요청에서 PPT 주제, 내용, 대상 프로젝트를 파악
- 프로젝트 아키텍처 분석이 필요한 경우 해당 프로젝트 탐색

### 2. YAML 설정 파일 생성

기본 구조:
```yaml
settings:
  show_page_numbers: true
  theme_name: default
  card_style: "gradient"  # 9가지 스타일 중 선택

cover:
  title: "제목"
  date: "YYYY. MM. DD"
  author: "미래융합설계센터 알고리즘개발팀 강민규 선임"
  report_type: "정보공유"

slides:
  - type: content
    title: "슬라이드 제목"
    content:
      - "내용 1"
      - "내용 2"
```

### 3. 사용 가능한 슬라이드 타입

| 타입 | 설명 | 주요 속성 |
|------|------|-----------|
| `content` | 기본 텍스트 슬라이드 | title, content(리스트) |
| `content_boxed` | 소주제별 박스 구분 | title, sections, columns(1/2) |
| `cards` | 카드형 레이아웃 | title, cards, card_style |
| `comparison` | 좌우 비교 | left_title, left_items, right_title, right_items |
| `table` | 표 | headers, rows |
| `architecture` | 아키텍처 다이어그램 | components, connections |
| `flowchart` | 플로우차트 | flow_type, steps |
| `timeline` | 타임라인 | milestones, style |
| `tree` | 디렉토리 구조 | tree_structure |
| `chart` | 차트 | chart_type, categories, series |
| `two_column` | 2단 레이아웃 | left_content, right_content |
| `stats` | 통계 카드 | stats |

### 4. 카드 스타일 (card_style) - 9가지

| 스타일 | 설명 |
|--------|------|
| `classic` | [백업] 기존 디자인 - 좌측 컬러바 + 원형 아이콘 |
| `gradient` | **기본값** - 상단 컬러 헤더 + 사각 아이콘 |
| `modern` | 좌측 큰 원형 아이콘 + 연한 배경 |
| `solid` | 전체 컬러 배경 + 골드 제목 |
| `outline` | 두꺼운 테두리 강조 + 테두리 아이콘 |
| `minimal` | 미니멀 - 하단 컬러 라인 + 작은 아이콘 |
| `banner` | 상단 배너 + 튀어나온 큰 아이콘 |
| `split` | 상단 40% 컬러 / 하단 60% 화이트 |
| `accent` | 좌측 두꺼운 악센트 바 + 큰 아이콘 |

**전역 설정:**
```yaml
settings:
  card_style: "gradient"
```

**슬라이드별 오버라이드:**
```yaml
- type: cards
  card_style: "banner"
  cards: [...]
```

### 5. 콘텐츠 작성 규칙

#### 데이터베이스 스키마 표현

DB 스키마/테이블 작성 시 **반드시 테이블 또는 박스**로 시각적 구분:

```yaml
- type: table
  title: "users 테이블"
  headers: ["필드명", "타입", "설명"]
  rows:
    - ["id", "INT", "기본키"]
    - ["name", "VARCHAR(100)", "사용자 이름"]
```

#### 핵심 내용 분석

**주제만 나열하지 않고, 주제 아래에 핵심 포인트 함께 작성:**

```yaml
- type: content_boxed
  title: "핵심 분석"
  sections:
    - title: "성능 최적화"
      items:
        - "쿼리 응답시간 50% 단축"
        - "캐싱 레이어 도입"
```

### 6. PPT 생성

```bash
python ppt_generator.py -c [config.yaml] -o output/[output.pptx]
```

**주의: 항상 output 폴더에 PPT를 생성합니다.**

## 색상 팔레트 (ref_1.pptx 기반)

| 색상명 | HEX | 용도 |
|--------|-----|------|
| primary | #28374E | 다크 네이비 - 기본/헤더 |
| secondary | #4F81BD | 미드 블루 - 강조 |
| accent | #1F497D | 딥 블루 |
| content_box | #E8EDF4 | 콘텐츠 박스 배경 |
| success | #35A29F | 티일 그린 |
| warning | #FFA76D | 소프트 오렌지 |
| danger | #ED6666 | 소프트 레드 |
| text | #333333 | 메인 텍스트 |

## 폰트

- 대주제: 현대하모니 M (20pt, 굵게)
- 섹션 헤더: 현대하모니 M (14pt, ● 불릿)
- 카드 제목: 현대하모니 M (14pt)
- 본문: 현대하모니 L (11-12pt)
- 표지 제목: 현대하모니 M (44pt, 굵게)

## 참고 파일

- 설정 예시: `example_config.yaml`
- 아키텍처 예시: `sl_sw_document_generate_arch.yaml`
- 카드 스타일 테스트: `test_card_styles.yaml`
- **디자인 참조**: `templates/ref_1.pptx`
