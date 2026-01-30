# PPT Generator 프로젝트

SL 템플릿을 사용하여 전문적인 PowerPoint 프레젠테이션을 자동 생성하는 Python 프로젝트입니다.

## 프로젝트 구조

```
ppt-generator/
├── ppt_generator.py          # 메인 PPT 생성 모듈
├── output/                   # 생성된 PPT 출력 폴더
├── templates/                # PPT 템플릿 파일
└── .claude/
    ├── agents/               # Sub-agents 정의
    │   ├── orchestrator.md       # 마스터 에이전트
    │   ├── backend-architect.md  # 아키텍처 분석
    │   ├── code-analyzer.md      # 코드 상세 분석
    │   ├── python-backend.md     # 프로그래밍 구현
    │   ├── e2e-api-tester.md     # 스타일 테스트
    │   └── prompt-engineer.md    # 프롬프트/문서 작성
    ├── commands/             # 슬래시 명령어
    └── skills/sl-ppt/        # PPT 생성 스킬
```

## 에이전트 구조

| 에이전트 | 역할 |
|---------|------|
| `orchestrator` | 마스터 에이전트 - 업무 분배 및 조율 |
| `backend-architect` | 전체 아키텍처 분석 |
| `code-analyzer` | 코드 상세 분석 (함수, 클래스) |
| `python-backend` | 프로그래밍 구현만 담당 |
| `e2e-api-tester` | 스타일 테스트 및 품질 검증 |
| `prompt-engineer` | 프롬프트 및 문서 작성 |

## PPT 생성 가이드라인

### 지원 슬라이드 타입

| 타입 | 용도 | 우선순위 |
|------|------|---------|
| `cards` | 핵심 기능/개요 소개 | ★★★ |
| `flowchart` | 프로세스/데이터 흐름 | ★★★ |
| `architecture` | 시스템 구조 다이어그램 | ★★★ |
| `timeline` | 단계별 처리 흐름 (박스형) | ★★★ |
| `tree` | 디렉토리 구조 (tree 명령어 스타일) | ★★★ |
| `stats` | 핵심 수치/요약 | ★★★ |
| `comparison` | 좌우 비교 | ★★☆ |
| `two_column` | 2단 레이아웃 | ★★☆ |
| `table` | 정보 비교표 | ★★☆ |
| `chart` | 데이터 시각화 | ★★☆ |
| `content_boxed` | 박스형 콘텐츠 | ★☆☆ |
| `content` | 일반 텍스트 (최소화) | ★☆☆ |

### 카드 스타일 (card_style) - 9가지

| 스타일 | 설명 | 특징 |
|--------|------|------|
| `classic` | [백업] 기존 디자인 | 좌측 컬러바 + 상단 원형 아이콘 |
| `gradient` | **기본값** | 상단 컬러 헤더 + 좌측 사각 아이콘 |
| `modern` | 모던 좌측 아이콘형 | 좌측 큰 원형 아이콘 + 연한 배경 |
| `solid` | 전체 컬러 카드 | 배경 전체 컬러 + 골드 제목 |
| `outline` | 테두리 강조 | 두꺼운 컬러 테두리 + 테두리 아이콘 |
| `minimal` | 미니멀 | 하단 컬러 라인 + 좌상단 작은 아이콘 |
| `banner` | 배너 스타일 | 상단 배너 + 튀어나온 큰 아이콘 |
| `split` | 분할 카드 | 상단 컬러 40% / 하단 화이트 60% |
| `accent` | 악센트 바 강조 | 좌측 두꺼운 바 + 큰 원형 아이콘 |

**전역 설정:**
```yaml
settings:
  card_style: "gradient"  # 모든 카드에 적용
```

**슬라이드별 오버라이드:**
```yaml
- type: cards
  title: "특정 슬라이드"
  card_style: "banner"  # 이 슬라이드만 다른 스타일
  cards: [...]
```

### 폰트 크기 규칙 (ref_1.pptx 기준)

- **최소 9pt**: 작은 텍스트
- **대주제**: 20pt (슬라이드 제목)
- **섹션 헤더**: 14pt (● 불릿 마커)
- **카드 제목**: 14pt
- **본문 내용**: 11-12pt

### 색상 팔레트 (ref_1.pptx 기반 개선)

| 색상 | RGB (HEX) | 용도 |
|------|-----------|------|
| primary | (40, 55, 78) `#28374E` | 다크 네이비 - 기본/헤더 |
| secondary | (79, 129, 189) `#4F81BD` | 미드 블루 - 강조 |
| accent | (31, 73, 125) `#1F497D` | 딥 블루 |
| content_box | (232, 237, 244) `#E8EDF4` | 콘텐츠 박스 배경 |
| success | (53, 162, 159) `#35A29F` | 티일 그린 |
| warning | (255, 167, 109) `#FFA76D` | 소프트 오렌지 |
| danger | (237, 102, 102) `#ED6666` | 소프트 레드 |
| teal | (11, 102, 105) `#0B6669` | 다크 티일 |
| highlight | (255, 192, 0) | 골드 |
| text | (51, 51, 51) `#333333` | 메인 텍스트 |

### 참조 디자인

`templates/ref_1.pptx` 파일을 참조하여 품질 개선됨:
- 콘텐츠 박스 배경색 (#E8EDF4)
- 섹션 헤더 스타일 (● 불릿)
- 개선된 그림자 효과
- 타이포그래피 계층 구조

### PPT 구성 원칙

1. **개요**: cards 또는 stats로 시작
2. **구조/아키텍처**: architecture 다이어그램 사용
3. **프로세스/흐름**: flowchart 또는 timeline 사용
4. **비교/대조**: comparison 또는 two_column 사용
5. **요약**: stats로 핵심 수치 강조

### 콘텐츠 작성 규칙

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
    - title: "보안 강화"
      items:
        - "JWT 기반 인증 구현"
        - "SQL Injection 방어"
```

## 사용 방법

```bash
python ppt_generator.py -c [config.yaml] -o output/[output.pptx]
```

## 참고 파일

- 설정 예시: `example_config.yaml`
- 테스트 파일: `test_phase3.yaml`
- 아키텍처 예시: `sl_sw_document_generate_arch.yaml`
- 카드 스타일 테스트: `test_card_styles.yaml`

---

## 외부 프로젝트 분석 → PPT 생성 가이드

외부 프로젝트를 분석하여 PPT를 생성할 때 다음 절차를 따릅니다.

### 분석 파이프라인

```
1. 구조 파악 (backend-architect)
   ├── 디렉토리 구조
   ├── 파일 수/LOC 메트릭스
   └── 레이어/모듈 식별

2. 상세 분석 (code-analyzer)
   ├── 핵심 클래스/함수
   ├── 데이터 흐름 추적
   └── 기술 스택 적용 위치

3. PPT YAML 생성
   └── 슬라이드 타입 매핑
```

### 필수 수집 정보

| 정보 | 슬라이드 타입 | 비고 |
|------|-------------|------|
| 프로젝트 규모 | `stats` | 파일 수, LOC, 클래스 수 |
| 핵심 기능 | `cards` | 3-6개 카드 |
| 디렉토리 구조 | `tree` | 주요 폴더 + 역할 |
| 데이터 흐름 | `flowchart` | 입력 → 출력 |
| 시스템 아키텍처 | `architecture` | 레이어 구조 |
| 기술 스택 | `table` | 기술 + 적용 위치 |
| 디자인 패턴 | `cards` | Pipeline, Strategy 등 |

### 권장 슬라이드 순서

1. `title` - 프로젝트명 + 한줄 소개
2. `toc` - 목차 (선택)
3. `stats` - 프로젝트 규모
4. `cards` - 핵심 기능
5. `tree` - 디렉토리 구조
6. `table` - 디렉토리별 역할
7. `flowchart` - 전체 데이터 흐름
8. `timeline` - 상세 파이프라인
9. `architecture` - 시스템 아키텍처
10. `table` - 기술 스택
11. `cards` - 기술 적용 위치
12. `stats` - 코드 메트릭스
13. `cards` - 디자인 패턴
14. `cards` - 학습 포인트
15. `closing` - Q&A

### 분석 예시 (sl-sw-document-generate-python)

```yaml
# 프로젝트 규모 → stats
- type: stats
  stats:
    - value: "151"
      label: "Python 파일"
    - value: "50,337"
      label: "총 LOC"

# 디렉토리 구조 → tree
- type: tree
  root: "project/"
  items:
    - name: "routes/"
      description: "API 엔드포인트"
    - name: "orchestrator/"
      description: "파이프라인 실행"

# 데이터 흐름 → flowchart
- type: flowchart
  nodes:
    - id: "input"
      label: "소스코드"
    - id: "parse"
      label: "코드 파싱"
    - id: "llm"
      label: "LLM 분석"
    - id: "output"
      label: "문서"

# 기술 스택 → table
- type: table
  headers: ["기술", "적용 위치", "용도"]
  rows:
    - ["FastAPI", "routes/", "REST API"]
    - ["LangChain", "agent/", "LLM 통합"]
```
