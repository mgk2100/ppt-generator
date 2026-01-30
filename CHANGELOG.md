# PPT Generator 변경 이력

## 프로젝트 개요
SL 템플릿 기반 PowerPoint 프레젠테이션 자동 생성 도구

---

## v2.0.0 (2025-01-28)

### 폰트 시스템 개선
- **현대하모니 폰트 적용**
  - 제목: 현대하모니 M (Medium)
  - 본문: 현대하모니 L (Light)
- **자동 폰트 설치 기능 추가**
  - `install_fonts()`: 시스템에 폰트 자동 설치
  - `check_fonts_installed()`: 폰트 설치 여부 확인
  - Linux, macOS, Windows 지원
- **폰트 파일 포함** (`fonts/` 폴더)
  - 현대하모니 B.ttf
  - 현대하모니 M.ttf
  - 현대하모니 L.ttf

### 표지 슬라이드 개선
- 주제: 현대하모니 M, 48pt, 굵게
- 날짜: 현대하모니 L, 24pt
- 작성자: 현대하모니 L, 24pt, 굵게 (고정: "미래융합설계센터 알고리즘개발팀 강민규 선임")
- 보고 유형: 현대하모니 L, 14pt (박스 제거, 텍스트만 표시)

### 새로운 슬라이드 타입

#### `content_boxed` - 소주제별 박스 구분 슬라이드
```yaml
- type: content_boxed
  title: "슬라이드 제목"
  columns: 2  # 1 또는 2
  sections:
    - title: "소주제"
      color: "primary"
      items:
        - "항목 1"
        - "항목 2"
```

**디자인 특징:**
- 소주제 제목: 악센트 색상 배경 + 흰색 텍스트 + 그림자
- 각 항목: 개별 박스 (흰색 배경 + 연한 테두리 + 그림자)
- 좌측 악센트 바로 시각적 강조
- 테두리 색상: 각 악센트 색상에 어울리는 연한 톤

### 헬퍼 함수 추가
- `_add_shadow_box()`: 그림자 효과가 있는 박스 생성
  - 커스텀 그림자 오프셋
  - 둥근 모서리 옵션
  - 커스텀 테두리/배경 색상

### 출력 경로 자동 처리
- 파일명만 입력 시 자동으로 `output/` 폴더에 저장
- 예: `-o result.pptx` → `output/result.pptx`

---

## v1.5.0 (2025-01-27)

### 카드 슬라이드 디자인 개선
- 프로페셔널 카드 디자인
- 원형 아이콘 배경
- 좌측 컬러 바
- 그림자 효과

### 아키텍처 다이어그램 개선
- 자동 스케일링으로 슬라이드 범위 내 유지
- `SLIDE_BOUNDS` 상수 추가
- 컴포넌트/연결선 자동 조정

### 플레이스홀더 제거 기능
- `_clear_unused_placeholders()`: 사용하지 않는 플레이스홀더 자동 제거
- "마스터 텍스트 스타일 편집" 등 기본 텍스트 제거

---

## v1.0.0 (2025-01-27)

### 초기 릴리즈
- 템플릿 기반 PPT 생성
- YAML/JSON 설정 파일 지원
- 다양한 슬라이드 타입:
  - `content`: 기본 텍스트
  - `cards`: 카드형 레이아웃
  - `comparison`: 좌우 비교
  - `table`: 표
  - `architecture`: 아키텍처 다이어그램
  - `flowchart`: 플로우차트
  - `timeline`: 타임라인
  - `chart`: 차트 (column, bar, line, pie)
  - `org_chart`: 조직도
  - `two_column`: 2단 레이아웃
  - `stats`: 통계 카드
  - `image`: 이미지

### 테마 시스템
- 사전 정의 테마: default, dark, green, purple, warm
- 커스텀 테마 파일 지원

---

## 사용 가능한 색상

| 색상명 | RGB | 용도 |
|--------|-----|------|
| primary | (0, 51, 102) | 주요 강조 (진한 네이비) |
| secondary | (0, 112, 192) | 보조 강조 (파랑) |
| accent | (68, 114, 196) | 중간 강조 |
| success | (0, 128, 0) | 성공/완료 (녹색) |
| warning | (255, 128, 0) | 주의 (주황) |
| danger | (192, 0, 0) | 위험/에러 (빨강) |
| highlight | (255, 192, 0) | 특별 강조 (골드) |

---

## 파일 구조

```
ppt-generator/
├── ppt_generator.py      # 메인 생성기 (130KB+)
├── requirements.txt      # 의존성
├── example_config.yaml   # 설정 예시
├── test_phase3.yaml      # 테스트 설정
├── CHANGELOG.md          # 변경 이력 (이 파일)
├── fonts/                # 현대하모니 폰트
│   ├── 현대하모니 B.ttf
│   ├── 현대하모니 M.ttf
│   └── 현대하모니 L.ttf
├── templates/            # PPT 템플릿
│   └── 표지.pptx
└── output/               # 생성된 PPT 저장
```

---

## 사용 예시

### 기본 사용
```bash
cd /home/ubuntu/Share/ppt-generator
python ppt_generator.py -c config.yaml -o output.pptx
```

### PDF 동시 생성
```bash
python ppt_generator.py -c config.yaml -o output.pptx --pdf
```

### 테마 지정
```bash
python ppt_generator.py -c config.yaml -o output.pptx --theme dark
```

---

## Claude Code 스킬

`/sl-ppt` 명령으로 PPT 생성 가능:
- 스킬 위치: `~/.claude/skills/sl-ppt/SKILL.md`
- 사용법: "SL템플릿으로 PPT 만들어줘" 또는 `/sl-ppt [주제]`
