# Document to HTML Converter

Word(.docx) 및 PDF 문서를 웹북용 HTML로 변환하는 도구입니다.

## 특징

- **DOCX 변환**: Word 스타일/폰트 크기 기반 제목 자동 감지 (h1~h6)
- **PDF 변환**: 문서 내 목차(TOC)를 파싱하여 제목 구조 판별
- 간편한 GUI 및 CLI 지원
- 단일 파일 / 배치(폴더) 변환
- 이미지 자동 추출 (문서별 독립 폴더)
- 테이블 자동 변환 (`<thead>/<tbody>` 구조)
- 기술문서 특화 (WARNING/CAUTION/NOTE 박스 지원)
- PDF 변환 시 매칭 정확도 리포트 자동 생성
- 폐쇄망 환경 지원 (오프라인 설치 가능)

## 폴더 구조

```
docx-to-html-converter/
├── src/
│   ├── main.py            # 진입점 (GUI/CLI)
│   ├── gui.py             # Tkinter GUI
│   ├── converter.py       # DOCX 변환 로직
│   ├── pdf_converter.py   # PDF 변환 로직 (TOC 기반)
│   └── utils.py           # 유틸리티 함수
├── config.json            # 변환 설정 (수정 가능)
├── samples/               # 테스트용 샘플 문서
├── packages/              # 오프라인 설치용 패키지
├── requirements.txt       # 의존성 목록
├── install.bat            # 설치 스크립트
└── SETUP-GUIDE.md         # 설치 가이드
```

## 출력 구조

변환된 HTML과 이미지는 문서별로 독립된 폴더에 저장됩니다.

```
output/
├── document.html
├── document_images/
│   ├── image_abc123.png
│   └── image_def456.jpg
└── document_report.txt      ← PDF 변환 시 매칭 리포트
```

## 설치 방법

### 온라인 환경

```cmd
pip install -r requirements.txt
```

### 오프라인/폐쇄망 환경

`install.bat` 더블클릭 또는:

```cmd
pip install --no-index --find-links=./packages -r requirements.txt
```

자세한 내용은 `SETUP-GUIDE.md` 참고

## 사용 방법

### GUI 모드 (기본)

```cmd
python src/main.py
```

1. 변환할 문서 파일(.docx / .pdf) 또는 폴더 선택
2. 출력 위치 선택
3. 옵션 설정 (이미지 추출, 빈 문단 제거)
4. [변환 실행] 클릭

### CLI 모드

```cmd
# 단일 파일 변환
python src/main.py document.pdf
python src/main.py report.docx -o output/report.html

# 폴더 배치 변환
python src/main.py ./docs/ -r -o ./output/

# 문서 구조 분석
python src/main.py document.pdf --analyze
```

## 설정 파일 (config.json)

### DOCX 관련 설정

| 항목 | 설명 |
|------|------|
| `style_mapping.by_style` | Word 스타일명 → HTML 태그 매핑 |
| `style_mapping.by_font_size` | 폰트 크기 → HTML 태그 매핑 |
| `special_blocks` | WARNING/CAUTION/NOTE 키워드 |
| `options` | 변환 옵션 (이미지 추출, 빈 문단 제거 등) |

### PDF 관련 설정

| 항목 | 설명 |
|------|------|
| `pdf.toc_keywords` | TOC 페이지 탐지 키워드 |
| `pdf.non_heading_prefixes` | 제목에서 제외할 접두사 (Figure, Table 등) |
| `pdf.matching.fuzzy_threshold` | 퍼지 매칭 임계값 (기본 0.80) |
| `pdf.options.generate_report` | 매칭 리포트 생성 여부 |

## PDF 변환 알고리즘

1. 문서 앞부분에서 목차(TOC) 페이지 자동 탐지
2. TOC 항목 파싱 → 번호 패턴으로 제목 레벨(h1~h6) 결정
3. 본문 텍스트와 TOC 항목을 3단계 매칭:
   - 정규화 접두사 매칭 → 인접 블록 연결 → 퍼지 매칭
4. 매칭된 텍스트를 `<h1>`~`<h6>` 태그로 변환
5. 매칭 정확도 리포트(`_report.txt`) 자동 생성

## 요구사항

- Python 3.11+
- python-docx (DOCX 처리)
- PyMuPDF (PDF 처리)
- tkinter (Python 기본 포함)
