# DOCX to HTML Converter

Word 문서(.docx)를 웹북용 HTML로 변환하는 도구입니다.

## 특징

- 간편한 GUI 제공
- 단일 파일 / 배치(폴더) 변환 지원
- 설정 파일(config.json)로 변환 규칙 커스터마이징
- 이미지 자동 추출 (문서별 독립 폴더)
- 기술문서 특화 (WARNING/CAUTION/NOTE 박스 지원)

## 폴더 구조

```
docx-to-html-converter/
├── src/
│   ├── main.py            # 진입점 (실행 파일)
│   ├── gui.py             # GUI 화면
│   ├── converter.py       # 변환 로직
│   └── utils.py           # 유틸리티 함수
├── config.json            # 변환 설정 (수정 가능)
├── templates/
│   └── output-template.html
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
├── introduction.html
├── introduction_images/
│   ├── image_abc123.png
│   └── image_def456.png
├── chapter01.html
└── chapter01_images/
    └── image_789ghi.png
```

**규칙**: `{문서명}_images/` 형식으로 HTML 파일과 1:1 매핑됩니다.

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

```cmd
python src/main.py
```

1. 변환할 Word 파일 또는 폴더 선택
2. 출력 위치 선택
3. 옵션 설정 (이미지 추출, 빈 문단 제거)
4. [변환 실행] 클릭

## 설정 파일 (config.json)

변환 규칙을 수정하려면 `config.json` 파일을 편집하세요.

### 주요 설정

- `style_mapping.by_style`: Word 스타일명 → HTML 태그 매핑
- `style_mapping.by_font_size`: 폰트 크기 → HTML 태그 매핑
- `special_blocks`: WARNING/CAUTION/NOTE로 인식할 키워드
- `options`: 변환 옵션

## 요구사항

- Python 3.11+
- python-docx
- tkinter (Python 기본 포함)
