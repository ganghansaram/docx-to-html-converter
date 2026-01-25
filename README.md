# DOCX to HTML Converter

Word 문서(.docx)를 웹북용 HTML로 변환하는 도구입니다.

## 특징

- Python 설치 없이 EXE 파일로 즉시 실행
- 간단한 GUI 제공
- 설정 파일(config.json)로 변환 규칙 커스터마이징
- 이미지 자동 추출
- 기술문서 특화 (WARNING/CAUTION/NOTE 박스 지원)

## 폴더 구조

```
docx-to-html-converter/
├── converter.exe          # 실행 파일
├── config.json            # 변환 설정 (수정 가능)
├── templates/
│   └── output-template.html   # 출력 템플릿 (수정 가능)
├── output/                # 변환 결과 저장 폴더
└── src/                   # 소스 코드
    ├── main.py
    ├── converter.py
    └── gui.py
```

## 사용 방법

1. `converter.exe` 실행
2. 변환할 Word 파일 선택
3. 출력 위치 선택
4. 옵션 설정
5. [변환 실행] 클릭

## 설정 파일 (config.json)

변환 규칙을 수정하려면 `config.json` 파일을 편집하세요.

### 주요 설정

- `heading_styles`: 제목으로 인식할 Word 스타일 이름
- `special_blocks`: WARNING/CAUTION/NOTE로 인식할 키워드
- `options`: 변환 옵션

## 빌드 방법

```cmd
pip install pyinstaller python-docx
pyinstaller --onefile --windowed src/main.py -n converter
```

## 요구사항 (개발 환경)

- Python 3.8+
- python-docx
- tkinter (Python 기본 포함)
