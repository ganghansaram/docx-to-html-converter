# 설치 및 실행 가이드 (폐쇄망 환경)

## 1. 압축 해제

받은 `docx-to-html-converter.zip` 파일을 원하는 위치에 압축 해제합니다.

```
예: C:\Projects\docx-to-html-converter\
```

---

## 2. 패키지 파일 복원 (중요!)

메일 전송을 위해 `.whl` 확장자가 `.whl_`로 변경되어 있습니다.

`packages` 폴더의 `restore.bat`을 더블클릭하여 복원하세요.

또는 수동으로:
```cmd
cd packages
ren *.whl_ *.whl
```

---

## 3. 의존성 설치

### 방법 A: 배치 파일 사용 (권장)

1. `install.bat` 더블클릭
2. "성공" 메시지 확인
3. 아무 키나 눌러 종료

### 방법 B: 수동 설치

명령 프롬프트(cmd)에서:

```cmd
cd C:\Projects\docx-to-html-converter
pip install --no-index --find-links=./packages -r requirements.txt
```

### 필요 패키지

| 패키지 | 용도 |
|--------|------|
| python-docx | DOCX 문서 파싱 |
| PyMuPDF | PDF 문서 파싱, 이미지/테이블 추출 |

---

## 4. PyCharm에서 프로젝트 열기

1. PyCharm 실행
2. `File` → `Open` → 압축 해제한 폴더 선택
3. 가상환경 생성 여부 묻는 창이 뜨면:
   - **"OK"** 클릭 (가상환경 생성)
   - 또는 **"Cancel"** 후 시스템 Python 사용

### 가상환경을 만든 경우

PyCharm 하단 터미널에서 다시 의존성 설치:

```cmd
pip install --no-index --find-links=./packages -r requirements.txt
```

---

## 5. 프로그램 실행

### GUI 모드 (권장)

PyCharm에서 `src/main.py` 열고 실행 (Shift+F10)

또는 터미널에서:

```cmd
python src/main.py
```

### CLI 모드

```cmd
# 단일 파일 변환
python src/main.py document.docx
python src/main.py document.pdf -o output/document.html

# 폴더 배치 변환 (하위 폴더 포함)
python src/main.py ./docs/ -r -o ./output/

# 문서 구조 분석만 수행
python src/main.py document.pdf --analyze
```

### 테스트 변환

1. GUI에서 `samples/` 폴더의 .docx 또는 .pdf 파일 선택
2. 출력 위치 선택
3. [변환 실행] 클릭
4. PDF의 경우 `_report.txt` 리포트에서 매칭 정확도 확인

---

## 6. 폴더 구조

```
docx-to-html-converter/
├── src/
│   ├── main.py            # 진입점 (GUI/CLI)
│   ├── gui.py             # Tkinter GUI
│   ├── converter.py       # DOCX 변환 로직
│   ├── pdf_converter.py   # PDF 변환 로직 (TOC 기반)
│   └── utils.py           # 유틸리티 함수
├── config.json            # 변환 설정 (스타일 매핑, PDF 옵션)
├── samples/               # 테스트용 샘플 문서
├── packages/              # 오프라인 설치용 패키지
├── requirements.txt       # 의존성 목록
└── install.bat            # 설치 스크립트
```

---

## 7. 주요 수정 포인트

| 파일 | 역할 | 수정 시점 |
|------|------|----------|
| `config.json` | 스타일→태그 매핑 규칙, PDF 설정 | 새로운 스타일 추가 시 |
| `src/converter.py` | DOCX 변환 로직 | DOCX 변환 규칙 변경 시 |
| `src/pdf_converter.py` | PDF 변환 로직 | PDF 변환 규칙 변경 시 |
| `src/utils.py` | 공통 함수 | 유틸리티 추가 시 |

---

## 8. 출력 구조

변환 시 이미지는 문서별 독립 폴더에 저장됩니다:

```
output/
├── document.html
├── document_images/
│   ├── image_abc123.png
│   └── image_def456.jpg
└── document_report.txt      ← PDF 변환 시 매칭 리포트
```

### PDF 매칭 리포트 예시

```
[변환 리포트] FY1-DoD-FLEX-4.pdf
──────────────────────────────────────────────────
✓ h1: "1. Overview" (p.4) → 매칭 (유사도: 1.00)
✓ h2: "3.1 Aggregated Funding" (p.9) → 매칭 (유사도: 1.00)
✗ h2: "3.3 Implementation Challenges" (p.13) → 매칭 실패
   유사 후보: "3.3 Implementation Challenges and..." (p.13, 0.82)
- [건너뜀] "Figure 1 – Breakdown..." (Figure/Table 참조)
──────────────────────────────────────────────────
섹션 제목: 10개 | 매칭 성공: 9 (90.0%) | 실패: 1
```

---

## 9. 문제 해결

### "pip를 찾을 수 없습니다"

Python이 PATH에 등록되지 않은 경우:

```cmd
C:\Python311\python.exe -m pip install --no-index --find-links=./packages -r requirements.txt
```

### "lxml 설치 실패" 또는 "PyMuPDF 설치 실패"

Python 버전이 3.11.x인지 확인:

```cmd
python --version
```

다른 버전이면 해당 버전용 wheel 파일이 필요합니다.

### PDF 변환 시 "TOC를 찾을 수 없습니다" 경고

문서에 목차(Table of Contents) 페이지가 없는 경우 발생합니다.
이 경우 모든 텍스트가 본문(`<p>`)으로 처리됩니다.
`config.json`의 `pdf.toc_keywords`에 해당 문서의 목차 키워드를 추가하면 해결될 수 있습니다.
