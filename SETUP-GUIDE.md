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

### 테스트 변환

1. GUI에서 `samples/sample_document.docx` 선택
2. 출력 위치 선택
3. [변환 실행] 클릭

---

## 6. 폴더 구조

```
docx-to-html-converter/
├── src/
│   ├── main.py          # 진입점 (실행 파일)
│   ├── gui.py           # GUI 화면
│   ├── converter.py     # 변환 로직 (핵심)
│   └── utils.py         # 유틸리티 함수
├── config.json          # 변환 설정 (스타일 매핑 등)
├── templates/           # HTML 템플릿
├── samples/             # 테스트용 샘플 문서
├── packages/            # 오프라인 설치용 패키지
├── requirements.txt     # 의존성 목록
└── install.bat          # 설치 스크립트
```

---

## 7. 주요 수정 포인트

| 파일 | 역할 | 수정 시점 |
|------|------|----------|
| `config.json` | 스타일→태그 매핑 규칙 | 새로운 Word 스타일 추가 시 |
| `src/converter.py` | 변환 로직 | 변환 규칙 변경 시 |
| `src/utils.py` | 공통 함수 | 유틸리티 추가 시 |

---

## 8. 출력 구조

변환 시 이미지는 문서별 독립 폴더에 저장됩니다:

```
output/
├── introduction.html
├── introduction_images/
│   └── image_xxx.png
├── chapter01.html
└── chapter01_images/
    └── image_yyy.png
```

---

## 9. 문제 해결

### "pip를 찾을 수 없습니다"

Python이 PATH에 등록되지 않은 경우:

```cmd
C:\Python311\python.exe -m pip install --no-index --find-links=./packages -r requirements.txt
```

### "lxml 설치 실패"

Python 버전이 3.11.x인지 확인:

```cmd
python --version
```

다른 버전이면 해당 버전용 wheel 파일이 필요합니다.
