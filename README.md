# TD Scanner 🔍

Excel 파일에서 특정 문자열을 빠르게 검색하는 GUI 기반 스캐너 프로그램

## ✨ 주요 기능

### 🎨 다양한 테마
- **10가지 테마** 지원 (모던, 레트로, 다크모드 등)
- 시스템 다크모드 **자동 감지**
- 테마별 맞춤 폰트 및 색상

### ⚡ 고성능 검색
- **멀티프로세싱** 기반 병렬 처리 (최대 4배 속도 향상)
- **조기 종료** 최적화로 불필요한 검색 스킵
- `read_only` 모드로 메모리 사용량 감소

### 🔧 고급 검색 옵션
- ✅ **대소문자 구분** 검색
- ✅ **정규식(Regex)** 패턴 매칭 지원
- ✅ **여러 타겟** 동시 검색

### 📊 사용자 친화적 UI
- **실시간 프로그레스 바**
- **스캔 취소** 버튼
- 검색 결과 **파일별 정리**
- 타겟별 **매칭 통계**

## 🖥️ 스크린샷

### 테마 예시
- **Clean Studio**: 깔끔한 미니멀 디자인
- **Dark Modern**: 다크모드 지원
- **Modern Y2K**: 세련된 Y2K 감성
- **Lavender Dream**: 부드러운 라벤더 테마

## 📦 설치 방법

### Option 1: 실행 파일 (.exe)
1. [Releases](https://github.com/CheHyeonYeong/TDScanner/releases)에서 최신 버전 다운로드
2. `TDScanner.exe` 실행

### Option 2: Python 소스코드
```bash
# 저장소 클론
git clone https://github.com/CheHyeonYeong/TDScanner.git
cd TDScanner

# 가상환경 생성 (선택사항)
python -m venv .venv
.venv\Scripts\activate  # Windows
# source .venv/bin/activate  # Mac/Linux

# 필요한 패키지 설치
pip install openpyxl

# 실행
python main.py
```

## 🚀 사용 방법

1. **Search Targets** 입력
   - 검색할 문자열 입력
   - `+ Add Target` 버튼으로 여러 타겟 추가 가능

2. **Scan Directory** 설정
   - `Browse` 버튼으로 검색할 디렉토리 선택

3. **검색 옵션** 선택 (선택사항)
   - `Case Sensitive`: 대소문자 구분
   - `Use Regex`: 정규식 패턴 사용

4. **Start Scan** 클릭
   - 진행률 바로 진행 상황 확인
   - 필요시 `Cancel` 버튼으로 중단

## 🔨 빌드 방법

```bash
# PyInstaller 설치
pip install pyinstaller

# exe 파일 생성
pyinstaller --onefile --windowed --name "TDScanner" main.py

# 생성된 파일 위치
# dist/TDScanner.exe
```

## 🎯 정규식 예시

```
# 패턴 예시
org.*Detail          # "org"로 시작하고 "Detail" 포함
^TD\d+              # "TD" 뒤에 숫자
[A-Z]{3}\d{4}       # 대문자 3자 + 숫자 4자
(certDetail|empCert) # "certDetail" 또는 "empCert"
```

## 🛠️ 기술 스택

- **Python 3.10+**
- **tkinter**: GUI 프레임워크
- **openpyxl**: Excel 파일 처리
- **multiprocessing**: 병렬 처리
- **re**: 정규식 지원
- **PyInstaller**: 실행 파일 빌드

## 📋 시스템 요구사항

- **OS**: Windows 10/11 (다크모드 자동 감지 지원)
- **Python**: 3.10 이상 (소스코드 실행 시)
- **메모리**: 최소 4GB RAM 권장

## 🎨 테마 목록

1. **Clean Studio** - 미니멀 화이트
2. **Modern Y2K** - 세련된 핑크
3. **Modern Minimal** - 그레이 미니멀
4. **Dark Modern** - 다크 모드
5. **Lavender Dream** - 라벤더 퍼플
6. **Mint Fresh** - 민트 그린
7. **Y2K Pink** - 레트로 핑크
8. **Cyber Purple** - 사이버 퍼플
9. **Retro Green** - 레트로 그린
10. **Neon Blue** - 네온 블루
11. **Sunset Orange** - 선셋 오렌지

## 📝 라이선스

MIT License

## 👤 개발자

- GitHub: [@CheHyeonYeong](https://github.com/CheHyeonYeong)

## 🤝 기여

이슈 제보 및 Pull Request 환영합니다!

---

🤖 Generated with [Claude Code](https://claude.com/claude-code)
