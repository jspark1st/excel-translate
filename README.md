# 엑셀 파일 번역 프로그램

구글 번역 API를 활용하여 엑셀 파일의 내용을 한국어로 번역하는 GUI 프로그램입니다.

## 주요 기능

- 📄 엑셀 파일 읽기: 모든 시트 자동 처리
- 🌐 구글 번역 API 연동: 한국어로 자동 번역
- 💾 원본 형식 유지: 숫자, 빈 셀은 그대로 보존
- 📊 진행 상황 표시: 번역 진행률 실시간 표시
- 🛡️ 에러 처리: 번역 실패 시 원문 유지
- 🐛 상세한 디버깅: 로그 파일 자동 생성 및 상세한 오류 추적
- ⏱️ 성능 측정: 각 단계별 소요 시간 측정
- 📈 통계 정보: 번역된 셀 수, 건너뛴 셀 수, 오류 수 등 상세 통계

## 설치 방법

### 1. 저장소 클론
```bash
git clone https://github.com/jspark1st/excel-translate.git
cd excel-translate
```

### 2. 가상 환경 생성 및 활성화
```bash
# Windows PowerShell
python -m venv venv
.\venv\Scripts\Activate.ps1

# Windows CMD
python -m venv venv
venv\Scripts\activate.bat

# Linux/Mac
python -m venv venv
source venv/bin/activate
```

### 3. 패키지 설치
```bash
pip install -r requirements.txt
```

## 사용 방법

### GUI 프로그램 실행

#### 방법 1: 배치 파일 사용 (Windows)
```bash
run.bat
```

#### 방법 2: PowerShell 스크립트 사용
```bash
.\run.ps1
```

#### 방법 3: 직접 실행
```bash
python gui_translate.py
```

### 명령줄 사용 (CLI)
```bash
python translate_excel.py <엑셀파일경로> [출력파일경로]
```

예시:
```bash
python translate_excel.py input.xlsx
python translate_excel.py input.xlsx output.xlsx
```

## 프로그램 구조

```
excel-translate/
├── gui_translate.py      # GUI 버전 번역 프로그램
├── translate_excel.py    # CLI 버전 번역 프로그램
├── logger_config.py      # 로깅 설정 모듈
├── requirements.txt      # 필요한 패키지 목록
├── run.bat              # Windows 실행 스크립트
├── run.ps1              # PowerShell 실행 스크립트
├── .gitignore           # Git 제외 파일 목록
├── logs/                # 로그 파일 저장 폴더 (자동 생성)
└── README.md            # 프로젝트 설명서
```

## 필요한 패키지

- `pandas` - 엑셀 파일 처리
- `openpyxl` - 엑셀 파일 읽기/쓰기
- `googletrans` - 구글 번역 API

## 디버깅 기능

프로그램은 개발 단계를 위해 상세한 디버깅 기능을 제공합니다:

### 로그 파일
- 모든 작업이 `logs/` 폴더에 자동으로 기록됩니다
- 로그 파일명: `{프로그램명}_{날짜}.log` (예: `translate_excel_20250101.log`)
- 로그 파일은 최대 10MB까지 저장되며, 5개까지 보관됩니다

### 로그 레벨
- **DEBUG**: 상세한 디버깅 정보 (모든 작업 단계)
- **INFO**: 일반 정보 (진행 상황, 통계)
- **WARNING**: 경고 메시지
- **ERROR**: 오류 발생 시 상세 정보
- **CRITICAL**: 치명적 오류

### 기록되는 정보
- 프로그램 시작/종료 시간
- 파일 경로 및 크기
- 각 시트/컬럼 처리 시간
- 번역된 텍스트 (원문 → 번역문)
- 오류 발생 시 전체 스택 트레이스
- 성능 통계 (소요 시간, 처리 셀 수 등)

### 통계 정보
프로그램 실행 후 다음 정보를 제공합니다:
- 총 소요 시간
- 총 처리 셀 수
- 번역된 셀 수
- 건너뛴 셀 수 (숫자, 빈 셀, 이미 한글 포함 등)
- 번역 오류 수

## 주의사항

- 인터넷 연결이 필요합니다 (구글 번역 API 사용)
- 대용량 파일의 경우 번역에 시간이 걸릴 수 있습니다
- API 제한으로 인해 일부 번역이 실패할 수 있습니다 (원문 유지)
- 로그 파일은 자동으로 생성되며, 디스크 공간을 차지할 수 있습니다

## 라이선스

이 프로젝트는 자유롭게 사용할 수 있습니다.

