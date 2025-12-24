# HTS 자동화 프로그램 - 사용 가이드

10개 증권사의 증거금률 및 신용/대출 가능 종목 정보를 자동으로 수집하고 통합하는 시스템입니다.

## 🎯 주요 기능

- 10개 증권사 HTS 자동 로그인 및 데이터 다운로드
- 증거금률별 종목 조회 (20%, 30%, 40%, 50%, 60%, 100%)
- 신용가능/대출가능 종목 조회
- 데이터 자동 통합 및 Excel 파일 생성
- 로깅 시스템 (실행 로그 자동 저장)
- 에러 발생 시 스크린샷 자동 저장

## 📋 지원 증권사

1. 키움증권
2. KB증권
3. 신한투자증권
4. 대신증권
5. 삼성증권
6. 한국투자증권
7. 하나증권
8. 메리츠증권
9. 미래에셋증권
10. NH투자증권

## 🚀 설치 및 설정

### 1단계: 폴더 구조 생성

Windows PowerShell 또는 CMD에서 실행:

```cmd
mkdir C:\HTS_Automation
cd C:\HTS_Automation
mkdir config images logs data backup
cd images
mkdir 키움 KB 신한 대신 삼성 한투 하나 메리츠 미래 NH
```

### 2단계: 파일 복사

Git 저장소의 파일들을 `C:\HTS_Automation`로 복사:

- `config/hts_config.json` → `C:\HTS_Automation\config\`
- `config/hts_info.json` → `C:\HTS_Automation\config\`
- `hts_utils.py` → `C:\HTS_Automation\`
- `auto_mouse_total_25.12.24.ipynb` → `C:\HTS_Automation\`

### 3단계: 계정 정보 입력

`C:\HTS_Automation\config\hts_info.json` 파일을 열어서 각 증권사 계정 정보 입력:

```json
{
  "Kiwoom": {
    "id": "실제_키움_ID",
    "password": "실제_키움_비밀번호"
  },
  "KB": {
    "id": "실제_KB_ID",
    "password": "실제_KB_비밀번호"
  },
  ...
}
```

**⚠️ 주의**: 이 파일은 보안에 주의하세요! Git에 커밋하지 마세요.

### 4단계: Python 패키지 설치

```cmd
pip install pyautogui pyperclip pywin32 numpy opencv-python Pillow pandas openpyxl pykrx
```

## 📁 폴더 구조

```
C:\HTS_Automation\
├── config\
│   ├── hts_config.json      # HTS 설정 (경로, 화면번호 등)
│   └── hts_info.json         # 계정 정보 (ID, 비밀번호)
├── images\                   # 화면 인식용 이미지 파일
│   ├── 키움\
│   ├── KB\
│   └── ...
├── logs\                     # 실행 로그 및 에러 스크린샷
├── data\                     # 다운로드된 데이터 파일
├── backup\                   # 백업 파일
├── hts_utils.py             # 유틸리티 모듈
└── auto_mouse_total_25.12.24.ipynb  # 메인 노트북
```

## 💡 사용 방법

### Jupyter Notebook 실행

1. Jupyter Notebook 실행:
   ```cmd
   cd C:\HTS_Automation
   jupyter notebook
   ```

2. `auto_mouse_total_25.12.24.ipynb` 파일 열기

3. 셀 실행:
   - **Cell 0**: 함수 및 클래스 정의
   - **Cell 1**: `run()` - 10개 증권사 데이터 다운로드
   - **Cell 2**: `total_output_save()` - 데이터 통합
   - **Cell 3**: `check_output()` - 결과 검증

### 테스트 모드 (특정 증권사만 실행)

```python
# 키움만 테스트
run(test_mode=True, test_securities=['키움'])

# 3개 증권사만 실행
run(test_mode=True, test_securities=['키움', 'KB', '신한'])
```

## 📊 출력 파일

### 증권사별 원본 파일
`C:\Users\kiwoom\Documents\`에 저장:
- 증거금20종목키움_2025.12.24.csv
- 증거금30종목키움_2025.12.24.csv
- ...

### 통합 파일
`C:\Users\kiwoom\Documents\`에 저장:
- total_output_2025.12.24.xlsx

컬럼:
- 종목코드, 종목명, 종목구분
- 각 증권사별: 증거금률, 신용가능여부, 대출가능여부

## 🔧 설정 파일 수정

### HTS 경로 변경

`C:\HTS_Automation\config\hts_config.json` 파일 수정:

```json
{
  "Kiwoom": {
    "실행경로": "C:\\KiwoomHero4\\bin\\nkstarter.exe",  // 여기 수정
    ...
  }
}
```

### 화면 번호 변경

증권사 HTS 화면이 변경되면 설정 파일에서 수정:

```json
{
  "Kiwoom": {
    "화면번호": ["0191"],  // 여기 수정
    ...
  }
}
```

## 📝 로그 확인

실행 로그는 `C:\HTS_Automation\logs\`에 자동 저장:
- `2025-12-24.log` - 실행 로그
- `키움_error_20251224_103015.png` - 에러 스크린샷

## ⚠️ 주의사항

1. **실행 중 PC 사용 금지**: 마우스/키보드 자동화가 진행되므로 방해하지 마세요
2. **화면 해상도 고정**: 마우스 좌표가 고정되어 있어 해상도 변경 시 오류 발생
3. **절전 모드 해제**: 작업 중 화면이 꺼지지 않도록 설정
4. **안정적인 인터넷**: HTS 로그인 및 데이터 조회에 필요
5. **인증서 유효기간**: 공인인증서 만료 여부 확인

## 🐛 문제 해결

### Q: HTS가 실행되지 않아요
A: `config/hts_config.json`에서 실행 파일 경로를 확인하세요

### Q: 이미지를 찾을 수 없다는 오류가 나요
A: 해당 증권사의 이미지 파일을 캡처해서 `images/증권사명/` 폴더에 저장하세요

### Q: 로그인이 실패해요
A: `config/hts_info.json`에서 ID/비밀번호를 확인하세요

### Q: 특정 증권사만 다시 실행하고 싶어요
A: 테스트 모드를 사용하세요: `run(test_mode=True, test_securities=['증권사명'])`

## 📞 문의

문제가 발생하면 `logs/` 폴더의 로그 파일과 스크린샷을 확인하세요.
