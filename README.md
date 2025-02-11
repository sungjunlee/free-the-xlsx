# Excel Processor (엑셀 변환기)

엑셀 파일을 읽기 쉬운 형식(CSV, JSON, Markdown)으로 변환하는 도구입니다.

## 설치 방법

### Windows에서 설치하기

1. Python 설치 (처음 하시는 분)
   - [Python 공식 사이트](https://www.python.org/downloads/)에서 Python 3.8 이상 버전 다운로드
   - 설치 시 "Add Python to PATH" 옵션 반드시 체크!
   - 설치 완료 후 명령 프롬프트(cmd)를 열고 `python --version` 입력하여 설치 확인

2. 프로그램 설치
   ```bash
   # 1. 원하는 폴더에 프로그램 파일 다운로드
   # excel_processor.py와 requirements.txt를 같은 폴더에 저장

   # 2. 명령 프롬프트(cmd)를 관리자 권한으로 실행하여 해당 폴더로 이동
   cd C:\작업폴더경로

   # 3. 필요한 패키지 설치
   pip install -r requirements.txt
   ```

## 사용 방법

### 방법 1: 엑셀 파일을 프로그램 폴더로 복사

1. 변환하고 싶은 엑셀 파일을 `excel_processor.py`가 있는 폴더 안에 input으로 복사
2. 명령 프롬프트에서 해당 폴더로 이동
3. 변환 실행
   ```bash
   python excel_processor.py input output
   ```

### 방법 2: 다른 폴더의 엑셀 파일 변환

1. 명령 프롬프트에서 `excel_processor.py`가 있는 폴더로 이동
2. 엑셀 파일이 있는 폴더 경로를 지정하여 실행
   ```bash
   # Windows 예시
   python excel_processor.py "C:\Users\사용자이름\Desktop\엑셀파일들" "C:\Users\사용자이름\Desktop\결과"

   # 상대 경로 예시
   python excel_processor.py ..\엑셀파일들 ..\결과
   ```
## 실행 파일로 사용하기 (초보자용)

### Windows 사용자

1. [Releases](https://github.com/your-username/excel-processor/releases) 페이지에서 최신 버전의 `excel_processor.exe` 다운로드
2. 다운로드한 파일을 원하는 폴더에 저장
3. 명령 프롬프트(cmd)를 열고 해당 폴더로 이동
4. 실행:
   ```bash
   # 기본 사용법
   excel_processor.exe input output

   # 옵션 사용
   excel_processor.exe input output --format json
   ```

### macOS 사용자

1. [Releases](https://github.com/your-username/excel-processor/releases) 페이지에서 최신 버전의 `excel_processor_mac` 다운로드
2. 터미널을 열고 다운로드 폴더로 이동
3. 실행 권한 부여:
   ```bash
   chmod +x excel_processor_mac
   ```
4. 실행:
   ```bash
   ./excel_processor_mac input output
   ```


### 자주 사용하는 예시

1. 바탕화면의 '엑셀' 폴더 처리하기:
   ```bash
   # excel_processor.py가 있는 폴더에서
   python excel_processor.py "%USERPROFILE%\Desktop\엑셀" "%USERPROFILE%\Desktop\결과"
   ```

2. 다운로드 폴더의 엑셀 처리하기:
   ```bash
   python excel_processor.py "%USERPROFILE%\Downloads" "%USERPROFILE%\Downloads\결과"
   ```

### 상세 사용법

```bash
# JSON 형식으로 변환
python excel_processor.py 입력폴더경로 출력폴더경로 --format json

# Markdown 형식으로 변환
python excel_processor.py 입력폴더경로 출력폴더경로 -f markdown

# 모든 형식으로 변환
python excel_processor.py 입력폴더경로 출력폴더경로 --format all

# 시트별로 파일 분리 (기본값)
python excel_processor.py 입력폴더경로 출력폴더경로 --no-combine

# 모든 시트를 하나의 파일로 합치기
python excel_processor.py 입력폴더경로 출력폴더경로 --combine
```

## 출력 형식 설명

### CSV 형식 (기본값)
- 가장 단순하고 널리 사용되는 형식
- Excel, Google Sheets 등에서 바로 열기 가능
- 시트별로 별도의 파일로 저장됨 (예: `파일명_시트1.csv`, `파일명_시트2.csv`)

### JSON 형식
- 프로그래밍에 적합한 형식
- 데이터의 구조를 그대로 유지
- 시트별로 별도의 파일로 저장됨 (예: `파일명_시트1.json`)

### Markdown 형식
- 읽기 쉬운 테이블 형식
- GitHub 등에서 바로 보기 가능
- 시트별로 별도의 파일로 저장됨 (예: `파일명_시트1.md`)

## 주요 기능

- 여러 엑셀 파일 일괄 처리
- 다양한 출력 형식 지원 (CSV, JSON, Markdown)
- 시트별 개별 파일 저장 또는 통합 저장
- 한글 지원 (UTF-8 인코딩)
- 날짜, 시간 데이터 자동 변환
- 중복 컬럼명 자동 처리

## 자주 발생하는 문제 해결

### Windows에서 자주 발생하는 문제

1. "python을 찾을 수 없습니다" 오류
   - Python이 제대로 설치되었는지 확인
   - "Add Python to PATH" 옵션이 체크되어 있는지 확인
   - 컴퓨터를 재시작해보세요

2. 엑셀 파일을 열 수 없는 경우
   - 엑셀 파일이 다른 프로그램에서 열려있지 않은지 확인
   - 관리자 권한으로 명령 프롬프트 실행 후 시도

3. 한글이 깨지는 경우
   - Windows에서는 메모장으로 CSV 파일을 열 때 "UTF-8" 인코딩 선택
   - Excel에서는 "데이터" 탭의 "텍스트/CSV 파일" 기능으로 열기

### 필요한 패키지가 설치되지 않는 경우
```bash
# 관리자 권한으로 실행
pip install --upgrade pip
pip install -r requirements.txt
```

## 라이선스

MIT License