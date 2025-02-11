# Contributing to Excel Processor

엑셀 프로세서 프로젝트에 기여해주셔서 감사합니다!

## 기여 방법

1. 이 저장소를 Fork 합니다
2. 새로운 Branch를 만듭니다 (`git checkout -b feature/amazing-feature`)
3. 변경사항을 Commit 합니다 (`git commit -m 'Add some amazing feature'`)
4. Branch에 Push 합니다 (`git push origin feature/amazing-feature`)
5. Pull Request를 생성합니다

## 개발 환경 설정

1. 저장소 클론
   ```bash
   git clone https://github.com/your-username/excel-processor.git
   cd excel-processor
   ```

2. 가상환경 생성 및 활성화
   ```bash
   # Windows
   python -m venv venv
   venv\Scripts\activate

   # macOS/Linux
   python3 -m venv venv
   source venv/bin/activate
   ```

3. 의존성 설치
   ```bash
   pip install -r requirements.txt
   ```

## 코드 스타일

- PEP 8 가이드라인을 따릅니다
- 의미 있는 변수명과 함수명을 사용합니다
- 한글 주석을 권장합니다

## 테스트

새로운 기능을 추가하거나 버그를 수정한 경우, 관련 테스트를 추가해주세요.

## 라이선스

이 프로젝트에 기여함으로써, 귀하의 기여가 MIT 라이선스 하에 배포됨에 동의하게 됩니다. 