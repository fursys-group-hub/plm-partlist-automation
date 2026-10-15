# PLM PartList Automation

퍼시스 그룹 사내 PLM 시스템 기반 부품리스트(PartList) 자동 생성 프로그램입니다.

---

## 🌟 주요 기능

**PLM BOM 실시간 자동 조회**:
   - 단품코드 일괄 입력을 통해 PLM 시스템에서 최신 BOM 데이터를 자동으로 추출합니다.
**결과 다운로드**:
   - 웹 브라우저에서 생성된 엑셀(.xlsx) 및 PDF 파일을 개별 또는 일괄(ZIP)로 즉시 다운로드할 수 있습니다.

---

## 🚀 실행 환경

- **언어/프레임워크**: Python 3.11, Streamlit
- **데이터 처리**: OpenPyXL, Selenium
- **배포 방식**: Docker Container (Streamlit Port 8501)
