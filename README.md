# PLM PartList Automation (V8)

퍼시스 그룹 사내 PLM 시스템 기반 부품리스트(PartList) 자동 생성 및 조립/용접 전개 웹 서비스입니다.

---

## 🌟 주요 기능

1. **PLM BOM 실시간 자동 조회**:
   - 단품코드 일괄 입력을 통해 PLM 시스템에서 최신 BOM 데이터를 자동으로 추출합니다.
2. **사내 표준 5단계 버킷 정렬**:
   - 조립_용접 전개 시 사내 우선순위에 따라 부품을 정렬합니다:
     $$\text{목재(WDWW)} \longrightarrow \text{철제(STPG/STCG)} \longrightarrow \text{기타(볼트/사출/하드웨어)} \longrightarrow \text{하위 빌드 세트} \longrightarrow \text{하위 웰드 세트}$$
3. **조립_용접 시트 전개 및 2페이지 규격 출력**:
   - 빌드/웰드 자재는 메인 시트에서 1개 열로 요약(비고에 구성품 기재)되고, `[조립_용접]` 전용 시트에 상세 전개됩니다.
   - 인쇄 영역 A4 가로 2페이지(`1/2`, `2/2`) 자동 분할 서식 적용.
4. **결과 다운로드**:
   - 웹 브라우저에서 생성된 엑셀(.xlsx) 및 PDF 파일을 개별 또는 일괄(ZIP)로 즉시 다운로드할 수 있습니다.

---

## 🚀 실행 환경

- **언어/프레임워크**: Python 3.11, Streamlit
- **데이터 처리**: OpenPyXL, Selenium
- **배포 방식**: Docker Container (Streamlit Port 8501)
