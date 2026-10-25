# PLM PART LIST Automation - Version 8 인수인계서

**작성일자**: 2026-10-13 (초안) / **최종 업데이트**: 2026-10-26 (사내 서버 배포 완료)  
**담당자**: 권영조 (yeongjo_kwon@fursys.com)  
**작업 환경**: Windows, Python 3.11, OpenPyXL, Win32COM (pywin32), Selenium, CustomTkinter

---

## 1. 개요 및 V8 핵심 변경 목적

기존 V7.1 버전에서는 품목명에 `빌드(build)` 또는 `웰드(weld)`가 포함된 경우 해당 모자재를 메인 목록에서 제외하고 하위 자식 자재들을 펼쳐서 기입하는 방식이었습니다.  
**V8 버전**에서는 3대 핵심 요구사항 및 실무 피드백을 반영하여 개편되었습니다:

1. **메인 시트(`[단품코드]`)의 조립 단위 단일화**:
   - `빌드(build)` 또는 `웰드(weld)`로 표기된 품목은 하위 자식들을 펼치지 않고 **해당 모자재 1개 열만 등재**.
   - 직속 하위 자식 자재들의 정보는 `비고`란에 `자재코드 : 수량` 형식으로 줄바꿈(`\n`) 표기.
2. **별도 `[조립_용접]` 시트 신설**:
   - 빌드/웰드 모자재와 그 하위 구성 부품들을 상세 전개하여 확인할 수 있는 전용 시트 생성.
   - 상단 구분 행을 3개 층(`div_level1`, `div_level2`, `div_level3`)으로 계층화하여 모자재-자식 관계 및 별도품목을 시각화.
   - 조립 헤드(`is_asm_head`) 품목은 세로 2행 병합, 하위 자식들은 가로 병합 및 레벨별 색상 틴트(#FCD5B4, #CCC0DA 등) 적용.
3. **`[출력 페이지]` 및 PDF 2페이지(1/2, 2/2) 자동 분할 출력**:
   - **Page 1 (1/2)**: 메인 부품리스트(`[단품코드]`) + 우측 색상표(있을 경우). 표제란 Rev, 부품명, 단품코드, 페이지번호(`1/2`).
   - **Page 2 (2/2)**: `[조립_용접]` 시트의 연결된 그림(Live Link Paste), 표제란 부품명(`조립/용접 자재`), 페이지번호(`2/2`).
   - Excel 인쇄 영역을 `$A$1:$Y$66`으로 확장하고 34행(`A34`)에 페이지 구분선(HPageBreak) 적용.
   - Excel COM을 통해 A4 가로(Landscape, Zoom=48%) 2페이지 규격의 PDF 자동 내보내기.
   - 조립/용접 자재가 없는 단품은 기존 규격인 1페이지(`1/1`, `$A$1:$Y$33`)로 호환 유지.

---

## 2. 파일 구성 체계

- **V7.1 유지 파일 (원본 보존)**:
  - `plm_part_list_generator_v7.py`
  - `plm_part_list_automation_v7.py`
  - `plm_part_list_gui_v7.py`
- **V8 활성 파일**:
  - `plm_part_list_generator_v8.py`: 부품 수집, 엑셀 테이블 작성, 이미지 삽입, 서식 및 PDF 변환
  - `plm_part_list_automation_v8.py`: PLM 웹 드라이버 제어, 데이터 조회 및 추출, 팝업 처리
  - `plm_part_list_gui_v8.py`: 다중 단품 일괄 처리 GUI 애플리케이션
- **결과물 저장 위치**:
  - 프로그램 실행 기준 현재 작업 디렉토리 (`c:\Users\yeongjo_kwon\PART LIST_automation\`)
  - 파일 형식: `{단품코드}_{품목명}_PartList.xlsx` 및 동일한 이름의 `.pdf`

---

## 3. 2026-10-14 작업 내역 및 수정 완료 사항

오늘 실무 테스트 과정에서 발견된 이슈들을 분석하고 다음과 같이 완벽히 조치 완료했습니다:

### 3.1 PLM 조회 결과 클릭 타임아웃 (`result_click_fail`) 해결
- **문제 원인**: 단품코드 조회 시 PLM 서버 지연으로 중앙에 "처리중입니다. 잠시만 기다려 주시기 바랍니다." 로딩 마스크가 10초 이상 떠 있는 상태에서, 기존 스크립트가 빈 테이블 요소를 조기 인식하여 클릭 타임아웃 발생.
- **조치 사항 (`plm_part_list_automation_v8.py`)**:
  1. `AUIGrid.getRowCount()` API를 활용하여 실제 데이터 행이 반환될 때까지 최대 60초간 동적 대기.
  2. 마우스 클릭 좌표/마스크 충돌을 원천 회피하기 위해, PLM 내부의 `cellDoubleClick` 핸들러 로직(`_popup("/partlist/create", ...)`)을 **JavaScript로 직접 호출**하는 1단계 방식 도입.
  3. 미반응 시 Selenium 마우스 더블클릭(2단계) 및 상단 '등록' 버튼 클릭(3단계)으로 이어지는 3중 안전망 구축.

### 3.2 색상표 시트 셀 병합 오타 (`NameError`) 해결
- **문제 원인**: `plm_part_list_generator_v8.py` 1030행에서 `start_r` 변수를 `start_row`로 참조하는 오타로 인해 Excel 생성 중단.
- **조치 사항**: `ws2.merge_cells(start_row=start_r, ...)` 로 수정 완료 및 문법 검증 통과.

### 3.3 비표준 BOM 품목 (`DSPAAE06031`) 검토 및 방안 A 확정
- **현상**: 서랍장 품목에서 단품 바로 아래에 `W590 3단 서랍장 ASS'Y (ETET002902)`라는 중간 모자재가 단독으로 감싸고 있어 하위 자재가 누락되는 현상 발생.
- **결론**: 사내 공식 규격서 상 "2레벨 이하부터 부품 선정" 원칙을 준수하고 시스템 경량화 및 신뢰성을 유지하기 위해, 코드를 불필요하게 복잡하게 늘리지 않고 **방안 A (사내 표준에 맞추어 현상 유지)**로 최종 채택.

### 3.4 2페이지 "조립_용접" 시트 열 너비 통일 (18 규격)
- **문제 원인**: 메인 시트는 열 너비가 `18`이었으나, `조립_용접` 시트는 C열만 18이고 나머지 열(D~N)이 13으로 좁게 설정되어 열 간격이 제각각이고 사진이 겹침.
- **조치 사항 (`plm_part_list_generator_v8.py`)**:
  - `ws_assy`의 모든 부품 열(C열부터 끝까지)을 메인 시트와 동일하게 `width = 18`로 완전히 일치화.

### 3.5 부품 사진 셀 피팅(Fit-in-box) 크기 확대 조정
- **문제 원인**: 셀 너비(18 ≈ 144px), 높이(72pt = 96px) 대비 사진이 126x70px로 제한되어 상하 여백(갭)이 큼.
- **조치 사항 (`plm_part_list_generator_v8.py`)**:
  - 최대 허용 크기를 **가로 최대 `140px`, 세로 최대 `90px`**로 확대 조정.
  - 원본 종횡비를 유지하면서 셀 안에 꽉 차도록 피팅되어 옆 열 침범 없이 선명하게 표시.

### 3.6 비고란(ROW_REM) 높이 동적 자동 확장 (`IPP00AQ01A` 대응)
- **문제 원인**: `ETET001222(빌드)`처럼 하위에 20종류의 자재가 있는 경우, 비고란 높이가 66.0pt(약 4줄)로 고정되어 4줄까지만 보이고 나머지가 잘림.
- **조치 사항 (`plm_part_list_generator_v8.py`)**:
  - **4줄 이하**: 기존 디폴트 높이인 **`66.0pt` 유지**.
  - **4줄 초과**: `줄 수 × 14.5pt + 8.0pt` 로 동적 자동 계산 (20줄 기준 `298.0pt`).
  - 메인 시트 및 1페이지 출력 라이브 링크 그림에서도 20줄 전체가 상하 잘림 없이 온전히 인쇄됨.

---

## 4. 2026-10-15 작업 내역: "조립_용접" 시트 나열 순서 정렬 적용 완료

사용자와 협의된 **옵션 B 변형 정렬 규칙**을 [plm_part_list_generator_v8.py](file:///c:/Users/yeongjo_kwon/PART%20LIST_automation/plm_part_list_generator_v8.py)의 [collect_parts_v8](file:///c:/Users/yeongjo_kwon/PART%20LIST_automation/plm_part_list_generator_v8.py#L436-L540)에 성공적으로 적용 및 검증 완료했습니다.

### 4.1 정렬 순서 정의
사내 8번 배열 규칙(`WDWW` > `STPG/STCG` > `기타` > `별도`)을 계승하며, V8 상단 3단 구분행 병합을 완벽히 유지하는 6단계 정렬:

$$\text{목재(WDWW)} \longrightarrow \text{철제(STPG/STCG)} \longrightarrow \text{기타(볼트/하드웨어)} \longrightarrow \text{일반 빌드 그룹} \longrightarrow \text{일반 웰드 그룹} \longrightarrow \text{별도품목 하위 빌드 그룹}$$

### 4.2 분류 그룹 상세 (Bucket 정의)
1. **Group 1 (목재 단품 자재)**: ERP 코드가 `WDWW`로 시작하는 일반 단품
2. **Group 2 (철제 단품 자재)**: ERP 코드가 `STPG`, `STCG`로 시작하는 일반 단품
3. **Group 3 (기타 단품 자재)**: 나사/볼트(`ETFS`), 레일(`ETHW`), 사출물(`OPPW`), 스트랩 등 일반 단품
4. **Group 4 (일반 [빌드] 세트)**: 품목명에 `빌드/build`가 포함된 서브 모자재 헤드 + 직속 자식들 (헤드-자식 연속성 및 상단 병합 유지)
5. **Group 5 (일반 [용접] 세트)**: 품목명에 `웰드/weld`가 포함된 서브 모자재 헤드 + 직속 자식들
6. **Group 6 (별도품목 하위 빌드 세트)**: `별도품목(ETPJ)` 하위에 속한 빌드/용접 자재 및 그 자식들 (`assy_baldo`)

### 4.3 적용 완료 및 검증 결과
- **적용 로직**:
  ```python
  assy_parts = (
      bucket_wdww +
      bucket_steel +
      bucket_hardware +
      bucket_build +
      bucket_weld +
      assy_baldo
  )
  ```
- **검증 단품 테스트 통과**:
  - `IPP00AQ01A` (아일랜드 서랍몸통 1000폭): 목재(6) $\rightarrow$ 철제(2) $\rightarrow$ 기타(8) $\rightarrow$ 빌드(23) $\rightarrow$ 용접(3) 총 42개 자재 순서 배치 및 Excel/PDF 2페이지 정상 생성 확인.
  - `HSKD1420LR` (익스텐션 책상 다리): 기타(1) $\rightarrow$ 빌드(3) $\rightarrow$ 용접(3) $\rightarrow$ 별도품목(10) 총 17개 자재 순서 배치 및 Excel/PDF 2페이지 정상 생성 확인.

### 4.4 WDWW 하위 자재 재귀 탐색 및 뒤쪽 나열 (사내 규칙 2-2 복원)
- **문제 원인**: V8 개편 과정에서 WDWW 자재를 리프(단품)로만 취급하여 그 하위 자식 노드를 탐색하지 않아, 케이싱, 케이싱스크류, 스티커, 네임플레이트 등이 누락되는 현상 발생.
- **조치 사항 (`plm_part_list_generator_v8.py`)**:
  1. `extract_wdww_sub_items_recursive()` 함수 구현:
     - 품목명에 `목제`, `목재`, `패널`, `엣지`, `edge` 또는 `-`인 항목은 자기 자신 및 하위 전체 제외(`_is_wdww_excluded`).
     - 비제외 항목(`케이싱`, `케이싱스크류`, `스티커`, `네임플레이트` 등)은 자재코드(ERP)별로 수량 누적 합산 및 비고란에 `부모자재코드: 수량` 기록.
  2. **메인 시트**: 2레벨 단품 자재들 뒤에 나열 (`main_normal` 끝, `main_baldo` 앞).
  3. **조립_용접 시트**: `bucket_hardware` (Group 3: 기타 단품)에 병합되어 WDWW, 철제 뒤에 나열.
- **검증 단품 (`DSCC0145V`)**:
  - 6개 WDWW 자재 뒤에 `ETFS001264 케이싱` (24개, 비고 3줄 누적), `ETFS001266 케이싱스크류` (10개, 비고 4줄 누적)가 정확히 등재되고 우측 색상표 및 1페이지 PDF 정상 출력 확인 완료.

### 4.5 다중 단품코드 일괄 처리 시 이전 결과 잔류(Stale Grid) 및 1칸 밀림 버그 해결
- **문제 현상**:
  - GUI에서 여러 단품코드를 일괄 입력하고 실행했을 때, 두 번째 파일명(`DSCDAD0807_...`) 안에 첫 번째 단품(`DSCC0145V`)의 부품리스트가 들어가고, 이후 모든 항목이 직전 단품의 정보로 1칸씩 밀려 생성되는 치명적 불일치 발생.
- **원인 분석**:
  1. **AUIGrid 잔류 데이터 조기 인식**: 직전 단품 검색 결과가 그리드에 남아있는 상태에서 새 단품 검색을 누르면, PLM 서버 응답(AJAX)이 오기도 전에 `AUIGrid.getRowCount() > 0` 조건을 즉시 통과함.
  2. **직전 단품의 팝업 호출**: 아직 새 결과로 갱신되지 않은 `data[0]`(직전 단품)의 `oid`로 팝업이 열려 직전 단품의 PartList가 생성됨.
  3. **BOM 메뉴 재클릭 레이스 컨디션**: 매 단품마다 좌측 메뉴를 재클릭하면서 `body` 프레임이 불필요하게 리로드되어 입력 상태가 꼬임.
- **조치 사항 (`plm_part_list_automation_v8.py`)**:
  1. **화면 위치 감지 (`is_already_on_partlist`)**: 이미 PARTLIST 관리 프레임에 있으면 좌측 메뉴 재클릭을 생략하여 불필요한 iframe 새로고침 방지.
  2. **검색 전 완전 초기화 (Clean Slate)**: 검색 전 `#searchInit` 버튼 클릭 및 JS로 `erpCode`, `creator`, `receptionist` 등 모든 필터 초기화, `AUIGrid.clearGridData()`로 그리드 행을 0개로 완전 초기화.
  3. **타겟 코드 일치 검증 대기**: 단순히 `rowCount > 0`이 아니라, 그리드 데이터 중 `erpCode`가 `target_code`와 일치하는 행이 실제로 로드되었는지(`matchCount > 0`)를 확인 후 대기 종료.
  4. **정확한 타겟 행 인덱스 지정**: 감지된 `matched_row_idx`를 전달하여 정확한 단품의 팝업을 열도록 개선 (JS 및 마우스 더블클릭 모두 타겟 행 매칭).
  5. **추출 단품코드 유효성 검증**: 최종 데이터 추출 시 루트 품목의 ERP 코드가 요청한 `target_code`와 일치하는지 교차 검증 및 안내 로그 출력.

### 4.6 비고란(ROW_REM) 높이 계산식 개선 (10줄 이상 줄바꿈 시 잘림 현상 방지)
- **문제 현상**:
  - `DSABPA10001`의 `WDWW120722`처럼 빌드 하위에 10개의 자식 자재가 있을 때, 비고 셀에 10줄의 텍스트가 모두 기재되어 있음에도 엑셀 화면 및 출력물에서 9줄까지만 보이고 10번째 줄(`ETFS000104 : 4`)이 아래 테두리에 가려지는 현상 발생.
- **원인 분석**:
  - 기존 높이 산출식(`max_rem_lines * 14.5 + 8.0`)은 10줄 기준 153.0pt로 계산됨.
  - Excel의 10pt 폰트 기본 상하 여백 및 줄 간격(~16.5~17.5pt/줄)으로 인해 153.0pt로는 10번째 줄의 박스가 미세하게 잘림.
- **조치 사항 (`plm_part_list_generator_v8.py`)**:
  - 4줄 초과 시 높이 산출식을 **`max(66.0, max_rem_lines * 18.5 + 10.0)`** 로 상향 조정.
  - 10줄 기준 기존 `153.0pt` $\rightarrow$ **`195.0pt`** 로 약 42pt의 충분한 숨통이 확보되어 10번째 줄 및 그 이상의 다중 항목도 하단 잘림 없이 시원하게 표시됨.

### 4.7 비-WDWW 자재(ETEP 모듈 등) 하위 스티커(PPGW) 누락 방지 및 PLM 모델링 가이드
- **문제 현상**:
  - `DSABPA10001`의 `WDWW120722` 하위에 속한 `무선충전허브모듈 2.0 (ETEP000393)`에 부착되는 `KC인증스티커 (PPGW001697)`가 부품리스트 및 `[조립_용접]` 시트에서 완전히 누락됨.
- **원인 분석**:
  1. 기존 규칙 2-2 재귀 탐색은 모자재가 `WDWW`로 시작하는 경우에만 하위 자식을 탐색하도록 한정되어 있었음.
  2. `ETEP000393`은 ERP 코드가 `ETEP`이고 품목명에도 `(빌드)`가 표기되어 있지 않아, 단순 리프(단품) 자재로 분류되어 하위 자식 노드를 확인하지 않고 본인만 등재함.
- **조치 사항 및 해결 방안**:
  1. **프로그램 안전망 강화 (`plm_part_list_generator_v8.py`)**:
     - `WDWW`뿐만 아니라 철제(`STPG/STCG`) 및 기타 단품 자재(`ETEP`, 사출물 등)라도 **하위 자식이 매달려 있는 경우(`ach.get('children')`)**, 스티커(`PPGW`), 나사, 네임플레이트 등 부속 자재를 자동으로 재귀 탐색하여 `[조립_용접]` 시트의 `bucket_hardware`에 정상 수집(비고: `부모코드: 수량`).
     - PLM 품목명을 바꾸지 않아도 누락 없이 자동 수집되도록 보장.
  2. **PLM 모델링 권장 가이드**:
     - `ETEP000393`의 Creo 원본 파일은 `..._CHARGE_ASSY.ASM`으로 어셈블리 구조이며, 모듈 본체에 KC스티커를 작업자가 부착(조립)하는 공정이 수반됩니다.
     - 따라서 BOM 표준 관리상 **PLM 품목명에 `(빌드)` 또는 `(조립)`을 표기하는 것이 가장 정석적**인 모델링 방식입니다.
     - 품목명에 `(빌드)`가 들어가면 프로그램이 이를 서브 조립품으로 인식하여, `[조립_용접]` 시트에서 **`[조립]` 오렌지 헤더와 함께 모듈-스티커가 하나의 조립 그룹(Group 4)으로 시각화**됩니다.

### 4.8 빌드(A) 하위 빌드(B) 구성 부품의 5단계 버킷 정렬 규칙 적용 (`IPP00AQ01A`)
- **문제 현상**:
  - `IPP00AQ01A`와 같이 복잡한 조립 구조에서 빌드(A) `ETET001222(조립부품)` 하위에 중첩 빌드(B)인 `ETET001223(아일랜드 1000_L 속서랍 조립도)`이 존재할 때, 빌드(B)의 하위 자재들이 사내 배열 규칙을 따르지 않고 PLM BOM 원본 등록 순서(`나사(ETFS)` $\rightarrow$ `철제(STCG)` $\rightarrow$ `목재(WDWW)`) 그대로 나열되는 현상 발생.
- **원인 분석**:
  - 기존 로직은 최상위 2레벨 조립품의 직속 하위 자재들은 6단계 버킷으로 분류했으나, 서브 조립품(`ach`) 내부의 자식(`ach_children`)들을 리스트에 추가할 때는 `for a_sub in ach_children:` 반복문으로 PLM 순서 그대로 순차 추가하고 있었음.
- **조치 사항 (`plm_part_list_generator_v8.py`)**:
  1. `_collect_assembly_group()` 재귀 헬퍼 함수 구현:
     - 서브 조립품(빌드/웰드)의 자식 노드들을 5단계 표준 버킷으로 분류 및 정렬:
       $$\text{모자재 본인(헤드)} \longrightarrow \text{목재(WDWW)} \longrightarrow \text{철제(STPG/STCG)} \longrightarrow \text{기타(볼트/사출/하드웨어)} \longrightarrow \text{하위 빌드 세트} \longrightarrow \text{하위 웰드 세트}$$
     - 모자재 헤드는 2행 세로 병합(`is_asm_head=True`)을 유지하고, 그 직후에 정렬된 하위 자재들이 연속 배치됨.
     - 하위 자재 중 또 다시 빌드/웰드가 포함된 경우 재귀적으로 탐색하여 하위 조립 세트로 그룹화.
  2. 별도품목(`assy_baldo`) 하위 조립품의 자식 자재에도 동일한 5단계 정렬 일괄 적용.
- **검증 결과 (`IPP00AQ01A`)**:
  - `ETET001223 (빌드 헤드)` 직후에:
    1. 목재(WDWW): 뒤판(`WDWW128758`), 앞판(`WDWW128760`), 밑판(`WDWW115619`), 측판L(`WDWW122485`), 측판R(`WDWW122487`) 5개 자재 선행 배치.
    2. 철제(STCG): 보강브라켓-450-N2(`STCG002706`) 배치.
    3. 기타(ETFS): 나사 3종류(`ETFS000109`, `ETFS000236`, `ETFS000239`) 후행 배치.
  - 상단 3단 구분행 병합 및 틴트 색상(#FCD5B4, #FDE9D9)이 온전히 유지되며 Excel 및 2페이지 PDF 정상 출력 확인 완료.

---

## 5. 2026-10-15 오후 작업 내역: 사내 AI Hub 웹 앱 구축 및 로컬 테스트

사내 AI Hub / Coolify 클라우드 배포를 위한 웹 애플리케이션 구축, GitHub 조직 연동, 도커화 및 로컬 웹 브라우저 검증을 완료했습니다.

### 5.1 사내 GitHub 연동 및 리포지토리 구성
1. **GitHub 계정 세팅 및 사내 조직 가입**:
   - 계정명: `fg-yeongjo-kwon` (사내 표기 원칙 준수)
   - 이름: `권영조` / 이메일: `yeongjo_kwon@fursys.com` (인증 완료)
   - 사내 공식 조직 가입 완료: `fursys-group-hub`
2. **원격 저장소 생성 및 푸시**:
   - 리포지토리 주소: `https://github.com/fursys-group-hub/plm-partlist-automation`
   - 브랜치: `main`
   - `.gitignore` 정돈: 불필요한 캐시, 구버전(v2~v7) 스크립트, 임시 산출물 제외, V8 핵심 코드 및 `template.xlsx`만 추적.

### 5.2 Streamlit 기반 웹 애플리케이션 (`app.py`) 개발
1. **사내 UI/UX 브랜딩**:
   - 퍼시스 블루(#1E3A8A) 테마, 반응형 레이아웃, 이용 가이드 및 5단계 버킷 정렬 규칙 안내 탑재.
2. **기능 구현**:
   - **PLM 로그인 계정 입력**: 사이드바에서 아이디/비밀번호 안전 입력.
   - **단품코드 다중 일괄 입력**: 엑셀 열 복사-붙여넣기 지원 (줄바꿈/쉼표 자동 파싱).
   - **실시간 작업 현황(Live Streaming Log)**: Python stdout과 웹 UI를 실시간 연결하여 현재 진행 단계(`[1/4] 브라우저 드라이버 시작`, `[2/4] 로그인`, `BOM 조회`, `그리드 추출`, `서식 작성`)를 한눈에 모니터링 가능.
   - **결과물 다운로드**: 단품별 개별 Excel/PDF 다운로드 및 복수 단품 시 **전체 일괄 ZIP 압축 다운로드** 지원.
   - **로컬 테스트 옵션 탑재**: Windows 환경에서 브라우저 화면 표시(Headless 해제) 및 Edge/Chrome 선택 토글 제공.

### 5.3 컨테이너화 (`Dockerfile`) 구성
- 사내 Coolify(Linux) 환경에서 실행될 수 있도록 `python:3.11-slim` 기반 이미지 작성.
- Selenium용 Chromium 및 Chromedriver 설치 및 환경 변수 자동 바인딩.

### 5.4 로컬 웹 테스트 검증 및 버그 조치 완료
1. **테스트 진행 (`http://localhost:8502`)**:
   - 테스트 대상: `HSKD1420LR`, `IPP00AQ01A`, `IFP00AQ04A` 3종 일괄 실행.
2. **1번 단품 (`HSKD1420LR`) 결과**: **100% 정상 완료**
   - PLM 자동 로그인 $\rightarrow$ BOM 데이터 추출 $\rightarrow$ V8 서식 적용 $\rightarrow$ Excel(`HSKD1420LR_..._PartList.xlsx`) 및 PDF(`HSKD1420LR_..._PartList.pdf`, 2페이지 가로) 생성 완벽 성공.
3. **2번 단품 (`IPP00AQ01A`) 타임아웃 원인 규명 및 조치**:
   - **현상**: `process_single_code` 261행 `erpCode` 대기 중 `TimeoutException` 발생.
   - **원인 분석**: 1번 단품 완료 후 PLM이 띄운 자식 팝업 창(`partListDetailPopup`)에 브라우저 포커스가 머물러 있어, 메인 윈도우의 `erpCode` 입력창을 찾지 못함 (데스크톱 GUI 버전에는 있던 팝업 닫기 및 `main_window_handle` 복귀 로직이 `app.py` 루프에서 누락됨).
   - **조치 사항 (`app.py`)**:
     - `main_window_handle` 저장 후, 매 단품 시작 전 메인 윈도우 및 최상위 컨텐트로 강제 전환(`driver.switch_to.window(main_window_handle)`).
     - 각 단품 완료 후 파생된 모든 자식 팝업 창 자동 닫기(`driver.close()`) 및 메인 창 복귀 로직 추가.
     - GitHub 원격 저장소(`main`)에 즉시 반영 푸시 완료 (`commit: b61e3a5`).

---

## 6. 2026-10-16 작업 내역 및 AI Hub 앱 등록 완료

10월 16일 진행된 다중 단품 연속 실행 안정화, 웹 UI/UX 최적화, 다운로드 상태 유지 버그 해결 및 사내 AI Hub 등록 결과입니다:

### 6.1 다중 단품 연속 실행 안정화 및 에러 격리 (`app.py`, `plm_part_list_automation_v8.py`)
- **다중 실행 루프 강화**: 단품별 `try...except...finally` 구조를 도입하여 중간 1개 단품에서 오류가 발생해도 후속 단품 작업이 중단되지 않고 끝까지 완료되도록 개선.
- **자식 팝업 정리 및 핸들러 복귀**: 단품 처리 전/후 잔류 팝업 창 자동 닫기 및 메인 윈도우 컨텍스트 복귀 로직을 상시 적용.
- **로컬 계정 자동 로드**: Windows 환경에서 `%APPDATA%\PLM_PartList_Automation\login_config.json` 계정 자동 불러오기 지원.

### 6.2 웹 애플리케이션 UI/UX 최적화 (`app.py`)
- **조직 명칭 변경**: 사이드바 타이틀을 '퍼시스그룹 생산기술 자동화' $\rightarrow$ **'퍼시스그룹 업무 자동화'** 로 변경.
- **시스템 정보 단순화**: 사이드바에 불필요한 세부 스펙 제거하고 **`버전: V8.0`** 만 깔끔하게 표기.
- **이용 가이드 4단계 표준화 & 디자인 개선**:
  - 0단계(사전 준비)부터 3단계(다운로드)까지 번호별 간격을 일정하게 통일하고 상하 여백 대칭 배치.
  - 사전 준비 위키 링크를 모던한 사각 링크 뱃지(`🔗 가이드 링크 ↗`)로 스타일링 (`Atlassian Wiki` 연결).
- **상단 Deploy 버튼 숨김**: Streamlit 기본 배포 버튼을 CSS로 완전 은폐 처리하여 사용자 혼선 차단.

### 6.3 결과 다운로드 세션 상태(`st.session_state`) 유지 및 오류 해결
- **문제 원인**:
  - `st.download_button` 클릭 시 Streamlit 스크립트가 재실행(Rerun)되면서, 시작 조건문(`if start_clicked:`) 안의 다운로드 UI가 통째로 사라지는 현상 발생.
  - 이로 인해 개별 파일 다운로드 버튼 컴포넌트가 언마운트되며 서버의 임시 파일 주소가 만료되어 브라우저에서 `사용할 수 없는 파일 (b04d...txt)` 에러 발생.
- **조치 사항**:
  - `st.session_state['generated_files']` 및 `st.session_state['zip_buffer']`를 도입하여 작업 결과와 압축 데이터를 세션 메모리에 영구 보존.
  - 다운로드 카드 섹션을 독립 분리하여 몇 번 다운로드를 클릭하더라도 화면과 파일 링크가 사라지지 않고 항상 유지되도록 수정 완료.

### 6.4 사내 AI Hub Appstore 등록 완료
- **등록 페이지**: `https://ai-hub.fursys.com/workspace/apps/339`
- **앱 ID**: **`339`**
- **앱 이름**: **`부품리스트 자동화 프로그램 (PLM기반)`**
- **업무 도메인**: **`생산·품질`** (외주 생산업체 배포 및 단품/포장 단위 자재 검수 목적)
- **앱 유형**: **`수집·자동화`**, **`업무 관리`**
- **GitHub 저장소**: `https://github.com/fursys-group-hub/plm-partlist-automation` (main 브랜치 최신화 완료)
- **프리뷰 이미지**: 배포 버튼이 숨겨진 고화질 UI 스크린샷 2장 등록 완료 ([ai_hub_preview](file:///c:/Users/yeongjo_kwon/PART%20LIST_automation/ai_hub_preview))

### 6.5 V8.0 독립 실행 파일 (`.exe`) 빌드 완료
- **빌드 방식**: PyInstaller 기반 단일 실행 파일 (`--onefile`, `template.xlsx` 서식 및 의존성 전체 번들링)
- **생성 파일**: `PARTLIST_Automation_v8.0.exe` (약 149 MB)
- **저장 위치**:
  - 작업 폴더 루트: [PARTLIST_Automation_v8.0.exe](file:///c:/Users/yeongjo_kwon/PART%20LIST_automation/PARTLIST_Automation_v8.0.exe)
  - 배포 폴더: [dist/PARTLIST_Automation_v8.0.exe](file:///c:/Users/yeongjo_kwon/PART%20LIST_automation/dist/PARTLIST_Automation_v8.0.exe)
- **사용법**: 파이썬 환경이 없는 일반 업무용 PC에서도 더블클릭만으로 즉시 구동 가능 (다중 단품 연속 실행 팝업 안전 정리 및 암호화 계정 기억 기능 탑재).

### 6.6 계정 보안 및 Git 저장소 격리 검증 완료
- **로컬 자격 증명 분리**: 로컬 편의를 위한 계정 저장(`login_config.json`)은 프로젝트 내부가 아닌 사용자 개인 로컬 디렉토리(`%APPDATA%\PLM_PartList_Automation\`)에만 저장됨.
- **GitHub 보안 보장**: `.gitignore`에 `login_config.json`, `.env`, `*.exe`가 모두 등록되어 있어, GitHub 원격 저장소 및 사내 배포 서버에는 개인 계정이나 비밀번호가 전혀 포함되지 않음 (공용 서버 실행 시 모든 사용자에게 깨끗한 빈 입력창으로 제공됨).

---

## 7. 사내 서버 배포 완료 및 운영 정보 (2026-10-26 완료)

2026년 10월 26일, 사내 AI 도구(`Claude`) 사용 승인 완료 후 배포 전 보안/환경 검토 통과 및 사내 전용 서버(Coolify) 배포를 최종 완료했습니다:

### 7.1 배포 진행 내역
1. **Claude Code 인증 및 도구 세팅**:
   - 사내 승인 계정(`yeongjo_kwon@fursys.com`, `Fursys Group 4`)으로 OAuth 연동 완료.
   - 사내 배포 플러그인 `fursys-deploy-hub (v4.14.0)` 설치 및 활성화.
2. **배포 전 보안 및 환경 검토 (`/deploy-check`, `/deploy-fix`)**:
   - `app.py`: 최근 파일 자동 복구 로직을 로컬(Windows) 환경으로 한정 격리하여 타 사용자 결과물 노출 방지(보안 등급 통과).
   - `Dockerfile`: `HEALTHCHECK` 대상을 `127.0.0.1`로 수정하고, `ENTRYPOINT`에 `--server.headless=true` 옵션을 추가하여 리눅스 컨테이너 환경에서 프롬프트 대기 없이 즉시 기동되도록 안정화.
   - 최종 검사 결과: **치명 0, 높음 0 (보안 통과 ✅ / 배포 준비 완료 ✅)**
   - 온라인 검토 리포트: `https://deploy.hub.fursys.com/report/zsuoXazHV8nQAQogek3FWS10lg9qvwqx`
3. **사내 서버 배포 (`/deploy`)**:
   - 서브도메인: `plm-partlist`
   - **사내 전용 접속 주소**: **`https://plm-partlist.app3.hub.fursys.com`**
   - 사내 계정 SSO 보호 연동 적용 완료.
4. **향후 운영 및 지속적 배포 (CI/CD)**:
   - 향후 코드 수정 시 추가 배포 명령어 실행 없이, GitHub 원격 저장소(`main` 브랜치)에 **`git push`** 만 수행하면 사내 Coolify 서버가 자동으로 변경 사항을 감지하여 재빌드 및 무중단 배포를 진행합니다.

---

## 8. 부록: Antigravity IDE 사용 팁

- **도구 자동 승인(Auto-Approval) 설정**:
  명령어 실행 시 매번 나타나는 컨펌 창을 건너뛰고 자율 실행하려면:
  `좌측 하단 ⚙️ 설정(Settings) > Agent Settings > Tool Execution Policy` 를 **`always-proceed`** 로 변경.





