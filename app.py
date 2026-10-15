"""
Fursys PLM PartList Automation - Web Application (Streamlit)
- 사내 AI Hub / Coolify 배포용 웹 인터페이스
- 단품코드 일괄 입력 -> PLM 자동 조회 -> 사내 표준 V8 규격 Excel 및 PDF 생성 -> 다운로드 제공
"""
import streamlit as st
import os
import sys
import io
import zipfile
import time
from urllib.parse import urlparse

# Selenium & Generator import
from plm_part_list_automation_v8 import setup_driver, login, process_single_code
from plm_part_list_generator_v8 import collect_parts_v8, generate_part_list_excel_v3

# Streamlit Page Config
st.set_page_config(
    page_title="PLM PartList 자동화 (V8)",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for modern, premium Fursys UI
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #4B5563;
        margin-bottom: 1.5rem;
    }
    .status-card {
        background-color: #F8FAFC;
        border-radius: 10px;
        padding: 1.2rem;
        border-left: 5px solid #3B82F6;
        margin-bottom: 1rem;
    }
    .stButton>button {
        background-color: #1E3A8A;
        color: white;
        font-weight: 600;
        border-radius: 8px;
        height: 3rem;
        width: 100%;
        border: none;
    }
    .stButton>button:hover {
        background-color: #1D4ED8;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

# Sidebar: Credentials & System Info
with st.sidebar:
    st.image("https://www.fursys.com/resources/images/common/logo.png", width=140)
    st.markdown("### 🔐 PLM 계정 설정")
    plm_id = st.text_input("PLM 아이디", placeholder="예: yeongjo_kwon", help="사내 PLM 로그인 아이디")
    plm_pw = st.text_input("PLM 비밀번호", type="password", placeholder="비밀번호 입력", help="사내 PLM 로그인 비밀번호")
    
    st.divider()
    st.markdown("### ⚙️ 시스템 정보")
    st.markdown("- **규격 버전**: `V8.0`")
    st.markdown("- **정렬 규칙**: 사내 표준 5단계 버킷 정렬")
    st.markdown("- **지원 규격**: A4 가로 2페이지 (메인 + 조립_용접)")
    st.markdown("- **환경**: 사내 AI Hub 클라우드 서버")
    
    st.caption("© 2026 Fursys Group. All rights reserved.")

# Main Header
st.markdown('<div class="main-header">📋 PLM 부품리스트 자동 생성기</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">단품코드를 입력하면 PLM BOM을 실시간 조회하여 사내 표준 서식에 맞춘 엑셀 및 PDF 부품리스트를 자동으로 생성합니다.</div>', unsafe_allow_html=True)

# Main Form
col_input, col_info = st.columns([2, 1])

with col_input:
    st.markdown("#### 📝 단품코드 목록 입력")
    raw_codes = st.text_area(
        "단품코드들을 입력하세요 (줄바꿈 또는 쉼표로 구분, 엑셀 열 복사-붙여넣기 가능)",
        placeholder="IPP00AQ01A\nDSABPA10001\nDSCC0145V",
        height=180
    )
    
    start_clicked = st.button("🚀 부품리스트 자동 생성 시작", use_container_width=True)

with col_info:
    st.markdown("#### 💡 이용 가이드")
    st.info("""
    1. **단품코드 입력**: 여러 개의 단품코드를 엑셀에서 복사해 그대로 붙여넣을 수 있습니다.
    2. **조립/용접 자동 전개**: 품목명에 `(빌드)`나 `(웰드)`가 포함된 품목은 별도의 `[조립_용접]` 시트에 자동 전개됩니다.
    3. **출력 페이지 자동 분할**: 조립/용접 자재가 있으면 2페이지(`1/2`, `2/2`), 없으면 1페이지(`1/1`)로 자동 분할 생성됩니다.
    4. **결과 다운로드**: 작업 완료 후 생성된 엑셀과 PDF를 즉시 다운로드하실 수 있습니다.
    """)

# Execution logic
if start_clicked:
    if not plm_id or not plm_pw:
        st.error("⚠️ 좌측 사이드바에서 PLM 아이디와 비밀번호를 모두 입력해 주세요.")
    elif not raw_codes.strip():
        st.error("⚠️ 생성할 단품코드를 최소 1개 이상 입력해 주세요.")
    else:
        # Parse target codes
        lines = [c.strip().upper() for c in raw_codes.replace(',', '\n').split('\n') if c.strip()]
        # Remove duplicates while preserving order
        target_codes = []
        for c in lines:
            if c not in target_codes:
                target_codes.append(c)
                
        st.success(f"총 {len(target_codes)}개 단품코드에 대해 생성을 시작합니다: {', '.join(target_codes)}")
        
        status_box = st.status("🚀 PLM 브라우저 초기화 및 접속 중...", expanded=True)
        log_placeholder = st.empty()
        progress_bar = st.progress(0.0)
        
        logs = []
        def log_msg(msg):
            logs.append(msg)
            # Display last 8 log lines
            log_placeholder.code('\n'.join(logs[-8:]), language="bash")

        driver = None
        generated_files = []
        
        try:
            log_msg("[1/4] 헤드리스 브라우저 드라이버 시작 중...")
            driver = setup_driver(browser='chrome', headless=True)
            
            log_msg("[2/4] PLM 시스템 자동 로그인 시도 중...")
            login(driver, "https://plm.fursys.com", plm_id, plm_pw)
            log_msg("-> PLM 로그인 성공!")
            
            total = len(target_codes)
            for idx, code in enumerate(target_codes):
                status_box.update(label=f"[{idx+1}/{total}] 단품코드 '{code}' BOM 조회 및 부품리스트 작성 중...")
                log_msg(f"\n--- [{idx+1}/{total}] 단품코드 '{code}' 처리 시작 ---")
                
                res = process_single_code(driver, code)
                if res:
                    excel_path = res[0]
                    pdf_path = excel_path.replace(".xlsx", ".pdf")
                    generated_files.append((code, excel_path, pdf_path if os.path.exists(pdf_path) else None))
                    log_msg(f"-> '{code}' 부품리스트 생성 완료: {os.path.basename(excel_path)}")
                else:
                    log_msg(f"-> [오류] '{code}' 처리 실패 (단품코드 또는 PLM 상태 확인 필요)")
                    
                progress_bar.progress((idx + 1) / total)
                
            status_box.update(label="🎉 모든 단품코드 처리 완료!", state="complete", expanded=False)
            
        except Exception as e:
            status_box.update(label=f"❌ 처리 중 에러 발생: {e}", state="error")
            st.error(f"오류 상세: {e}")
            import traceback
            st.code(traceback.format_exc())
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass

        # Results & Download Section
        if generated_files:
            st.divider()
            st.markdown("### 📥 생성된 부품리스트 다운로드")
            
            # ZIP download for all files if multiple
            if len(generated_files) > 1 or any(f[2] for f in generated_files):
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for code, xl, pdf in generated_files:
                        if xl and os.path.exists(xl):
                            zip_file.write(xl, os.path.basename(xl))
                        if pdf and os.path.exists(pdf):
                            zip_file.write(pdf, os.path.basename(pdf))
                zip_buffer.seek(0)
                
                st.download_button(
                    label="📦 전체 결과 파일 일괄 다운로드 (ZIP)",
                    data=zip_buffer,
                    file_name=f"Fursys_PartLists_{time.strftime('%Y%m%d_%H%M%S')}.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                st.markdown("<br>", unsafe_allow_html=True)
            
            # Individual file download cards
            for code, xl, pdf in generated_files:
                with st.expander(f"📄 {code} 부품리스트 파일", expanded=True):
                    c1, c2 = st.columns(2)
                    with c1:
                        if xl and os.path.exists(xl):
                            with open(xl, "rb") as f:
                                st.download_button(
                                    label=f"📊 Excel 다운로드 ({os.path.basename(xl)})",
                                    data=f.read(),
                                    file_name=os.path.basename(xl),
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key=f"dl_xl_{code}"
                                )
                    with c2:
                        if pdf and os.path.exists(pdf):
                            with open(pdf, "rb") as f:
                                st.download_button(
                                    label=f"📑 PDF 다운로드 ({os.path.basename(pdf)})",
                                    data=f.read(),
                                    file_name=os.path.basename(pdf),
                                    mime="application/pdf",
                                    key=f"dl_pdf_{code}"
                                )
                        else:
                            st.caption("(PDF는 서버 환경에 따라 생성되지 않을 수 있으며, 엑셀에서 바로 인쇄 가능합니다.)")
