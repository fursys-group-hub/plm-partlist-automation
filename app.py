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
from plm_part_list_automation_v8 import setup_driver, login, process_single_code, check_plm_network
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
        margin-bottom: 1.2rem;
    }
    /* Deploy 버튼 숨김 (사내 배포 환경 최적화) */
    .stDeployButton, [data-testid="stDeployButton"], .stAppDeployButton, [data-testid="stAppDeployButton"] {
        display: none !important;
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
    .guide-label {
        font-size: 14px;
        color: #31333F;
        margin-bottom: 0.5rem;
        font-weight: 400;
        line-height: 1.5;
        height: 21px;
    }
    .guide-card {
        background-color: #F0F7FF;
        border: 1px solid #BFDBFE;
        border-radius: 8px;
        padding: 0 1.25rem;
        height: 220px;
        box-sizing: border-box;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    .guide-list {
        margin: 0;
        padding-left: 1.25rem;
        font-size: 0.85rem;
        color: #1E293B;
    }
    .guide-list li {
        margin-bottom: 0.52rem;
        line-height: 1.35;
    }
    .guide-list li:last-child {
        margin-bottom: 0;
    }
    .guide-item-title {
        font-weight: 700;
        color: #0F172A;
        font-size: 0.875rem;
    }
    .guide-item-desc {
        color: #475569;
        font-size: 0.81rem;
        margin-top: 2px;
        line-height: 1.3;
    }
    .link-badge {
        display: inline-flex;
        align-items: center;
        gap: 3px;
        background-color: #2563EB;
        color: #FFFFFF !important;
        font-size: 0.72rem;
        font-weight: 600;
        padding: 2px 7px;
        border-radius: 4px;
        text-decoration: none !important;
        vertical-align: middle;
        margin-left: 6px;
        box-shadow: 0 1px 2px rgba(37,99,235,0.25);
        transition: all 0.15s ease;
    }
    .link-badge:hover {
        background-color: #1D4ED8;
        color: #FFFFFF !important;
        box-shadow: 0 2px 4px rgba(29,78,216,0.3);
        transform: translateY(-1px);
    }
</style>
""", unsafe_allow_html=True)

# Sidebar: Credentials & System Info
with st.sidebar:
    st.markdown("<h2 style='color: #1E3A8A; font-weight: 800; margin: 0; padding-top: 0;'>🏢 FURSYS</h2>", unsafe_allow_html=True)
    st.markdown("<div style='font-size: 0.85rem; color: #64748B; margin-bottom: 1rem;'>퍼시스그룹 업무 자동화</div>", unsafe_allow_html=True)
    # Windows 환경에서 기존 저장된 자격 증명이 있을 경우 기본값으로 로드
    default_id = ""
    default_pw = ""
    is_linux = (sys.platform != 'win32')
    if not is_linux:
        config_path = os.path.join(
            os.environ.get("APPDATA", os.path.expanduser("~")),
            "PLM_PartList_Automation",
            "login_config.json"
        )
        if os.path.exists(config_path):
            try:
                import json, base64
                with open(config_path, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                if cfg.get("remember", False):
                    default_id = cfg.get("id", "")
                    saved_pw_b64 = cfg.get("pw", "")
                    if saved_pw_b64:
                        default_pw = base64.b64decode(saved_pw_b64.encode("utf-8")).decode("utf-8")
            except Exception:
                pass

    st.markdown("### 🔐 PLM 계정 설정")
    plm_id = st.text_input("PLM 아이디", value=default_id, placeholder="예: yeongjo_kwon", help="사내 PLM 로그인 아이디")
    plm_pw = st.text_input("PLM 비밀번호", value=default_pw, type="password", placeholder="비밀번호 입력", help="사내 PLM 로그인 비밀번호")

    if not is_linux:
        with st.expander("🛠️ 실행 옵션 (로컬 테스트용)", expanded=False):
            headless_mode = st.checkbox("브라우저 백그라운드 실행 (Headless)", value=False, help="체크 해제 시 실제 브라우저가 화면에 열려 작업 진행 과정을 직접 볼 수 있습니다.")
            browser_type = st.selectbox("사용 브라우저", ["edge", "chrome"], index=0)
    else:
        headless_mode = True
        browser_type = 'chrome'
        
    st.divider()
    st.markdown("### ⚙️ 시스템 정보")
    st.markdown("- **버전**: `V8.0`")
    
    st.caption("© 2026 Fursys Group. All rights reserved.")

# Main Header
st.markdown('<div class="main-header">📋 PLM 부품리스트 자동 생성기</div>', unsafe_allow_html=True)

# Main Form
col_input, col_info = st.columns([1, 1])

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
    st.markdown('<div class="guide-label">사전 준비 작업 및 시스템 이용 절차 안내</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="guide-card">
        <ol start="0" class="guide-list">
            <li>
                <div class="guide-item-title">깔끔한 부품리스트 생성을 위해선 사전 준비 작업이 필요합니다.</div>
                <div class="guide-item-desc">: 사전 준비 작업 가이드 <a href="https://fursys.atlassian.net/wiki/external/YTc0MzUzOWYxMDIxNDZjY2I4MDExMWM5N2M3MGE3ZmI" target="_blank" class="link-badge"><span>🔗 가이드 링크 ↗</span></a></div>
            </li>
            <li>
                <div class="guide-item-title">단품코드 입력</div>
                <div class="guide-item-desc">: 직접입력 or 여러 개의 단품코드를 엑셀에서 복사해 그대로 붙여넣을 수 있습니다.</div>
            </li>
            <li>
                <div class="guide-item-title">부품리스트 자동 생성 시작 버튼 클릭</div>
                <div class="guide-item-desc">: 버튼 클릭 후 대기해 주세요.</div>
            </li>
            <li>
                <div class="guide-item-title">결과 다운로드</div>
                <div class="guide-item-desc">: 작업 완료 후 생성된 엑셀과 PDF를 즉시 다운로드하실 수 있습니다.</div>
            </li>
        </ol>
    </div>
    """, unsafe_allow_html=True)

# ==========================================
# Session State Initialization for Persistent Downloads
# ==========================================
if 'generated_files' not in st.session_state:
    st.session_state['generated_files'] = []
    # 앱 재로딩 시 최근 생성된 부품리스트 파일을 감지하여 다운로드 섹션 유지
    if not is_linux:
        try:
            recent = []
            for f in os.listdir("."):
                if f.endswith("_PartList.xlsx") and not f.startswith("~$"):
                    code = f.split("_")[0]
                    xl_p = os.path.abspath(f)
                    pdf_p = xl_p.replace(".xlsx", ".pdf")
                    recent.append((code, xl_p, pdf_p if os.path.exists(pdf_p) else None))
            if recent:
                recent.sort(key=lambda x: os.path.getmtime(x[1]), reverse=True)
                st.session_state['generated_files'] = recent[:10]
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for c, xl, pdf in st.session_state['generated_files']:
                        if xl and os.path.exists(xl):
                            zf.write(xl, os.path.basename(xl))
                        if pdf and os.path.exists(pdf):
                            zf.write(pdf, os.path.basename(pdf))
                st.session_state['zip_buffer'] = zip_buf.getvalue()
                st.session_state['zip_name'] = f"Fursys_PartLists_{time.strftime('%Y%m%d_%H%M%S')}.zip"
        except Exception:
            pass

if 'zip_buffer' not in st.session_state:
    st.session_state['zip_buffer'] = None

# Execution logic
if start_clicked:
    if not plm_id or not plm_pw:
        st.error("⚠️ 좌측 사이드바에서 PLM 아이디와 비밀번호를 모두 입력해 주세요.")
    elif not raw_codes.strip():
        st.error("⚠️ 생성할 단품코드를 최소 1개 이상 입력해 주세요.")
    else:
        # 새 작업 시작 시 이전 세션 결과 초기화
        st.session_state['generated_files'] = []
        st.session_state['zip_buffer'] = None

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
            # Display last 10 log lines
            log_placeholder.code('\n'.join(logs[-10:]), language="bash")

        class LogCapturer:
            def __init__(self, callback, original_stdout):
                self.callback = callback
                self.original_stdout = original_stdout
            def write(self, s):
                self.original_stdout.write(s)
                line = s.strip()
                if line:
                    self.callback(line)
            def flush(self):
                self.original_stdout.flush()

        driver = None
        generated_files = []
        orig_stdout = sys.stdout
        sys.stdout = LogCapturer(log_msg, orig_stdout)
        
        try:
            log_msg(f"[1/4] 브라우저 드라이버 시작 중... ({browser_type.upper()}{', 헤드리스' if headless_mode else ''})")
            driver = setup_driver(browser=browser_type, headless=headless_mode)
            
            log_msg("[2/4] PLM 시스템 자동 로그인 시도 중...")
            PLM_URL = "https://plm.fursys.com/Windchill/platform/main#"
            
            # 사전 네트워크 진단 (Coolify 배포 서버 -> PLM 서버 통신 가능 여부)
            net_ok, diag_msg = check_plm_network("plm.fursys.com", 443, timeout=5)
            log_msg(f"  -> 네트워크 진단: {diag_msg}")
            if not net_ok:
                raise ConnectionError(
                    f"사내 배포 서버에서 PLM 서버(plm.fursys.com:443)로의 네트워크 연결이 불가능합니다.\n"
                    f"상세 사유: {diag_msg}\n"
                    f"💡 원인 분석 및 조치 방안:\n"
                    f"1. Coolify 배포 서버(10.x / 클라우드 환경)에서 사내 인트라넷 PLM 서버(192.9.200.12)로의 아웃바운드 443 포트 방화벽 오픈(또는 사내망 VPN/라우팅)이 필요합니다.\n"
                    f"2. 사내 IT/인프라팀에 'Coolify 서버 -> PLM 서버(192.9.200.12:443) 통신 허용'을 요청하시거나, 로컬 PC용 'PARTLIST_Automation_v8.0.exe'를 사용해 주세요."
                )

            login(driver, PLM_URL, plm_id, plm_pw)
            log_msg("-> PLM 로그인 단계 완료!")
            
            main_window_handle = driver.current_window_handle
            
            total = len(target_codes)
            for idx, code in enumerate(target_codes):
                status_box.update(label=f"[{idx+1}/{total}] 단품코드 '{code}' BOM 조회 및 부품리스트 작성 중...")
                log_msg(f"\n--- [{idx+1}/{total}] 단품코드 '{code}' 처리 시작 ---")
                
                # 시작 전 남아있는 자식 팝업 정리 및 메인 창 컨텍스트 복귀
                try:
                    for handle in list(driver.window_handles):
                        if handle != main_window_handle:
                            try:
                                driver.switch_to.window(handle)
                                try:
                                    driver.switch_to.alert.accept()
                                except Exception:
                                    pass
                                driver.close()
                            except Exception:
                                pass
                    if main_window_handle in driver.window_handles:
                        driver.switch_to.window(main_window_handle)
                    elif driver.window_handles:
                        main_window_handle = driver.window_handles[0]
                        driver.switch_to.window(main_window_handle)
                    driver.switch_to.default_content()
                except Exception:
                    pass
                
                try:
                    res = process_single_code(driver, code)
                    if res:
                        excel_path = res[0]
                        pdf_path = excel_path.replace(".xlsx", ".pdf")
                        generated_files.append((code, excel_path, pdf_path if os.path.exists(pdf_path) else None))
                        log_msg(f"-> '{code}' 부품리스트 생성 완료: {os.path.basename(excel_path)}")
                    else:
                        log_msg(f"-> [오류] '{code}' 처리 실패 (단품코드 또는 PLM 상태 확인 필요)")
                except Exception as code_err:
                    log_msg(f"-> [오류] '{code}' 처리 중 예외 발생: {code_err}")
                finally:
                    # 파생 팝업 창 닫기 및 메인 창으로 안전 복귀
                    for handle in list(driver.window_handles):
                        if handle != main_window_handle:
                            try:
                                driver.switch_to.window(handle)
                                try:
                                    driver.switch_to.alert.accept()
                                except Exception:
                                    pass
                                driver.close()
                            except Exception:
                                pass
                    try:
                        if main_window_handle in driver.window_handles:
                            driver.switch_to.window(main_window_handle)
                        elif driver.window_handles:
                            main_window_handle = driver.window_handles[0]
                            driver.switch_to.window(main_window_handle)
                        driver.switch_to.default_content()
                    except Exception:
                        pass
                    time.sleep(1)
                    
                progress_bar.progress((idx + 1) / total)
                
            status_box.update(label="🎉 모든 단품코드 처리 완료!", state="complete", expanded=False)
            
            # 생성된 파일 목록 및 ZIP을 세션 상태에 저장하여 다운로드 클릭 시 화면 유지 보장
            if generated_files:
                st.session_state['generated_files'] = generated_files
                if len(generated_files) > 1 or any(f[2] for f in generated_files):
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                        for c, xl, pdf in generated_files:
                            if xl and os.path.exists(xl):
                                zf.write(xl, os.path.basename(xl))
                            if pdf and os.path.exists(pdf):
                                zf.write(pdf, os.path.basename(pdf))
                    st.session_state['zip_buffer'] = zip_buf.getvalue()
                    st.session_state['zip_name'] = f"Fursys_PartLists_{time.strftime('%Y%m%d_%H%M%S')}.zip"
            
        except Exception as e:
            status_box.update(label=f"❌ 처리 중 에러 발생: {e}", state="error")
            st.error(f"오류 상세: {e}")
            import traceback
            st.code(traceback.format_exc())
        finally:
            sys.stdout = orig_stdout
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass

# ==========================================
# Results & Download Section (Persistent via Session State)
# ==========================================
if st.session_state.get('generated_files'):
    results = st.session_state['generated_files']
    st.divider()
    st.markdown("### 📥 생성된 부품리스트 다운로드")
    
    # 1. 전체 ZIP 일괄 다운로드
    if st.session_state.get('zip_buffer'):
        st.download_button(
            label="📦 전체 결과 파일 일괄 다운로드 (ZIP)",
            data=st.session_state['zip_buffer'],
            file_name=st.session_state.get('zip_name', f"Fursys_PartLists_{time.strftime('%Y%m%d_%H%M%S')}.zip"),
            mime="application/zip",
            use_container_width=True,
            key="dl_zip_all"
        )
        st.markdown("<br>", unsafe_allow_html=True)
    
    # 2. 단품별 개별 파일 다운로드
    for code, xl, pdf in results:
        with st.expander(f"📄 {code} 부품리스트 파일", expanded=True):
            c1, c2 = st.columns(2)
            with c1:
                if xl and os.path.exists(xl):
                    with open(xl, "rb") as f:
                        xl_bytes = f.read()
                    st.download_button(
                        label=f"📊 Excel 다운로드 ({os.path.basename(xl)})",
                        data=xl_bytes,
                        file_name=os.path.basename(xl),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_xl_{code}"
                    )
                else:
                    st.warning("엑셀 파일을 찾을 수 없습니다.")
            with c2:
                if pdf and os.path.exists(pdf):
                    with open(pdf, "rb") as f:
                        pdf_bytes = f.read()
                    st.download_button(
                        label=f"📑 PDF 다운로드 ({os.path.basename(pdf)})",
                        data=pdf_bytes,
                        file_name=os.path.basename(pdf),
                        mime="application/pdf",
                        key=f"dl_pdf_{code}"
                    )
                else:
                    st.caption("(PDF 파일이 생성되지 않은 단품입니다.)")

