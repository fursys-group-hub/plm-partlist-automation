"""
PLM Part List 자동화 메인 스크립트 - Version 8
- PLM 로그인 -> BOM > PARTLIST 관리 -> 단품코드 일괄 검색 ->
  단품별 자재 색상 등록 -> PART LIST 작성 페이지에서 데이터 추출 -> Excel 저장
- v5: 다중 코드 일괄 처리(Batch), V5 명명 규칙 적용
"""
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.chrome import ChromeDriverManager

import time

import getpass
import os

try:
    import pyautogui
except ImportError:
    pyautogui = None

from plm_part_list_generator_v8 import (
    extract_grid_data, process_and_generate,
)


def setup_driver(browser=None, headless=None):
    """
    브라우저 드라이버 초기화:
    - Windows 데스크톱 기본: browser='edge', headless=False
    - Linux / 서버(Docker) 기본: browser='chrome', headless=True
    """
    is_linux = (sys.platform != 'win32')
    if headless is None:
        headless = is_linux
    if browser is None:
        browser = 'chrome' if is_linux else 'edge'

    if browser == 'edge' and not is_linux:
        try:
            options = webdriver.EdgeOptions()
            if headless:
                options.add_argument('--headless=new')
                options.add_argument('--disable-gpu')
                options.add_argument('--window-size=1920,1080')
            service = EdgeService()
            driver = webdriver.Edge(service=service, options=options)
            if not headless:
                driver.maximize_window()
            print("[드라이버] Microsoft Edge 사용")
            return driver
        except Exception as e:
            print(f"[경고] Edge 드라이버 실패: {e}")
            print("Chrome으로 전환합니다...")
            browser = 'chrome'

    # Chrome
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument('--headless=new')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')
        options.add_argument('--window-size=1920,1080')
        options.add_argument('--disable-extensions')
        options.add_argument('--disable-setuid-sandbox')

    chrome_bin = os.environ.get("CHROME_BIN")
    if chrome_bin and os.path.exists(chrome_bin):
        options.binary_location = chrome_bin

    driver_path = os.environ.get("CHROMEDRIVER_PATH")
    if driver_path and os.path.exists(driver_path):
        service = ChromeService(executable_path=driver_path)
    else:
        service = ChromeService(ChromeDriverManager().install())

    driver = webdriver.Chrome(service=service, options=options)
    if not headless:
        driver.maximize_window()
    print(f"[드라이버] Google Chrome 사용 (headless={headless})")
    return driver


def dump_debug_info(driver, step_name):
    """디버그 정보(HTML, 스크린샷) 저장"""
    print(f"  [디버그 저장] {step_name}...")
    try:
        with open(f"debug_{step_name}.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        driver.save_screenshot(f"debug_{step_name}.png")
        print(f"  -> debug_{step_name}.html / .png 저장됨")
    except Exception as e:
        print(f"  [경고] 디버그 저장 실패: {e}")


def switch_to_frame_containing(driver, text):
    """지정된 텍스트를 포함한 iframe으로 전환"""
    frames = driver.find_elements(By.TAG_NAME, "iframe")
    for i, frame in enumerate(frames):
        try:
            driver.switch_to.frame(frame)
            if text in driver.page_source:
                return True
            driver.switch_to.default_content()
        except Exception:
            driver.switch_to.default_content()
    return False


def login(driver, url, username, password):
    """
    PLM 최초 1회 로그인 로직
    """
    print("[1] PLM 접속 중...")
    
    # URL에 아이디/비밀번호를 포함하여 접속하면 브라우저가 해당 도메인에 대한 인증 정보를 
    # 세션 동안 캐시하게 되므로, 이후 새 창(팝업)이 열려도 추가 로그인 창이 뜨지 않습니다.
    from urllib.parse import urlparse, quote
    parsed = urlparse(url)
    auth_url = f"{parsed.scheme}://{quote(username)}:{quote(password)}@{parsed.netloc}{parsed.path}"
    if parsed.query: auth_url += f"?{parsed.query}"
    if parsed.fragment: auth_url += f"#{parsed.fragment}"

    driver.get(auth_url)
    print(f"  -> 페이지 요청 완료")

    print("  -> 메인 페이지 로드 완료 대기 중 (최대 60초)...")
    main_loaded = False
    main_selectors = [
        (By.XPATH, "//*[contains(text(), 'BOM')]"),
        (By.XPATH, "//*[contains(text(), 'PARTLIST')]"),
        (By.ID,    "mainMenu"),
        (By.ID,    "GNB"),
        (By.XPATH, "//*[contains(@class, 'gnb') or contains(@class, 'GNB')]"),
    ]
    for selector in main_selectors:
        try:
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located(selector)
            )
            main_loaded = True
            print(f"  -> 메인 페이지 로드 완료! ({selector[1]})")
            break
        except Exception:
            pass

    if not main_loaded:
        print("  [경고] 메인 페이지 자동 감지 실패. 현재 상태에서 계속 진행합니다.")
    time.sleep(2)


def process_single_code(driver, target_code):
    """
    단일 코드 처리.
    Returns: (output_file, pdf_row, pdf_col) on success, None on failure.
    """
    print(f"\n[2] BOM 메뉴 탐색 및 PARTLIST 관리 진입 중... (타겟: '{target_code}')")
    version_str = "R000"

    # 1. 이미 PARTLIST 관리 화면(iframe#body)에 진입해 있는지 확인 (불필요한 메뉴 재클릭 및 iframe 리로드 방지)
    driver.switch_to.default_content()
    is_already_on_partlist = False
    try:
        driver.switch_to.frame("body")
        if driver.find_elements(By.NAME, "erpCode"):
            is_already_on_partlist = True
            print("  -> 이미 PARTLIST 관리 화면에 위치해 있습니다. (메뉴 재진입 생략)")
        driver.switch_to.default_content()
    except Exception:
        driver.switch_to.default_content()

    if not is_already_on_partlist:
        if not switch_to_frame_containing(driver, "PARTLIST 관리"):
            switch_to_frame_containing(driver, "BOM")

        partlist_clicked = False
        try:
            partlist_menu = WebDriverWait(driver, 2).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//*[contains(text(), 'PARTLIST 관리') or contains(text(), 'PartList') or contains(text(), 'PART LIST')]")
                )
            )
            partlist_menu.click()
            partlist_clicked = True
            print("  -> 이미 펼쳐져 있는 메뉴에 바로 진입했습니다.")
        except Exception:
            pass

        if not partlist_clicked:
            print("  -> 닫혀있는 메뉴(BOM) 펼치기 시도...")
            try:
                bom_menu = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//*[contains(text(), 'BOM') or contains(text(), 'bom')]")
                    )
                )
                bom_menu.click()
            except Exception:
                pass

            time.sleep(1)

            print("[3] PARTLIST 관리 메뉴 클릭...")
            try:
                partlist_menu = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//*[contains(text(), 'PARTLIST 관리') or contains(text(), 'PartList') or contains(text(), 'PART LIST')]")
                    )
                )
                partlist_menu.click()
            except Exception:
                dump_debug_info(driver, "bom_expanded")
                raise Exception("PARTLIST 관리 메뉴를 찾지 못했습니다")

        time.sleep(2)

    print(f"[4] 단품코드 '{target_code}' 검색 준비 중...")
    driver.switch_to.default_content()

    try:
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "body"))
        )
    except Exception:
        frames = driver.find_elements(By.TAG_NAME, "iframe")
        if frames:
            driver.switch_to.frame(frames[-1])

    # 4-1. 이전 검색 상태 및 AUIGrid 완전 초기화 (이전 단품 데이터 잔류로 인한 오작동 방지)
    try:
        init_btn = driver.find_elements(By.ID, "searchInit")
        if init_btn and init_btn[0].is_displayed():
            driver.execute_script("arguments[0].click();", init_btn[0])
            time.sleep(0.2)
    except Exception:
        pass

    driver.execute_script("""
        try {
            $('input[name="erpCode"]').val('');
            $('input[name="number"]').val('');
            $('#creator').val('');
            $('#creatorOid').val('');
            $('#receptionist').val('');
            $('#receptionistOid').val('');
            var pid = (typeof myGridID !== 'undefined') ? myGridID : '#grid_wrap';
            try { AUIGrid.clearGridData(pid); } catch(e1) { AUIGrid.setGridData(pid, []); }
        } catch(e) {}
    """)
    time.sleep(0.3)

    # 4-2. 단품코드 입력
    try:
        code_input = WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located((By.NAME, "erpCode"))
        )
    except Exception:
        code_input = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.NAME, "erpCode"))
        )

    driver.execute_script("arguments[0].scrollIntoView(true);", code_input)
    code_input.click()
    code_input.send_keys(Keys.CONTROL + "a")
    code_input.send_keys(Keys.BACKSPACE)
    code_input.send_keys(target_code)
    driver.execute_script("""
        arguments[0].value = arguments[1];
        $(arguments[0]).trigger('change');
    """, code_input, target_code)
    print(f"  -> '{target_code}' 입력 완료")

    # 4-3. 사용자 클리어 버튼(axi-close2) fallback
    try:
        clear_btns = driver.find_elements(
            By.CSS_SELECTOR, ".axi-close2.clearUser, .axicon.clearUser, .axi-close2"
        )
        for btn in clear_btns:
            try:
                if btn.is_displayed():
                    driver.execute_script("arguments[0].click();", btn)
                    time.sleep(0.2)
            except Exception:
                continue
    except Exception:
        pass

    # 4-4. 검색 버튼 클릭
    try:
        search_btn = driver.find_element(By.ID, "searchBtn")
    except Exception:
        search_btn = driver.find_element(
            By.XPATH, "//button[contains(text(), '검색') or contains(text(), '조회') or @type='submit']"
        )
    driver.execute_script("arguments[0].click();", search_btn)
    print("  -> 검색 버튼 클릭")

    print(f"[5] '{target_code}' 검색 결과 로딩 대기 중 (최대 60초)...")
    time.sleep(1)

    # 마스크 레이어 해제 대기
    try:
        driver.switch_to.default_content()
        WebDriverWait(driver, 5).until(
            EC.invisibility_of_element_located((By.XPATH, "//*[contains(text(), '처리중입니다')]"))
        )
    except Exception:
        pass
    finally:
        try:
            WebDriverWait(driver, 10).until(
                EC.frame_to_be_available_and_switch_to_it((By.ID, "body"))
            )
        except Exception:
            frames = driver.find_elements(By.TAG_NAME, "iframe")
            if frames:
                driver.switch_to.frame(frames[-1])

    # AUIGrid 행 데이터 로드 및 타겟 단품코드 일치 확인 루프
    check_search_js = """
        try {
            var target = String(arguments[0] || '').trim().toUpperCase();
            var pid = (typeof myGridID !== 'undefined') ? myGridID : '#grid_wrap';
            
            var loader = document.querySelector('.aui-grid-ajax-loader, .aui-grid-loader');
            var isLoading = loader && loader.offsetParent !== null && loader.style.display !== 'none';
            
            var noData = false;
            var noDataEls = document.getElementsByClassName('aui-grid-nodata-msg-layer');
            for (var i = 0; i < noDataEls.length; i++) {
                if (noDataEls[i].offsetParent !== null && noDataEls[i].style.display !== 'none') {
                    noData = true;
                    break;
                }
            }
            if (document.body && document.body.innerText.indexOf("No Data to display") >= 0) {
                noData = true;
            }
            
            var data = null;
            try { data = AUIGrid.getGridData(pid); } catch(e) {}
            if (!data) data = [];
            
            var matchCount = 0;
            var matchIdx = -1;
            for (var i = 0; i < data.length; i++) {
                var ec = String(data[i].erpCode || '').trim().toUpperCase();
                var no = String(data[i].number || '').trim().toUpperCase();
                if (ec === target || ec.indexOf(target) >= 0 || target.indexOf(ec) >= 0 || no.indexOf(target) >= 0) {
                    matchCount++;
                    if (matchIdx === -1) matchIdx = i;
                }
            }
            
            return {
                isLoading: isLoading,
                noData: noData,
                rowCount: data.length,
                matchCount: matchCount,
                matchIdx: matchIdx
            };
        } catch(e) {
            return { error: e.toString(), matchCount: 0 };
        }
    """

    has_data = False
    matched_row_idx = 0
    for wait_i in range(60):
        try:
            status = driver.execute_script(check_search_js, target_code)
            if isinstance(status, dict):
                match_cnt = status.get('matchCount', 0)
                is_loading = status.get('isLoading', False)
                no_data = status.get('noData', False)
                row_cnt = status.get('rowCount', 0)

                if match_cnt > 0:
                    has_data = True
                    matched_row_idx = status.get('matchIdx', 0)
                    print(f"  -> 검색 결과 로딩 완료: '{target_code}' 일치 행 감지 (총 {row_cnt}개 행 중 {match_cnt}개 일치, 행 인덱스: {matched_row_idx})")
                    break

                if not is_loading and no_data and wait_i >= 2:
                    print(f"  [결과없음] '{target_code}' 검색 결과 없음 (No Data)")
                    return None
        except Exception:
            pass
        time.sleep(1)

    if not has_data:
        # DOM 요소 직접 체크 (fallback)
        try:
            xpath_match = (
                f"//div[contains(@class,'aui-grid-body-panel')]//table//tr"
                f"[contains(@style,'height') and not(contains(@style,'display: none')) and contains(., '{target_code}')]"
            )
            rows = driver.find_elements(By.XPATH, xpath_match)
            if rows:
                has_data = True
                print(f"  -> DOM 기준 타겟 일치 행 감지 ({len(rows)}개)")
        except Exception:
            pass

    if not has_data:
        print(f"  [타임아웃] '{target_code}' 검색 결과를 시간 내에 가져오지 못했습니다.")
        dump_debug_info(driver, "result_click_fail")
        return None

    print(f"[6] 결과 항목 열기 (행 {matched_row_idx} 타겟)...")
    popup_opened = False
    init_handles = set(driver.window_handles)

    # 방법 A: AUIGrid cellDoubleClick 핸들러 로직 직접 JS 실행 (타겟 행 인덱스 명시)
    try:
        js_trigger_result = driver.execute_script("""
            try {
                var pid = (typeof myGridID !== 'undefined') ? myGridID : '#grid_wrap';
                var data = AUIGrid.getGridData(pid);
                var targetIdx = arguments[0];
                if (data && data.length > 0) {
                    var rowItem = (targetIdx >= 0 && targetIdx < data.length) ? data[targetIdx] : data[0];
                    var oid = "";
                    if (rowItem.oid && rowItem.oid.indexOf("EBOMMaster") > 0) {
                        var groupList = AUIGrid.getItemsByValue(pid, "groupId", rowItem.groupId);
                        groupList = groupList.filter(function(item){
                            return item.oid.indexOf("PARTLISTMaster") > 0;
                        });
                        for (var i = 0; i < groupList.length; i++) {
                            if (i == 0) oid = groupList[i].oid;
                            else oid += "," + groupList[i].oid;
                        }
                    } else {
                        oid = rowItem.oid;
                    }
                    if (rowItem.type != "ebom" && oid) {
                        var url = _url("/partlist/create", oid);
                        _popup(url, "", "", "f");
                        return true;
                    }
                }
            } catch(e) {
                return false;
            }
            return false;
        """, matched_row_idx)
        if js_trigger_result:
            print("  -> JS를 통한 팝업 호출 실행 성공")
            for _ in range(8):
                time.sleep(0.5)
                if len(driver.window_handles) > len(init_handles):
                    popup_opened = True
                    print("  -> 새 팝업 창 감지 완료 (방법 A)")
                    break
    except Exception as e_js:
        print(f"  [알림] JS 팝업 트리거 실패 ({e_js}), 마우스 더블클릭 시도...")

    # 방법 B: JS 팝업 실패 시 Selenium ActionChains 더블클릭 시도 (타겟 코드 행 우선)
    if not popup_opened:
        try:
            xpath_match = (
                f"//div[contains(@class,'aui-grid-body-panel')]//table//tr"
                f"[contains(@style,'height') and not(contains(@style,'display: none')) and contains(., '{target_code}')]"
            )
            try:
                result_row = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, xpath_match))
                )
            except Exception:
                result_row = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(
                        (By.XPATH,
                         "//div[contains(@class,'aui-grid-body-panel')]//table//tr"
                         "[contains(@style,'height') and not(contains(@style,'display: none'))]")
                    )
                )
            tds = result_row.find_elements(By.TAG_NAME, "td")
            click_target = tds[4] if len(tds) > 4 else result_row
            ActionChains(driver).double_click(click_target).perform()
            print("  -> 마우스 더블클릭 수행 완료 (방법 B)")
        except Exception as e_dc:
            print(f"  [경고] 더블클릭 시도 실패: {e_dc}")
            # 방법 C: 상단 '등록' 버튼 클릭 fallback
            try:
                driver.execute_script("""
                    var pid = (typeof myGridID !== 'undefined') ? myGridID : '#grid_wrap';
                    var data = AUIGrid.getGridData(pid);
                    var targetIdx = arguments[0];
                    if (data && data.length > 0) {
                        var row = (targetIdx >= 0 && targetIdx < data.length) ? data[targetIdx] : data[0];
                        AUIGrid.setCheckedRowsByValue(pid, "oid", [row.oid]);
                        $('#createBtn').click();
                    }
                """, matched_row_idx)
                print("  -> '등록' 버튼 클릭 시도 (방법 C)")
            except Exception:
                dump_debug_info(driver, "result_click_fail")
                raise e_dc

    print("[7] 더블클릭 후 창 탐색 중...")
    new_window_appeared = False
    for w_i in range(60):
        time.sleep(1)
        if len(driver.window_handles) > len(init_handles):
            new_window_appeared = True
            break

    if new_window_appeared:
        new_handles = [h for h in driver.window_handles if h not in init_handles]
        if new_handles:
            driver.switch_to.window(new_handles[-1])
        else:
            driver.switch_to.window(driver.window_handles[-1])
        for _ in range(60):
            try:
                cur_url = driver.current_url.lower()
                if cur_url != 'about:blank' and cur_url != '':
                    break
            except Exception:
                pass
            time.sleep(1)

    all_urls = []
    for h in driver.window_handles:
        try:
            driver.switch_to.window(h)
            all_urls.append((h, driver.current_url))
        except Exception:
            pass

    batch_handle = None
    for h, u in all_urls:
        if 'partlist/batch' in u.lower():
            batch_handle = h
            print("  -> [경로 A] partlist/batch 발견 (클릭 생략)")
            break

    if batch_handle:
        driver.switch_to.window(batch_handle)
    else:
        print("  -> [경로 B] 팝업에서 버튼 클릭 시도...")
        if len(driver.window_handles) > 1:
            driver.switch_to.window(driver.window_handles[-1])

        try:
            popup_row = WebDriverWait(driver, 60).until(
                EC.element_to_be_clickable(
                    (By.XPATH,
                     "//div[contains(@class,'aui-grid-body-panel')]//table//tr"
                     "[contains(@style,'height') and not(contains(@style,'display: none'))]")
                )
            )
            
            try:
                td_elements = popup_row.find_elements(By.TAG_NAME, "td")
                for td_idx in [10, 9, 11]:
                    if len(td_elements) > td_idx:
                        v_text = td_elements[td_idx].text.strip()
                        if v_text.isdigit():
                            version_str = f"R{int(v_text):03d}"
                            break
            except Exception:
                pass

            popup_row.click()
            time.sleep(1)

            color_btn = WebDriverWait(driver, 60).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//*[contains(text(), '단품별 자재 색상 등록')]")
                )
            )
            color_btn.click()
        except Exception as e:
            dump_debug_info(driver, "popup_fail")
            raise e

    print("[8] PART LIST 작성 창 탐색 중...")
    target_handle = None
    for wait_i in range(60):
        time.sleep(1)
        for handle in driver.window_handles:
            try:
                driver.switch_to.window(handle)
                cur_url = driver.current_url.lower()
                if 'partlist/batch' in cur_url:
                    target_handle = handle
                    break
            except Exception:
                pass
        if target_handle:
            break

    if not target_handle:
        for wait_i2 in range(30):
            time.sleep(1)
            for handle in driver.window_handles:
                try:
                    driver.switch_to.window(handle)
                    cur_url = driver.current_url.lower()
                    if 'partlist/create' in cur_url:
                        target_handle = handle
                        break
                except Exception:
                    pass
            if target_handle:
                break

    if not target_handle:
        driver.switch_to.window(driver.window_handles[-1])
    else:
        driver.switch_to.window(target_handle)

    print("  -> 페이지 로드 대기 중...")
    try:
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.ID, "grid_wrap"))
        )
    except Exception:
        dump_debug_info(driver, "grid_extraction_fail")

    check_js = """
    try {
        if (typeof AUIGrid === 'undefined') return -1;
        var pid = '#grid_wrap';
        if (typeof myGridID !== 'undefined' && myGridID) pid = myGridID;
        var rc = -3;
        try { rc = AUIGrid.getRowCount(pid); } catch(e1) {}
        if (rc >= 0) return rc;
        var d = null;
        try { d = AUIGrid.getGridData(pid); } catch(e2) {}
        if (!d || d.length === 0) return 0;
        var cnt = 0;
        function countAll(items) {
            items.forEach(function(it) { cnt++; if (it.children) countAll(it.children); });
        }
        countAll(d);
        return cnt;
    } catch(e) { return -2; }
    """
    prev_count = -99
    stable_count = 0
    for poll_i in range(60):
        time.sleep(4)
        cur_count = driver.execute_script(check_js)
        if cur_count > 0 and cur_count == prev_count:
            stable_count += 1
            if stable_count >= 2:
                print(f"  -> 로딩 완료 확인! (총 {cur_count}개)")
                break
        else:
            stable_count = 0
        prev_count = cur_count

    print("[9] AUIGrid 데이터 추출 중...")
    driver_result = extract_grid_data(driver)

    try:
        auth_cookies = {c['name']: c['value'] for c in driver.get_cookies()}
        driver_result['cookies'] = auth_cookies
    except Exception as e:
        print(f"  [경고] 쿠키 추출 실패: {e}")

    grid_data = driver_result.get('grid_data', [])
    if grid_data:
        import re
        product_name = ""
        extracted_code = ""
        try:
            true_roots = [it for it in grid_data if int(it.get('_$depth', 0)) <= 1]
            if true_roots:
                product_name = str(true_roots[0].get('partName') or '').strip()
                extracted_code = str(true_roots[0].get('erpCode') or true_roots[0].get('partNo') or '').strip()
        except Exception:
            pass

        print(f"  -> 추출 데이터 확인: 타겟='{target_code}', 추출단품코드='{extracted_code}', 단품명='{product_name}'")
        if extracted_code and (target_code.upper() not in extracted_code.upper() and extracted_code.upper() not in target_code.upper()):
            print(f"  [주의!] 타겟 단품코드({target_code})와 추출된 코드({extracted_code})가 일치하지 않습니다!")

        safe_product_name = re.sub(r'[\\/:*?"<>|]', '_', product_name)

        if safe_product_name:
            output_file = f"{target_code}_{safe_product_name}_PartList.xlsx"
        else:
            output_file = f"{target_code}_PartList.xlsx"

        result = process_and_generate(
            driver_result, output_file,
            product_code=target_code, product_name=product_name, driver=driver, version_str=version_str
        )
        if result:
            result_file, pdf_row, pdf_col = result
            print(f"[완료] {result_file} 생성 완료!")
            return (result_file, pdf_row, pdf_col)
        else:
            print("[실패] 항목 선정 실패")
            dump_debug_info(driver, "grid_extraction_fail")
    else:
        print("[실패] 데이터 추출 실패")
        dump_debug_info(driver, "grid_extraction_fail")

    return None


if __name__ == "__main__":
    URL = "https://plm.fursys.com/Windchill/platform/main#"

    print("=" * 55)
    print("  PLM PART LIST 다중 단품코드 자동화 (v8)")
    print("  ※ 조립도 PDF 자동 첨부 기능 포함")
    print("=" * 55)
    plm_id = input("PLM 아이디: ").strip()
    plm_pw = getpass.getpass("PLM 비밀번호: ")

    print("\n■ 타겟 단품코드 입력 ■")
    print("여러 개일 경우 엑셀에서 세로로 여러 셀을 드래그해 복사한 후 창에 붙여넣기 하세요.")
    print("모두 붙여넣은 뒤, 엔터를 한 번 더 치면 작업이 시작됩니다.")
    lines = []
    while True:
        try:
            line = input()
            if not line.strip():
                break
            lines.append(line)
        except EOFError:
            break

    target_codes = []
    for line in lines:
        for code in line.replace(',', ' ').split():
            c = code.strip()
            if c and c not in target_codes:
                target_codes.append(c)

    print("=" * 55)

    if not plm_id or not plm_pw or not target_codes:
        print("[오류] 지정된 입력 정보가 부족합니다. 종료합니다.")
        sys.exit(1)

    driver = setup_driver()
    try:
        # 로그인 파트 (최초 1회만 수행)
        login(driver, URL, plm_id, plm_pw)
        main_window_handle = driver.current_window_handle

        print(f"\n총 {len(target_codes)}개의 단품코드에 대해 일괄 연속 작업을 시작합니다.")

        for code_idx, code in enumerate(target_codes):
            print(f"\n{'='*55}")
            print(f"  [{code_idx+1}/{len(target_codes)}] >>>> 단품코드: '{code}' <<<<")
            print(f"{'='*55}")

            result_info = None
            try:
                driver.switch_to.default_content()
                result_info = process_single_code(driver, code)
            except Exception as e:
                print(f"[오류] '{code}' 처리 중 에러 발생: {e}")
                import traceback
                traceback.print_exc()

            # 파생 창 모두 닫기 (메인 창만 유지)
            for handle in driver.window_handles:
                if handle != main_window_handle:
                    try:
                        driver.switch_to.window(handle)
                        driver.close()
                    except Exception:
                        pass

            try:
                driver.switch_to.window(main_window_handle)
            except Exception as e:
                print(f"  [치명적 오류] 메인 창 손실 (세션 종료). {e}")
                break

            time.sleep(1)

    finally:
        input("\n>>> 모든 일괄 작업이 끝났습니다. 브라우저를 닫으려면 Enter를 누르세요...")
        try:
            driver.quit()
        except Exception:
            pass
