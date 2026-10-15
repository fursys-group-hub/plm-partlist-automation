import customtkinter as ctk
import threading
import sys
import os
from plm_part_list_automation_v8 import setup_driver, login, process_single_code
import time
import json
import base64

class ConsoleRedirector:
    def __init__(self, textbox):
        self.textbox = textbox

    def write(self, string):
        self.textbox.configure(state="normal")
        self.textbox.insert("end", string)
        self.textbox.see("end")
        self.textbox.configure(state="disabled")

    def flush(self):
        pass

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Apple-like aesthetic
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        
        # Set fonts based on user request ("Pretendard" handles both English and Korean very well)
        font_title = ctk.CTkFont(family="Proxima Nova", size=28, weight="bold")
        font_main = ctk.CTkFont(family="Pretendard", size=14)
        font_bold = ctk.CTkFont(family="Pretendard", size=15, weight="bold")
        font_btn = ctk.CTkFont(family="Pretendard", size=16, weight="bold")
        font_console = ctk.CTkFont(family="Pretendard", size=13)
        
        self.title("PLM Part List Automation")
        self.geometry("600x780")
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        self.main_frame = ctk.CTkFrame(self, corner_radius=15, fg_color="transparent")
        self.main_frame.grid(row=0, column=0, padx=30, pady=30, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        
        # Title
        self.title_label = ctk.CTkLabel(self.main_frame, text="PLM PartList Automation", font=font_title)
        self.title_label.grid(row=0, column=0, pady=(10, 30), sticky="w")
        


        # Config path for remembering credentials
        self.config_path = os.path.join(
            os.environ.get("APPDATA", os.path.expanduser("~")),
            "PLM_PartList_Automation",
            "login_config.json"
        )
        
        # ID / PW
        self.id_entry = ctk.CTkEntry(self.main_frame, placeholder_text="PLM 아이디", height=45, font=font_main, corner_radius=10)
        self.id_entry.grid(row=1, column=0, pady=(0, 15), sticky="ew")
        
        self.pw_entry = ctk.CTkEntry(self.main_frame, placeholder_text="PLM 비밀번호", show="*", height=45, font=font_main, corner_radius=10)
        self.pw_entry.grid(row=2, column=0, pady=(0, 12), sticky="ew")
        
        # Remember ID/PW Checkbox
        self.remember_var = ctk.BooleanVar(value=True)
        self.remember_cb = ctk.CTkCheckBox(
            self.main_frame,
            text="아이디 / 비밀번호 저장",
            variable=self.remember_var,
            font=font_main,
            corner_radius=6
        )
        self.remember_cb.grid(row=3, column=0, pady=(0, 20), sticky="w")
        
        # Target Codes Label
        self.codes_label = ctk.CTkLabel(self.main_frame, text="단품코드 목록 (엑셀에서 여러 셀 복사하여 붙여넣기)", font=font_bold)
        self.codes_label.grid(row=4, column=0, pady=(0, 10), sticky="w")
        
        self.codes_textbox = ctk.CTkTextbox(self.main_frame, height=180, font=font_main, corner_radius=10)
        self.codes_textbox.grid(row=5, column=0, pady=(0, 25), sticky="ew")
        
        # Start button
        self.start_btn = ctk.CTkButton(self.main_frame, text="자동화 시작", height=50, font=font_btn, corner_radius=10, command=self.start_automation)
        self.start_btn.grid(row=6, column=0, pady=(0, 20), sticky="ew")
        
        # Console output
        self.console_textbox = ctk.CTkTextbox(self.main_frame, height=150, font=font_console, fg_color="#F5F5F7", text_color="#1D1D1F", corner_radius=10)
        self.console_textbox.grid(row=7, column=0, pady=(0, 10), sticky="nsew")
        self.main_frame.grid_rowconfigure(7, weight=1)
        self.console_textbox.configure(state="disabled")
        
        # Footer Label (Version info)
        self.footer_label = ctk.CTkLabel(self.main_frame, text="업데이트: 2026.10.12 | 버전: V8", font=ctk.CTkFont(family="Pretendard", size=13), text_color="#5577AA")
        self.footer_label.grid(row=8, column=0, pady=(5, 0), sticky="e")
        
        # Redirect stdout
        sys.stdout = ConsoleRedirector(self.console_textbox)
        sys.stderr = sys.stdout

        # Load saved credentials if any
        self.load_saved_credentials()

        print("준비 완료! 아이디와 비밀번호, 단품코드를 붙여넣고 [자동화 시작] 버튼을 눌러주세요.")

    def load_saved_credentials(self):
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if data.get("remember", False):
                    saved_id = data.get("id", "")
                    saved_pw_b64 = data.get("pw", "")
                    saved_pw = ""
                    if saved_pw_b64:
                        saved_pw = base64.b64decode(saved_pw_b64.encode("utf-8")).decode("utf-8")
                    if saved_id:
                        self.id_entry.insert(0, saved_id)
                    if saved_pw:
                        self.pw_entry.insert(0, saved_pw)
                    self.remember_var.set(True)
                else:
                    self.remember_var.set(False)
        except Exception:
            pass

    def save_credentials(self, user_id, user_pw):
        try:
            if self.remember_var.get():
                os.makedirs(os.path.dirname(self.config_path), exist_ok=True)
                pw_b64 = base64.b64encode(user_pw.encode("utf-8")).decode("utf-8")
                with open(self.config_path, "w", encoding="utf-8") as f:
                    json.dump({"remember": True, "id": user_id, "pw": pw_b64}, f, ensure_ascii=False)
            else:
                if os.path.exists(self.config_path):
                    os.remove(self.config_path)
        except Exception:
            pass

    def start_automation(self):
        user_id = self.id_entry.get().strip()
        user_pw = self.pw_entry.get()
        codes_text = self.codes_textbox.get("1.0", "end")
        
        target_codes = []
        for line in codes_text.split('\n'):
            for code in line.replace(',', ' ').split():
                c = code.strip()
                if c and c not in target_codes:
                    target_codes.append(c)
                    
        if not user_id or not user_pw or not target_codes:
            print("[오류] 아이디, 비밀번호, 단품코드를 올바르게 입력해야 합니다.")
            return
            
        self.save_credentials(user_id, user_pw)
        self.start_btn.configure(state="disabled", text="작업 진행 중...")
        
        thread = threading.Thread(target=self.run_automation_task, args=(user_id, user_pw, target_codes))
        thread.daemon = True
        thread.start()

    def run_automation_task(self, plm_id, plm_pw, target_codes):
        URL = "https://plm.fursys.com/Windchill/platform/main#"
        driver = None
        try:
            print("="*50)
            print("  PLM PART LIST 다중 단품코드 자동화 (v8 GUI)")
            print("="*50)
            
            driver = setup_driver()
            
            login(driver, URL, plm_id, plm_pw)
            main_window_handle = driver.current_window_handle
            
            print(f"\n총 {len(target_codes)}개의 자재에 대해 작업을 시작합니다.")
            
            for code_idx, code in enumerate(target_codes):
                print(f"\n{'='*40}")
                print(f"[{code_idx+1}/{len(target_codes)}] '{code}' 작업 시작")
                print(f"{'='*40}")
                
                try:
                    driver.switch_to.default_content()
                    process_single_code(driver, code)
                except Exception as e:
                    print(f"[오류] '{code}' 처리 중 에러 발생: {e}")
                
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
                    print(f"[치명적 오류] 메인 창 손실. {e}")
                    break
                    
                time.sleep(1)
                
            print("\n>>> 모든 일괄 작업이 완료되었습니다! 브라우저를 닫습니다.")
        except Exception as final_e:
            print(f"[오류] 자동화 중 치명적 에러: {final_e}")
        finally:
            try:
                if driver:
                    driver.quit()
            except:
                pass
            
            self.after(0, lambda: self.start_btn.configure(state="normal", text="자동화 시작"))


if __name__ == "__main__":
    app = App()
    app.mainloop()
