import os
import sys
import json
import re
import time
import glob
import traceback
from datetime import datetime, timedelta

# ==========================================
# 0. 윈도우 cp949 인코딩 오류 방지
# ==========================================
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except (AttributeError, ValueError):
    pass

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager

import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

# ==========================================
# 1. 경로 설정 (exe / 스크립트 겸용)
# ==========================================
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DOWNLOAD_DIR = os.path.join(BASE_DIR, "temp_downloads")

# ==========================================
# 2. 이 스크립트 전용 설정 (개별 관리)
# ==========================================
SPREADSHEET_ID = "1gkQVbbgSq1TNJJQVWijoQMYx_vWciyN1M3uUgxAbE94"

ERP_URL = "http://211.253.8.106:8080"
ERP_TAB = "erp_입금입력조회"

ECOUNT_TAB_ACCOUNT = "ecnt_입출금계좌조회"
ECOUNT_TAB_REPORT = "ecnt_입금보고서집계"

ECOUNT_ACCOUNT_KEY = "ecount_main"
ERP_ACCOUNT_KEY = "erp_main"

LOGIN_URL = "https://login.ecount.com/LOGIN?lan_type=ko-KR"

# 새 기기 등록 팝업 '등록' 버튼 (중복 ID 회피용 control-set 한정)
_REGIST_BTN_XPATH = (
    "//div[contains(@class, 'control-set')]"
    "//*[@id='toolbar_sid_toolbar_item_regist']/button"
)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ==========================================
# 3. config.json 로더
# 스크립트 자신의 폴더를 먼저 탐색하고, 없으면 상위 폴더로 폴백
# ==========================================
def load_config():
    candidates = [
        os.path.join(BASE_DIR, "config.json"),
        os.path.join(os.path.dirname(BASE_DIR), "config.json"),
    ]
    for path in candidates:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    raise FileNotFoundError(
        "config.json을 찾을 수 없습니다. (탐색 경로: " + ", ".join(candidates) + ")"
    )

# ==========================================
# 4. 다운로드 폴더 유틸
# ==========================================
def clear_download_dir(download_dir):
    if not os.path.isdir(download_dir):
        return
    for filename in os.listdir(download_dir):
        file_path = os.path.join(download_dir, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
        except Exception:
            pass

def wait_for_download(download_dir, timeout=60):
    print("다운로드 완료를 대기 중입니다...")
    seconds = 0
    while seconds < timeout:
        time.sleep(1)
        files = glob.glob(os.path.join(download_dir, "*"))
        if files:
            # .crdownload 와 .tmp 파일이 모두 없을 때 다운로드 완료로 간주
            unfinished = [f for f in files if f.endswith(".crdownload") or f.endswith(".tmp")]
            if not unfinished:
                return max(files, key=os.path.getctime)
        seconds += 1
    return None

def build_driver(download_dir, extra_args=None, extra_prefs=None, insecure=False):
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)
    clear_download_dir(download_dir)

    options = Options()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--start-maximized")

    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
    }
    if extra_prefs:
        prefs.update(extra_prefs)
    options.add_experimental_option("prefs", prefs)

    if insecure:
        options.add_argument("--ignore-certificate-errors")
        options.add_argument("--allow-running-insecure-content")

    if extra_args:
        for arg in extra_args:
            options.add_argument(arg)

    try:
        service = Service(ChromeDriverManager().install())
        return webdriver.Chrome(service=service, options=options)
    except Exception as e:
        raise RuntimeError(f"Chrome 드라이버 생성 실패: {e}") from e


# ==========================================
# 5. 이카운트 클라이언트
# ==========================================
class EcountClient:
    def __init__(self, driver, default_timeout=15):
        self.driver = driver
        self.wait = WebDriverWait(driver, default_timeout)

    def _click_xpath(self, xpath, timeout=None):
        wait = self.wait if timeout is None else WebDriverWait(self.driver, timeout)
        wait.until(EC.element_to_be_clickable((By.XPATH, xpath))).click()

    def login(self, account, wait_after=4):
        try:
            print("이카운트 로그인 페이지로 이동합니다...")
            self.driver.get(LOGIN_URL)

            self.wait.until(
                EC.presence_of_element_located((By.ID, "com_code"))
            ).send_keys(account["com_code"])
            self.driver.find_element(By.ID, "id").send_keys(account["user_id"])

            pw_input = self.driver.find_element(By.ID, "passwd")
            pw_input.send_keys(account["password"])
            pw_input.send_keys(Keys.RETURN)

            time.sleep(wait_after)
            print("🎉 로그인 성공!")

            self.handle_new_device_popup()
        except KeyError as e:
            raise KeyError(f"계정 정보 누락: {e}") from e
        except Exception as e:
            raise RuntimeError(f"이카운트 로그인 실패: {e}") from e

    def handle_new_device_popup(self, timeout=5):
        try:
            print("📋 새 기기 로그인 팝업 여부 확인 중...")
            WebDriverWait(self.driver, timeout).until(
                EC.element_to_be_clickable((By.XPATH, _REGIST_BTN_XPATH))
            ).click()
            print("✅ 새 기기 등록 팝업 → '등록' 버튼 클릭 완료")
            time.sleep(2)
        except Exception:
            print("ℹ️ 새 기기 팝업 없음 → 다음 단계로 진행")

    def click(self, xpath, delay=0.0, timeout=None):
        self._click_xpath(xpath, timeout=timeout)
        if delay:
            time.sleep(delay)


# ==========================================
# 6. 구글 시트 업로더
# ==========================================
class GoogleSheet:
    def __init__(self, service_account_info, spreadsheet_id):
        try:
            creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)
            self._gc = gspread.authorize(creds)
            self._doc = self._gc.open_by_key(spreadsheet_id)
        except Exception as e:
            raise RuntimeError(f"구글 시트 인증/연결 실패: {e}") from e

    @staticmethod
    def _strip_all_whitespace(value):
        """문자열 셀의 공백을 (중간 공백 포함) 완전히 제거한다. 문자열이 아니면 그대로 반환."""
        if isinstance(value, str):
            return re.sub(r"\s+", "", value)
        return value

    @classmethod
    def _read_tabular(cls, file_path):
        # 엑셀로 먼저 시도하고, 실패하면 이카운트 특유의 'xls 빙자 HTML' 파일로 파싱
        try:
            df = pd.read_excel(file_path)
        except Exception:
            print("💡 HTML 형식의 엑셀 파일로 감지되어 변환하여 읽습니다.")
            dfs = pd.read_html(file_path, encoding="utf-8")
            df = dfs[0]
        df = df.fillna("")

        # 모든 셀(헤더 포함)의 공백을 완전 제거 (중간 공백 포함)
        df = df.apply(lambda col: col.map(cls._strip_all_whitespace))
        df.columns = [
            cls._strip_all_whitespace(str(c)) if not isinstance(c, str) else cls._strip_all_whitespace(c)
            for c in df.columns
        ]
        return df

    def upload_file(self, file_path, worksheet_name, add_timestamp=False, cleanup=True):
        try:
            print(f"\n--- [{worksheet_name}] 구글 스프레드시트 업로드 준비 ---")
            worksheet = self._doc.worksheet(worksheet_name)

            df = self._read_tabular(file_path)
            data = [df.columns.values.tolist()] + df.values.tolist()

            worksheet.batch_clear(["A1:Z20000"])
            # USER_ENTERED 사용: 숫자 값 앞에 불필요한 아포스트로피(') 방지
            worksheet.update(range_name="A1", values=data, value_input_option="USER_ENTERED")
            print(f"🎉 [{worksheet_name}] 업로드 완료!")

            if add_timestamp:
                next_row = len(data) + 1
                stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                worksheet.update(
                    values=[[f"업데이트 일시: {stamp}"]],
                    range_name=f"A{next_row}",
                )

        except Exception as e:
            print(f"❌ 구글 시트 업로드 중 오류가 발생했습니다: {e}")
        finally:
            if cleanup and os.path.exists(file_path):
                try:
                    os.remove(file_path)
                    print("🗑️ 처리가 완료된 로컬 임시 파일을 삭제했습니다.")
                except Exception:
                    pass

    def format_columns(self, worksheet_name, formats):
        try:
            worksheet = self._doc.worksheet(worksheet_name)
            sheet_id = worksheet.id

            type_map = {
                "TEXT": {"type": "TEXT", "pattern": "@"},
                "NUMBER": {"type": "NUMBER", "pattern": "#,##0"},
                "NUMBER_DECIMAL": {"type": "NUMBER", "pattern": "#,##0.00"},
            }

            requests = []
            for col_index, fmt in formats.items():
                fmt_key = str(fmt).upper()
                if fmt_key not in type_map:
                    raise ValueError(
                        f"지원하지 않는 형식입니다: {fmt} (TEXT, NUMBER 또는 NUMBER_DECIMAL)"
                    )
                requests.append(
                    {
                        "repeatCell": {
                            "range": {
                                "sheetId": sheet_id,
                                "startColumnIndex": col_index,
                                "endColumnIndex": col_index + 1,
                            },
                            "cell": {"userEnteredFormat": {"numberFormat": type_map[fmt_key]}},
                            "fields": "userEnteredFormat.numberFormat",
                        }
                    }
                )

            if requests:
                self._doc.batch_update({"requests": requests})
                print(f"🎨 [{worksheet_name}] 열 표시 형식 지정 완료!")

        except Exception as e:
            print(f"❌ 열 표시 형식 지정 중 오류가 발생했습니다: {e}")


# ==========================================
# 7. 영림원 ERP 새 기기 로그인 팝업 처리
# ==========================================
def handle_new_device_popup_erp(driver):
    try:
        print("📋 새 기기 로그인 팝업 여부 확인 중...")
        regist_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, _REGIST_BTN_XPATH))
        )
        regist_btn.click()
        print("✅ 새 기기 등록 팝업 → '등록' 버튼 클릭 완료")
        time.sleep(2)
    except Exception:
        print("ℹ️ 새 기기 팝업 없음 → 다음 단계로 진행")


# ==========================================
# 8. 이카운트 자동화
# 프로세스: 로그인 → 입출금계좌조회(입금보고서 비교, 전체 검색) 다운로드/업로드
#          → 입금보고서집계(입금보고서비교 조건) 다운로드/업로드
# ==========================================
def run_ecount_automation(account, gsheet):
    print("\n[이카운트] 프로세스를 시작합니다.")
    driver = build_driver(DOWNLOAD_DIR, extra_prefs={
        "profile.default_content_setting_values.insecure_content": 1
    })

    try:
        client = EcountClient(driver)
        client.login(account)

        # [1] 입출금계좌조회
        client.click('//*[@id="link_depth1_MENUTREE_000001"]')
        client.click('//*[@id="link_depth2_MENUTREE_002562"]', delay=2)

        client.click('/html/body/div[2]/div[5]/div[3]/div/div[1]/div[2]/div[2]/ul/li[1]/a', delay=3)
        print("검색범위 : 전체")
        client.click('//*[@id="tgHeaderSearch"]')
        print("Search")
        client.click('//*[@id="custom3"]')
        print("조건 : 입금보고서 비교")
        client.click('//*[@id="multiSearch"]', delay=3)
        print("검색")
        client.click('//*[@id="excel"]')
        print("엑셀 다운로드")

        file1 = wait_for_download(DOWNLOAD_DIR, timeout=60)
        if file1:
            gsheet.upload_file(file1, ECOUNT_TAB_ACCOUNT)
            gsheet.format_columns(ECOUNT_TAB_ACCOUNT, {2: "TEXT", 4: "TEXT", 7: "NUMBER", 8: "NUMBER"})

        # [2] 입금보고서집계
        client.click('//*[@id="link_depth1_MENUTREE_000001"]')
        client.click('//*[@id="link_depth2_MENUTREE_000017"]')
        client.click('/html/body/div[2]/div[5]/div[3]/div/div[1]/div[2]/div[2]/div[2]/ul/li[1]/ul/li[14]/a', delay=2)

        client.click('//*[@id="custom3"]')
        print("조건 : 입금보고서비교")
        client.click('//*[@id="searchGroup"]', delay=8)
        print("검색")
        client.click('//*[@id="outputExcel"]')
        print("엑셀 다운로드")

        file2 = wait_for_download(DOWNLOAD_DIR, timeout=60)
        if file2:
            gsheet.upload_file(file2, ECOUNT_TAB_REPORT)
            gsheet.format_columns(ECOUNT_TAB_REPORT, {2: "TEXT", 5: "NUMBER", 8: "TEXT"})

        print("\n[이카운트] 성공: 이카운트 자동화 작업이 종료되었습니다.")

    except Exception as e:
        print(f"[이카운트] 에러 발생: {e}")
        traceback.print_exc()
    finally:
        driver.quit()


# ==========================================
# 9. 영림원 ERP 자동화
# 프로세스: 로그인 → 새 기기 팝업 처리 → 메뉴 이동 → iframe 전환
#          → 전월 1일~오늘 날짜 검색 → 조회 → 톱니바퀴 우클릭 → 엑셀 내보내기 → 업로드
# ==========================================
def run_erp_automation(account, gsheet):
    print("\n[영림원] 프로세스를 시작합니다.")

    # HTTP 사이트이므로 인증서/혼합콘텐츠 허용 + 안전 출처 예외 처리
    driver = build_driver(
        DOWNLOAD_DIR,
        insecure=True,
        extra_args=[f"--unsafely-treat-insecure-origin-as-secure={ERP_URL}"],
        extra_prefs={"profile.default_content_setting_values.insecure_content": 1},
    )
    wait = WebDriverWait(driver, 10)

    try:
        print("[영림원] 1단계: 사이트 접속 및 로그인")
        driver.get(ERP_URL)
        wait.until(EC.presence_of_element_located((By.ID, "txtLoginId"))).send_keys(account["user_id"])
        driver.find_element(By.ID, "inputLoginPwd").send_keys(account["password"])
        driver.find_element(By.ID, "btnLogin").click()
        time.sleep(4)
        handle_new_device_popup_erp(driver)

        print("[영림원] 2단계: 메뉴 이동")
        menu1 = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/article/section[3]/div/div[2]/ul/li[4]/a')))
        driver.execute_script("arguments[0].click();", menu1)
        time.sleep(1)
        menu2 = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/article/section[4]/section/aside[1]/section[1]/div[2]/ul/li[7]')))
        driver.execute_script("arguments[0].click();", menu2)
        time.sleep(1)
        menu3 = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/article/section[4]/section/aside[1]/section[1]/div[2]/ul/li[7]/ul/li[1]')))
        driver.execute_script("arguments[0].click();", menu3)
        time.sleep(1)
        menu4 = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/article/section[4]/section/aside[1]/section[1]/div[2]/ul/li[7]/ul/li[1]/ul/li[3]/a/p')))
        driver.execute_script("arguments[0].click();", menu4)
        time.sleep(5)

        print("[영림원] 3단계: 엑티브 페이지(iframe) 선택")
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "500945_iframe")))
        time.sleep(2)

        print("[영림원] 4단계: 검색 조건(날짜) 입력")
        today = datetime.today()
        first_day_of_this_month = today.replace(day=1)
        last_day_of_prev_month = first_day_of_this_month - timedelta(days=1)
        first_day_of_month = last_day_of_prev_month.replace(day=1).strftime("%Y-%m-%d")
        today_str = today.strftime("%Y-%m-%d")

        start_date = wait.until(EC.presence_of_element_located((By.ID, 'txtReceiptDateFr_dat')))
        start_date.clear()
        start_date.send_keys(first_day_of_month)

        end_date = driver.find_element(By.ID, 'txtReceiptDateTo_dat')
        end_date.clear()
        end_date.send_keys(today_str)

        print("[영림원] 5단계: 조회 버튼 클릭")
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/article/section[1]/div/ul/li[1]/a[1]/span"))).click()
        time.sleep(3)

        print("[영림원] 6단계: 톱니바퀴 영역 우클릭")
        gear = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/article/section[2]/form/div/div[2]/div/div/a")))
        ActionChains(driver).context_click(gear).perform()
        time.sleep(2)

        print("[영림원] 7단계: 엑셀로 내보내기")
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/article/section[2]/div/div/section/ul/li[4]/a[1]"))).click()

        file3 = wait_for_download(DOWNLOAD_DIR, timeout=60)
        if file3:
            gsheet.upload_file(file3, ERP_TAB, add_timestamp=True)
            gsheet.format_columns(ERP_TAB, {2: "TEXT", 11: "NUMBER", 12: "NUMBER"})

        print("\n[영림원] 성공: 영림원 자동화 작업이 완료되었습니다.")

    except Exception:
        print("\n[영림원] 에러 발생: 진행 중 문제가 발생했습니다.")
        traceback.print_exc()
    finally:
        driver.quit()


# ==========================================
# 10. 메인 실행부
# ==========================================
def main():
    config = load_config()

    for key in (ECOUNT_ACCOUNT_KEY, ERP_ACCOUNT_KEY):
        if key not in config:
            raise KeyError(f"config.json에 '{key}' 계정 정보가 없습니다.")
    if "google" not in config or "service_account" not in config["google"]:
        raise KeyError("config.json에 'google.service_account' 정보가 없습니다.")

    gsheet = GoogleSheet(config["google"]["service_account"], SPREADSHEET_ID)

    # 1. 이카운트 자동화 실행
    run_ecount_automation(config[ECOUNT_ACCOUNT_KEY], gsheet)

    # 2. 영림원 자동화 실행
    run_erp_automation(config[ERP_ACCOUNT_KEY], gsheet)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ 실행 중 오류가 발생했습니다: {e}")
    finally:
        input("\n종료하려면 Enter를 누르세요...")