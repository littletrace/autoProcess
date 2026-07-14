"""
income_report.py - 완전 독립 실행 스크립트 (PyInstaller 빌드 대응)
이카운트(ecount_sub, 670989) 입금보고서집계 / 판매현황 → Google Sheets 값 덮어쓰기

절차:
  [작업1] 재고1 > 영업관리 > Search > 입금보고서체크 > 검색꺽쇠 > Excel
          → '입금보고서집계' 탭에 값만 붙여넣기(덮어쓰기)
  [작업2] 회계1 > 출력 > 입금보고서집계 > 입금보고서확인(카드) > 검색 > Excel
          → '판매현황' 탭에 값만 붙여넣기(덮어쓰기)
"""

import os
import sys
import json
import time
import glob
import traceback
from datetime import datetime

# ── Windows cp949 인코딩 오류 방지 ──────────────────────────────────────────
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except (AttributeError, ValueError):
    pass

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

import gspread
from google.oauth2.service_account import Credentials

# ════════════════════════════════════════════════════════════════
# 1. 상수
# ════════════════════════════════════════════════════════════════

ACCOUNT_KEY = "ecount_sub"          # config.json 내 로그인 정보 키 (670989)
SPREADSHEET_ID = "1gkQVbbgSq1TNJJQVWijoQMYx_vWciyN1M3uUgxAbE94"

TAB1 = "입금보고서집계"
TAB2 = "판매현황"

# 탭별 열 서식: {0-based 열 인덱스: "TEXT" | "NUMBER"}  (A=0, B=1, C=2, ...)
COLUMN_FORMATS = {
    TAB1: {
        0: "TEXT",    # A열
        2: "TEXT",    # C열
        5: "TEXT",    # F열
        7: "NUMBER",  # H열
        8: "NUMBER",  # I열
    },
    TAB2: {
        2: "TEXT",    # C열
        5: "NUMBER",    # F열
    },
}

LOGIN_URL = "https://login.ecount.com/LOGIN?lan_type=ko-KR"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# 새 기기 등록 팝업 '등록' 버튼
_REGIST_BTN_XPATH = (
    "//div[contains(@class, 'control-set')]"
    "//*[@id='toolbar_sid_toolbar_item_regist']/button"
)
# 새 기기 팝업이 없을 때 대신 노출되는 로그인 확인 버튼
_SAVE_BTN_ID = "save"

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DOWNLOAD_DIR = os.path.join(BASE_DIR, "temp_downloads")


# ════════════════════════════════════════════════════════════════
# 2. config.json 로더 (자신의 폴더 우선, 없으면 상위 폴더)
# ════════════════════════════════════════════════════════════════

def load_config() -> dict:
    candidates = [
        os.path.join(BASE_DIR, "config.json"),
        os.path.join(os.path.dirname(BASE_DIR), "config.json"),
    ]
    for path in candidates:
        if os.path.exists(path):
            print(f"[설정] config.json 로드: {path}")
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    raise FileNotFoundError(
        "config.json을 찾을 수 없습니다. (탐색 경로: " + ", ".join(candidates) + ")"
    )


# ════════════════════════════════════════════════════════════════
# 3. 다운로드 폴더 유틸
# ════════════════════════════════════════════════════════════════

def clear_download_dir(download_dir: str):
    if not os.path.isdir(download_dir):
        return
    for filename in os.listdir(download_dir):
        file_path = os.path.join(download_dir, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
        except Exception:
            pass


def wait_for_download(download_dir: str, timeout: int = 60) -> str:
    print("[다운로드] 완료 대기 중...")
    deadline = time.time() + timeout
    while time.time() < deadline:
        time.sleep(1)
        files = glob.glob(os.path.join(download_dir, "*"))
        if files:
            unfinished = [f for f in files if f.endswith(".crdownload") or f.endswith(".tmp")]
            if not unfinished:
                latest = max(files, key=os.path.getctime)
                print(f"[다운로드] 완료: {os.path.basename(latest)}")
                return latest
    raise TimeoutError(f"다운로드 {timeout}초 초과: {download_dir}")


def build_driver(download_dir: str) -> webdriver.Chrome:
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
    }
    options.add_experimental_option("prefs", prefs)

    try:
        service = Service(ChromeDriverManager().install())
        return webdriver.Chrome(service=service, options=options)
    except Exception as e:
        raise RuntimeError(f"Chrome 드라이버 생성 실패: {e}") from e


# ════════════════════════════════════════════════════════════════
# 4. 이카운트 클라이언트
# ════════════════════════════════════════════════════════════════

class EcountClient:
    def __init__(self, driver: webdriver.Chrome, default_timeout: int = 15):
        self.driver = driver
        self.wait = WebDriverWait(driver, default_timeout)

    def _click_xpath(self, xpath: str, timeout: int = None):
        wait = self.wait if timeout is None else WebDriverWait(self.driver, timeout)
        wait.until(EC.element_to_be_clickable((By.XPATH, xpath))).click()

    def click(self, xpath: str, delay: float = 0.0, timeout: int = None):
        self._click_xpath(xpath, timeout=timeout)
        if delay:
            time.sleep(delay)

    def login(self, account: dict, wait_after: int = 4):
        """
        이카운트 로그인.
        로그인 후 두 가지 케이스로 분기:
          1) 새 기기 등록 팝업이 뜨는 경우 → '등록' 버튼 클릭
          2) 팝업이 뜨지 않는 경우 → 로그인 확인 버튼(id="save") 클릭
             (기존에는 이 케이스를 처리하지 않아 로그인 페이지에 그대로
              머물러 있는 버그가 있었음)
        """
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
            print("🎉 로그인 요청 완료")

            self._handle_login_confirmation()
        except KeyError as e:
            raise KeyError(f"계정 정보 누락: {e}") from e
        except Exception as e:
            raise RuntimeError(f"이카운트 로그인 실패: {e}") from e

    def _handle_login_confirmation(self, timeout: int = 5):
        print("📋 새 기기 로그인 팝업 여부 확인 중...")
        try:
            WebDriverWait(self.driver, timeout).until(
                EC.element_to_be_clickable((By.XPATH, _REGIST_BTN_XPATH))
            ).click()
            print("✅ 새 기기 등록 팝업 → '등록' 버튼 클릭 완료")
            time.sleep(2)
            return
        except TimeoutException:
            print("ℹ️ 새 기기 팝업 없음 → 로그인 확인 버튼(save) 확인 중...")

        try:
            WebDriverWait(self.driver, timeout).until(
                EC.element_to_be_clickable((By.ID, _SAVE_BTN_ID))
            ).click()
            print("✅ 로그인 확인 버튼(save) 클릭 완료")
            time.sleep(2)
        except TimeoutException:
            print("ℹ️ 로그인 확인 버튼 없음 → 정상 로그인으로 간주하고 진행")


# ════════════════════════════════════════════════════════════════
# 5. 구글 시트 - 값만 붙여넣기(덮어쓰기)
# ════════════════════════════════════════════════════════════════

class GoogleSheet:
    def __init__(self, service_account_info: dict, spreadsheet_id: str):
        try:
            creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)
            self._gc = gspread.authorize(creds)
            self._doc = self._gc.open_by_key(spreadsheet_id)
        except Exception as e:
            raise RuntimeError(f"구글 시트 인증/연결 실패: {e}") from e

    @staticmethod
    def _read_tabular(file_path: str) -> pd.DataFrame:
        # 엑셀로 먼저 시도하고, 실패하면 이카운트 특유의 'xls 빙자 HTML' 파일로 파싱
        try:
            df = pd.read_excel(file_path)
        except Exception:
            print("💡 HTML 형식의 엑셀 파일로 감지되어 변환하여 읽습니다.")
            dfs = pd.read_html(file_path, encoding="utf-8")
            df = dfs[0]
        return df.fillna("")

    def paste_values(self, worksheet_name: str, file_path: str,
                      col_formats: dict = None, cleanup: bool = True):
        """
        지정 탭을 전체 삭제하고 다운로드 파일의 값(헤더+데이터)만 덮어쓰기.

        주의: 열 서식(TEXT/NUMBER)은 반드시 값 update() 이전에 적용해야 한다.
        USER_ENTERED로 값을 넣는 순간 Sheets가 숫자로 보이는 문자열을
        자동으로 숫자 타입으로 확정해버리므로, update() 이후에 TEXT 서식을
        걸어도 표시형식만 바뀔 뿐 내부 값 타입은 숫자로 남아 합계가 계속
        표시된다. 따라서 clear() → format_columns() → update() 순서로 진행한다.
        """
        try:
            print(f"\n--- [{worksheet_name}] 구글 스프레드시트 값 덮어쓰기 준비 ---")
            worksheet = self._doc.worksheet(worksheet_name)

            df = self._read_tabular(file_path)
            data = [df.columns.values.tolist()] + df.values.tolist()

            worksheet.clear()

            if col_formats:
                self.format_columns(worksheet_name, col_formats)

            worksheet.update(range_name="A1", values=data, value_input_option="USER_ENTERED")
            print(f"🎉 [{worksheet_name}] 값 덮어쓰기 완료! ({len(data)}행)")

        except Exception as e:
            print(f"❌ [{worksheet_name}] 구글 시트 업로드 중 오류: {e}")
            raise
        finally:
            if cleanup and os.path.exists(file_path):
                try:
                    os.remove(file_path)
                    print("🗑️ 처리가 완료된 로컬 임시 파일을 삭제했습니다.")
                except Exception:
                    pass

    def format_columns(self, worksheet_name: str, col_formats: dict):
        """
        열 단위 표시 형식 지정 (Sheets API v4 repeatCell).
        col_formats: {0-based 열 인덱스: "TEXT" | "NUMBER"}
        """
        if not col_formats:
            return
        try:
            worksheet = self._doc.worksheet(worksheet_name)
            sheet_id = worksheet.id

            type_map = {
                "TEXT":   {"type": "TEXT",   "pattern": "@"},
                "NUMBER": {"type": "NUMBER", "pattern": "#,##0"},
            }

            requests = []
            for col_idx, fmt in col_formats.items():
                fmt_key = str(fmt).upper()
                if fmt_key not in type_map:
                    raise ValueError(f"지원 형식: TEXT / NUMBER (입력값: {fmt})")
                requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_id,
                            "startColumnIndex": col_idx,
                            "endColumnIndex":   col_idx + 1,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "numberFormat": type_map[fmt_key]
                            }
                        },
                        "fields": "userEnteredFormat.numberFormat",
                    }
                })

            if requests:
                self._doc.batch_update({"requests": requests})
                print(f"🎨 [{worksheet_name}] 열 서식 적용 완료")

        except Exception as e:
            print(f"❌ [{worksheet_name}] 열 서식 지정 중 오류: {e}")


# ════════════════════════════════════════════════════════════════
# 6. 작업 정의
# ════════════════════════════════════════════════════════════════

def task1_deposit_report_summary(client: EcountClient, gsheet: GoogleSheet):
    """
    [작업1] 재고1 > 영업관리 > Search > 입금보고서체크 > 검색꺽쇠 > Excel
    → '판매현황' 탭 덮어쓰기
    """
    print("\n[작업1] 판매현황 시작")

    client.click('//*[@id="link_depth1_MENUTREE_000004"]', delay=1.5)   # 재고1
    client.click('//*[@id="link_depth2_MENUTREE_000030"]', delay=1.5)   # 영업관리
    client.click(
        "/html/body/div[2]/div[5]/div[3]/div/div[1]/div/div[1]/div[1]/div[1]/div[2]/div[2]/div/button",
        delay=1.5,
    )  # Search (하위 메뉴 열기)
    client.click("//*[contains(text(), '입금보고서체크')]", delay=1.5)
    client.click(
        "/html/body/div[2]/div[5]/div[3]/div/div[1]/div/div[1]/div[2]/div[2]/div[1]/div[1]/div/div/button[2]",
        delay=2,
    )  # 검색 꺽쇠
    client.click("//*[contains(text(), 'Excel')]", delay=2)

    file_path = wait_for_download(DOWNLOAD_DIR)
    gsheet.paste_values(TAB2, file_path, col_formats=COLUMN_FORMATS[TAB2])
    print("[작업1] 완료")


def task2_deposit_report_check(client: EcountClient, gsheet: GoogleSheet):
    """
    [작업2] 회계1 > 출력 > 입금보고서집계 > 입금보고서확인(카드) > 검색 > Excel
    → '입금보고서집계' 탭 덮어쓰기
    """
    print("\n[작업2] 입금보고서집계 시작")

    client.click('//*[@id="link_depth1_MENUTREE_000001"]', delay=1.5)   # 회계1
    client.click('//*[@id="link_depth2_MENUTREE_000017"]', delay=1.5)   # 출력
    client.click(
        "/html/body/div[2]/div[5]/div[3]/div/div[1]/div[2]/div[2]/div[2]/ul/li[1]/ul/li[14]/a",
        delay=1.5,
    )  # 입금보고서집계 메뉴
    client.click("//*[contains(text(), '입금보고서확인(카드')]", delay=1.5)
    client.click('//*[@id="searchGroup"]', delay=5)   # 검색
    client.click("//*[contains(text(), 'Excel')]", delay=2)

    file_path = wait_for_download(DOWNLOAD_DIR)
    gsheet.paste_values(TAB1, file_path, col_formats=COLUMN_FORMATS[TAB1])
    print("[작업2] 완료")


# ════════════════════════════════════════════════════════════════
# 7. 메인
# ════════════════════════════════════════════════════════════════

def main():
    print("=" * 55)
    print(" income_report.py 시작")
    print(f" {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 55)

    config = load_config()

    if ACCOUNT_KEY not in config:
        raise KeyError(f"config.json에 '{ACCOUNT_KEY}' 계정 정보가 없습니다.")
    if "google" not in config or "service_account" not in config["google"]:
        raise KeyError("config.json에 'google.service_account' 정보가 없습니다.")

    account = config[ACCOUNT_KEY]
    gsheet = GoogleSheet(config["google"]["service_account"], SPREADSHEET_ID)
    print("[Sheets] 인증 완료")

    driver = build_driver(DOWNLOAD_DIR)
    task1_ok = False
    task2_ok = False

    try:
        client = EcountClient(driver)
        client.login(account)

        try:
            task1_deposit_report_summary(client, gsheet)
            task1_ok = True
        except Exception:
            print("[작업1] 실패:")
            traceback.print_exc()

        try:
            task2_deposit_report_check(client, gsheet)
            task2_ok = True
        except Exception:
            print("[작업2] 실패:")
            traceback.print_exc()

    finally:
        driver.quit()
        print("🔒 브라우저 종료")

    print("\n" + "=" * 55)
    print(f" 작업1(판매현황)       : {'✅ 성공' if task1_ok else '❌ 실패'}")
    print(f" 작업2(입금보고서집계) : {'✅ 성공' if task2_ok else '❌ 실패'}")
    print("=" * 55)

    if not (task1_ok and task2_ok):
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ 실행 중 오류가 발생했습니다: {e}")
        traceback.print_exc()
    finally:
        input("\n종료하려면 Enter를 누르세요...")