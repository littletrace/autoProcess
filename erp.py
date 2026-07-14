"""
erp.py - 완전 독립 실행 스크립트 (PyInstaller 빌드 대응)
영림원 ERP + 이카운트 판매현황 → Google Sheets 업로드
"""

import sys
import os
import json
import time
import glob
import traceback
import shutil
from datetime import datetime, timedelta
from pathlib import Path

# ── 외부 의존성 ──────────────────────────────────────────────
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, WebDriverException
)

import openpyxl
import gspread
from google.oauth2.service_account import Credentials

# ════════════════════════════════════════════════════════════════
# 1. 경로 / 설정 로더
# ════════════════════════════════════════════════════════════════

def get_base_path() -> Path:
    """PyInstaller exe / 일반 .py 실행 모두 대응"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


def load_config() -> dict:
    """config.json 로드"""
    config_path = get_base_path() / "config.json"
    if not config_path.exists():
        raise FileNotFoundError(f"config.json 없음: {config_path}")
    with open(config_path, encoding="utf-8") as f:
        return json.load(f)


# ════════════════════════════════════════════════════════════════
# 2. Google Sheets 클라이언트
# ════════════════════════════════════════════════════════════════

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
SPREADSHEET_ID = "1z_Hn9GGMQvFjGdwQp1qa4ycadcbD_dHweKPelV7pkHM"


class GoogleSheet:
    def __init__(self, service_account_info: dict):
        creds = Credentials.from_service_account_info(
            service_account_info, scopes=SCOPES
        )
        self.client = gspread.authorize(creds)
        self.spreadsheet = self.client.open_by_key(SPREADSHEET_ID)

    def get_worksheet(self, tab_name: str):
        return self.spreadsheet.worksheet(tab_name)

    def upload_user_entered(self, tab_name: str, rows: list[list]):
        """
        탭 전체를 rows 로 교체 (USER_ENTERED: 수식·자동변환 허용).
        이카운트 탭처럼 '나머지는 user_entered' 요건에 사용.
        """
        ws = self.get_worksheet(tab_name)
        ws.clear()
        if rows:
            ws.update(
                range_name="A1",
                values=rows,
                value_input_option="USER_ENTERED",
            )
        print(f"  [Sheets] '{tab_name}' 탭 업로드 완료 ({len(rows)}행, USER_ENTERED)")

    def format_columns(self, tab_name: str, col_formats: dict):
        """
        열 단위 표시 형식 강제 지정 (Sheets API v4 repeatCell).

        Args:
            tab_name    : 대상 워크시트 탭 이름
            col_formats : { 열문자(str) 또는 0-based 인덱스(int) : 'TEXT' | 'NUMBER' }

        열 문자 예) 'F', 'H', 'I'  →  내부적으로 0-based 인덱스로 변환
        """
        _TYPE_MAP = {
            "TEXT":   {"type": "TEXT",   "pattern": "@"},
            "NUMBER": {"type": "NUMBER", "pattern": "#,##0"},
        }

        def _col_to_idx(col) -> int:
            """'A'→0, 'B'→1, ... 'Z'→25, 'AA'→26 / 정수는 그대로"""
            if isinstance(col, int):
                return col
            col = str(col).upper().strip()
            idx = 0
            for ch in col:
                idx = idx * 26 + (ord(ch) - ord("A") + 1)
            return idx - 1

        ws = self.get_worksheet(tab_name)
        sheet_id = ws.id

        requests = []
        for col, fmt in col_formats.items():
            fmt_key = str(fmt).upper()
            if fmt_key not in _TYPE_MAP:
                raise ValueError(f"지원 형식: TEXT / NUMBER  (입력값: {fmt})")
            col_idx = _col_to_idx(col)
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startColumnIndex": col_idx,
                        "endColumnIndex":   col_idx + 1,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": _TYPE_MAP[fmt_key]
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat",
                }
            })

        if requests:
            self.spreadsheet.batch_update({"requests": requests})
            cols_str = ", ".join(str(c) for c in col_formats.keys())
            print(f"  [Sheets] '{tab_name}' 열 서식 지정 완료 ({cols_str})")

    def append_timestamp(self, tab_name: str):
        """데이터 마지막 행 다음에 업데이트 일시 기록"""
        ws = self.get_worksheet(tab_name)
        all_vals = ws.get_all_values()
        next_row = len(all_vals) + 1
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ws.update(
            range_name=f"A{next_row}",
            values=[[f"업데이트: {now_str}"]],
            value_input_option="RAW",
        )


# ════════════════════════════════════════════════════════════════
# 3. Selenium 드라이버 팩토리
# ════════════════════════════════════════════════════════════════

def build_driver(download_dir: str) -> webdriver.Chrome:
    """
    Chrome 드라이버 생성.
    - 다운로드 경로 고정
    - HTTP insecure / mixed-content 허용
    - 팝업 차단 해제
    """
    opts = Options()

    # 다운로드 설정
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": False,
        "safebrowsing.disable_download_protection": True,
    }
    opts.add_experimental_option("prefs", prefs)

    # HTTP / insecure / mixed-content 허용
    opts.add_argument("--allow-running-insecure-content")
    opts.add_argument("--disable-web-security")
    opts.add_argument("--unsafely-treat-insecure-origin-as-secure=http://211.253.8.106:8080")
    opts.add_argument("--allow-insecure-localhost")
    opts.add_argument("--ignore-certificate-errors")

    # 기타
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)

    driver = webdriver.Chrome(options=opts)
    driver.implicitly_wait(10)
    return driver


# ════════════════════════════════════════════════════════════════
# 4. 파일 다운로드 대기 유틸
# ════════════════════════════════════════════════════════════════

def wait_for_download(download_dir: str, timeout: int = 60) -> str:
    """
    download_dir 에서 .crdownload / .tmp 가 사라지고
    새 xlsx/xls 파일이 생길 때까지 대기.
    완료된 파일 경로 반환.
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        # 진행 중인 임시 파일 확인
        tmp_files = (
            glob.glob(os.path.join(download_dir, "*.crdownload"))
            + glob.glob(os.path.join(download_dir, "*.tmp"))
        )
        if tmp_files:
            time.sleep(1)
            continue

        # 완료된 엑셀 파일 탐색 (최신순)
        xl_files = sorted(
            glob.glob(os.path.join(download_dir, "*.xlsx"))
            + glob.glob(os.path.join(download_dir, "*.xls")),
            key=os.path.getmtime,
            reverse=True,
        )
        if xl_files:
            return xl_files[0]
        time.sleep(1)

    raise TimeoutError(f"다운로드 {timeout}초 초과: {download_dir}")


def read_excel_to_rows(filepath: str) -> list[list]:
    """
    openpyxl 로 엑셀 읽기.
    모든 값을 문자열로 변환 (None → "").
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    rows = []
    for row in ws.iter_rows(values_only=True):
        rows.append([("" if v is None else str(v)) for v in row])
    # 완전히 빈 행 제거
    rows = [r for r in rows if any(c.strip() for c in r)]
    return rows


# ════════════════════════════════════════════════════════════════
# 5. 이카운트 클라이언트
# ════════════════════════════════════════════════════════════════

# 기존 ecount_client.py 검증 로그인 URL / 팝업 XPath
_ECOUNT_LOGIN_URL = "https://login.ecount.com/LOGIN?lan_type=ko-KR"
_REGIST_BTN_XPATH = (
    "//div[contains(@class, 'control-set')]"
    "//*[@id='toolbar_sid_toolbar_item_regist']/button"
)


class EcountClient:
    """
    config.json ecount_main 키 구조:
        {"com_code": "...", "user_id": "...", "password": "..."}
    """

    def __init__(self, driver: webdriver.Chrome, account: dict):
        self.driver = driver
        self.account = account
        self.wait = WebDriverWait(driver, 20)

    def _click_xpath(self, xpath: str, wait_sec: int = 15):
        WebDriverWait(self.driver, wait_sec).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        ).click()

    def login(self, wait_after: int = 4):
        """
        이카운트 로그인 (기존 ecount_client.py 방식 그대로).
        필드 ID: com_code / id / passwd
        """
        from selenium.webdriver.common.keys import Keys

        print("  [이카운트] 로그인 페이지 이동...")
        self.driver.get(_ECOUNT_LOGIN_URL)

        self.wait.until(
            EC.presence_of_element_located((By.ID, "com_code"))
        ).send_keys(self.account["com_code"])

        self.driver.find_element(By.ID, "id").send_keys(self.account["user_id"])

        pw = self.driver.find_element(By.ID, "passwd")
        pw.send_keys(self.account["password"])
        pw.send_keys(Keys.RETURN)

        time.sleep(wait_after)
        print("  [이카운트] 로그인 완료")

        self._handle_new_device_popup()

    def _handle_new_device_popup(self, timeout: int = 5):
        """
        새 기기 등록 팝업 → '등록' 버튼 클릭.
        팝업 없으면 조용히 통과.
        """
        try:
            print("  [이카운트] 새 기기 팝업 확인 중...")
            WebDriverWait(self.driver, timeout).until(
                EC.element_to_be_clickable((By.XPATH, _REGIST_BTN_XPATH))
            ).click()
            print("  [이카운트] 새 기기 팝업 → '등록' 클릭 완료")
            time.sleep(2)
        except TimeoutException:
            print("  [이카운트] 새 기기 팝업 없음 → 계속 진행")

    def navigate_to_sales(self):
        """판매현황 메뉴 순서대로 클릭 (요구사항 1~4단계)"""
        steps = [
            "/html/body/div[2]/div[3]/div/div[4]/ul/li[3]/div[1]/a",
            "/html/body/div[2]/div[3]/div/div[4]/ul/li[3]/div[2]/ul/li[7]/a",
            "/html/body/div[2]/div[5]/div[3]/div/div[1]/div[2]/div[2]/div[2]/ul/li[2]/ul/li[1]/a",
            "/html/body/div[2]/div[5]/div[3]/div/div[1]/div/div[1]/div[2]/ul/li[5]/a",
        ]
        for i, xpath in enumerate(steps, 1):
            print(f"  [이카운트] 메뉴 {i}/{len(steps)} 클릭")
            self._click_xpath(xpath)
            time.sleep(1.5)

    def download_excel(self):
        """조회 버튼 → Excel 다운로드 (요구사항 5~6단계)"""
        print("  [이카운트] 조회 버튼 클릭")
        self._click_xpath(
            "/html/body/div[2]/div[5]/div[3]/div/div[1]/div/div[1]/div[2]/div[2]/div[1]/div[1]/div/div/button[2]"
        )
        time.sleep(3)

        print("  [이카운트] Excel 다운로드 클릭")
        self._click_xpath("//*[contains(text(), 'Excel')]")
        time.sleep(2)


# ════════════════════════════════════════════════════════════════
# 6. 영림원 ERP 작업
# ════════════════════════════════════════════════════════════════

def run_erp_youngimwon(config: dict, gs: GoogleSheet, download_dir: str):
    print("\n[영림원 ERP] 시작")
    account = config["erp_main"]
    driver = build_driver(download_dir)

    try:
        wait = WebDriverWait(driver, 20)

        # ── 로그인 ──────────────────────────────────────────
        driver.get("http://211.253.8.106:8080")
        time.sleep(2)

        wait.until(EC.presence_of_element_located((By.ID, "txtLoginId")))
        driver.find_element(By.ID, "txtLoginId").clear()
        driver.find_element(By.ID, "txtLoginId").send_keys(account["user_id"])
        driver.find_element(By.ID, "inputLoginPwd").clear()
        driver.find_element(By.ID, "inputLoginPwd").send_keys(account["password"])
        driver.find_element(By.ID, "btnLogin").click()
        time.sleep(3)
        print("  [영림원] 로그인 완료")

        # ── 메뉴 순서대로 클릭 ──────────────────────────────
        menu_xpaths = [
            # 1. 영업관리
            "/html/body/article/section[3]/div/div[2]/ul/li[4]/a",
            # 2. 매출관리
            "/html/body/article/section[4]/section/aside[1]/section[1]/div[2]/ul/li[5]/a/p",
            # 3. 거래명세서
            "/html/body/article/section[4]/section/aside[1]/section[1]/div[2]/ul/li[5]/ul/li[1]/a/p",
            # 4. 거래명세서품목조회
            "/html/body/article/section[4]/section/aside[1]/section[1]/div[2]/ul/li[5]/ul/li[1]/ul/li[3]/a/p",
        ]
        for i, xpath in enumerate(menu_xpaths, 1):
            el = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            el.click()
            time.sleep(1.5)
            print(f"  [영림원] 메뉴 {i}/4 클릭")

        # ── iframe 전환 ──────────────────────────────────────
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "500630_iframe")))
        print("  [영림원] iframe 전환 완료")

        # ── 날짜 / 담당자 입력 ──────────────────────────────
        today = datetime.now()
        date_from = (today - timedelta(days=5)).strftime("%Y%m%d")
        date_to   = today.strftime("%Y%m%d")

        def set_date_field(field_id: str, value: str):
            el = wait.until(EC.presence_of_element_located((By.ID, field_id)))
            driver.execute_script("arguments[0].value = '';", el)
            el.send_keys(value)

        set_date_field("datInvoiceDateFr_dat", date_from)
        set_date_field("datInvoiceDateTo_dat", date_to)
        print(f"  [영림원] 날짜 설정: {date_from} ~ {date_to}")

        emp_field = wait.until(EC.presence_of_element_located((By.ID, "txtEmpName_txt")))
        emp_field.clear()
        emp_field.send_keys("박혜성")
        print("  [영림원] 담당자 입력: 박혜성")

        # ── 조회 버튼 ────────────────────────────────────────
        query_btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/article/section[1]/div/ul/li[1]/a[1]/span")
        ))
        query_btn.click()
        time.sleep(3)
        print("  [영림원] 조회 완료")

        # ── 우클릭 메뉴(톱니바퀴) → 엑셀 다운로드 ───────────
        gear_btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/article/section[2]/form/div/div[2]/div/div/a")
        ))
        ActionChains(driver).context_click(gear_btn).perform()
        time.sleep(1.5)

        dl_btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/article/section[2]/div/div/section/ul/li[4]/a[1]")
        ))
        dl_btn.click()
        print("  [영림원] 엑셀 다운로드 클릭")

        # ── 다운로드 대기 ────────────────────────────────────
        filepath = wait_for_download(download_dir)
        print(f"  [영림원] 다운로드 완료: {filepath}")

        # ── 구글 시트 업로드 ─────────────────────────────────
        rows = read_excel_to_rows(filepath)
        gs.upload_user_entered("영림원ERP", rows)
        gs.format_columns("영림원ERP", {
            "F": "TEXT",
            "H": "NUMBER",
            "I": "NUMBER",
        })
        gs.append_timestamp("영림원ERP")
        print("  [영림원] 시트 업로드 완료")

        # ── 임시 파일 정리 ───────────────────────────────────
        try:
            os.remove(filepath)
        except Exception:
            pass

    except Exception:
        print("[영림원 ERP] 오류 발생:")
        traceback.print_exc()
        raise
    finally:
        driver.quit()


# ════════════════════════════════════════════════════════════════
# 7. 이카운트 작업
# ════════════════════════════════════════════════════════════════

def run_ecount_sales(config: dict, gs: GoogleSheet, download_dir: str):
    print("\n[이카운트] 시작")
    account = config["ecount_main"]
    driver = build_driver(download_dir)

    try:
        ec_client = EcountClient(driver, account)

        # ── 로그인 ──────────────────────────────────────────
        ec_client.login()
        print("  [이카운트] 로그인 완료")

        # ── 판매현황 메뉴 이동 ───────────────────────────────
        ec_client.navigate_to_sales()
        print("  [이카운트] 메뉴 이동 완료")

        # ── 엑셀 다운로드 ────────────────────────────────────
        ec_client.download_excel()
        print("  [이카운트] 엑셀 다운로드 클릭")

        filepath = wait_for_download(download_dir)
        print(f"  [이카운트] 다운로드 완료: {filepath}")

        # ── 구글 시트 업로드 ─────────────────────────────────
        rows = read_excel_to_rows(filepath)
        gs.upload_user_entered("이카운트 판매(영림원전송완료)", rows)
        gs.format_columns("이카운트 판매(영림원전송완료)", {
            "C": "TEXT",
            "I": "TEXT",
            "J": "TEXT",
            "G": "NUMBER",
            "H": "NUMBER",
        })
        print("  [이카운트] 시트 업로드 완료")

        try:
            os.remove(filepath)
        except Exception:
            pass

    except Exception:
        print("[이카운트] 오류 발생:")
        traceback.print_exc()
        raise
    finally:
        driver.quit()


# ════════════════════════════════════════════════════════════════
# 8. 진입점
# ════════════════════════════════════════════════════════════════

def main():
    print("=" * 55)
    print(" ERP 자동화 스크립트 시작")
    print(f" {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 55)

    # ── config 로드 ─────────────────────────────────────────
    try:
        config = load_config()
    except FileNotFoundError as e:
        print(f"[FATAL] {e}")
        sys.exit(1)

    # ── 다운로드 임시 폴더 ───────────────────────────────────
    download_dir = str(get_base_path() / "_downloads_tmp")
    os.makedirs(download_dir, exist_ok=True)

    # ── Google Sheets 초기화 ────────────────────────────────
    try:
        sa_info = config["google"]["service_account"]
        gs = GoogleSheet(sa_info)
        print("[Sheets] 인증 완료")
    except Exception:
        print("[FATAL] Google Sheets 인증 실패:")
        traceback.print_exc()
        sys.exit(1)

    # ── 작업 1: 영림원 ERP ──────────────────────────────────
    erp_ok = True
    try:
        run_erp_youngimwon(config, gs, download_dir)
    except Exception:
        print("[영림원 ERP] 실패 - 이카운트 작업은 계속 진행합니다.")
        erp_ok = False

    # ── 작업 2: 이카운트 ────────────────────────────────────
    ecount_ok = True
    try:
        run_ecount_sales(config, gs, download_dir)
    except Exception:
        print("[이카운트] 실패")
        ecount_ok = False

    # ── 임시 폴더 정리 ───────────────────────────────────────
    try:
        shutil.rmtree(download_dir, ignore_errors=True)
    except Exception:
        pass

    # ── 최종 결과 ────────────────────────────────────────────
    print("\n" + "=" * 55)
    print(f" 영림원 ERP : {'✅ 성공' if erp_ok    else '❌ 실패'}")
    print(f" 이카운트   : {'✅ 성공' if ecount_ok else '❌ 실패'}")
    print("=" * 55)

    if not (erp_ok and ecount_ok):
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ 실행 중 오류가 발생했습니다: {e}")
    finally:
        input("\n종료하려면 Enter를 누르세요...")