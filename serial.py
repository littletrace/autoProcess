import os
import sys
import json
import glob
import time
from datetime import datetime

# ── Selenium (웹 파트) ──────────────────────────────────────
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ── Windows GUI (엑셀 애드인 파트) ──────────────────────────
import win32com.client
import win32gui
import win32con
import win32api
import win32process
import psutil
import pyautogui
from pywinauto import Application

# ── Google Sheets ───────────────────────────────────────────
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

pyautogui.FAILSAFE = True

# ==============================================================
# 0. 윈도우 cp949 인코딩 오류 방지
# ==============================================================
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except (AttributeError, ValueError):
    pass

# ==============================================================
# 1. config.json 로더
#    - 로그인 정보(ecount_main)와 구글 서비스 계정 정보만 참조
#    - 스프레드시트ID, XPath 등 개별 실행 정보는 이 파일 하단에서 상수로 관리
# ==============================================================
if getattr(sys, "frozen", False):
    current_dir = os.path.dirname(sys.executable)  # --onefile exe 실제 위치 폴더
else:
    current_dir = os.path.dirname(os.path.abspath(__file__))  # .py 직접 실행 시 기존 방식

def get_config_path():
    own_folder = os.path.join(current_dir, "config.json")
    if os.path.exists(own_folder):
        return own_folder
    parent_folder = os.path.join(os.path.dirname(current_dir), "config.json")
    if os.path.exists(parent_folder):
        return parent_folder
    raise FileNotFoundError(f"[ERROR] config.json을 찾을 수 없습니다: {own_folder}")


# 1. 설정 로드
def load_config():
    with open(get_config_path(), "r", encoding="utf-8") as f:
        return json.load(f)
 
 
CONFIG = load_config()
ECOUNT_ACCOUNT = CONFIG["ecount_main"]
GOOGLE_SERVICE_ACCOUNT = CONFIG["google"]["service_account"]
 
if not all(ECOUNT_ACCOUNT.get(k) for k in ("com_code", "user_id", "password")):
    raise ValueError("[ERROR] config.json의 ecount_main 항목(com_code/user_id/password) 중 비어있는 값이 있습니다.")
 
print(f"[INFO] 로그인 정보 로드 완료 (com_code: {ECOUNT_ACCOUNT['com_code']}, user_id: {ECOUNT_ACCOUNT['user_id']})")


# ==============================================================
# 2. 개별 실행 정보 (스프레드시트 ID, XPath 등)
#    - 웹 파트/엑셀 파트가 동일 스프레드시트를 참조하므로 ID 하나로 통합 관리
# ==============================================================
SPREADSHEET_ID = "1BkaU2LrkEomXubn6GY7qlBDYHjrZ-X5pBQAc828M76Q"
SPREADSHEET_URL = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit"

DOWNLOAD_DIR = os.path.join(current_dir, "temp_downloads")

ECOUNT_EXCEL_XPATH = "//*[contains(text(), 'Excel')]"
SEARCH_BTN_XPATH = "//button[@id='searchGroup']"

# [작업 1] 재고현황 → 시리얼정렬
STOCK_TAB = "재고"
STOCK_MENU_XPATHS = [
    "/html/body/div[2]/div[3]/div/div[4]/ul/li[3]/div[1]/a",                               # 재고 I
    "/html/body/div[2]/div[3]/div/div[4]/ul/li[3]/div[2]/ul/li[7]/a",                       # 출력물
    "/html/body/div[2]/div[5]/div[3]/div/div[1]/div[2]/div[2]/div[2]/ul/li[1]/ul/li[1]/a",  # 재고현황
    "/html/body/div[2]/div[5]/div[3]/div/div[1]/div[2]/div[1]/div[2]/ul/li[2]/a[1]",        # 시리얼정렬
]

# [작업 2] 시리얼/로트No.내역현황 → 시리얼정렬
SERIAL_TAB = "2월이후이카운트"
SERIAL_MENU_XPATHS = [
    "/html/body/div[2]/div[3]/div/div[4]/ul/li[4]/div[1]/a",                          # 재고 II
    "/html/body/div[2]/div[3]/div/div[4]/ul/li[4]/div[2]/ul/li[2]/a",                 # 시리얼/로트No.
    "/html/body/div[2]/div[4]/div/ul/li[1]/ul/li[3]/a",                               # 시리얼/로트No.내역현황
    "/html/body/div[2]/div[5]/div[3]/div/div[1]/div[2]/div[1]/div[2]/ul/li[2]/a[1]",  # 시리얼정렬
]

# 엑셀 애드인 파트: 시리얼/로트No.재고조정 → 구글시트 반영 탭
EXCEL_SHEET_NAME = "시리얼재고조정"

SHEETS_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ==============================================================
# 3. 웹 파트 — Selenium 공통 유틸 (기존 browser.py 인라인)
# ==============================================================
def clear_download_dir(download_dir):
    """다운로드 전 이전 잔여 파일을 정리하여 꼬임 방지."""
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
    """.crdownload 임시 확장자가 사라질 때까지 대기 후 최신 파일 경로 반환."""
    print("다운로드 완료를 대기 중입니다...")
    seconds = 0
    while seconds < timeout:
        time.sleep(1)
        files = glob.glob(os.path.join(download_dir, "*"))
        if files:
            crdownloads = [f for f in files if f.endswith(".crdownload")]
            if not crdownloads:
                return max(files, key=os.path.getctime)
        seconds += 1
    return None


def build_driver(download_dir, extra_args=None, extra_prefs=None, insecure=False):
    """격리된 다운로드 경로가 적용된 Chrome 드라이버 생성."""
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


# ==============================================================
# 4. 웹 파트 — 이카운트 Selenium 클라이언트 (기존 ecount_client.py 인라인)
# ==============================================================
LOGIN_URL = "https://login.ecount.com/LOGIN?lan_type=ko-KR"
_REGIST_BTN_XPATH = (
    "//div[contains(@class, 'control-set')]"
    "//*[@id='toolbar_sid_toolbar_item_regist']/button"
)


class EcountWebClient:
    """이카운트 웹(Selenium) 로그인/네비게이션 래퍼."""

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
            print("이카운트 웹 로그인 성공")

            self.handle_new_device_popup()
        except KeyError as e:
            raise KeyError(f"계정 정보 누락: {e}") from e
        except Exception as e:
            raise RuntimeError(f"이카운트 로그인 실패: {e}") from e

    def handle_new_device_popup(self, timeout=5):
        try:
            print("새 기기 로그인 팝업 여부 확인 중...")
            WebDriverWait(self.driver, timeout).until(
                EC.element_to_be_clickable((By.XPATH, _REGIST_BTN_XPATH))
            ).click()
            print("새 기기 등록 팝업 → '등록' 버튼 클릭 완료")
            time.sleep(2)
        except Exception:
            print("새 기기 팝업 없음 → 다음 단계로 진행")

    def navigate(self, xpaths, delay=1.0):
        for idx, xpath in enumerate(xpaths, start=1):
            print(f"  → 메뉴 이동 {idx}/{len(xpaths)}")
            self._click_xpath(xpath)
            time.sleep(delay)

    def click(self, xpath, delay=0.0, timeout=None):
        self._click_xpath(xpath, timeout=timeout)
        if delay:
            time.sleep(delay)


# ==============================================================
# 5. 웹 파트 — 구글시트 "파일 업로드" 클래스 (기존 google_sheet.py 인라인)
#    엑셀/HTML 다운로드 파일을 통째로 시트에 업로드하는 용도.
#    (엑셀 애드인 파트의 값-리스트 갱신 클래스와는 성격이 달라 별도 클래스로 유지)
# ==============================================================
class GoogleSheetUploader:
    """서비스 계정 기반 구글 스프레드시트 '파일 업로드' 래퍼."""

    def __init__(self, service_account_info, spreadsheet_id):
        try:
            creds = Credentials.from_service_account_info(
                service_account_info, scopes=SHEETS_SCOPES
            )
            self._gc = gspread.authorize(creds)
            self._doc = self._gc.open_by_key(spreadsheet_id)
        except Exception as e:
            raise RuntimeError(f"구글 시트 인증/연결 실패: {e}") from e

    @staticmethod
    def _read_tabular(file_path):
        """엑셀로 먼저 시도하고, 실패하면 HTML 테이블로 파싱."""
        try:
            df = pd.read_excel(file_path)
        except Exception:
            print("HTML 형식의 엑셀 파일로 감지되어 변환하여 읽습니다.")
            dfs = pd.read_html(file_path, encoding="utf-8")
            df = dfs[0]
        return df.fillna("")

    def upload_file(self, file_path, worksheet_name, add_timestamp=False, cleanup=True):
        try:
            print(f"\n--- [{worksheet_name}] 구글 스프레드시트 업로드 준비 ---")
            worksheet = self._doc.worksheet(worksheet_name)

            df = self._read_tabular(file_path)
            data = [df.columns.values.tolist()] + df.values.tolist()

            worksheet.batch_clear(["A1:Z20000"])
            worksheet.update(range_name="A1", values=data, value_input_option="USER_ENTERED")
            print(f"[완료] [{worksheet_name}] 업로드 완료")

            if add_timestamp:
                next_row = len(data) + 1
                stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                worksheet.update(
                    values=[[f"업데이트 일시: {stamp}"]],
                    range_name=f"A{next_row}",
                )
        except Exception as e:
            print(f"[오류] 구글 시트 업로드 중 오류가 발생했습니다: {e}")
        finally:
            if cleanup and os.path.exists(file_path):
                try:
                    os.remove(file_path)
                    print("처리가 완료된 로컬 임시 파일을 삭제했습니다.")
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
                            "cell": {
                                "userEnteredFormat": {"numberFormat": type_map[fmt_key]}
                            },
                            "fields": "userEnteredFormat.numberFormat",
                        }
                    }
                )

            if requests:
                self._doc.batch_update({"requests": requests})
                print(f"[완료] [{worksheet_name}] 열 표시 형식 지정 완료")

        except Exception as e:
            print(f"[오류] 열 표시 형식 지정 중 오류가 발생했습니다: {e}")


def _collect_and_upload(client, gsheet, menu_xpaths, worksheet_name, formats=None):
    """메뉴 이동 → 검색 → 엑셀 다운로드 → 구글 시트 업로드 단위 작업."""
    client.navigate(menu_xpaths, delay=2)
    client.click(SEARCH_BTN_XPATH, delay=5)
    client.click(ECOUNT_EXCEL_XPATH)

    latest_file = wait_for_download(DOWNLOAD_DIR, timeout=60)
    if latest_file:
        print(f"[완료] 다운로드 완료 확인: {os.path.basename(latest_file)}")
        gsheet.upload_file(latest_file, worksheet_name)
        if formats:
            gsheet.format_columns(worksheet_name, formats)
    else:
        print("[경고] 파일 다운로드 실패 또는 타임아웃")


def run_web_part():
    """[1/2] 이카운트 웹(Selenium) 다운로드 → 구글시트 업로드."""
    print("\n===== [1/2] 이카운트 웹 다운로드 & 구글시트 업로드 =====")
    driver = build_driver(DOWNLOAD_DIR)
    try:
        client = EcountWebClient(driver)
        client.login(ECOUNT_ACCOUNT)

        gsheet = GoogleSheetUploader(GOOGLE_SERVICE_ACCOUNT, SPREADSHEET_ID)

        print("\n[작업 1] 시리얼정렬(재고현황) 데이터를 수집합니다.")
        stock_formats = {0: 'TEXT', 1: 'TEXT', 3: 'NUMBER_DECIMAL'}
        _collect_and_upload(client, gsheet, STOCK_MENU_XPATHS, STOCK_TAB, formats=stock_formats)

        print("\n[작업 2] 시리얼/로트No.내역현황 데이터를 수집합니다.")
        serial_formats = {1: 'TEXT', 2: 'TEXT', 3: 'TEXT', 5: 'TEXT', 7: 'NUMBER_DECIMAL'}
        _collect_and_upload(client, gsheet, SERIAL_MENU_XPATHS, SERIAL_TAB, formats=serial_formats)

        print("\n[완료] 웹 파트 작업이 모두 종료되었습니다.")
    finally:
        driver.quit()


# ==============================================================
# 6. 엑셀 파트 — 구글시트 "값 리스트" 클래스 (기존 newxl.py 로직 인라인)
#    엑셀 시트 값을 직접 읽고/쓰는 용도. 웹 파트의 파일업로드 클래스와는
#    성격이 달라 별도 클래스로 유지.
# ==============================================================
class GoogleSheetValues:
    """서비스 계정 기반 구글 스프레드시트 '값 리스트' 읽기/쓰기 래퍼."""

    def __init__(self, service_account_info, spreadsheet_url):
        try:
            creds = Credentials.from_service_account_info(
                service_account_info, scopes=SHEETS_SCOPES
            )
            self._gc = gspread.authorize(creds)
            self._doc = self._gc.open_by_url(spreadsheet_url)
            print("[INFO] Google Sheets(값 리스트) 클라이언트 초기화 완료")
        except Exception as e:
            raise RuntimeError(f"구글 시트 인증/연결 실패: {e}") from e

    def paste_to_sheet(self, sheet_name, data):
        """기존 데이터 삭제 → 값 붙여넣기 → 마지막 행에 타임스탬프 기록."""
        worksheet = self._doc.worksheet(sheet_name)
        print(f"[INFO] 시트 연결 완료: '{sheet_name}'")

        worksheet.clear()
        print("[INFO] 기존 데이터 삭제 완료")

        worksheet.update(range_name="A1", values=data, value_input_option="USER_ENTERED")
        print(f"[INFO] 데이터 붙여넣기 완료 — {len(data)}행")

        next_row = len(data) + 1
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
        worksheet.update(
            range_name=f"A{next_row}",
            values=[[timestamp]],
            value_input_option="USER_ENTERED",
        )
        print(f"[INFO] 타임스탬프 입력 완료 — A{next_row}: {timestamp}")

    def read_upload_sheet(self, tab_name="업로드자료"):
        """[업로드자료] 탭의 A~F열 데이터 반환."""
        worksheet = self._doc.worksheet(tab_name)
        all_values = worksheet.get_all_values()
        data = [row[:6] for row in all_values if any(cell.strip() for cell in row[:6])]
        print(f"[INFO] [{tab_name}] 탭 읽기 완료 — {len(data)}행 x 6열")
        return data


# ==============================================================
# 7. 엑셀 파트 — 윈도우/엑셀 프로세스 제어 (기존 newxl.py 로직 인라인)
# ==============================================================
def kill_excel():
    killed = 0
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] and proc.info['name'].lower() == 'excel.exe':
            proc.kill()
            killed += 1

    if killed > 0:
        print(f"[INFO] 엑셀 프로세스 {killed}개 종료됨")
        time.sleep(3)
    else:
        print("[INFO] 실행 중인 엑셀 없음")


def wait_for_excel_ready(xl, timeout=30):
    start = time.time()
    while time.time() - start < timeout:
        try:
            if xl.Ready:
                print("[INFO] 엑셀 Ready 상태 확인됨")
                return True
        except Exception:
            pass
        time.sleep(0.5)
    raise TimeoutError("[ERROR] 엑셀이 Ready 상태가 되지 않았습니다 (timeout)")


def focus_excel_window():
    hwnd = None

    def enum_handler(h, _):
        nonlocal hwnd
        title = win32gui.GetWindowText(h)
        if 'Excel' in title or '통합 문서' in title or 'Book' in title:
            hwnd = h

    win32gui.EnumWindows(enum_handler, None)

    if not hwnd:
        raise RuntimeError("[ERROR] 엑셀 창을 찾을 수 없습니다")

    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)

    # SetForegroundWindow 보안 정책 우회: ALT 키 입력으로 포커스 잠금 해제
    win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
    win32gui.SetForegroundWindow(hwnd)
    win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)

    print(f"[INFO] 엑셀 창 포커스 완료 (hwnd: {hwnd})")
    time.sleep(0.5)


def open_new_excel_addin():
    kill_excel()

    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = True
    xl.WindowState = -4137  # xlMaximized
    xl.DisplayAlerts = False
    print("[INFO] 엑셀 새로 실행됨")

    wait_for_excel_ready(xl)

    workbook = xl.Workbooks.Add()
    print("[INFO] 새 통합문서 생성됨")

    workbook.Activate()
    xl.ActiveWindow.WindowState = -4137
    print("[INFO] 엑셀 창 활성화 완료")

    return xl, workbook


def find_hwnd_by_keyword(keyword: str) -> int:
    """창 제목에 keyword가 포함된 hwnd 반환."""
    result = []

    def enum_handler(hwnd, _):
        title = win32gui.GetWindowText(hwnd)
        if keyword in title:
            result.append(hwnd)

    win32gui.EnumWindows(enum_handler, None)

    if not result:
        raise RuntimeError(f"[ERROR] '{keyword}' 키워드가 포함된 창을 찾을 수 없습니다")

    return result[0]


def click_button_by_name(window_title_keyword: str, button_name: str, timeout: int = 10):
    """hwnd로 창을 직접 찾아 pywinauto로 연결 후 버튼 클릭 (한글 title_re 매칭 실패 우회)."""
    start = time.time()

    hwnd = None
    while time.time() - start < timeout:
        try:
            hwnd = find_hwnd_by_keyword(window_title_keyword)
            break
        except RuntimeError:
            time.sleep(0.5)

    if hwnd is None:
        raise RuntimeError(f"[ERROR] '{window_title_keyword}' 창을 {timeout}초 안에 찾지 못했습니다")

    print(f"  [INFO] 창 발견 hwnd={hwnd} | keyword='{window_title_keyword}'")

    app = Application(backend="uia").connect(handle=hwnd)
    win = app.window(handle=hwnd)

    btn = win.child_window(title=button_name, control_type="Button")
    btn.click_input()
    print(f"  [CLICK] '{window_title_keyword}' 창 > '{button_name}' 버튼 클릭")
    time.sleep(1)


def activate_ecount_tab():
    """pywinauto UIA로 Excel 리본의 '이카운트' TabItem 클릭."""
    hwnd = find_hwnd_by_keyword('Excel')
    app = Application(backend="uia").connect(handle=hwnd)
    excel_win = app.window(handle=hwnd)

    tab = excel_win.child_window(title="이카운트", control_type="TabItem")
    tab.click_input()
    print("  [CLICK] 이카운트 탭 활성화")
    time.sleep(0.5)


# ──────────────────────────────────────────────────────────────
# 리본 버튼 클릭: 텍스트 기반 탐색 우선 → Alt 키팁 시퀀스 폴백
# (기존 로그인/자료입력/자료전송 3개 버튼 트리거를 이 방식으로 교체.
#  paste_to_excel()의 셀보호해제 Alt→R→P→S 시퀀스는 별도 성격이라 유지)
# ──────────────────────────────────────────────────────────────
def _attach_focus_excel(hwnd):
    """AttachThreadInput으로 포그라운드 스레드와 연결 후 안전하게 포커스 전환."""
    fore_hwnd = win32gui.GetForegroundWindow()
    fore_tid = win32process.GetWindowThreadProcessId(fore_hwnd)[0]
    target_tid = win32process.GetWindowThreadProcessId(hwnd)[0]

    attached = False
    if fore_tid != target_tid:
        try:
            win32process.AttachThreadInput(target_tid, fore_tid, True)
            attached = True
        except Exception:
            attached = False

    try:
        win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
        win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)
        win32gui.SetForegroundWindow(hwnd)
        win32gui.BringWindowToTop(hwnd)
    finally:
        if attached:
            try:
                win32process.AttachThreadInput(target_tid, fore_tid, False)
            except Exception:
                pass

    return win32gui.GetForegroundWindow() == hwnd


def _ensure_excel_focus(hwnd, max_attempts=3, retry_delay=0.4):
    """
    [A안] 포커스 전환 성공 여부를 검증하고, 실패 시 재시도.
    이미 포커스가 맞아 있으면 Alt 토글을 다시 보내지 않는다
    (Alt를 반복 입력하면 이미 뜬 키팁이 꺼지는 부작용이 있으므로).
    """
    for attempt in range(1, max_attempts + 1):
        if win32gui.GetForegroundWindow() == hwnd:
            return True
        ok = _attach_focus_excel(hwnd)
        if ok:
            return True
        print(f"  [WARN] 엑셀 포커스 전환 실패 ({attempt}/{max_attempts}) — 재시도")
        time.sleep(retry_delay)
    return False


def send_ribbon_keytips(keys, key_delay=0.3, focus_attempts=3):
    """
    Escape로 기존 리본/키팁 상태를 완전히 해제한 뒤 Alt 키팁 시퀀스 전송.
    (Alt 단독 포커스 트릭과 리본 Alt 토글이 충돌해 오작동하는 문제 방지)

    [A안] 포커스 확보 여부를 매 단계 검증하며, 확보하지 못하면
    (엉뚱한 셀에 'y2y3' 등이 입력되는 오작동을 막기 위해) 키를 보내지 않고
    예외를 발생시켜 상위에서 실패로 처리하도록 한다.
    """
    hwnd = find_hwnd_by_keyword('Excel')

    if not _ensure_excel_focus(hwnd, max_attempts=focus_attempts):
        raise RuntimeError(
            "[ERROR] 엑셀 창 포커스를 확보하지 못해 리본 키 입력을 중단합니다 (오입력 방지)"
        )

    time.sleep(0.3)
    pyautogui.press('escape')
    time.sleep(0.2)
    pyautogui.press('escape')
    time.sleep(0.2)

    # Escape 처리 중 포커스가 이탈했는지 재확인 후에만 키 전송
    if win32gui.GetForegroundWindow() != hwnd:
        if not _ensure_excel_focus(hwnd, max_attempts=focus_attempts):
            raise RuntimeError(
                "[ERROR] Escape 처리 중 엑셀 포커스가 이탈했습니다 — 리본 키 입력을 중단합니다 (오입력 방지)"
            )

    for key in keys:
        pyautogui.press(key)
        print(f"  [KEY] '{key}' 입력")
        time.sleep(key_delay)


def click_ribbon_button_smart(button_title_re, fallback_keys, search_timeout=8, search_attempts=2):
    """
    이카운트 리본 버튼을 title_re 텍스트 기반으로 우선 탐색하여 클릭.

    [C안] 탐색 타임아웃을 늘리고(기존 5초→8초) search_attempts회까지 재시도해,
    UI 로딩 지연으로 인한 오탐(텍스트 탐색 실패 → 키팁 폴백行)을 줄인다.
    모든 시도가 실패한 경우에만 Alt 키팁 시퀀스(fallback_keys)로 폴백한다.
    """
    last_error = None
    for attempt in range(1, search_attempts + 1):
        try:
            hwnd = find_hwnd_by_keyword('Excel')
            app = Application(backend="uia").connect(handle=hwnd)
            excel_win = app.window(handle=hwnd)
            btn = excel_win.child_window(title_re=button_title_re, control_type="Button")
            btn.wait("exists enabled visible ready", timeout=search_timeout)
            btn.click_input()
            print(f"[CLICK] 텍스트 탐색으로 버튼 클릭 성공 (패턴: {button_title_re}, 시도 {attempt}/{search_attempts})")
            return True
        except Exception as e:
            last_error = e
            print(f"  [WARN] 텍스트 탐색 실패 (시도 {attempt}/{search_attempts}): {e}")
            time.sleep(0.5)

    print(f"[FALLBACK] 텍스트 탐색 {search_attempts}회 모두 실패({last_error}) → Alt 키팁 시퀀스로 전환")
    send_ribbon_keytips(fallback_keys)
    return False


def open_ecount_login():
    """이카운트 애드인 로그인 창 열기 (텍스트 탐색 우선, 실패 시 Alt>Y2Y4 폴백)."""
    print("[INFO] 이카운트 로그인 리본 메뉴 진입 시작")
    focus_excel_window()
    time.sleep(1)

    activate_ecount_tab()
    time.sleep(0.3)

    click_ribbon_button_smart(".*로그인.*", ['alt', 'y', '2', 'y', '4'])

    print("[INFO] 로그인 창 호출 완료 — 창 로딩 대기 중")
    time.sleep(2)


def login_ecount(com_code: str, user_id: str, password: str):
    """이카운트 애드인 로그인 팝업에 값 입력 (팝업 EDIT 필드 특성상 키 입력 방식 유지)."""
    print("[INFO] 로그인 정보 입력 시작")

    pyautogui.hotkey('ctrl', 'a')
    pyautogui.typewrite(com_code, interval=0.05)
    print(f"  [INPUT] CODE: {com_code}")
    time.sleep(0.2)

    pyautogui.press('tab')
    time.sleep(0.2)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.typewrite(user_id, interval=0.05)
    print(f"  [INPUT] USER ID: {user_id}")
    time.sleep(0.2)

    pyautogui.press('tab')
    time.sleep(0.2)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.typewrite(password, interval=0.05)
    print(f"  [INPUT] PASSWORD: {'*' * len(password)}")
    time.sleep(0.2)

    pyautogui.press('enter')
    print("[INFO] 로그인 버튼 Enter 입력")
    time.sleep(3)


def open_serial_stock():
    """자료 입력하기 → 재고 → 시리얼/로트No.재고조정 → 검색."""
    print("[INFO] '자료 입력하기' 리본 메뉴 진입")
    focus_excel_window()
    time.sleep(0.5)

    activate_ecount_tab()
    time.sleep(0.3)

    click_ribbon_button_smart(r".*자료\s*입력.*", ['alt', 'y', '2', 'y', '1'])

    print("[INFO] '자료 입력하기' 팝업 로딩 대기")
    time.sleep(2)

    click_button_by_name(window_title_keyword="자료", button_name="재고")
    time.sleep(1)

    click_button_by_name(window_title_keyword="자료", button_name="시리얼/로트No.재고조정")
    time.sleep(2)

    print("[INFO] '시리얼/로트No.재고조정' 검색 팝업 > 검색 실행")
    try:
        click_button_by_name(window_title_keyword="시리얼", button_name="검색")
    except Exception:
        print("  [FALLBACK] 검색 버튼 미발견 — Enter 키로 대체")
        pyautogui.press('enter')

    print("[INFO] 검색 실행 완료 — 데이터 로딩 대기 중")
    time.sleep(5)


def read_excel_data(wb) -> list:
    """활성 시트의 A~F열 데이터를 2차원 리스트로 반환 (헤더 포함)."""
    ws = wb.ActiveSheet

    last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  # -4162 = xlUp

    if last_row < 1:
        raise ValueError("[ERROR] 엑셀에 데이터가 없습니다")

    data = []
    for row in range(1, last_row + 1):
        row_data = []
        for col in range(1, 7):  # A=1 ~ F=6
            cell_val = ws.Cells(row, col).Value
            row_data.append(str(cell_val) if cell_val is not None else "")
        data.append(row_data)

    print(f"[INFO] 엑셀 데이터 읽기 완료 — {last_row}행 x 6열")
    return data


def paste_to_excel(xl, wb, data: list):
    """
    1. 셀보호 해제 (Alt > R > P > S) — 셀보호 해제 전용 시퀀스라 텍스트 탐색 대상이 아니므로 기존 방식 유지
    2. 기존 데이터 삭제
    3. 업로드자료 붙여넣기
    """
    ws = wb.ActiveSheet

    focus_excel_window()
    time.sleep(0.5)

    print("[INFO] 셀보호 해제 시작 (Alt > R > P > S)")
    pyautogui.press('alt')
    time.sleep(0.3)
    pyautogui.press('r')
    time.sleep(0.3)
    pyautogui.press('p')
    time.sleep(0.3)
    pyautogui.press('s')
    time.sleep(1)
    print("[INFO] 셀보호 해제 완료")

    ws.Cells(1, 1).Select()
    ws.UsedRange.ClearContents()
    print("[INFO] 엑셀 기존 데이터 삭제 완료")

    print(f"[INFO] 엑셀에 데이터 쓰기 시작 — {len(data)}행")
    for r_idx, row in enumerate(data, start=1):
        for c_idx, val in enumerate(row, start=1):
            ws.Cells(r_idx, c_idx).Value = val
    print(f"[INFO] 엑셀 데이터 쓰기 완료 — {len(data)}행 x 6열")
    time.sleep(1)


def wait_for_gas_done():
    """구글시트 GAS 스크립트 실행 완료 후 Enter 입력 대기."""
    print("")
    print("========================================================")
    print("  [대기] 구글시트 '시리얼정열' 탭에서")
    print("         GAS 스크립트 실행 버튼을 클릭하세요.")
    print("  데이터 생성이 완료되면 이 터미널 창에서 Enter를 누르세요.")
    print("========================================================")
    input()
    print("[INFO] Enter 입력 확인 — 다음 단계로 진행합니다")


def send_ecount_data():
    """자료전송하기 (텍스트 탐색 우선, 실패 시 Alt>Y2Y3 폴백)."""
    print("[INFO] 이카운트 자료전송하기 진입")
    focus_excel_window()
    time.sleep(1)

    activate_ecount_tab()
    time.sleep(0.3)

    click_ribbon_button_smart(r".*자료\s*전송.*", ['alt', 'y', '2', 'y', '3'])

    print("[INFO] 자료전송하기 완료")
    time.sleep(3)


def run_excel_part():
    """[2/2] 엑셀 애드인 자동화 → 구글시트 연동 → 자료전송."""
    print("\n===== [2/2] 엑셀 애드인 자동화 & 구글시트 연동 =====")
    com_code = ECOUNT_ACCOUNT["com_code"]
    user_id = ECOUNT_ACCOUNT["user_id"]
    password = ECOUNT_ACCOUNT["password"]

    xl, wb = open_new_excel_addin()
    print(f"[완료] 활성 시트: {wb.ActiveSheet.Name}")

    open_ecount_login()
    login_ecount(com_code, user_id, password)
    print("[완료] 이카운트 애드인 로그인 완료")

    open_serial_stock()
    print("[완료] 시리얼재고조정 탭 열림")

    data = read_excel_data(wb)
    gvalues = GoogleSheetValues(GOOGLE_SERVICE_ACCOUNT, SPREADSHEET_URL)
    gvalues.paste_to_sheet(EXCEL_SHEET_NAME, data)
    print("[완료] 구글시트 업데이트 완료")

    wait_for_gas_done()
    print("[완료] GAS 실행 확인")

    upload_data = gvalues.read_upload_sheet()
    print("[완료] 업로드자료 탭 읽기 완료")

    paste_to_excel(xl, wb, upload_data)
    print("[완료] 엑셀 붙여넣기 완료")

    send_ecount_data()

    print("\n[완료] 엑셀 파트 작업이 모두 종료되었습니다.")


# ==============================================================
# 8. 실행 진입점 — 웹 파트 → 엑셀 파트 순차 실행 (생략 분기 없음)
# ==============================================================
if __name__ == "__main__":
    try:
        print("이카운트 자동화 통합 프로세스를 시작합니다 (웹 → 엑셀 순차 진행)")

        run_web_part()

        print("\n[진행] 웹 작업이 완료되었습니다. 이어서 엑셀 애드인 프로세스를 진행합니다...")
        run_excel_part()

        print("\n✅ 전체 프로세스 완료")
    except Exception as e:
        print(f"\n❌ 실행 중 오류가 발생했습니다: {e}")
    finally:
        input("\n종료하려면 Enter를 누르세요...")