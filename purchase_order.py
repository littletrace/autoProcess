"""

─────────────────────────────────────────────────────────────────────────────
이카운트 ERP 판매/발주 데이터를 다운로드하여 Google Sheets에 업로드하는
독립 실행 스크립트. PyInstaller exe 빌드를 고려한 단일 파일 구조.
─────────────────────────────────────────────────────────────────────────────
"""

import os
import sys
import json
import time
import glob
import shutil

import openpyxl
import gspread
from google.oauth2.service_account import Credentials

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


# ─────────────────────────────────────────────────────────────────────────────
# 상수
# ─────────────────────────────────────────────────────────────────────────────
SPREADSHEET_ID = "1z_Hn9GGMQvFjGdwQp1qa4ycadcbD_dHweKPelV7pkHM"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

ECOUNT_LOGIN_URL = "https://login.ecount.com/Login/?lan_type=ko-KR"
DOWNLOAD_TIMEOUT = 60  # 초


# ─────────────────────────────────────────────────────────────────────────────
# 경로 설정 (PyInstaller exe 및 일반 실행 모두 지원)
# ─────────────────────────────────────────────────────────────────────────────
def get_script_dir() -> str:
    """스크립트(또는 exe) 자신이 위치한 폴더를 반환."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


SCRIPT_DIR = get_script_dir()
DOWNLOAD_DIR = os.path.join(SCRIPT_DIR, "temp_downloads")


# ─────────────────────────────────────────────────────────────────────────────
# 설정 로더
# ─────────────────────────────────────────────────────────────────────────────
def find_config() -> str:
    """
    config.json을 아래 순서로 탐색하여 첫 번째로 발견된 경로를 반환.
      1. 스크립트(exe)와 동일한 폴더
      2. 그 상위 폴더
    두 곳 모두 없으면 FileNotFoundError 발생.
    """
    candidates = [
        os.path.join(SCRIPT_DIR, "config.json"),
        os.path.join(os.path.dirname(SCRIPT_DIR), "config.json"),
    ]
    for path in candidates:
        if os.path.exists(path):
            print(f"  [설정] config.json 로드: {path}")
            return path
    searched = "\n  ".join(candidates)
    raise FileNotFoundError(
        f"config.json을 찾을 수 없습니다. 탐색한 경로:\n  {searched}"
    )


def load_config() -> dict:
    """config.json 로드."""
    with open(find_config(), encoding="utf-8") as f:
        return json.load(f)


# ─────────────────────────────────────────────────────────────────────────────
# GoogleSheet 클래스
# ─────────────────────────────────────────────────────────────────────────────
class GoogleSheet:
    """
    gspread 기반 구글 스프레드시트 래퍼.
    service_account_info 딕셔너리로 직접 인증 (파일 경로 불필요).
    """

    def __init__(self, service_account_info: dict, spreadsheet_id: str):
        creds = Credentials.from_service_account_info(
            service_account_info, scopes=SCOPES
        )
        gc = gspread.authorize(creds)
        self._doc = gc.open_by_key(spreadsheet_id)

    def clear_and_upload(self, sheet_name: str, rows: list[list]) -> None:
        """지정 탭을 초기화하고 rows 데이터를 RAW로 업로드."""
        ws = self._doc.worksheet(sheet_name)
        ws.batch_clear(["A1:Z20000"])
        ws.update(range_name="A1", values=rows, value_input_option="RAW")
        print(f"  [Sheets] '{sheet_name}' 업로드 완료 ({len(rows)}행)")

    def format_columns(self, sheet_name: str, formats: dict) -> None:
        """
        열별 숫자 서식 지정. Sheets API v4 batchUpdate 사용.

        Args:
            sheet_name : 대상 탭 이름
            formats    : {열번호(0-based): 'TEXT' | 'NUMBER' | 'USER_ENTERED'} dict
                         USER_ENTERED는 서식 미지정이므로 실질적으로
                         TEXT / NUMBER 열만 넘기면 됨.
        """
        ws = self._doc.worksheet(sheet_name)
        sheet_id = ws.id

        _FMT = {
            "TEXT":   {"type": "TEXT",   "pattern": "@"},
            "NUMBER": {"type": "NUMBER", "pattern": "#,##0"},
        }

        requests = []
        for col_index, fmt in formats.items():
            fmt_key = fmt.upper()
            if fmt_key not in _FMT:
                continue  # USER_ENTERED 등 미정의 키는 건너뜀
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startColumnIndex": col_index,
                        "endColumnIndex": col_index + 1,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": _FMT[fmt_key]
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat",
                }
            })

        if requests:
            self._doc.batch_update({"requests": requests})
            print(f"  [Sheets] '{sheet_name}' 열 서식 지정 완료")


# ─────────────────────────────────────────────────────────────────────────────
# EcountClient 클래스 (웹 자동화)
# ─────────────────────────────────────────────────────────────────────────────
class EcountClient:
    """
    Selenium 기반 이카운트 ERP 웹 자동화 클라이언트.
    win32gui / pyautogui 의존성 없음.
    """

    def __init__(self, download_dir: str):
        self.download_dir = download_dir
        self.driver: webdriver.Chrome | None = None

    # ── 드라이버 ──────────────────────────────────────────────────────────────
    def start(self) -> None:
        """Chrome 드라이버를 동적으로 설치·실행."""
        options = webdriver.ChromeOptions()
        prefs = {
            "download.default_directory": self.download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        }
        options.add_experimental_option("prefs", prefs)
        options.add_argument("--start-maximized")
        options.add_argument("--disable-notifications")

        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)
        self.wait = WebDriverWait(self.driver, 20)
        print("  [Chrome] 드라이버 시작 완료")

    def quit(self) -> None:
        if self.driver:
            self.driver.quit()
            self.driver = None
            print("  [Chrome] 드라이버 종료")

    # ── 공통 헬퍼 ─────────────────────────────────────────────────────────────
    def _click_xpath(self, xpath: str, desc: str = "") -> None:
        """XPath 요소가 클릭 가능해질 때까지 대기 후 클릭."""
        elem = self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        elem.click()
        label = f" ({desc})" if desc else ""
        print(f"  [클릭]{label} {xpath[:60]}...")
        time.sleep(0.8)

    def _wait_for_download(self, timeout: int = DOWNLOAD_TIMEOUT) -> str:
        """
        download_dir 내에 .tmp / .crdownload 가 사라지고
        새 파일(.xlsx 또는 .xls)이 생성될 때까지 대기.
        완성된 파일의 전체 경로를 반환.
        """
        print("  [다운로드] 완료 대기 중...", end="", flush=True)
        deadline = time.time() + timeout
        while time.time() < deadline:
            # 미완성 임시 파일이 존재하면 계속 대기
            tmp_files = (
                glob.glob(os.path.join(self.download_dir, "*.tmp"))
                + glob.glob(os.path.join(self.download_dir, "*.crdownload"))
            )
            if tmp_files:
                print(".", end="", flush=True)
                time.sleep(1)
                continue

            # 완성된 엑셀 파일 탐색
            excel_files = (
                glob.glob(os.path.join(self.download_dir, "*.xlsx"))
                + glob.glob(os.path.join(self.download_dir, "*.xls"))
            )
            if excel_files:
                # 가장 최근에 수정된 파일 반환
                latest = max(excel_files, key=os.path.getmtime)
                print(f"\n  [다운로드] 완료: {os.path.basename(latest)}")
                return latest

            time.sleep(1)

        raise TimeoutError(f"다운로드 타임아웃 ({timeout}초 초과)")

    # ── 로그인 ────────────────────────────────────────────────────────────────
    def login(self, company_code: str, user_id: str, password: str) -> None:
        """이카운트 로그인 + 새 기기 접속 팝업 자동 처리."""
        print("  [로그인] 이카운트 로그인 페이지로 이동합니다...")
        self.driver.get(ECOUNT_LOGIN_URL)

        self.wait.until(
            EC.presence_of_element_located((By.ID, "com_code"))
        ).send_keys(company_code)

        self.driver.find_element(By.ID, "id").send_keys(user_id)

        pw_input = self.driver.find_element(By.ID, "passwd")
        pw_input.send_keys(password)
        pw_input.send_keys(Keys.RETURN)

        print(f"  [로그인] {company_code} / {user_id} 로그인 시도")
        time.sleep(4)

        self._handle_new_device_popup()

    def _handle_new_device_popup(self) -> None:
        """
        새 기기 로그인 팝업이 뜨면 '등록' 버튼 클릭.
        팝업이 없으면 조용히 통과.
        """
        _REGIST_BTN_XPATH = (
            "//div[contains(@class, 'control-set')]"
            "//*[@id='toolbar_sid_toolbar_item_regist']/button"
        )
        try:
            print("  [팝업] 새 기기 로그인 팝업 확인 중...")
            WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, _REGIST_BTN_XPATH))
            ).click()
            print("  [팝업] 새 기기 등록 완료")
            time.sleep(2)
        except Exception:
            print("  [팝업] 새 기기 팝업 없음 → 계속 진행")

    # ── 작업 1: 판매현황(발주체크) 다운로드 ─────────────────────────────────
    def download_sales_status(self) -> str:
        """
        이카운트 메뉴 → 판매현황(발주체크) 엑셀 다운로드.
        완성된 파일 경로 반환.
        """
        print("\n  [작업1] 판매현황(발주체크) 다운로드 시작")

        steps = [
            ("/html/body/div[2]/div[3]/div/div[4]/ul/li[3]/div[1]/a",          "재고1"),
            ("/html/body/div[2]/div[3]/div/div[4]/ul/li[3]/div[2]/ul/li[7]/a", "출력물"),
            (
                "/html/body/div[2]/div[5]/div[3]/div/div[1]/div[2]/div[2]"
                "/div[2]/ul/li[2]/ul/li[1]/a",
                "판매현황",
            ),
            (
                "/html/body/div[2]/div[5]/div[3]/div/div[1]/div/div[1]"
                "/div[2]/ul/li[6]/a",
                "발주체크",
            ),
            (
                "/html/body/div[2]/div[5]/div[3]/div/div[1]/div/div[1]"
                "/div[2]/div[2]/div[1]/div[1]/div/div/button[2]",
                "꺽쇠버튼",
            ),
            ("//*[contains(text(), 'Excel')]", "엑셀다운로드"),
        ]

        for xpath, desc in steps:
            self._click_xpath(xpath, desc)

        return self._wait_for_download()

    # ── 작업 2: 발주서(670989) 다운로드 ─────────────────────────────────────
    def download_purchase_order(self) -> str:
        """
        이카운트 메뉴 → 발주서(670989) 엑셀 다운로드.
        완성된 파일 경로 반환.
        """
        print("\n  [작업2] 발주서(670989) 다운로드 시작")

        steps = [
            ("/html/body/div[2]/div[3]/div/div[4]/ul/li[3]/div[1]/a",          "재고1"),
            ("/html/body/div[2]/div[3]/div/div[4]/ul/li[3]/div[2]/ul/li[5]/a", "출력물"),
            (
                "/html/body/div[2]/div[5]/div[3]/div/div[1]/div[2]/div[2]"
                "/div[2]/ul/li[3]/ul/li[4]/a",
                "발주서현황",
            ),
            (
                "/html/body/div[2]/div[5]/div[3]/div/div[1]/div[2]/div[1]"
                "/div[2]/ul/li[2]/a[1]",
                "발주체크 탭",
            ),
            (
                "/html/body/div[2]/div[5]/div[3]/div/div[1]/div[2]/div[1]"
                "/div[2]/div[2]/div[1]/div/div[1]/button[1]",
                "검색",
            ),
            (
                "/html/body/div[2]/div[5]/div[3]/div/div[1]/div[2]/div[3]"
                "/div/div[1]/div[3]/div/button",
                "엑셀",
            ),
        ]

        for xpath, desc in steps:
            self._click_xpath(xpath, desc)

        return self._wait_for_download()


# ─────────────────────────────────────────────────────────────────────────────
# 유틸리티
# ─────────────────────────────────────────────────────────────────────────────
def prepare_download_dir(path: str) -> None:
    """다운로드 폴더를 비우고 생성."""
    if os.path.exists(path):
        shutil.rmtree(path)
    os.makedirs(path, exist_ok=True)
    print(f"  [폴더] temp_downloads 초기화: {path}")


def clean_download_dir(path: str) -> None:
    """작업 후 임시 다운로드 파일 전체 삭제."""
    if os.path.exists(path):
        shutil.rmtree(path)
        print(f"  [폴더] temp_downloads 삭제 완료")


def excel_to_rows(filepath: str) -> list[list]:
    """
    엑셀 파일을 읽어 2D 리스트로 변환.
    - None → 빈 문자열
    - int / float → 숫자 타입 그대로 유지 (문자열 변환 시 Sheets가 ' 접두사 삽입)
    - 나머지 → str 변환
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    rows = []
    for row in ws.iter_rows(values_only=True):
        converted = []
        for v in row:
            if v is None:
                converted.append("")
            elif isinstance(v, (int, float)):
                converted.append(v)          # 숫자 타입 유지
            else:
                converted.append(str(v))
        rows.append(converted)
    return rows


# ─────────────────────────────────────────────────────────────────────────────
# 메인 실행 흐름
# ─────────────────────────────────────────────────────────────────────────────
def main() -> None:
    print("=" * 60)
    print("  Order Form 자동화 시작")
    print("=" * 60)

    # ── 설정 로드 ─────────────────────────────────────────────────────────────
    cfg = load_config()
    sa_info = cfg["google"]["service_account"]          # 서비스 계정 딕셔너리
    main_acc = cfg["ecount_main"]                       # {company, id, pw}
    sub_acc  = cfg["ecount_sub"]                        # {company, id, pw}

    # ── Google Sheets 클라이언트 초기화 ───────────────────────────────────────
    print("\n[1/4] Google Sheets 인증 중...")
    gs = GoogleSheet(sa_info, SPREADSHEET_ID)
    print("  인증 완료")

    # ── 다운로드 폴더 준비 ────────────────────────────────────────────────────
    prepare_download_dir(DOWNLOAD_DIR)

    # ══════════════════════════════════════════════════════════════════════════
    # 작업 1: ecount_main → 판매현황(발주체크) → 시트 업로드
    # ══════════════════════════════════════════════════════════════════════════
    print("\n[2/4] 작업1: 판매현황(발주체크)")
    client1 = EcountClient(DOWNLOAD_DIR)
    try:
        client1.start()
        client1.login(
            company_code=main_acc["com_code"],
            user_id=main_acc["user_id"],
            password=main_acc["password"],
        )
        file1 = client1.download_sales_status()
    finally:
        client1.quit()

    rows1 = excel_to_rows(file1)
    gs.clear_and_upload("판매현황(발주체크)", rows1)
    # E열(4) 텍스트 / F,G,H,I열(5,6,7,8) 숫자 / 나머지 USER_ENTERED(생략)
    gs.format_columns("판매현황(발주체크)", {
        4: "TEXT",
        5: "NUMBER",
        6: "NUMBER",
        7: "NUMBER",
        8: "NUMBER",
    })
    os.remove(file1)
    print(f"  임시 파일 삭제: {os.path.basename(file1)}")

    # ══════════════════════════════════════════════════════════════════════════
    # 작업 2: ecount_sub → 발주서(670989) → 시트 업로드
    # ──────────────────────────────────────────────────────────────────────────
    # ※ 요구사항 명세상 업로드 대상 탭이 '판매현황(발주체크)'로 동일하게 기재됨.
    #   별도 탭(예: '발주서(670989)')에 올려야 한다면 아래 sheet_name을 수정하세요.
    # ══════════════════════════════════════════════════════════════════════════
    print("\n[3/4] 작업2: 발주서(670989)")
    client2 = EcountClient(DOWNLOAD_DIR)
    try:
        client2.start()
        client2.login(
            company_code=sub_acc["com_code"],
            user_id=sub_acc["user_id"],
            password=sub_acc["password"],
        )
        file2 = client2.download_purchase_order()
    finally:
        client2.quit()

    rows2 = excel_to_rows(file2)
    gs.clear_and_upload("발주서(670989)", rows2)
    # C,D,E,F열(2,3,4,5) 숫자 / 나머지 USER_ENTERED(생략)
    gs.format_columns("발주서(670989)", {
        2: "NUMBER",
        3: "NUMBER",
        4: "NUMBER",
        5: "NUMBER",
    })
    os.remove(file2)
    print(f"  임시 파일 삭제: {os.path.basename(file2)}")

    # ── 마무리 ────────────────────────────────────────────────────────────────
    print("\n[4/4] 임시 폴더 정리...")
    clean_download_dir(DOWNLOAD_DIR)

    print("\n" + "=" * 60)
    print("  모든 작업 완료")
    print("=" * 60)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ 실행 중 오류가 발생했습니다: {e}")
    finally:
        input("\n종료하려면 Enter를 누르세요...")