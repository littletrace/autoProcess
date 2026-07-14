"""
delivery_order.py
이카운트 ERP → 구글 스프레드시트 자동화
 - 판매현황       → 탭: '이카운트 판매'
 - 출하지시서현황 → 탭: '이카운트 출하지시서'
"""

import os
import sys
import json
import time
import glob
import shutil
from datetime import datetime

# ── Windows cp949 인코딩 오류 방지 ──────────────────────────────────────────
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except (AttributeError, ValueError):
    pass

# ── 외부 라이브러리 ──────────────────────────────────────────────────────────
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import gspread
from google.oauth2.service_account import Credentials

# ── 상수 ────────────────────────────────────────────────────────────────────
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
SPREADSHEET_ID = "1z_Hn9GGMQvFjGdwQp1qa4ycadcbD_dHweKPelV7pkHM"
LOGIN_URL      = "https://login.ecount.com/LOGIN?lan_type=ko-KR"

# ── config.json 경로 탐색 ────────────────────────────────────────────────────
def get_config_path():
    if getattr(sys, "frozen", False):
        # exe 실행: exe 파일 위치 → 한 단계 위 순으로 탐색
        candidates = [
            os.path.join(os.path.dirname(sys.executable), "config.json"),
            os.path.join(os.path.dirname(os.path.dirname(sys.executable)), "config.json"),
        ]
    else:
        # .py 실행: 스크립트 폴더 → 한 단계 위 순으로 탐색
        script_dir = os.path.dirname(os.path.abspath(__file__))
        candidates = [
            os.path.join(script_dir, "config.json"),
            os.path.join(os.path.dirname(script_dir), "config.json"),
        ]
    for path in candidates:
        if os.path.exists(path):
            return path
    # 못 찾으면 첫 번째 경로 반환 (FileNotFoundError 메시지용)
    return candidates[0]

# ── config 로드 ──────────────────────────────────────────────────────────────
def load_config():
    path = get_config_path()
    print(f"📂 config 경로: {path}")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

# ── ChromeDriver 설정 ────────────────────────────────────────────────────────
def build_driver(download_dir: str) -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    prefs = {
        "download.default_directory":         download_dir,
        "download.prompt_for_download":        False,
        "download.directory_upgrade":          True,
        "safebrowsing.enabled":                True,
    }
    options.add_experimental_option("prefs", prefs)
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)

# ── 이카운트 로그인 ──────────────────────────────────────────────────────────
def ecount_login(driver: webdriver.Chrome, cred: dict):
    print("🔐 이카운트 로그인 페이지로 이동합니다...")
    driver.get(LOGIN_URL)
    wait = WebDriverWait(driver, 20)

    wait.until(EC.presence_of_element_located((By.ID, "com_code"))).send_keys(cred["com_code"])
    driver.find_element(By.ID, "id").send_keys(cred["user_id"])
    pw = driver.find_element(By.ID, "passwd")
    pw.send_keys(cred["password"])
    pw.send_keys(Keys.RETURN)
    print("🎉 로그인 성공!")
    time.sleep(4)

    # 새 기기 등록 팝업 처리
    try:
        print("📋 새 기기 로그인 팝업 여부 확인 중...")
        popup_xpath = "//div[contains(@class, 'control-set')]//*[@id='toolbar_sid_toolbar_item_regist']/button"
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, popup_xpath))
        ).click()
        print("✅ 새 기기 등록 팝업 → '등록' 버튼 클릭 완료")
        time.sleep(2)
    except Exception:
        print("ℹ️  새 기기 팝업 없음 → 다음 단계로 진행")

# ── XPath 클릭 헬퍼 ──────────────────────────────────────────────────────────
def click_xpath(driver: webdriver.Chrome, xpath: str, label: str = "", wait_sec: float = 2):
    try:
        el = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        el.click()
        if label:
            print(f"  ✔ {label}")
        time.sleep(wait_sec)
    except Exception as e:
        raise RuntimeError(f"클릭 실패 [{label}]: {e}")

# ── 다운로드 대기 ─────────────────────────────────────────────────────────────
def wait_for_download(download_dir: str, timeout: int = 60) -> str:
    """
    temp_downloads/ 에서 .crdownload/.tmp 가 아닌 파일이 생길 때까지 대기.
    완성된 파일 경로 반환.
    """
    print("⏳ 파일 다운로드 대기 중...")
    deadline = time.time() + timeout
    while time.time() < deadline:
        files = [
            f for f in glob.glob(os.path.join(download_dir, "*"))
            if not f.endswith(".crdownload") and not f.endswith(".tmp")
            and os.path.isfile(f)
        ]
        if files:
            # 가장 최근 수정 파일
            latest = max(files, key=os.path.getmtime)
            print(f"  ✅ 다운로드 완료: {os.path.basename(latest)}")
            return latest
        time.sleep(1)
    raise TimeoutError(f"다운로드 타임아웃 ({timeout}초)")

# ── 엑셀 파일 읽기 ───────────────────────────────────────────────────────────
def read_excel_file(filepath: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(filepath, dtype=str)
        print("  📊 pd.read_excel() 성공")
    except Exception:
        try:
            df = pd.read_html(filepath)[0].astype(str)
            print("  📊 pd.read_html() 폴백 성공")
        except Exception as e:
            raise RuntimeError(f"파일 읽기 실패: {e}")
    df = df.fillna("")
    return df

# ── 열 서식 설정 ─────────────────────────────────────────────────────────────
# 탭별 열 서식: {0-based 열 인덱스: "TEXT" | "NUMBER"}
COLUMN_FORMATS = {
    "이카운트 출하지시서": {
        1: "TEXT",    # B열
        4: "NUMBER",  # E열
    },
    "이카운트 판매": {
        1: "TEXT",    # B열
        2: "TEXT",    # C열
        8: "TEXT",    # I열
        9: "TEXT",    # J열
        6: "NUMBER",  # G열
        7: "NUMBER",  # H열
    },
}

def format_columns(spreadsheet, sheet_tab: str):
    """탭별 열 서식을 Google Sheets API v4 batch_update로 적용."""
    fmt_map = COLUMN_FORMATS.get(sheet_tab)
    if not fmt_map:
        return

    ws       = spreadsheet.worksheet(sheet_tab)
    sheet_id = ws.id

    type_def = {
        "TEXT":   {"type": "TEXT",   "pattern": "@"},
        "NUMBER": {"type": "NUMBER", "pattern": "#,##0"},
    }

    requests = []
    for col_idx, fmt_key in fmt_map.items():
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId":          sheet_id,
                    "startColumnIndex": col_idx,
                    "endColumnIndex":   col_idx + 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": type_def[fmt_key]
                    }
                },
                "fields": "userEnteredFormat.numberFormat",
            }
        })

    spreadsheet.batch_update({"requests": requests})
    print(f"  🎨 [{sheet_tab}] 열 서식 적용 완료")


# ── 구글 시트 업로드 ──────────────────────────────────────────────────────────
def upload_to_sheet(gc: gspread.Client, sheet_tab: str, df: pd.DataFrame, filepath: str):
    print(f"📤 구글 시트 업로드 → 탭: [{sheet_tab}]")
    sh = gc.open_by_key(SPREADSHEET_ID)

    try:
        ws = sh.worksheet(sheet_tab)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=sheet_tab, rows=20000, cols=50)
        print(f"  ➕ 탭 [{sheet_tab}] 새로 생성")

    ws.batch_clear(["A1:Z20000"])

    # 서식 먼저 적용 (데이터 입력 전에 열 서식을 지정해야 효과적)
    format_columns(sh, sheet_tab)

    # 헤더 + 데이터 직렬화
    header = df.columns.tolist()
    rows   = df.values.tolist()
    data   = [header] + [[str(v) for v in row] for row in rows]

    ws.update(range_name="A1", values=data, value_input_option="USER_ENTERED")

    # 임시 파일 삭제
    try:
        os.remove(filepath)
        print(f"  🗑️  임시 파일 삭제: {os.path.basename(filepath)}")
    except Exception:
        pass

# ── 작업 1: 판매현황 ──────────────────────────────────────────────────────────
def task_sales(driver: webdriver.Chrome, download_dir: str, gc: gspread.Client):
    print("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    print("📋 작업 1: 판매현황 수집 시작")

    steps = [
        ("/html/body/div[2]/div[3]/div/div[4]/ul/li[3]/div[1]/a",                                                                    "재고 메뉴"),
        ("/html/body/div[2]/div[3]/div/div[4]/ul/li[3]/div[2]/ul/li[7]/a",                                                            "출력물"),
        ("/html/body/div[2]/div[5]/div[3]/div/div[1]/div[2]/div[2]/div[2]/ul/li[2]/ul/li[1]/a",                                       "판매현황"),
        ("/html/body/div[2]/div[5]/div[3]/div/div[1]/div/div[1]/div[2]/ul/li[4]/a",                                                   "출하지시서체크"),
        ("/html/body/div[2]/div[5]/div[3]/div/div[1]/div/div[1]/div[2]/div[2]/div[1]/div[1]/div/div/button[2]",                       "검색(꺽쇠)"),
        ("//*[contains(text(), 'Excel')]",                                                                                             "Excel 다운로드"),
    ]

    for xpath, label in steps:
        click_xpath(driver, xpath, label)

    filepath = wait_for_download(download_dir)
    df       = read_excel_file(filepath)
    upload_to_sheet(gc, "이카운트 판매", df, filepath)

# ── 작업 2: 출하지시서현황 ────────────────────────────────────────────────────
def task_release_order(driver: webdriver.Chrome, download_dir: str, gc: gspread.Client):
    print("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    print("📦 작업 2: 출하지시서현황 수집 시작")

    steps = [
        ("/html/body/div[2]/div[3]/div/div[4]/ul/li[3]/div[1]/a",                                                                    "재고 메뉴"),
        ("/html/body/div[2]/div[3]/div/div[4]/ul/li[3]/div[2]/ul/li[7]/a",                                                            "출력물"),
        ("/html/body/div[2]/div[5]/div[3]/div/div[1]/div[2]/div[2]/div[2]/ul/li[2]/ul/li[4]/a",                                       "출하지시서현황"),
        ("/html/body/div[2]/div[5]/div[3]/div/div[1]/div/div[1]/div[2]/ul/li[3]/a[1]",                                                "출하지시서체크1"),
        ("/html/body/div[2]/div[5]/div[3]/div/div[1]/div/div[1]/div[2]/div[2]/div[1]/div[1]/div/div/button[2]",                       "검색(꺽쇠)"),
        ("//*[contains(text(), 'Excel')]",                                                                                             "Excel 다운로드"),
    ]

    for xpath, label in steps:
        click_xpath(driver, xpath, label)

    filepath = wait_for_download(download_dir)
    df       = read_excel_file(filepath)
    upload_to_sheet(gc, "이카운트 출하지시서", df, filepath)

# ── 메인 ─────────────────────────────────────────────────────────────────────
def main():
    print("🚀 release_order.py 시작")

    # 1. config 로드
    config = load_config()
    cred   = config["ecount_main"]

    # 2. 구글 인증
    print("🔑 구글 서비스 계정 인증...")
    gc_creds = Credentials.from_service_account_info(
        config["google"]["service_account"], scopes=SCOPES
    )
    gc = gspread.authorize(gc_creds)
    print("  ✅ 구글 인증 완료")

    # 3. 다운로드 폴더 준비
    script_dir   = os.path.dirname(os.path.abspath(__file__))
    download_dir = os.path.join(script_dir, "temp_downloads")
    os.makedirs(download_dir, exist_ok=True)

    # 기존 파일 전체 삭제
    for f in glob.glob(os.path.join(download_dir, "*")):
        try:
            os.remove(f)
        except Exception:
            pass
    print(f"🗂️  다운로드 폴더 초기화: {download_dir}")

    # 4. 브라우저 실행
    driver = build_driver(download_dir)

    try:
        # 5. 로그인
        ecount_login(driver, cred)

        # 6. 작업 1: 판매현황
        task1_ok = False
        try:
            task_sales(driver, download_dir, gc)
            task1_ok = True
        except Exception as e:
            print(f"\n❌ 작업 1 실패: {e}")

        # 7. 작업 2: 출하지시서현황 (작업 1 실패 여부와 무관하게 실행)
        try:
            task_release_order(driver, download_dir, gc)
        except Exception as e:
            print(f"\n❌ 작업 2 실패: {e}")

        if task1_ok:
            print("\n🎉 모든 작업 완료!")
        else:
            print("\n⚠️  작업 2는 완료, 작업 1은 실패. 로그를 확인하세요.")

    finally:
        driver.quit()
        print("🔒 브라우저 종료")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ 실행 중 오류가 발생했습니다: {e}")
    finally:
        pass # input("\n종료하려면 Enter를 누르세요...")