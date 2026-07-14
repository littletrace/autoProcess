# sell.py
# 이카운트 판매현황 엑셀 다운로드 → 구글 스프레드시트 업로드 자동화
# 단일 파일 실행 / PyInstaller .exe 변환 가능

import os
import sys
import json
import time
import glob
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except (AttributeError, ValueError):
    pass

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from google.oauth2.service_account import Credentials
import gspread

# ──────────────────────────────────────────────
# 상수
# ──────────────────────────────────────────────
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

LOGIN_URL = "https://login.ecount.com/LOGIN?lan_type=ko-KR"

# 계정별 작업 정의
# col_formats: {0-based 열 인덱스: "TEXT" | "NUMBER" | "NUMBER_DECIMAL"}
# A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7, I=8, J=9
ACCOUNT_TASKS = [
    {
        "config_key": "ecount_sub",
        "sheet_tab": "670989[금월]",
        "sell_tab_xpath": "//*[contains(text(), '판매취합')]",
        "col_formats": {
            6: "TEXT",    # G열
            7: "NUMBER",  # H열
            8: "NUMBER",  # I열
        },
    },
    {
        "config_key": "ecount_main",
        "sheet_tab": "148713[금월]",
        "sell_tab_xpath": "//*[contains(text(), '판매취합')]",
        "col_formats": {
            6: "TEXT",    # G열
            7: "NUMBER",  # H열
            9: "NUMBER",  # J열
        },
    },
]
SPREADSHEET_ID = '1BvdpZAUq9jBudJ-PtBilKvltrLXrhEIorOr-hhPynTM'

# ──────────────────────────────────────────────
# 설정 로드
# ──────────────────────────────────────────────
def get_config_path():
    if getattr(sys, "frozen", False):
        current = os.path.dirname(sys.executable)
    else:
        current = os.path.dirname(os.path.abspath(__file__))

    # 1순위: 같은 폴더
    candidate = os.path.join(current, "config.json")
    if os.path.exists(candidate):
        return candidate

    # 2순위: 상위 폴더
    candidate = os.path.join(os.path.dirname(current), "config.json")
    if os.path.exists(candidate):
        return candidate

    raise FileNotFoundError(f"config.json을 찾을 수 없습니다. 탐색 위치: {current} 및 상위 폴더")


def load_config():
    path = get_config_path()
    print(f"📂 config.json 경로: {path}")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


# ──────────────────────────────────────────────
# 다운로드 폴더 관리
# ──────────────────────────────────────────────
def get_download_dir():
    script_dir = (
        os.path.dirname(sys.executable)
        if getattr(sys, "frozen", False)
        else os.path.dirname(os.path.abspath(__file__))
    )
    dl_dir = os.path.join(script_dir, "temp_downloads")
    os.makedirs(dl_dir, exist_ok=True)
    return dl_dir


def clear_download_dir(dl_dir):
    for f in glob.glob(os.path.join(dl_dir, "*")):
        try:
            os.remove(f)
        except Exception:
            pass
    print(f"🗑️  temp_downloads 초기화 완료: {dl_dir}")


# ──────────────────────────────────────────────
# 크롬 드라이버 생성
# ──────────────────────────────────────────────
def build_driver(dl_dir):
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--start-maximized")
    prefs = {
        "download.default_directory": dl_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    options.add_experimental_option("prefs", prefs)
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)


# ──────────────────────────────────────────────
# 이카운트 로그인
# ──────────────────────────────────────────────
def login_ecount(driver, com_code, user_id, password):
    print("🔐 이카운트 로그인 시도...")
    driver.get(LOGIN_URL)
    wait = WebDriverWait(driver, 15)

    wait.until(EC.presence_of_element_located((By.ID, "com_code"))).send_keys(com_code)
    driver.find_element(By.ID, "id").send_keys(user_id)
    pw_field = driver.find_element(By.ID, "passwd")
    pw_field.send_keys(password)
    pw_field.send_keys(Keys.RETURN)

    time.sleep(4)
    print("✅ 로그인 완료")

    try:
        popup_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//div[contains(@class, 'control-set')]"
                "//*[@id='toolbar_sid_toolbar_item_regist']/button",
            ))
        )
        popup_btn.click()
        print("🔔 새 기기 등록 팝업 닫음")
        time.sleep(1)
    except Exception:
        print("ℹ️  새 기기 등록 팝업 없음")


# ──────────────────────────────────────────────
# 메뉴 진입
# ──────────────────────────────────────────────
def navigate_to_sales(driver, sell_tab_xpath):
    wait = WebDriverWait(driver, 15)

    print("📌 재고 1 메뉴 진입...")
    wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@id='link_depth1_MENUTREE_000004']"))).click()
    time.sleep(1)

    print("📌 판매현황 메뉴 진입...")
    wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@id='link_depth4_MENUTREE_000494']"))).click()
    time.sleep(5)

    print("📌 판매취합 버튼 클릭...")
    wait.until(EC.element_to_be_clickable((By.XPATH, sell_tab_xpath))).click()
    time.sleep(1)

    print("🔍 검색 버튼 클릭...")
    search_btn_xpath = (
        "/html/body/div[2]/div[5]/div[3]/div/div[1]/div/div[1]"
        "/div[2]/div[2]/div[1]/div[1]/div/div/button[2]"
    )
    wait.until(EC.element_to_be_clickable((By.XPATH, search_btn_xpath))).click()
    time.sleep(1)


# ──────────────────────────────────────────────
# 엑셀 다운로드
# ──────────────────────────────────────────────
def wait_for_download(dl_dir, timeout=60):
    print("⬇️  엑셀 다운로드 대기 중...")
    end_time = time.time() + timeout
    while time.time() < end_time:
        files = [
            f for f in glob.glob(os.path.join(dl_dir, "*"))
            if not f.endswith(".crdownload") and not f.endswith(".tmp")
        ]
        if files:
            latest = max(files, key=os.path.getctime)
            print(f"✅ 다운로드 완료: {os.path.basename(latest)}")
            return latest
        time.sleep(3)
    raise TimeoutError(f"엑셀 다운로드 {timeout}초 초과")


def download_excel(driver, dl_dir):
    wait = WebDriverWait(driver, 15)
    print("📥 Excel 버튼 클릭...")
    wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Excel')]"))).click()
    time.sleep(2)
    return wait_for_download(dl_dir)


# ──────────────────────────────────────────────
# 엑셀 파일 읽기
# ──────────────────────────────────────────────
def read_excel_file(file_path):
    print(f"📖 파일 읽기: {os.path.basename(file_path)}")
    try:
        df = pd.read_excel(file_path)
        print("   → openpyxl 방식 성공")
    except Exception:
        print("   → HTML 형식 엑셀로 재시도...")
        df = pd.read_html(file_path, encoding="utf-8")[0]
    return df.fillna("")


# ──────────────────────────────────────────────
# 열 서식 지정 (Sheets API batchUpdate)
# ──────────────────────────────────────────────
def apply_column_formats(spreadsheet, worksheet, col_formats):
    """
    col_formats: {0-based 열 인덱스: "TEXT" | "NUMBER" | "NUMBER_DECIMAL"}
    TEXT           → numberFormat type=TEXT,   pattern=@
    NUMBER         → numberFormat type=NUMBER, pattern=#,##0
    NUMBER_DECIMAL → numberFormat type=NUMBER, pattern=#,##0.00
    나머지 열은 건드리지 않음
    """
    if not col_formats:
        return

    sheet_id = worksheet.id
    type_map = {
        "TEXT":   {"type": "TEXT",   "pattern": "@"},
        "NUMBER":         {"type": "NUMBER", "pattern": "#,##0"},
        "NUMBER_DECIMAL": {"type": "NUMBER", "pattern": "#,##0.00"},
    }

    requests = []
    for col_idx, fmt_key in col_formats.items():
        fmt = type_map.get(str(fmt_key).upper())
        if not fmt:
            print(f"   ⚠️  알 수 없는 서식 키: {fmt_key}, 건너뜀")
            continue
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startColumnIndex": col_idx,
                    "endColumnIndex": col_idx + 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": fmt
                    }
                },
                "fields": "userEnteredFormat.numberFormat",
            }
        })

    if requests:
        spreadsheet.batch_update({"requests": requests})
        labels = {idx: fmt for idx, fmt in col_formats.items()}
        print(f"   → 열 서식 적용 완료: {labels}")


# ──────────────────────────────────────────────
# 구글 시트 업로드
# ──────────────────────────────────────────────
def upload_to_sheets(config, df, sheet_tab, col_formats):
    print(f"☁️  구글 시트 업로드 → 탭: {sheet_tab}")
    creds = Credentials.from_service_account_info(
        config["google"]["service_account"], scopes=SCOPES
    )
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = sh.worksheet(sheet_tab)

    # 클리어
    ws.batch_clear(["A1:Z20000"])
    print("   → 시트 클리어 완료")

    # 헤더 + 데이터 구성
    headers = [str(c) for c in df.columns.tolist()]
    rows = [[str(v) if v != "" else "" for v in row] for row in df.values.tolist()]
    data = [headers] + rows

    # 데이터 업로드
    ws.update(range_name="A1", values=data, value_input_option="USER_ENTERED")
    print(f"   → {len(rows)}행 업로드 완료")

    # 열 서식 적용
    apply_column_formats(sh, ws, col_formats)
""" 
   # 업데이트 일시 기록
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.update(
        range_name=f"A{len(data) + 1}",
        values=[[f"업데이트 일시: {timestamp}"]],
        value_input_option="USER_ENTERED",
    )
    print(f"   → 업데이트 일시 기록: {timestamp}")
"""


# ──────────────────────────────────────────────
# 계정 1개 전체 처리
# ──────────────────────────────────────────────
def process_account(config, task):
    config_key   = task["config_key"]
    sheet_tab    = task["sheet_tab"]
    sell_xpath   = task["sell_tab_xpath"]
    col_formats  = task.get("col_formats", {})

    account  = config[config_key]
    dl_dir   = get_download_dir()
    clear_download_dir(dl_dir)

    driver    = build_driver(dl_dir)
    file_path = None

    try:
        login_ecount(driver, account["com_code"], account["user_id"], account["password"])
        navigate_to_sales(driver, sell_xpath)
        file_path = download_excel(driver, dl_dir)
        df = read_excel_file(file_path)
        upload_to_sheets(config, df, sheet_tab, col_formats)
        print(f"🎉 [{config_key}] 처리 완료!\n")
    except Exception as e:
        print(f"❌ [{config_key}] 오류 발생: {e}\n")
    finally:
        driver.quit()
        print(f"🔒 브라우저 종료 [{config_key}]")
        if file_path and os.path.exists(file_path):
            try:
                os.remove(file_path)
                print(f"🗑️  임시 파일 삭제: {os.path.basename(file_path)}")
            except Exception:
                pass


# ──────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────
def main():
    print("=" * 50)
    print("📊 이카운트 판매현황 업로드 시작")
    print("=" * 50)

    config = load_config()

    for task in ACCOUNT_TASKS:
        print(f"\n▶ 계정 처리 시작: {task['config_key']} → {task['sheet_tab']}")
        process_account(config, task)

    print("=" * 50)
    print("✅ 모든 계정 처리 완료")
    print("=" * 50)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ 실행 중 오류가 발생했습니다: {e}")
    finally:
        input("\n종료하려면 Enter를 누르세요...")