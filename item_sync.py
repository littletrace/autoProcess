"""
item_sync.py - 이카운트 품목/품목관계 자료내려받기 → Google Sheets 동기화
완전 독립 실행 스크립트 (PyInstaller 빌드 대응)

- credentials.json 미사용. config.json 의 google.service_account 로 인증.
- Selenium / EcountClient(웹 제어) 제거. 실제 다운로드는 엑셀 매크로(Win32/UIA)로 수행.
"""

# ==========================================
# 0. 표준 라이브러리
# ==========================================
import os
import sys
import time
import json
import ctypes

# 윈도우 cp949 인코딩 오류 방지
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except (AttributeError, ValueError):
    pass

# ==========================================
# 1. 외부 의존성
# ==========================================
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

import win32gui
import win32con
import win32clipboard
import win32com.client
import win32process
import win32api
import pyautogui


# ══════════════════════════════════════════════════════════════
#  경로 / 설정 로더
# ══════════════════════════════════════════════════════════════

def get_config_path():
    """
    config.json 경로 탐색: 자신의 폴더 우선, 없으면 한 단계 위 폴더.
    """
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


def load_config():
    path = get_config_path()
    print(f"📂 config 경로: {path}")
    if not os.path.exists(path):
        raise FileNotFoundError(f"설정 파일을 찾을 수 없습니다: {path}")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


# ══════════════════════════════════════════════════════════════
#  구글 스프레드시트
# ══════════════════════════════════════════════════════════════

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


class GoogleSheet:
    """
    서비스 계정 정보(dict)를 메모리 상에서 바로 인증.
    config["google"]["service_account"] 를 그대로 전달하면 된다.
    """

    def __init__(self, service_account_info: dict, spreadsheet_id: str):
        try:
            creds = Credentials.from_service_account_info(
                service_account_info, scopes=SCOPES
            )
            self._gc = gspread.authorize(creds)
            self._doc = self._gc.open_by_key(spreadsheet_id)
        except Exception as e:
            raise RuntimeError(f"구글 시트 인증/연결 실패: {e}") from e

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
                    raise ValueError(f"지원하지 않는 형식: {fmt}")
                requests.append({
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
                })
            if requests:
                self._doc.batch_update({"requests": requests})
        except Exception as e:
            print(f"❌ 열 표시 형식 지정 중 오류: {e}")

    def worksheet(self, name):
        return self._doc.worksheet(name)


# ══════════════════════════════════════════════════════════════
#  GUI 유틸리티 (Win32)
# ══════════════════════════════════════════════════════════════

def _get_window_text(hwnd):
    try:
        return win32gui.GetWindowText(hwnd) or ""
    except Exception:
        return ""


def _get_control_text(hwnd):
    try:
        buf_len = ctypes.windll.user32.SendMessageW(hwnd, win32con.WM_GETTEXTLENGTH, 0, 0)
        buf = ctypes.create_unicode_buffer(buf_len + 1)
        ctypes.windll.user32.SendMessageW(hwnd, win32con.WM_GETTEXT, buf_len + 1, buf)
        return buf.value
    except Exception:
        return ""


def set_edit_text(hwnd, text, delay=0.2):
    if not hwnd:
        return False
    try:
        ctypes.windll.user32.SendMessageW(hwnd, win32con.WM_SETTEXT, 0, "")
        time.sleep(delay)
        ctypes.windll.user32.SendMessageW(hwnd, win32con.WM_SETTEXT, 0, str(text))
        time.sleep(delay)
        return True
    except Exception:
        return False


def click_button(hwnd, delay=0.2):
    if not hwnd:
        return False
    try:
        parent = win32gui.GetParent(hwnd)
        if parent:
            try:
                win32gui.SetForegroundWindow(parent)
            except Exception:
                pass
        time.sleep(delay)
        ctypes.windll.user32.SendMessageW(hwnd, win32con.BM_CLICK, 0, 0)
        return True
    except Exception:
        return False


def click_button_by_text(window_keyword, button_text, timeout=15):
    for i in range(timeout):
        target_hwnd = None

        def find_window(hwnd, _):
            nonlocal target_hwnd
            if window_keyword not in _get_window_text(hwnd):
                return

            def find_btn(child_hwnd, _):
                nonlocal target_hwnd
                if _get_control_text(child_hwnd) == button_text:
                    target_hwnd = child_hwnd

            try:
                win32gui.EnumChildWindows(hwnd, find_btn, None)
            except Exception:
                pass

        try:
            win32gui.EnumWindows(find_window, None)
        except Exception:
            pass

        if target_hwnd:
            if click_button(target_hwnd):
                return True

        time.sleep(1)

    return False


def find_window_by_keyword(keyword, visible_only=False):
    found = None

    def enum_handler(hwnd, _):
        nonlocal found
        if found:
            return
        if keyword not in _get_window_text(hwnd):
            return
        if visible_only:
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return
            except Exception:
                return
        found = hwnd

    try:
        win32gui.EnumWindows(enum_handler, None)
    except Exception:
        pass
    return found


def ribbon_click(keys, label="", key_delay=0.5):
    if label:
        print(f"\n🖱️  리본: {label}")
    try:
        activate_excel_window()
        time.sleep(0.5)
        pyautogui.press('escape')
        time.sleep(0.3)
        pyautogui.press('escape')
        time.sleep(0.3)
        pyautogui.keyDown('alt')
        time.sleep(0.3)
        pyautogui.keyUp('alt')
        time.sleep(0.6)
        for key in keys:
            pyautogui.press(key.lower() if len(key) == 1 else key)
            time.sleep(key_delay)
        time.sleep(1)
        return True
    except Exception:
        return False


def activate_excel_window(restore=True, settle=1.0):
    target_hwnd = None

    def enum_handler(hwnd, _):
        nonlocal target_hwnd
        title = _get_window_text(hwnd)
        if "Excel" in title or "엑셀" in title:
            try:
                if win32gui.IsWindowVisible(hwnd):
                    target_hwnd = hwnd
            except Exception:
                pass

    try:
        win32gui.EnumWindows(enum_handler, None)
    except Exception:
        pass

    if not target_hwnd:
        return False

    success = False
    for attempt in range(5):
        try:
            if win32gui.IsIconic(target_hwnd) or restore:
                win32gui.ShowWindow(target_hwnd, win32con.SW_RESTORE)
            fore_hwnd = win32gui.GetForegroundWindow()
            fore_tid = win32process.GetWindowThreadProcessId(fore_hwnd)[0]
            target_tid = win32process.GetWindowThreadProcessId(target_hwnd)[0]
            attached = False
            if fore_tid != target_tid:
                try:
                    win32process.AttachThreadInput(target_tid, fore_tid, True)
                    attached = True
                except Exception:
                    pass
            try:
                win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
                win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)
                win32gui.SetForegroundWindow(target_hwnd)
                win32gui.BringWindowToTop(target_hwnd)
                win32gui.SetActiveWindow(target_hwnd)
            finally:
                if attached:
                    try:
                        win32process.AttachThreadInput(target_tid, fore_tid, False)
                    except Exception:
                        pass
        except Exception:
            pass
        try:
            if win32gui.GetForegroundWindow() == target_hwnd:
                success = True
                break
        except Exception:
            pass
        time.sleep(0.2)

    time.sleep(settle)
    return success


def close_and_save_excel(settle=2.0):
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        return False
    try:
        excel.DisplayAlerts = False
        for wb in excel.Workbooks:
            try:
                if wb.Path:
                    wb.Save()
            except Exception:
                pass
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        excel.Quit()
        time.sleep(settle)
        return True
    except Exception:
        return False


def close_excel(settle=1.0):
    if not activate_excel_window():
        return False
    try:
        pyautogui.hotkey('alt', 'F4')
        time.sleep(settle)
        pyautogui.hotkey('alt', 'n')
        time.sleep(0.5)
        try:
            pyautogui.hotkey('alt', 'n')
        except Exception:
            pass
        return True
    except Exception:
        return False


def open_new_excel(excel_exe, launch_wait=8, settle=2):
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.Workbooks.Add()
        time.sleep(settle)
        activate_excel_window()
        return True
    except Exception as e:
        print(f"Excel launch failed: {e}")
        return False


def clear_clipboard():
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        return True
    except Exception:
        return False
    finally:
        try:
            win32clipboard.CloseClipboard()
        except Exception:
            pass


def get_clipboard_text():
    try:
        win32clipboard.OpenClipboard()
        try:
            if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                return win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT) or ""
            return ""
        finally:
            win32clipboard.CloseClipboard()
    except Exception:
        return ""


def set_clipboard_text(text):
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, text)
        return True
    except Exception:
        return False
    finally:
        try:
            win32clipboard.CloseClipboard()
        except Exception:
            pass


# ══════════════════════════════════════════════════════════════
#  상수
# ══════════════════════════════════════════════════════════════

SPREADSHEET_ID = "1u6_lPc_snviUn0Yu3PeJiJV_spWoNeJp11AcxfbEF3c"

EXCEL_EXE = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"

# 자료 내려받기 팝업에서 찾을 버튼 텍스트
ITEM_BTN_TEXT = "품목 (자료올리기형태)"
ITEM_REL_BTN_TEXT = "품목관계 (자료올리기형태)"

# 데이터 로딩 대기 시간(초)
DATA_LOAD_WAIT = 8

# ──────────────────────────────────────────────────────────────
# 탭별 열 표시 형식 (0-based 열 인덱스 → 'TEXT' | 'NUMBER' | 'NUMBER_DECIMAL')
#   A=0, B=1, C=2, D=3, E=4, F=5, G=6, J=9, K=10, L=11
# ──────────────────────────────────────────────────────────────
TAB_FORMATS = {
    # 품목관계: A,B,D,E열 TEXT / F,G열 NUMBER
    "품목관계": {
        0: "TEXT", 1: "TEXT", 3: "TEXT", 4: "TEXT",
        5: "NUMBER_DECIMAL", 6: "NUMBER_DECIMAL",
    },
    # 품목(본사): A,B열 TEXT / K,L열 NUMBER
    "품목(본사)": {
        0: "TEXT", 1: "TEXT",
        10: "NUMBER_DECIMAL", 11: "NUMBER_DECIMAL",
    },
    # 품목(영업점): A,B열 TEXT / J,K열 NUMBER
    "품목(영업점)": {
        0: "TEXT", 1: "TEXT",
        9: "NUMBER_DECIMAL", 10: "NUMBER_DECIMAL",
    },
    # 품목관계(영업점): A,B,D,E열 TEXT / F,G열 NUMBER
    "품목관계(영업점)": {
        0: "TEXT", 1: "TEXT", 3: "TEXT", 4: "TEXT",
        5: "NUMBER_DECIMAL", 6: "NUMBER_DECIMAL",
    },
}


# ══════════════════════════════════════════════════════════════
#  이카운트 로그인 (Win32 GUI)
# ══════════════════════════════════════════════════════════════

def find_login_controls() -> dict | None:
    """
    이카운트 엑셀 로그인 팝업 컨트롤 탐색.
    대상 창: 타이틀에 'ECOUNT' 포함.
    구성: EDIT 3개(회사코드, 아이디, 비밀번호) + BUTTON(로그인).
    """
    login_hwnd = None

    def find_window(hwnd, _):
        nonlocal login_hwnd
        if "ECOUNT" in win32gui.GetWindowText(hwnd).upper():
            login_hwnd = hwnd

    win32gui.EnumWindows(find_window, None)
    if not login_hwnd:
        return None

    edit_list = []
    button_list = []

    def find_children(hwnd, _):
        cls = win32gui.GetClassName(hwnd)
        if "EDIT" in cls:
            rect = win32gui.GetWindowRect(hwnd)
            edit_list.append((rect[1], hwnd))
        if "BUTTON" in cls:
            text_len = ctypes.windll.user32.SendMessageW(
                hwnd, win32con.WM_GETTEXTLENGTH, 0, 0)
            button_list.append((text_len, hwnd))

    win32gui.EnumChildWindows(login_hwnd, find_children, None)

    edit_list.sort(key=lambda x: x[0])
    edit_hwnds = [h for _, h in edit_list]
    button_list.sort(key=lambda x: x[0])
    login_btn = button_list[0][1] if button_list else None

    return {
        "code": edit_hwnds[0],
        "user_id": edit_hwnds[1],
        "password": edit_hwnds[2],
        "login": login_btn,
    }


def ecount_auto_login(com_code: str, user_id: str, password: str) -> bool:
    """
    이카운트 로그인 자동화.
    10초 안에 로그인 창을 찾지 못하면 False 반환.
    (재시도 루프는 호출부인 open_excel_and_login() 에서 담당)
    """
    print("\n🔐 로그인 창 탐색 중... (최대 10초)")

    controls = None
    for i in range(10):
        controls = find_login_controls()
        if controls:
            break
        print(f"  대기 중... ({i + 1}초)")
        time.sleep(1)

    if not controls:
        print("  ❌ 10초 내 로그인 창 미발견")
        return False

    set_edit_text(controls["code"], com_code)
    print(f"  COM_CODE 입력: {com_code}")
    set_edit_text(controls["user_id"], user_id)
    print(f"  USER_ID  입력: {user_id}")
    set_edit_text(controls["password"], password)
    print("  PASSWORD 입력 완료")

    time.sleep(0.5)
    click_button(controls["login"])
    print("✅ 로그인 버튼 클릭")
    time.sleep(4)
    return True


def open_excel_and_login(com_code: str, user_id: str, password: str,
                         max_retry: int = 3) -> bool:
    """
    엑셀 새 통합 문서 열기 → 이카운트 로그인 리본 → 로그인 자동화.

    로그인 창 탐색 실패 시 '새 통합 문서 열기'부터 재시도.
    (원인: 엑셀 새 문서가 열리지 않아 이카운트 리본이 없는 경우)

    max_retry : 최대 재시도 횟수 (기본 3회)
    """
    for attempt in range(1, max_retry + 1):
        print(f"\n{'━'*55}")
        print(f"  🔄 엑셀 실행 + 로그인 시도 ({attempt}/{max_retry})")
        print(f"{'━'*55}")

        # ── Step 1. 엑셀 실행 → 새 통합 문서 ────────────────────
        if not open_new_excel(EXCEL_EXE):
            print(f"  ❌ 엑셀 실행 실패 (시도 {attempt})")
            if attempt < max_retry:
                print("  잠시 후 재시도합니다...")
                time.sleep(3)
            continue

        # ── Step 2. 이카운트 로그인 리본 클릭 ────────────────────
        activate_excel_window()
        time.sleep(3)
        open_ecount_login_ribbon()

        # ── Step 3. 로그인 창 탐색 + 자동 입력 ──────────────────
        success = ecount_auto_login(com_code, user_id, password)

        if success:
            print(f"  ✅ 로그인 성공 (시도 {attempt}회)")
            return True

        print("  ⚠️  로그인 창 탐색 실패 → '새 통합 문서 열기'부터 재시도합니다.")
        print("     (실패 원인: 이카운트 리본이 없는 시트 상태로 추정)")

    print(f"\n❌ {max_retry}회 재시도 후에도 로그인 실패. 프로세스를 중단합니다.")
    return False


# ══════════════════════════════════════════════════════════════
#  리본 메뉴 조작 (UIAutomation 기반)
# ══════════════════════════════════════════════════════════════

def _activate_ecount_tab_uia(hwnd):
    from pywinauto import Application
    app = Application(backend="uia").connect(handle=hwnd)
    excel_win = app.window(handle=hwnd)
    tab = excel_win.child_window(title="이카운트", control_type="TabItem")
    tab.click_input()
    return excel_win


def open_ecount_login_ribbon():
    """리본 메뉴 → 이카운트 탭 → '.*로그인.*' 버튼 클릭."""
    print("\n🖱️  리본: 이카운트 로그인 (UIA)")
    activate_excel_window()
    time.sleep(1)

    hwnd = find_window_by_keyword("Excel", visible_only=True)
    if not hwnd:
        hwnd = find_window_by_keyword("통합 문서", visible_only=True)

    if hwnd:
        try:
            excel_win = _activate_ecount_tab_uia(hwnd)
            time.sleep(0.5)
            btn = excel_win.child_window(title_re=".*로그인.*", control_type="Button")
            btn.click_input()
            time.sleep(1)
            return True
        except Exception as e:
            print(f"  [WARNING] UIA 클릭 실패, 매크로로 진입합니다: {e}")

    return ribbon_click(['Y', '2', 'Y', '4'], "이카운트 로그인")


def open_download_ribbon():
    """리본 메뉴 → 이카운트 탭 → '.*자료.*내려받기.*' 버튼 클릭."""
    print("\n🖱️  리본: 자료 내려받기 (UIA)")
    activate_excel_window()
    time.sleep(1)

    hwnd = find_window_by_keyword("Excel", visible_only=True)
    if not hwnd:
        hwnd = find_window_by_keyword("통합 문서", visible_only=True)

    if hwnd:
        try:
            excel_win = _activate_ecount_tab_uia(hwnd)
            time.sleep(0.5)
            btn = excel_win.child_window(title_re=".*자료.*내려받기.*", control_type="Button")
            btn.click_input()
            time.sleep(1)
            return True
        except Exception as e:
            print(f"  [WARNING] UIA 클릭 실패, 매크로로 진입합니다: {e}")

    return ribbon_click(['Y', '2', 'Y', '2'], "자료 내려받기")


# ══════════════════════════════════════════════════════════════
#  자료 내려받기 팝업 내 메뉴 클릭 (기초 → 항목)
# ══════════════════════════════════════════════════════════════

def click_download_item(item_btn_text: str) -> bool:
    """
    '자료' 팝업에서 '기초' 카테고리 버튼 클릭 → 이후 item_btn_text 버튼 클릭.
    (ITEM_BTN_TEXT 또는 ITEM_REL_BTN_TEXT)
    """
    if not click_button_by_text("자료", "기초"):
        print("❌ '기초' 버튼 미발견")
        return False
    time.sleep(1)

    if not click_button_by_text("자료", item_btn_text):
        print(f"❌ '{item_btn_text}' 버튼 미발견")
        return False

    return True


# ══════════════════════════════════════════════════════════════
#  엑셀 데이터 읽기 (전체 동적 범위)
# ══════════════════════════════════════════════════════════════

def read_excel_all_data() -> list[list[str]]:
    """
    현재 활성 엑셀 시트의 전체 데이터를 읽어 2D 리스트 반환.

    절차:
      1) A1 셀로 이동 (F5 → A1 → Enter)
      2) Ctrl+A 로 데이터 전체 선택
      3) Ctrl+C 로 복사
      4) 클립보드에서 읽어 2D 리스트로 변환
    """
    print("\n📊 엑셀 데이터 읽는 중...")
    activate_excel_window()

    pyautogui.press('F5')
    time.sleep(0.4)
    pyautogui.typewrite('A1', interval=0.05)
    pyautogui.press('enter')
    time.sleep(0.3)

    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.3)

    pyautogui.hotkey('ctrl', 'c')
    time.sleep(0.8)

    raw = get_clipboard_text()

    rows = []
    for line in raw.strip().split('\n'):
        cols = line.strip('\r').split('\t')
        if any(c.strip() for c in cols):
            rows.append(cols)

    print(f"  ✅ 읽기 완료: {len(rows)}행 × {len(rows[0]) if rows else 0}열")
    return rows


# ══════════════════════════════════════════════════════════════
#  스프레드시트 탭에 붙여넣기 (기존 내용 삭제 후 값만)
# ══════════════════════════════════════════════════════════════

def paste_to_sheet(gsheet: GoogleSheet, tab_name: str, data: list[list[str]]):
    """지정 탭의 내용을 전체 삭제하고 data 를 붙여넣기."""
    print(f"\n📋 [{tab_name}] 탭에 붙여넣는 중... ({len(data)}행)")

    try:
        ws = gsheet.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        print(f"  ❌ 탭 '{tab_name}' 없음 — 스프레드시트 탭 이름을 확인하세요.")
        raise

    ws.clear()
    time.sleep(1)

    if data:
        ws.update(data, "A1", value_input_option="USER_ENTERED")

    print(f"  ✅ [{tab_name}] 업로드 완료")


# ══════════════════════════════════════════════════════════════
#  자료 내려받기 완료 팝업 닫기
# ══════════════════════════════════════════════════════════════

def dismiss_download_complete_popup(timeout: int = 15) -> bool:
    """
    다운로드 완료 알림 팝업(#32770) 처리.
    '확인'/'OK'/'&확인' 버튼 클릭 또는 Enter 전송.
    """
    print("\n  🔔 다운로드 완료 팝업 대기 중...")

    for i in range(timeout):
        dialog_hwnd = None
        confirm_hwnd = None

        def find_dialog(hwnd, _):
            nonlocal dialog_hwnd, confirm_hwnd
            if not win32gui.IsWindowVisible(hwnd):
                return
            cls = win32gui.GetClassName(hwnd)
            title = win32gui.GetWindowText(hwnd)

            if cls == "#32770" or "대화" in title or title == "":
                def find_ok_btn(child, _):
                    nonlocal confirm_hwnd
                    buf_len = ctypes.windll.user32.SendMessageW(
                        child, win32con.WM_GETTEXTLENGTH, 0, 0)
                    buf = ctypes.create_unicode_buffer(buf_len + 1)
                    ctypes.windll.user32.SendMessageW(
                        child, win32con.WM_GETTEXT, buf_len + 1, buf)
                    if buf.value in ("확인", "OK", "&확인"):
                        nonlocal dialog_hwnd
                        dialog_hwnd = hwnd
                        confirm_hwnd = child

                try:
                    win32gui.EnumChildWindows(hwnd, find_ok_btn, None)
                except Exception:
                    pass

        win32gui.EnumWindows(find_dialog, None)

        # ── 1순위: '확인' 버튼 hwnd 직접 클릭 ─────────────────
        if confirm_hwnd:
            try:
                win32gui.SetForegroundWindow(dialog_hwnd)
                time.sleep(0.2)
                ctypes.windll.user32.SendMessageW(
                    confirm_hwnd, win32con.BM_CLICK, 0, 0)
                print("  ✅ 팝업 닫기 완료 (확인 버튼 클릭)")
                time.sleep(0.5)
                return True
            except Exception as e:
                print(f"  ⚠️  버튼 클릭 실패: {e}")

        # ── 2순위: 대화 상자 창에 Enter 전송 ──────────────────
        if dialog_hwnd:
            try:
                win32gui.SetForegroundWindow(dialog_hwnd)
                time.sleep(0.2)
                win32gui.PostMessage(dialog_hwnd, win32con.WM_KEYDOWN,
                                     win32con.VK_RETURN, 0)
                win32gui.PostMessage(dialog_hwnd, win32con.WM_KEYUP,
                                     win32con.VK_RETURN, 0)
                print("  ✅ 팝업 닫기 완료 (Enter 전송)")
                time.sleep(0.5)
                return True
            except Exception as e:
                print(f"  ⚠️  Enter 전송 실패: {e}")

        # ── 3순위: pyautogui Enter (포커스 있을 때만 유효) ─────
        if i >= 2:
            pyautogui.press('enter')
            time.sleep(0.5)
            still_open = False

            def recheck(hwnd, _):
                nonlocal still_open
                cls = win32gui.GetClassName(hwnd)
                if cls == "#32770" and win32gui.IsWindowVisible(hwnd):
                    def check_ok(child, _):
                        nonlocal still_open
                        buf_len = ctypes.windll.user32.SendMessageW(
                            child, win32con.WM_GETTEXTLENGTH, 0, 0)
                        buf = ctypes.create_unicode_buffer(buf_len + 1)
                        ctypes.windll.user32.SendMessageW(
                            child, win32con.WM_GETTEXT, buf_len + 1, buf)
                        if buf.value in ("확인", "OK", "&확인"):
                            still_open = True
                    try:
                        win32gui.EnumChildWindows(hwnd, check_ok, None)
                    except Exception:
                        pass

            win32gui.EnumWindows(recheck, None)

            if not still_open:
                print("  ✅ 팝업 닫기 완료 (pyautogui Enter)")
                return True

        print(f"  팝업 대기 중... ({i + 1}초)")
        time.sleep(1)

    print("  ⚠️  팝업을 찾지 못했습니다 (이미 닫혔거나 미발생). 계속 진행합니다.")
    return False


# ══════════════════════════════════════════════════════════════
#  블록 단위 실행 함수
# ══════════════════════════════════════════════════════════════

def download_and_paste(gsheet: GoogleSheet, item_btn_text: str, tab_name: str) -> bool:
    """
    공통 루틴:
      1) 엑셀 활성화
      2) 리본 → 자료 내려받기
      3) 기초 → item_btn_text 클릭
      4) 데이터 로딩 대기
      5) 엑셀 전체 선택 → 읽기
      6) 스프레드시트 tab_name 탭에 붙여넣기
      7) 탭별 열 표시 형식 적용
    """
    print(f"\n{'─'*55}")
    print(f"▶ 다운로드: {item_btn_text}  →  [{tab_name}]")
    print(f"{'─'*55}")

    activate_excel_window()
    time.sleep(1)

    open_download_ribbon()

    if not click_download_item(item_btn_text):
        return False

    print(f"  ⏳ 데이터 로딩 대기 ({DATA_LOAD_WAIT}초)...")
    time.sleep(DATA_LOAD_WAIT)

    dismiss_download_complete_popup()

    data = read_excel_all_data()
    if not data:
        print("  ❌ 읽어온 데이터 없음")
        return False

    paste_to_sheet(gsheet, tab_name, data)

    col_formats = TAB_FORMATS.get(tab_name)
    if col_formats:
        gsheet.format_columns(tab_name, col_formats)

    return True


# ══════════════════════════════════════════════════════════════
#  메인
# ══════════════════════════════════════════════════════════════

def main():
    # ── 초기 설정 ──────────────────────────────────────────────
    try:
        config = load_config()
    except Exception as e:
        print(f"설정 파일 로드 실패: {e}")
        sys.exit(1)

    # GoogleSheet 객체 생성 시점에 config 값을 명시적으로 주입
    gsheet = GoogleSheet(
        service_account_info=config["google"]["service_account"],
        spreadsheet_id=SPREADSHEET_ID,
    )

    # 본격적인 자동화 시작 전 기존 엑셀 프로세스 정리
    close_and_save_excel()

    # ══════════════════════════════════════════════════════════
    #  BLOCK 1 — ecount_main
    #    품목(본사), 품목관계
    # ══════════════════════════════════════════════════════════
    print("\n" + "═"*60)
    print("  BLOCK 1 : ecount_main")
    print("═"*60)

    acc_main = config.get("ecount_main", {})

    if not open_excel_and_login(
        acc_main["com_code"],
        acc_main["user_id"],
        acc_main["password"],
    ):
        raise SystemExit("❌ ecount_main 로그인 실패")

    if not download_and_paste(gsheet, ITEM_BTN_TEXT, "품목(본사)"):
        raise SystemExit("❌ 품목(본사) 다운로드 실패")

    if not download_and_paste(gsheet, ITEM_REL_BTN_TEXT, "품목관계"):
        raise SystemExit("❌ 품목관계 다운로드 실패")

    close_excel()

    print("\n✅ BLOCK 1 완료")
    time.sleep(2)

    # ══════════════════════════════════════════════════════════
    #  BLOCK 2 — ecount_sub
    #    품목(영업점), 품목관계(영업점)
    # ══════════════════════════════════════════════════════════
    print("\n" + "═"*60)
    print("  BLOCK 2 : ecount_sub")
    print("═"*60)

    acc_sub = config.get("ecount_sub", {})

    if not open_excel_and_login(
        acc_sub["com_code"],
        acc_sub["user_id"],
        acc_sub["password"],
    ):
        raise SystemExit("❌ ecount_sub 로그인 실패")

    if not download_and_paste(gsheet, ITEM_BTN_TEXT, "품목(영업점)"):
        raise SystemExit("❌ 품목(영업점) 다운로드 실패")

    if not download_and_paste(gsheet, ITEM_REL_BTN_TEXT, "품목관계(영업점)"):
        raise SystemExit("❌ 품목관계(영업점) 다운로드 실패")

    close_excel()

    print("\n" + "═"*60)
    print("🎉 모든 작업 완료!")
    print("═"*60)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ 실행 중 오류가 발생했습니다: {e}")
    finally:
        input("\n종료하려면 Enter를 누르세요...")