import os
import re
import pandas as pd
import win32com.client as win32
import gc
import time
import pywintypes
import sys

# ===========================
# 1) 설정
# ===========================
DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
SRC_PREFIX = "통합 문서1"

DST_START_ROW = 6
DST_START_COL = 2
PASTE_COLS = 23
CONTRACT_REL_IDX = 3  # B~X 기준 E열

PASSWORD = "nilla17()"
SHEET_DST = "삼성_DATA"


# ===========================
# 2) 환경 정보 출력
# ===========================
print("====== DEBUG: ENV INFO ======")
print("Python:", sys.version)
print("pandas:", pd.__version__)
print("Downloads:", DOWNLOAD_DIR)
print("=============================")


# ===========================
# 3) OneDrive / 고객파일
# ===========================
def get_onedrive_path():
    for env in ("OneDriveCommercial", "OneDrive"):
        p = os.environ.get(env)
        if p and os.path.exists(p):
            print("DEBUG-ONEDRIVE:", p)
            return p
    raise EnvironmentError("OneDrive 경로 없음")

def find_customer_file():
    onedrive = get_onedrive_path()
    for root, _, files in os.walk(onedrive):
        if "고객data_v101.xlsx" in files:
            path = os.path.join(root, "고객data_v101.xlsx")
            print("DEBUG-CUSTOMER-FILE:", path)
            return path
    raise FileNotFoundError("고객data_v101.xlsx 없음")

CUSTOMER_FILE = find_customer_file()


# ===========================
# 4) 유틸
# ===========================
def com_call_with_retry(fn, tries=20, delay=0.5):
    for _ in range(tries):
        try:
            return fn()
        except pywintypes.com_error:
            time.sleep(delay)
    raise

def clean_contract(x):
    if x is None:
        return ""
    s = str(x)
    s = s.replace("\u00a0", "")
    s = s.replace("\t", "")
    s = s.replace("\r", "").replace("\n", "")
    return s.strip()


def find_latest_source_file():
    files = [
        f for f in os.listdir(DOWNLOAD_DIR)
        if f.startswith(SRC_PREFIX) and f.lower().endswith((".xls", ".xlsx"))
    ]
    files.sort(key=lambda f: os.path.getmtime(os.path.join(DOWNLOAD_DIR, f)), reverse=True)
    path = os.path.join(DOWNLOAD_DIR, files[0])
    print("DEBUG-SOURCE-FILE:", path)
    return path


# ===========================
# 5) 삼성증권 파일 읽기 (디버깅 핵심)
# ===========================
def read_and_debug_source(src_path):
    df = pd.read_excel(src_path, dtype=str)

    print("\n====== DEBUG: CONTRACT RAW VALUES (상위 20개) ======")
    for v in df.iloc[:20, CONTRACT_REL_IDX + 1]:
        print("DEBUG-CONTRACT-RAW:", v)

    df_bx = df.iloc[:, 1:1 + PASTE_COLS].copy()
    df_bx["__contract_raw__"] = df_bx.iloc[:, CONTRACT_REL_IDX]
    df_bx["__contract__"] = df_bx["__contract_raw__"].map(clean_contract)

    print("\n====== DEBUG: CONTRACT REPR (마지막 10개) ======")
    for v in df_bx["__contract__"].tail(10):
        print("DEBUG-CONTRACT-REPR:", repr(v), "LEN:", len(v))

    before = len(df_bx)
    df_bx = df_bx[df_bx["__contract__"] != ""]
    after = len(df_bx)

    print(f"\nDEBUG-FILTER: before={before}, after={after}")

    df_bx = df_bx.sort_values("__contract__")

    values_df = df_bx.drop(columns=["__contract__", "__contract_raw__"]).fillna("").astype(str)
    values_list = values_df.values.tolist()
    contracts = df_bx["__contract__"].tolist()

    print("\n====== DEBUG: FINAL CONTRACT LIST (마지막 5개) ======")
    for c in contracts[-5:]:
        print("DEBUG-FINAL-CONTRACT:", repr(c))

    return values_list, contracts


# ===========================
# 6) parkpark 쓰기 (비교 디버그)
# ===========================
def write_debug(sorted_rows, sorted_contracts):
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    wb = None

    try:
        wb = excel.Workbooks.Open(CUSTOMER_FILE, False, False, None, PASSWORD)
        ws = wb.Worksheets(SHEET_DST)

        xlUp = -4162
        last_row = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row

        old_contracts = []
        rng = ws.Range(ws.Cells(DST_START_ROW, 5), ws.Cells(last_row, 5)).Value

        for r in rng:
            old_contracts.append(clean_contract(r[0]))

        print("\n====== DEBUG: NEW CONTRACT CHECK ======")
        new_set = set(sorted_contracts)
        old_set = set(old_contracts)

        print("NEW ONLY:", sorted(new_set - old_set))
        print("OLD ONLY:", sorted(old_set - new_set))

    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        try:
            excel.Quit()
        except Exception:
            pass
        del excel
        gc.collect()


# ===========================
# 7) main
# ===========================
def main():
    src = find_latest_source_file()
    rows, contracts = read_and_debug_source(src)
    write_debug(rows, contracts)

if __name__ == "__main__":
    main()