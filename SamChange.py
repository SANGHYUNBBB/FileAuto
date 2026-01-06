import os
import pandas as pd
import win32com.client as win32
import time
import gc
import pywintypes

# ===========================
# 1) 설정
# ===========================
DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
SRC_PREFIX = "통합 문서1"

SHEET_DST = "삼성_DATA"
DST_START_ROW = 6
DST_START_COL = 2
PASTE_COLS = 23
CONTRACT_REL_IDX = 3   # B기준 E열

PASSWORD = "nilla17()"

# ===========================
# 2) 유틸
# ===========================
def excel_date_to_str(x):
    """
    엑셀 날짜(serial) / 문자열 날짜 모두 처리
    """
    if pd.isna(x) or x == "":
        return ""
    try:
        # 엑셀 serial number
        if isinstance(x, (int, float)):
            return pd.to_datetime(x, unit="D", origin="1899-12-30").strftime("%Y/%m/%d")
        # 문자열 날짜
        return pd.to_datetime(x).strftime("%Y/%m/%d")
    except Exception:
        return ""
def com_call_with_retry(fn, tries=30, delay=0.5):
    for _ in range(tries):
        try:
            return fn()
        except pywintypes.com_error:
            time.sleep(delay)
    raise

def get_onedrive_path():
    for env in ("OneDriveCommercial", "OneDrive"):
        p = os.environ.get(env)
        if p and os.path.exists(p):
            return p
    raise EnvironmentError("OneDrive 경로 없음")

def find_customer_file():
    base = get_onedrive_path()
    for root, _, files in os.walk(base):
        if "고객data_v101.xlsx" in files:
            return os.path.join(root, "고객data_v101.xlsx")
    raise FileNotFoundError("고객data_v101.xlsx 없음")

CUSTOMER_FILE = find_customer_file()

def find_latest_source_file():
    files = [
        f for f in os.listdir(DOWNLOAD_DIR)
        if f.startswith(SRC_PREFIX) and f.lower().endswith((".xls", ".xlsx"))
    ]
    if not files:
        raise FileNotFoundError("증권사 파일 없음")
    files.sort(key=lambda f: os.path.getmtime(os.path.join(DOWNLOAD_DIR, f)), reverse=True)
    path = os.path.join(DOWNLOAD_DIR, files[0])
    print(f"📂 최신 증권사 파일: {path}")
    return path

def convert_xls_to_xlsx(path):
    if path.lower().endswith(".xlsx"):
        return path
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(path)
    new_path = path.replace(".xls", ".xlsx")
    wb.SaveAs(new_path, FileFormat=51)
    wb.Close(False)
    excel.Quit()
    return new_path

# ===========================
# 3) 증권사 파일 읽기
# ===========================
def read_and_sort_source(src_path):
    src_xlsx = convert_xls_to_xlsx(src_path)

    df = pd.read_excel(src_xlsx)
    df_bx = df.iloc[:, 1:1 + PASTE_COLS].copy()

    # 계약번호: PLVA로 시작하는 행만
    df_bx["__contract__"] = df_bx.iloc[:, CONTRACT_REL_IDX].astype(str).str.strip()
    df_bx = df_bx[df_bx["__contract__"].str.startswith("PLVA")]

    # 날짜 컬럼 처리
    DATE_COLS = {"최초계약일", "연장계약일", "만료일"}
    for col in df_bx.columns:
        if str(col).strip() in DATE_COLS:
            df_bx[col] = df_bx[col].apply(excel_date_to_str)

    df_bx = df_bx.sort_values("__contract__")

    values_df = df_bx.drop(columns="__contract__").fillna("").astype(str)
    values = values_df.values.tolist()

    # 🔥 무조건 텍스트 처리할 컬럼
    TARGET_COLS = {"계좌번호", "수수료출금계좌"}
    target_idx = [
        i for i, c in enumerate(values_df.columns)
        if str(c).strip() in TARGET_COLS
    ]

    for row in values:
        for i in target_idx:
            s = row[i].strip()
            if not s:
                continue
            if "E+" in s or "e+" in s:
                s = format(int(float(s)), "d")
            if s.endswith(".0"):
                s = s[:-2]
            row[i] = "'" + s   # ✅ 무조건 텍스트

    contracts = df_bx["__contract__"].tolist()
    print(f"✅ 유효 계약 수: {len(contracts)}")

    return values, contracts

# ===========================
# 4) 기존 비고 + 계약 목록
# ===========================
def build_remark_map(ws):
    xlUp = -4162
    last_row = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row

    name_map = {} 
    remark_map = {}
    old_contracts = []

    if last_row < DST_START_ROW:
        return remark_map, name_map,old_contracts

    rng = ws.Range(ws.Cells(DST_START_ROW, 1), ws.Cells(last_row, 6)).Value
    for r in rng:
        contract = "" if r[4] is None else str(r[4]).strip()
        name = "" if r[5] is None else str(r[5]).strip()   # 🔹 C열 = 이름 (필요시 수정)
        if contract.startswith("PLVA"):
            remark_map[contract] = r[0] or ""
            name_map[contract] = name
            old_contracts.append(contract)

    print(f"📝 기존 계약 수: {len(old_contracts)}")
    return remark_map,name_map, old_contracts

# ===========================
# 5) parkpark 쓰기
# ===========================
def write_to_parkpark(rows, contracts):
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False

    wb = excel.Workbooks.Open(CUSTOMER_FILE, False, False, None, PASSWORD)
    ws = wb.Worksheets(SHEET_DST)

    remark_map,  name_map ,old_contracts = build_remark_map(ws)

    new_set = set(contracts)
    old_set = set(old_contracts)

    added = sorted(new_set - old_set)
    removed = sorted(old_set - new_set)

    print("🔍 변경 내역")
    print(f"   ➕ 신규 추가: {len(added)}건")
    print(f"   ➖ 삭제/해지: {len(removed)}건")
    # ✅ 여기 추가
    if removed:
        print("🚫 해지된 계약 목록")
        for c in removed:
            print(f"   - {name_map.get(c, '이름없음')} / {c}")
    # 5행 헤더 유지, 데이터만 삭제
    last_used = ws.UsedRange.Row + ws.UsedRange.Rows.Count
    ws.Range(
        ws.Cells(DST_START_ROW, 1),
        ws.Cells(last_used, 24)
    ).ClearContents()

    # 비고
    remarks = [remark_map.get(c, "") for c in contracts]
    ws.Range(
        ws.Cells(DST_START_ROW, 1),
        ws.Cells(DST_START_ROW + len(rows) - 1, 1)
    ).Value = tuple((v,) for v in remarks)

    # 본 데이터
    ws.Range(
        ws.Cells(DST_START_ROW, DST_START_COL),
        ws.Cells(DST_START_ROW + len(rows) - 1, DST_START_COL + PASTE_COLS - 1)
    ).Value = tuple(tuple(r) for r in rows)

    print("💾 저장 중...")
    wb.Save()
    wb.Close(False)
    excel.Quit()
    gc.collect()
    print("📁 완료")

# ===========================
# 6) main
# ===========================
def main():
    src = find_latest_source_file()
    rows, contracts = read_and_sort_source(src)
    write_to_parkpark(rows, contracts)

if __name__ == "__main__":
    main()
