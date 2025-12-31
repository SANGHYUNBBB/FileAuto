import os
import pandas as pd
import win32com.client as win32

# ===========================
# 1. 기본 설정
# ===========================
download_path = os.path.join(os.path.expanduser("~"), "Downloads")
FILE_PREFIX = "file_"

def get_onedrive_path():
    for env in ("OneDriveCommercial", "OneDrive"):
        p = os.environ.get(env)
        if p and os.path.exists(p):
            return p
    raise EnvironmentError("OneDrive 경로를 찾을 수 없습니다.")

def find_customer_file():
    onedrive = get_onedrive_path()
    for root, _, files in os.walk(onedrive):
        if "고객data_v101.xlsx" in files:
            return os.path.join(root, "고객data_v101.xlsx")
    raise FileNotFoundError("고객data_v101.xlsx 파일을 찾을 수 없습니다.")

CUSTOMER_FILE = find_customer_file()
PASSWORD = "nilla17()"

KEY_COL = "계약번호"
ASSET_COL = "계좌자산"
RET_COL = "수익률"
STATUS_COL = "계약요청상태"
NAME_COL = "고객명"

# ===========================
# 2. xls -> xlsx 변환
# ===========================
def convert_xls_to_xlsx(xls_path: str) -> str:
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(xls_path)
        xlsx_path = os.path.splitext(xls_path)[0] + ".xlsx"
        wb.SaveAs(xlsx_path, FileFormat=51)
        wb.Close()
    finally:
        excel.Quit()
    return xlsx_path

def normalize_key(val) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s

# ===========================
# 3. 최신 증권사 파일 읽기
# ===========================
xls_files = [
    f for f in os.listdir(download_path)
    if f.startswith(FILE_PREFIX) and f.endswith(".xls")
]

xls_files.sort(
    key=lambda name: os.path.getmtime(os.path.join(download_path, name)),
    reverse=True,
)

latest_xlsx = convert_xls_to_xlsx(os.path.join(download_path, xls_files[0]))
df_new = pd.read_excel(latest_xlsx, dtype={KEY_COL: str})
# 계약번호: 텍스트 → 숫자 변환 (가능한 경우만)
def to_int_if_possible(x):
    if x is None:
        return x
    s = str(x).strip()
    if s.isdigit():
        return int(s)   # 텍스트 숫자 → int
    return x           # 숫자 아닌 건 그대로

df_new[KEY_COL] = df_new[KEY_COL].apply(to_int_if_possible)

df_new.columns = df_new.columns.map(lambda x: str(x).replace(" ", ""))

for col in [KEY_COL, ASSET_COL, RET_COL, STATUS_COL]:
    if col not in df_new.columns:
        raise KeyError(f"증권사 파일에 '{col}' 컬럼이 없습니다.")

df_new[KEY_COL] = df_new[KEY_COL].map(normalize_key)
df_new = df_new[df_new[KEY_COL] != ""]

df_new = df_new[~df_new.duplicated(subset=[KEY_COL], keep="last")]
df_new_idx = df_new.set_index(KEY_COL)

asset_map = df_new_idx[ASSET_COL].to_dict()
ret_map = df_new_idx[RET_COL].to_dict()
status_map = df_new_idx[STATUS_COL].to_dict()
row_map = df_new_idx.to_dict("index")

# ===========================
# 4. FOK_DATA 업데이트
# ===========================
excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False

xlUp = -4162
xlToLeft = -4159

updated_rows = 0
cancelled_count = 0
status_changed_count = 0
cancelled_infos = []          # (계약번호, 이름)
status_changed_infos = []     # (계약번호, 이름)
new_infos = []   # (계약번호, 이름)
try:
    wb = excel.Workbooks.Open(CUSTOMER_FILE, False, False, None, PASSWORD)
    ws = wb.Worksheets("FOK_DATA")

    last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    last_col = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    header_names = [None] * last_col
    col_key = col_asset = col_ret = col_status = None

    for c in range(1, last_col + 1):
        h = ws.Cells(1, c).Value
        if h:
            h = str(h).replace(" ", "")
            header_names[c - 1] = h
            if h == KEY_COL:
                col_key = c
            elif h == ASSET_COL:
                col_asset = c
            elif h == RET_COL:
                col_ret = c
            elif h == STATUS_COL:
                col_status = c

    idx_key = col_key - 1
    idx_asset = col_asset - 1
    idx_ret = col_ret - 1
    idx_status = col_status - 1

    data = ws.Range(ws.Cells(2, 1), ws.Cells(last_row, last_col)).Value
    data_list = [list(r) for r in data] if data else []

    existing_rows = []
    existing_keys = set()

    for row in data_list:
        key = normalize_key(row[idx_key])
        if not key:
            continue

        if key in asset_map:
            row[idx_asset] = asset_map[key]
            row[idx_ret] = ret_map[key]

            # 🔴 계약요청상태 변경
            if (
                status_map.get(key) == "계약해지"
                and row[idx_status] == "계약완료(승인)"
            ):
                row[idx_status] = "계약해지"
                status_changed_count += 1
                name = row_map[key].get(NAME_COL, "")
                status_changed_infos.append((key, name))

            updated_rows += 1
            existing_rows.append(row)
            existing_keys.add(key)
        else:
            cancelled_count += 1
            name = row_map.get(key, {}).get(NAME_COL, "")
            cancelled_infos.append((key, name))

    new_rows = []
    for k in row_map.keys():
        if k in existing_keys:
            continue

        row_dict = row_map[k]
        row = [None] * last_col

        name = row_dict.get(NAME_COL, "")
        new_infos.append((k, name))

        for i, h in enumerate(header_names):
            if h and h in row_dict:
                row[i] = row_dict[h]

        row[idx_key] = k
        row[idx_asset] = asset_map.get(k)
        row[idx_ret] = ret_map.get(k)
        row[idx_status] = status_map.get(k)

        new_rows.append(row)

    final_rows = existing_rows + new_rows

    ws.Range(ws.Cells(2, 1), ws.Cells(last_row, last_col)).ClearContents()
    ws.Range(ws.Cells(2, 1), ws.Cells(1 + len(final_rows), last_col)).Value = tuple(
        tuple(r) for r in final_rows
    )

    print(f"✅ 기존 고객 업데이트: {updated_rows}")
    print(f"🔁 계약완료 → 계약해지 변경: {status_changed_count}")

    # ❌ 해지(삭제)
    if cancelled_count > 0:
        print(f"❌ 삭제(해지): {cancelled_count}")
        print("=== ❌ 해지된 고객 ===")
        for k, name in cancelled_infos:
            print(f" - {k} / {name}")

    # ➕ 신규
    if len(new_infos) > 0:
        print(f"➕ 신규 추가: {len(new_infos)}")
        print("=== ➕ 신규 고객 ===")
        for k, name in new_infos:
            print(f" - {k} / {name}")

    # 🔁 상태 변경
    if status_changed_count > 0:
        print("=== 🔁 계약완료 → 계약해지 변경 ===")
        for k, name in status_changed_infos:
            print(f" - {k} / {name}")
    

    wb.Save()

finally:
    wb.Close(False)
    excel.Quit()