import os
import pandas as pd
import win32com.client as win32
from config import get_fixed_customer_path

# ===========================
# 1. 기본 설정
# ===========================
download_path = os.path.join(os.path.expanduser("~"), "Downloads")
FILE_PREFIX = "file_"

CUSTOMER_FILE = get_fixed_customer_path()
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
    excel = win32.DispatchEx("Excel.Application")
    try:
        excel.Visible = False
    except:
        pass  # Ignore if can't set Visible property
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

# Simple approach - keep first occurrence
df_new_unique = df_new.drop_duplicates()
df_new_idx = df_new_unique.set_index(KEY_COL)

asset_map = {}
ret_map = {}
status_map = {}
row_map = {}

for _, row in df_new_unique.iterrows():
    key = row[KEY_COL]
    asset_map[key] = row[ASSET_COL]
    ret_map[key] = row[RET_COL] 
    status_map[key] = row[STATUS_COL]
    row_map[key] = row.to_dict()

# ===========================
# 4. FOK_DATA 업데이트
# ===========================
excel = None
wb = None

try:
    excel = win32.DispatchEx("Excel.Application")
    try:
        excel.Visible = False
    except:
        pass
    
    xlUp = -4162
    xlToLeft = -4159
    
    updated_rows = 0
    cancelled_count = 0
    status_changed_count = 0
    cancelled_infos = []          # (계약번호, 이름)
    status_changed_infos = []     # (계약번호, 이름)
    new_infos = []   # (계약번호, 이름)
    
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

    if col_key is None:
        raise KeyError(f"'{KEY_COL}' 컬럼을 찾을 수 없습니다.")
    if col_asset is None:
        raise KeyError(f"'{ASSET_COL}' 컬럼을 찾을 수 없습니다.")
    if col_ret is None:
        raise KeyError(f"'{RET_COL}' 컬럼을 찾을 수 없습니다.")
    if col_status is None:
        raise KeyError(f"'{STATUS_COL}' 컬럼을 찾을 수 없습니다.")
    
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
    try:
        saved_path = wb.FullName
        print(f"📂 엑셀 실제 저장 위치: {saved_path}")
    except Exception as e:
        print("⚠ 저장 위치를 확인하지 못했습니다:", e)
finally:
    try:
        if wb is not None:
            wb.Close(False)
    except Exception as e:
        print(f"⚠ 워크북 닫기 오류: {e}")
    
    try:
        if excel is not None:
            excel.Quit()
    except Exception as e:
        print(f"⚠ Excel 종료 오류: {e}")