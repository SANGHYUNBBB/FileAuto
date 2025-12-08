import os
import pandas as pd
import win32com.client as win32

# ===========================
# 1. 기본 설정
# ===========================
download_path = r"C:\Users\pc\Downloads"
FILE_PREFIX = "file_"   # 증권사 파일 접두사

CUSTOMER_FILE = r"C:\Users\pc\OneDrive - 주식회사 플레인바닐라\LEEJAEWOOK의 파일 - 플레인바닐라 업무\Customer\고객data\고객data_v101_parkpark.xlsx"
PASSWORD = "nilla17()"

KEY_COL = "계약번호"
ASSET_COL = "계좌자산"
RET_COL = "수익률"


# ===========================
# 2. xls -> xlsx 변환
# ===========================
def convert_xls_to_xlsx(xls_path: str) -> str:
    if not os.path.exists(xls_path):
        raise FileNotFoundError(f"xls 파일을 찾을 수 없습니다: {xls_path}")

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(xls_path)
        xlsx_path = os.path.splitext(xls_path)[0] + ".xlsx"
        wb.SaveAs(xlsx_path, FileFormat=51)  # xlsx
        wb.Close()
    finally:
        excel.Quit()

    print(f"[변환 완료] {xls_path} -> {xlsx_path}")
    return xlsx_path


def normalize_key(val) -> str:
    """계약번호를 문자열로 통일 (.0 제거, 공백 제거)"""
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

if not xls_files:
    raise FileNotFoundError(f"{download_path}에 '{FILE_PREFIX}*.xls' 파일이 없습니다.")

xls_files.sort(
    key=lambda name: os.path.getmtime(os.path.join(download_path, name)),
    reverse=True,
)
latest_xls = os.path.join(download_path, xls_files[0])
print("📂 가장 최근 다운로드 xls 파일:", latest_xls)

latest_xlsx = convert_xls_to_xlsx(latest_xls)

print("📖 증권사 xlsx 읽는 중...")
df_new = pd.read_excel(latest_xlsx)
df_new.columns = df_new.columns.map(lambda x: str(x).replace(" ", ""))

need_cols = [KEY_COL, ASSET_COL, RET_COL]
for col in need_cols:
    if col not in df_new.columns:
        raise KeyError(f"증권사 파일에 '{col}' 컬럼이 없습니다. 실제 컬럼 목록: {list(df_new.columns)}")

df_new = df_new[need_cols].copy()
df_new[KEY_COL] = df_new[KEY_COL].map(normalize_key)

asset_map = df_new.set_index(KEY_COL)[ASSET_COL].to_dict()
ret_map = df_new.set_index(KEY_COL)[RET_COL].to_dict()

print(f"✅ 증권사 파일에서 읽은 계약번호 수: {len(asset_map)}")


# ===========================
# 4. parkpark FOK_DATA 업데이트 (기존 업데이트 + 해지 삭제 + 신규 추가)
# ===========================
excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False

xlUp = -4162
xlToLeft = -4159

updated_rows = 0

try:
    print("📘 parkpark 파일 여는 중...")
    wb = excel.Workbooks.Open(CUSTOMER_FILE, False, False, None, PASSWORD)
    ws = wb.Worksheets("FOK_DATA")

    header_row = 1
    last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    last_col = ws.Cells(header_row, ws.Columns.Count).End(xlToLeft).Column

    # 헤더 위치 잡기
    col_key = col_asset = col_ret = None
    for c in range(1, last_col + 1):
        header = ws.Cells(header_row, c).Value
        if header is None:
            continue
        h = str(header).replace(" ", "")
        if h == KEY_COL:
            col_key = c
        elif h == ASSET_COL:
            col_asset = c
        elif h == RET_COL:
            col_ret = c

    if col_key is None or col_asset is None or col_ret is None:
        raise RuntimeError(f"FOK_DATA 시트에서 '{KEY_COL}', '{ASSET_COL}', '{RET_COL}' 헤더를 찾지 못했습니다.")

    print(f"🔎 헤더 위치 - 계약번호: {col_key}, 계좌자산: {col_asset}, 수익률: {col_ret}")
    print(f"📊 FOK_DATA 데이터 행 범위: 2 ~ {last_row}")

    # 데이터 읽기
    data_list = []
    if last_row > 1:
        data_range = ws.Range(ws.Cells(2, 1), ws.Cells(last_row, last_col))
        data = data_range.Value
        data_list = [list(row) for row in data]

    idx_key = col_key - 1
    idx_asset = col_asset - 1
    idx_ret = col_ret - 1

    existing_rows = []
    existing_keys = set()
    cancelled_count = 0

    # 1) 기존 고객 업데이트 + 해지 고객 삭제
    for row in data_list:
        raw_key = row[idx_key]
        if raw_key is None:
            continue

        key = normalize_key(raw_key)
        if not key:
            continue

        if key in asset_map:
            row[idx_asset] = asset_map[key]
            row[idx_ret] = ret_map[key]
            updated_rows += 1
            existing_rows.append(row)
            existing_keys.add(key)
        else:
            cancelled_count += 1   # 해지 고객 → 삭제 처리 (append 안함)

    # 2) 신규 고객 추가
    new_keys = [k for k in asset_map if k not in existing_keys]
    new_rows = []

    for k in new_keys:
        row = [None] * last_col
        row[idx_key] = k
        row[idx_asset] = asset_map.get(k)
        row[idx_ret] = ret_map.get(k)
        new_rows.append(row)

    # 3) 최종 데이터 구성
    final_rows = existing_rows + new_rows

    # 기존 데이터 전부 삭제
    ws.Range(ws.Cells(2, 1), ws.Cells(last_row, last_col)).ClearContents()

    # 새 데이터 쓰기
    if final_rows:
        write_range = ws.Range(ws.Cells(2, 1), ws.Cells(1 + len(final_rows), last_col))
        write_range.Value = tuple(tuple(r) for r in final_rows)

    print(f"✅ 업데이트: {updated_rows}행")
    print(f"❌ 해지로 삭제된 고객: {cancelled_count}행")
    print(f"➕ 신규 고객 추가: {len(new_rows)}행")
    print("🎉 최종적으로 FOK_DATA가 최신 증권사 데이터 기준으로 정리되었습니다.")

    wb.Save()

finally:
    try:
        wb.Close(False)
    except:
        pass
    excel.Quit()
    print("📁 엑셀 프로세스 종료")