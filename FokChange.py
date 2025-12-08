import os
import pandas as pd
import win32com.client as win32

# ===========================
# 1. 기본 설정
# ===========================
download_path = r"C:\Users\pc\Downloads"
FILE_PREFIX = "file_"   # 증권사 파일 접두사 (file_066..., file_1297... 등)

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
# 최신 file_*.xls 찾기
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
# 4. parkpark FOK_DATA 업데이트 (배열로 한 번에)
# ===========================
excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False  # True로 바꾸면 엑셀 화면 보이면서 진행됨

xlUp = -4162        # xlUp
xlToLeft = -4159    # xlToLeft

updated_rows = 0
total_rows = 0

try:
    print("📘 parkpark 파일 여는 중...")
    # parkpark 파일은 반드시 엑셀에서 닫혀 있어야 함
    wb = excel.Workbooks.Open(CUSTOMER_FILE, False, False, None, PASSWORD)
    ws = wb.Worksheets("FOK_DATA")

    header_row = 1
    last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    last_col = ws.Cells(header_row, ws.Columns.Count).End(xlToLeft).Column

    # 헤더 위치 찾기
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
        raise RuntimeError(
            f"FOK_DATA 시트에서 '{KEY_COL}', '{ASSET_COL}', '{RET_COL}' 헤더를 찾지 못했습니다."
        )

    print(f"🔎 헤더 위치 - 계약번호: {col_key}, 계좌자산: {col_asset}, 수익률: {col_ret}")
    print(f"📊 FOK_DATA 데이터 행 범위: 2 ~ {last_row}")

    # ----★ 핵심: Range 전체를 한 번에 배열로 읽어오기 ----
    data_range = ws.Range(ws.Cells(2, 1), ws.Cells(last_row, last_col))
    data = data_range.Value  # 2차원 튜플 (row, col)

    # 튜플 → 리스트로 변환 (수정 가능하게)
    if last_row == 1:
        print("데이터 행이 없습니다.")
    else:
        rows = last_row - 1  # 헤더 제외
        cols = last_col
        data_list = [list(row) for row in data]

        total_rows = rows
        print(f"⚙ 총 {total_rows}개 행 업데이트 시도 중...")

        # 인덱스 보정: 엑셀 열 번호는 1부터, 파이썬 인덱스는 0부터
        idx_key = col_key - 1
        idx_asset = col_asset - 1
        idx_ret = col_ret - 1

        for i, row in enumerate(data_list):
            raw_key = row[idx_key]
            if raw_key is None:
                continue

            key = normalize_key(raw_key)
            if not key:
                continue

            changed = False

            if key in asset_map:
                row[idx_asset] = asset_map[key]
                changed = True
            if key in ret_map:
                row[idx_ret] = ret_map[key]
                changed = True

            if changed:
                updated_rows += 1

            # ★ 진행 상황 로그 (500행마다 한 번씩)
            if (i + 1) % 500 == 0 or (i + 1) == total_rows:
                print(f"   → {i+1}/{total_rows} 행 처리 완료 (현재까지 업데이트 {updated_rows}행)")

        # ----★ 수정된 배열을 엑셀에 한 번에 다시 쓰기 ----
        data_range.Value = tuple(tuple(row) for row in data_list)

    wb.Save()
    print(f"✅ 최종 업데이트 완료: {updated_rows}개 행의 계좌자산/수익률을 최신 증권사 데이터로 반영했습니다.")

finally:
    try:
        wb.Close(False)
    except Exception:
        pass
    excel.Quit()
    print("📁 엑셀 프로세스 종료")