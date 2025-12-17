import os
import pandas as pd
import win32com.client as win32

# ===========================
# 1. 기본 설정
# ===========================
download_path = os.path.join(os.path.expanduser("~"), "Downloads")
FILE_PREFIX = "file_"   # 증권사 파일 접두사 (file_066..., file_1297... 등)

CUSTOMER_FILE = r"C:\Users\pc\OneDrive - 주식회사 플레인바닐라\LEEJAEWOOK의 파일 - 플레인바닐라 업무\Customer\고객data\고객data_v101.xlsx"
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

# 가장 최근 파일 선택
xls_files.sort(
    key=lambda name: os.path.getmtime(os.path.join(download_path, name)),
    reverse=True,
)
latest_xls = os.path.join(download_path, xls_files[0])
print("📂 가장 최근 다운로드 xls 파일:", latest_xls)

latest_xlsx = convert_xls_to_xlsx(latest_xls)

print("📖 증권사 xlsx 읽는 중...")
df_new = pd.read_excel(latest_xlsx)

# 컬럼 이름 공백 제거
df_new.columns = df_new.columns.map(lambda x: str(x).replace(" ", ""))

# 필수 컬럼 체크
need_cols = [KEY_COL, ASSET_COL, RET_COL]
for col in need_cols:
    if col not in df_new.columns:
        raise KeyError(f"증권사 파일에 '{col}' 컬럼이 없습니다. 실제 컬럼 목록: {list(df_new.columns)}")

# 계약번호 정규화
df_new[KEY_COL] = df_new[KEY_COL].map(normalize_key)
df_new = df_new[df_new[KEY_COL] != ""]  # 계약번호 빈 값 제거

# 🔴 (핵심) 계약번호 중복 제거: 같은 계약번호가 여러 번 나오면 마지막 행만 사용
dup_mask = df_new.duplicated(subset=[KEY_COL], keep="last")
dup_cnt = dup_mask.sum()
if dup_cnt > 0:
    print(f"⚠ 중복 계약번호 {dup_cnt}개 발견 → 마지막 행 기준으로만 사용합니다.")
    df_new = df_new[~dup_mask]

# 계약번호를 인덱스로 사용
df_new_idx = df_new.set_index(KEY_COL)

# 기존 업데이트용 (계좌자산 / 수익률)
asset_map = df_new_idx[ASSET_COL].to_dict()
ret_map = df_new_idx[RET_COL].to_dict()

# 신규 고객 전체 데이터용: key -> {컬럼명: 값, ...}
row_map = df_new_idx.to_dict("index")

print(f"✅ 증권사 파일에서 읽은 계약번호 수: {len(asset_map)}")


# ===========================
# 4. parkpark FOK_DATA 업데이트
#    (기존 업데이트 + 해지 삭제 + 신규 전체 데이터 추가)
# ===========================
excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False

xlUp = -4162
xlToLeft = -4159

updated_rows = 0

try:
    print("📘 파일 여는 중...")
    wb = excel.Workbooks.Open(CUSTOMER_FILE, False, False, None, PASSWORD)
    ws = wb.Worksheets("FOK_DATA")

    header_row = 1
    last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    last_col = ws.Cells(header_row, ws.Columns.Count).End(xlToLeft).Column

    # 헤더 위치 및 전체 헤더 이름(공백 제거) 저장
    col_key = col_asset = col_ret = None
    header_names = [None] * last_col  # 인덱스: 0 ~ last_col-1

    for c in range(1, last_col + 1):
        header = ws.Cells(header_row, c).Value
        if header is None:
            header_names[c - 1] = None
            continue
        h = str(header).replace(" ", "")
        header_names[c - 1] = h

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



    # 인덱스 보정 (엑셀 1-based → 파이썬 0-based)
    idx_key = col_key - 1
    idx_asset = col_asset - 1
    idx_ret = col_ret - 1

    data_list = []
    if last_row > 1:
        data_range = ws.Range(ws.Cells(2, 1), ws.Cells(last_row, last_col))
        data = data_range.Value  # 2차원 튜플
        data_list = [list(row) for row in data]

    existing_rows = []
    existing_keys = set()
    cancelled_count = 0

    # 1) 기존 고객: 업데이트 or 해지 삭제
    for row in data_list:
        raw_key = row[idx_key]
        if raw_key is None:
            continue

        key = normalize_key(raw_key)
        if not key:
            continue

        if key in asset_map:
            # 기존 고객: 계좌자산 / 수익률 업데이트
            row[idx_asset] = asset_map[key]
            row[idx_ret] = ret_map[key]
            updated_rows += 1

            existing_rows.append(row)
            existing_keys.add(key)
        else:
            # 증권사 데이터에 없는 계약번호 → 해지 고객 → 삭제
            cancelled_count += 1
            # append 하지 않음 = 삭제 효과

     # 2) 신규 고객: FOK_DATA에 없는 계약번호들
    new_keys = [k for k in row_map.keys() if k not in existing_keys]
    new_rows = []

    for k in new_keys:
        row_dict = row_map.get(k, {})  # {컬럼명: 값}
        row = [None] * last_col        # FOK_DATA 열 개수만큼 빈 리스트

        # 2-1. FOK_DATA 헤더 이름과 증권사 컬럼 이름이 같은 곳은 그대로 채우기
        for idx, h in enumerate(header_names):
            if not h:
                continue
            if h in row_dict:
                row[idx] = row_dict[h]

        # 2-2. 계약번호, 계좌자산, 수익률은 확실히 채워 넣기
        # (계약번호는 인덱스라 row_dict 안에 없으므로 직접 넣어줘야 함)
        row[idx_key] = k

        # 혹시 위에서 이미 들어갔어도 다시 한 번 확실히 세팅
        if ASSET_COL in row_dict:
            row[idx_asset] = row_dict[ASSET_COL]
        else:
            row[idx_asset] = asset_map.get(k)

        if RET_COL in row_dict:
            row[idx_ret] = row_dict[RET_COL]
        else:
            row[idx_ret] = ret_map.get(k)

        new_rows.append(row)

    # 3) 최종 데이터 = 기존(해지 제거 후) + 신규
    final_rows = existing_rows + new_rows

    # 기존 데이터 지우기
    if last_row > 1:
        ws.Range(ws.Cells(2, 1), ws.Cells(last_row, last_col)).ClearContents()

    # 새 데이터 쓰기
    if final_rows:
        write_range = ws.Range(ws.Cells(2, 1), ws.Cells(1 + len(final_rows), last_col))
        write_range.Value = tuple(tuple(r) for r in final_rows)

    print(f"✅ 업데이트된 기존 고객 수: {updated_rows}행")
    print(f"❌ 해지로 삭제된 고객 수: {cancelled_count}행")
    print(f"➕ 신규 고객 추가 수: {len(new_rows)}행")
    print("🎉 FOK_DATA가 증권사 데이터 기준으로 완전히 동기화되었습니다.")

    wb.Save()

finally:
    try:
        wb.Close(False)
    except Exception:
        pass
    excel.Quit()
    print("📁 엑셀 프로세스 종료")