import os
import pandas as pd
import win32com.client as win32

# ============================================
# 1. 기본 설정
# ============================================
download_path = r"C:\Users\pc\Downloads"

# 증권사 엑셀 파일 이름 접두사
# (현재는 file_ 로 시작하는 최신 xls 를 사용. 필요하면 "file_066" 등으로 바꿔도 됨)
FILE_PREFIX = "file_"

# 기존 고객 엑셀 (비밀번호 걸려 있음)
CUSTOMER_FILE = r"C:\Users\pc\OneDrive - 주식회사 플레인바닐라\LEEJAEWOOK의 파일 - 플레인바닐라 업무\Customer\고객data\고객data_v101_parkpark.xlsx"
PASSWORD = "nilla17()"

# 비교 결과 리포트 저장 경로
DIFF_REPORT = r"C:\Code\FOK_DIFF_REPORT.xlsx"

# 비교 키 컬럼
KEY_COL = "계약번호"


# ============================================
# 2. 증권사 xls → xlsx 변환 함수
# ============================================
def convert_xls_to_xlsx(xls_path: str) -> str:
    """xls 파일을 xlsx로 변환해서 경로를 반환"""
    if not os.path.exists(xls_path):
        raise FileNotFoundError(f"xls 파일을 찾을 수 없습니다: {xls_path}")

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(xls_path)
        xlsx_path = os.path.splitext(xls_path)[0] + ".xlsx"
        wb.SaveAs(xlsx_path, FileFormat=51)  # 51 = xlsx
        wb.Close()
    finally:
        excel.Quit()

    print(f"[변환 완료] {xls_path} -> {xlsx_path}")
    return xlsx_path


# ============================================
# 3. 비밀번호 걸린 엑셀에서 FOK_DATA 시트를 pandas로 읽기
# ============================================
def read_fok_data_from_protected(path: str, password: str, sheet_name: str = "FOK_DATA") -> pd.DataFrame:
    """
    win32com으로 비밀번호 걸린 엑셀을 열고,
    sheet_name의 UsedRange를 읽어서 pandas DataFrame으로 반환.
    """
    if not os.path.exists(path):
        raise FileNotFoundError(f"고객 파일이 없습니다: {path}")

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    try:
        # ReadOnly=True 로 열기 (저장 안 함)
        wb = excel.Workbooks.Open(path, False, True, None, password)
        ws = wb.Worksheets(sheet_name)

        used = ws.UsedRange
        values = used.Value  # 2차원 튜플 (헤더 + 데이터)
    finally:
        wb.Close(False)
        excel.Quit()

    # values → 2차원 리스트로 정규화
    if not isinstance(values, tuple):
        data = [[values]]
    else:
        if isinstance(values[0], tuple):
            data = [list(row) for row in values]
        else:
            data = [list(values)]

    header = data[0]
    rows = data[1:]

    df = pd.DataFrame(rows, columns=header)

    # ------ 🔥 None 컬럼 제거 ------
    df = df.loc[:, df.columns.notnull()]

    # ------ 🔥 컬럼명 공백 제거 (중간 공백 포함) ------
    df.columns = df.columns.map(lambda x: str(x).replace(" ", ""))

    return df


# ============================================
# 4. 계약번호 정규화 함수
# ============================================
def normalize_contract(df: pd.DataFrame) -> pd.DataFrame:
    """
    계약번호 컬럼을 문자열로 통일하고, .0, 공백 등을 제거해서
    새 데이터/기존 데이터가 동일하게 비교되도록 함.
    """
    df[KEY_COL] = (
        df[KEY_COL]
        .astype(str)
        .str.replace(r"\.0$", "", regex=True)  # 123.0 -> 123
        .str.strip()
    )
    return df


# ============================================
# 5. 최신 증권사 데이터 DataFrame 준비
# ============================================
# 최신 file_* .xls 찾기
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
print("가장 최근 다운로드 xls 파일:", latest_xls)

latest_xlsx = convert_xls_to_xlsx(latest_xls)

# pandas로 새 데이터 읽기
df_new = pd.read_excel(latest_xlsx)

# 컬럼명 공백 제거 (새 데이터도 동일하게)
df_new.columns = df_new.columns.map(lambda x: str(x).replace(" ", ""))

# 예수금, 평가금액 제거 (있으면)
drop_cols = ["예수금", "평가금액"]
df_new = df_new.drop(columns=drop_cols, errors="ignore")

# NaN은 None으로 (엑셀 빈칸처럼 보이도록)
df_new = df_new.where(pd.notna(df_new), None)

print("새 데이터 컬럼:", list(df_new.columns))


# ============================================
# 6. 기존 FOK_DATA DataFrame 읽기
# ============================================
df_old = read_fok_data_from_protected(CUSTOMER_FILE, PASSWORD, "FOK_DATA")
print("기존 FOK_DATA 컬럼:", list(df_old.columns))

# ============================================
# 7. 계약번호 정규화
# ============================================
df_new = normalize_contract(df_new)
df_old = normalize_contract(df_old)

# ============================================
# 8. '계약번호' 기준으로 두 DataFrame 비교
# ============================================
df_new_key = df_new.set_index(KEY_COL)
df_old_key = df_old.set_index(KEY_COL)

# 두 쪽에 공통으로 있는 컬럼만 비교 대상으로 사용
common_cols = [c for c in df_new_key.columns if c in df_old_key.columns]
df_new_sub = df_new_key[common_cols].copy()
df_old_sub = df_old_key[common_cols].copy()

# 추가/삭제/공통 키 구하기
added_keys = df_new_sub.index.difference(df_old_sub.index)     # 새로 생긴 계약번호
removed_keys = df_old_sub.index.difference(df_new_sub.index)   # 기존에만 있던 계약번호
common_keys = df_new_sub.index.intersection(df_old_sub.index)  # 둘 다 있는 계약번호

print("추가된 계약번호 수:", len(added_keys))
print("삭제된 계약번호 수:", len(removed_keys))
print("공통 계약번호 수:", len(common_keys))

if len(common_keys) < 50:
    print("⚠ 공통 계약번호가 너무 적습니다. 계약번호 형식(공백/문자열) 문제일 수 있습니다.")

# 공통 키에서 셀 단위로 값이 다른 부분만 추출
diff_records = []

def to_scalar(x):
    """Series → 단일값, NaN → None 변환"""
    if isinstance(x, pd.Series):
        x = x.iloc[0]
    if pd.isna(x):
        return None
    return x

for key in common_keys:

    old_row = df_old_sub.loc[key]
    new_row = df_new_sub.loc[key]

    # --------------------------
    # 1) 계약번호가 유일하여 Series로 나오는 경우
    # --------------------------
    if isinstance(old_row, pd.Series):

        for col in common_cols:
            old_v = to_scalar(old_row[col])
            new_v = to_scalar(new_row[col])

            # 값이 다를 때만 기록
            if old_v != new_v:
                diff_records.append({
                    KEY_COL: key,
                    "컬럼": col,
                    "기존값": old_v,
                    "신규값": new_v,
                })

    # --------------------------
    # 2) 계약번호가 중복되어 DataFrame으로 나오는 경우
    # --------------------------
    else:
        old_df = old_row
        new_df = new_row

        common_idx = old_df.index.intersection(new_df.index)

        for ridx in common_idx:
            o = old_df.loc[ridx]
            n = new_df.loc[ridx]

            for col in common_cols:
                old_v = to_scalar(o[col])
                new_v = to_scalar(n[col])

                if old_v != new_v:
                    diff_records.append({
                        KEY_COL: key,
                        "row_id": ridx,
                        "컬럼": col,
                        "기존값": old_v,
                        "신규값": new_v,
                    })

df_diff = pd.DataFrame(diff_records)

# 추가/삭제 키도 DataFrame으로 정리
df_added = df_new_sub.loc[added_keys].reset_index()
df_removed = df_old_sub.loc[removed_keys].reset_index()
from datetime import datetime

def strip_timezone(df: pd.DataFrame) -> pd.DataFrame:
    """
    DataFrame 안의 timezone 포함 datetime 컬럼/값들에서 tz 정보를 제거.
    - datetime64[ns, tz] 타입 컬럼
    - object 컬럼 안의 tz-aware datetime 객체
    둘 다 처리.
    """
    df = df.copy()

    # 1) datetime64[ns, tz] 타입 컬럼 처리
    for col in df.columns:
        col_data = df[col]
        # pandas의 tz-aware datetime 컬럼
        if hasattr(col_data.dtype, "tz") and col_data.dtype.tz is not None:
            # tz 정보를 날리고 naive datetime으로
            df[col] = col_data.dt.tz_localize(None)

    # 2) object 타입 컬럼에 섞인 tz-aware datetime 처리
    for col in df.columns:
        if df[col].dtype == "object":
            def _strip_tz(v):
                # pandas Timestamp
                if isinstance(v, pd.Timestamp) and v.tz is not None:
                    return v.tz_localize(None)
                # 파이썬 datetime
                if isinstance(v, datetime) and v.tzinfo is not None:
                    return v.replace(tzinfo=None)
                return v
            df[col] = df[col].map(_strip_tz)

    return df
# 타임존 포함 datetime 제거
df_diff = strip_timezone(df_diff)
df_added = strip_timezone(df_added)
df_removed = strip_timezone(df_removed)
df_new_sub = strip_timezone(df_new_sub)
df_old_sub = strip_timezone(df_old_sub)
# ============================================
# 9. 리포트 엑셀로 저장
# ============================================
with pd.ExcelWriter(DIFF_REPORT, engine="openpyxl") as writer:
    df_diff.to_excel(writer, sheet_name="변경된_셀", index=False)
    df_added.to_excel(writer, sheet_name="추가된_계약번호", index=False)
    df_removed.to_excel(writer, sheet_name="삭제된_계약번호", index=False)
    df_new_sub.reset_index().to_excel(writer, sheet_name="신규기준_전체데이터", index=False)
    df_old_sub.reset_index().to_excel(writer, sheet_name="기존_FOK_DATA", index=False)

# ============================================
# 🔥 10. 변경된 값으로 FOK_DATA 업데이트 생성
# ============================================

df_updated = df_old_sub.copy()   # 기존 데이터 기반으로 복사

for rec in diff_records:
    key = str(rec[KEY_COL])
    col = rec["컬럼"]
    new_val = rec["신규값"]

    # 해당 key가 기존 df에 있을 때만 업데이트
    if key in df_updated.index:
        df_updated.at[key, col] = new_val

# df_updated 를 엑셀로 저장 또는 Win32로 FOK_DATA에 Write 가능
df_updated_reset = df_updated.reset_index()
df_updated_reset.to_excel("C:/Code/FOK_UPDATED.xlsx", index=False)
print("✅ 비교 완료. 리포트 저장:", DIFF_REPORT)
print("  - 변경된_셀 : 같은 계약번호인데 값이 달라진 셀 목록")
print("  - 추가된_계약번호 : 새 파일에만 존재하는 계약번호 행")
print("  - 삭제된_계약번호 : 기존 FOK_DATA에만 있던 계약번호 행")
print("  - 신규기준_전체데이터 : 새 증권사 데이터를 기준으로 정리한 전체")
print("  - 기존_FOK_DATA : 비교에 사용한 FOK_DATA 스냅샷")