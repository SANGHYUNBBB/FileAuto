import os
import re
import pandas as pd
import win32com.client as win32

# ===========================
# 1. 기본 설정
# ===========================
HTS_FOLDER = os.path.join(
    os.path.expanduser("~"),
    "Downloads",
    "hts"
)
HTS_PREFIX = "Excel"  # NH HTS 파일 접두사

CUSTOMER_FILE = r"C:\Users\pc\OneDrive - 주식회사 플레인바닐라\LEEJAEWOOK의 파일 - 플레인바닐라 업무\Customer\고객data\고객data_v101.xlsx"
PASSWORD = "nilla17()"

SHEET_NH_DATA = "NH_DATA"
SHEET_DAILY = "Daily"

# 두 번째 파일에서 사용할 컬럼 이름 (공백 제거 후 기준)
COL_CODE = "상품유형"
COL_ASSET = "전일평가금액"


# ===========================
# 2. 공통 유틸
# ===========================
def convert_xls_to_xlsx(path: str) -> str:
    """첫 번째 HTS 파일(고객정보)에만 사용.
       .xlsx면 그대로 리턴, .xls면 Excel로 열어서 xlsx로 저장."""
    base, ext = os.path.splitext(path)
    if ext.lower() != ".xls":
        # 이미 xlsx이면 그대로 사용
        return path

    if not os.path.exists(path):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {path}")

    print(f"[변환 시작] {path} -> xlsx")

    import win32com.client as win32_local
    excel = win32_local.DispatchEx("Excel.Application")
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(path)
        xlsx_path = base + ".xlsx"
        wb.SaveAs(xlsx_path, FileFormat=51)
        wb.Close()
    finally:
        excel.Quit()

    print(f"[변환 완료] {path} -> {xlsx_path}")
    return xlsx_path

def extract_number_from_filename(name: str) -> int:
    """파일명에서 숫자만 뽑아서 int로 반환 (없으면 0)"""
    nums = re.findall(r"\d+", name)
    if not nums:
        return 0
    return int(nums[-1])


HTS_FOLDER = r"C:\Users\pc\Downloads\hts"
HTS_PREFIX = "Excel"


def find_two_hts_files(folder: str, prefix: str = "Excel"):
    """
    HTS 폴더 안의 Excel*.xls 파일 중
    - 숫자가 더 작은 파일 → 고객정보 파일
    - 숫자가 더 큰 파일 → 잔고파일
    로 구분해서 (customer_path, balance_path)를 반환한다.
    (xlsx는 완전히 무시)
    """
    xls_files = [
        f for f in os.listdir(folder)
        if f.startswith(prefix) and f.lower().endswith(".xls")
    ]

    if len(xls_files) < 2:
        raise FileNotFoundError(f"{folder}에 '{prefix}*.xls' 파일이 2개 이상 있어야 합니다. 현재: {xls_files}")

    def extract_number(name: str) -> int:
        m = re.search(r"(\d+)", name)
        return int(m.group(1)) if m else 0

    # 숫자 기준으로 정렬
    xls_files.sort(key=extract_number)

    # 숫자가 작은 게 고객, 큰 게 잔고
    customer_file = os.path.join(folder, xls_files[0])
    balance_file = os.path.join(folder, xls_files[-1])

    print(f"📂 HTS 고객정보 파일(작은 번호): {customer_file}")
    print(f"📂 HTS 잔고파일(큰 번호): {balance_file}")

    return customer_file, balance_file

# ===========================
# 3. 첫 번째 파일 → NH_DATA 시트 채우기
# ===========================
SHEET_NH_DATA = "NH_DATA"   # 시트 이름 다르면 여기만 바꿔줘


def update_nh_data_sheet(excel_app, parkpark_wb, customer_file_path: str):
    """
    증권사 HTS 고객파일에서
    - '자문사' 열부터 '자문관리사원명' 열까지 전체 데이터를 읽어서
    - parkpark NH_DATA 시트의 A열(자문사) ~ AW열까지 A2부터 그대로 붙여넣기
    (엑셀에서 사람 손으로 복붙하는 것과 동일한 효과)
    """


    df = pd.read_excel(customer_file_path)

    # 1) 컬럼 이름 정리 (줄바꿈, CR/LF, 공백 제거)
    def norm_col(s: str) -> str:
        s = str(s)
        for token in ["_x000D_", "\r", "\n"]:
            s = s.replace(token, "")
        return s.strip()

    original_cols = list(df.columns)
    df.columns = [norm_col(c) for c in df.columns]



    # 2) '자문사' ~ '자문관리사원명' 구간만 사용
    try:
        start_idx = df.columns.get_loc("자문사")
        end_idx = df.columns.get_loc("자문관리사원명")
    except KeyError as e:
        raise KeyError(
            "고객정보 파일에서 '자문사' 또는 '자문관리사원명' 컬럼을 찾지 못했습니다.\n"
            f"원본 컬럼: {original_cols}\n"
            f"정리 후 컬럼: {df.columns.tolist()}"
        ) from e

    df_use = df.iloc[:, start_idx:end_idx + 1]

    # 완전히 빈 행은 제거
    df_use = df_use.dropna(how="all")
    # --- 상품코드 3자리 변환 추가 ---
    # 고객파일 컬럼 이름에 '상품'이 있으니, 그 열을 001,002,003 형식으로 통일
    if "상품" in df_use.columns:
 
        df_use["상품"] = (
            df_use["상품"]
            .astype(str)
            .str.replace(".0", "", regex=False)  # 1.0 → 1
            .str.strip()
        )

        def pad_code(x: str) -> str:
            # 숫자가 아니면 그대로 두고, 숫자면 3자리로 패딩
            if not x.isdigit():
                return x
            return x.zfill(3)

        df_use["상품"] = df_use["상품"].map(pad_code)
    rows, cols = df_use.shape

    if rows == 0:
        print("⚠ 사용할 고객 데이터 행이 없습니다. NH_DATA 갱신 건너뜀.")
        return

    # 3) NaN → 빈 문자열로 바꾼 뒤 파이썬 기본 타입으로 변환
    df_use = df_use.astype(object).where(pd.notnull(df_use), "")



    # 4) NH_DATA 시트에 써 넣기 (A2부터, 행 단위로)
    nh_ws = parkpark_wb.Worksheets(SHEET_NH_DATA)

 
    nh_ws.Range("A2:AW1048576").ClearContents()



    start_row = 2  # A2에서 시작
    for i, (_, row) in enumerate(df_use.iterrows(), start=start_row):
        # 현재 행의 값들을 파이썬 리스트로 변환
        row_values = list(row.values)

        # A열부터 연속으로 cols개 셀에 한 줄씩 세팅
        nh_ws.Range(
            nh_ws.Cells(i, 1),  # A{i}
            nh_ws.Cells(i, cols)  # (A+cols-1){i}
        ).Value = row_values

        # 진행 상황 가끔 찍기
        if (i - start_row + 1) % 200 == 0 or i == start_row + rows - 1:
            print(f"   → {i - start_row + 1}/{rows} 행 붙여넣기 완료")

    # 5) 확인용 로그

    print("✅ NH_DATA 시트 업데이트 완료.")
# ===========================
# 4. 두 번째 파일 → Daily 시트 수치 업데이트
# ===========================
def update_daily_sheet_from_second(balance_file_path: str, customer_wb):

    df = pd.read_excel(balance_file_path)

    def norm_col(s: str) -> str:
        s = str(s)
        for token in ["_x000D_", "\r", "\n", " "]:
            s = s.replace(token, "")
        return s

    original_cols = list(df.columns)
    df.columns = [norm_col(c) for c in df.columns]


    code_col = "상품코드"
    asset_col = "총합계"
    if code_col not in df.columns or asset_col not in df.columns:
        raise KeyError(
            "잔고파일에서 '상품코드' 또는 '총합계' 컬럼을 찾지 못했습니다.\n"
            f"원본 컬럼: {original_cols}\n정규화 후 컬럼: {df.columns.tolist()}"
        )

    df2 = df[[code_col, asset_col]].copy()
    df2[code_col] = pd.to_numeric(df2[code_col], errors="coerce")
    df2[asset_col] = pd.to_numeric(df2[asset_col], errors="coerce")
    df2 = df2.dropna(subset=[code_col, asset_col])

    sum_4_5_won = df2.loc[df2[code_col].isin([4, 5]), asset_col].sum()
    sum_1_4_5_won = df2.loc[df2[code_col].isin([1, 4, 5]), asset_col].sum()

    print(f"📊 코드 4,5 총합계(원): {sum_4_5_won:,.0f}")
    print(f"📊 코드 1,4,5 총합계(원): {sum_1_4_5_won:,.0f}")

    sum_4_5_억 = sum_4_5_won / 100_000_000.0
    sum_1_4_5_억 = sum_1_4_5_won / 100_000_000.0

    print(f"📊 코드 4,5 총합계(억): {sum_4_5_억}")
    print(f"📊 코드 1,4,5 총합계(억): {sum_1_4_5_억}")

    daily_ws = customer_wb.Worksheets(SHEET_DAILY)
    daily_ws.Range("B14").Value = float(sum_4_5_억)   # 4,5번 합계(억)
    daily_ws.Range("C6").Value = float(sum_1_4_5_억)  # 1,4,5번 합계(억)

    print("✅ Daily 시트 B14(4·5억), C6(1·4·5억) 업데이트 완료.")
# ===========================
# 5. main 실행부
# ===========================
def main():
    # 1) HTS 폴더에서 두 개 xls 파일 찾기 (작은 번호=고객, 큰 번호=잔고)
    customer_hts, balance_hts = find_two_hts_files(HTS_FOLDER, HTS_PREFIX)

    excel = None
    wb = None

    try:
        excel = win32.DispatchEx("Excel.Application")
        try:
            excel.Visible = False
        except Exception as e:
            print(f"⚠ Excel.Visible 설정 실패, 무시하고 진행합니다: {e}")

        print("📘 parkpark 고객 파일 여는 중...")
        wb = excel.Workbooks.Open(CUSTOMER_FILE, False, False, None, PASSWORD)

        # 2) NH_DATA : 고객정보 파일 붙여넣기
        update_nh_data_sheet(excel, wb, customer_hts)

        # 3) Daily : 잔고파일로 B14, C6 업데이트
        update_daily_sheet_from_second(balance_hts, wb)

        wb.Save()
        print("💾 parkpark 파일 저장 완료.")

    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass

        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

        print("📁 엑셀 프로세스 종료")

if __name__ == "__main__":
    main()