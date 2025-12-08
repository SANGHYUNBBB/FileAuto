import os
import re
import pandas as pd
import win32com.client as win32

# ===========================
# 1. 기본 설정
# ===========================
HTS_FOLDER = r"C:\Users\pc\Downloads\hts"
HTS_PREFIX = "Excel"  # NH HTS 파일 접두사

CUSTOMER_FILE = r"C:\Users\pc\OneDrive - 주식회사 플레인바닐라\LEEJAEWOOK의 파일 - 플레인바닐라 업무\Customer\고객data\고객data_v101_parkpark.xlsx"
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
    """xls면 xlsx로 변환해서 xlsx 경로를 리턴, 이미 xlsx면 그대로 리턴"""
    base, ext = os.path.splitext(path)
    if ext.lower() == ".xlsx":
        return path

    if not os.path.exists(path):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {path}")

    print(f"[변환 시작] {path} -> xlsx")

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(path)
        xlsx_path = base + ".xlsx"
        wb.SaveAs(xlsx_path, FileFormat=51)  # xlsx
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


def find_two_hts_files(folder: str, prefix: str):
    """폴더에서 prefix로 시작하는 파일들 중 가장 최근 2개를 찾고,
    그 둘을 숫자 기준으로 작은 것 / 큰 것으로 나눠서 리턴"""
    files = [
        f for f in os.listdir(folder)
        if f.startswith(prefix) and f.lower().endswith((".xls", ".xlsx"))
    ]
    if len(files) < 2:
        raise FileNotFoundError(f"{folder} 안에 '{prefix}*' 형식의 엑셀 파일이 2개 이상 필요합니다. 현재: {files}")

    # 수정 시간 기준으로 최근 2개
    files.sort(key=lambda n: os.path.getmtime(os.path.join(folder, n)), reverse=True)
    latest_two = files[:2]

    # 두 개 중 숫자 기준으로 작은/큰 파일 나누기
    nums = [extract_number_from_filename(n) for n in latest_two]
    if nums[0] <= nums[1]:
        smaller, larger = latest_two[0], latest_two[1]
    else:
        smaller, larger = latest_two[1], latest_two[0]

    first_path = os.path.join(folder, smaller)  # 고객 정보 파일
    second_path = os.path.join(folder, larger)  # 계좌 잔고 파일

    print("📂 HTS 첫 번째 파일(고객정보):", first_path)
    print("📂 HTS 두 번째 파일(잔고파일):", second_path)

    return first_path, second_path


# ===========================
# 3. 첫 번째 파일 → NH_DATA 시트 채우기
# ===========================
def update_nh_data_sheet(excel_app, customer_wb, first_xlsx_path: str):
    """
    1) 첫 번째 HTS 파일에서 AG:AQ 열 삭제
    2) A열부터 마지막 사용 열까지(자문사~자동주문여부)를 모두 복사
    3) parkpark의 NH_DATA 시트 A2~ 에 붙여넣기 (기존 데이터 삭제 후)
    """
    print("📘 첫 번째 HTS 파일 여는 중 (NH 고객정보)...")
    src_wb = excel_app.Workbooks.Open(first_xlsx_path)
    src_ws = src_wb.Worksheets(1)  # 보통 첫 번째 시트 사용

    xlUp = -4162
    xlToLeft = -4159

    # 1) AG~AQ 열 삭제 (오른쪽 데이터가 왼쪽으로 밀려서 마지막 열이 AY가 됨)
    print("✂ AG:AQ 열 삭제 중...")
    src_ws.Range("AG:AQ").Delete()

    # 2) 마지막 행/열 동적으로 찾기
    #   - 행: A열 기준 마지막 데이터 행
    #   - 열: 헤더가 있는 1행에서 맨 오른쪽 사용 열
    last_row = src_ws.Cells(src_ws.Rows.Count, "A").End(xlUp).Row
    last_col = src_ws.Cells(1, src_ws.Columns.Count).End(xlToLeft).Column

    if last_row < 2:
        print("⚠ 고객 데이터가 없습니다. (A열 기준 데이터 행 없음)")
        src_wb.Close(False)
        return

    # 자문사(열 A)부터 마지막 열까지 전체 고객 데이터 범위 설정
    first_col_idx = 1  # A열
    src_range = src_ws.Range(
        src_ws.Cells(2, first_col_idx),
        src_ws.Cells(last_row, last_col)
    )

    rows = last_row - 1
    cols = last_col - first_col_idx + 1
    print(f"✅ HTS 고객 데이터 범위: A2:{chr(64+last_col)}{last_row} (rows={rows}, cols={cols})")

    # 3) parkpark NH_DATA 시트에 붙여넣기
    nh_ws = customer_wb.Worksheets(SHEET_NH_DATA)

    # 기존 데이터 지우기 (A열~마지막 열, 2행 이후)
    print("🧹 NH_DATA 기존 데이터 삭제 중...")
    nh_ws.Range("A2:AZ1048576").ClearContents()  # 넉넉하게 삭제

    print("📥 NH_DATA 시트에 고객 데이터 붙여넣는 중...")
    dest_range = nh_ws.Cells(2, 1).Resize(rows, cols)  # A2부터 시작
    dest_range.Value = src_range.Value

    src_wb.Close(False)
    print("✅ NH_DATA 시트 업데이트 완료.")

# ===========================
# 4. 두 번째 파일 → Daily 시트 수치 업데이트
# ===========================
def update_daily_sheet_from_second(second_xlsx_path: str, customer_wb):
    print("📖 두 번째 HTS xlsx 읽는 중 (잔고파일)...")
    df = pd.read_excel(second_xlsx_path)

    # 1) 컬럼 이름 정규화 함수 정의
    def norm_col(s: str) -> str:
        s = str(s)
        # 줄바꿈, 캐리지리턴, _x000D_ , 공백 제거
        for token in ["_x000D_", "\r", "\n", " "]:
            s = s.replace(token, "")
        return s

    # 2) 정규화된 컬럼 이름 적용
    original_cols = list(df.columns)
    df.columns = [norm_col(c) for c in df.columns]

    print("🔎 정규화된 컬럼 목록:", list(df.columns))

    # 3) 코드 / 잔고 컬럼 후보 지정
    code_candidates = ["상품코드", "상품유형"]
    asset_candidates = ["총자산", "전일평가금액", "순자산", "총합계"]

    # 실제 존재하는 컬럼 찾기
    code_col = next((c for c in code_candidates if c in df.columns), None)
    asset_col = next((c for c in asset_candidates if c in df.columns), None)

    if code_col is None or asset_col is None:
        raise KeyError(
            "두 번째 파일에서 상품코드/잔고 컬럼을 찾지 못했습니다.\n"
            f"원본 컬럼 목록: {original_cols}\n"
            f"정규화 후 컬럼 목록: {list(df.columns)}"
        )

    print(f"✅ 사용 컬럼 - 코드: {code_col}, 자산: {asset_col}")

    # 4) 필요한 컬럼만 사용
    df2 = df[[code_col, asset_col]].copy()

    # 숫자로 변환
    df2[code_col] = pd.to_numeric(df2[code_col], errors="coerce")
    df2[asset_col] = pd.to_numeric(df2[asset_col], errors="coerce")

    # 코드/자산이 NaN인 행 제거
    df2 = df2.dropna(subset=[code_col, asset_col])

    # 5) 합계 계산
    sum_4_5 = df2.loc[df2[code_col].isin([4, 5]), asset_col].sum()
    sum_1_4_5 = df2.loc[df2[code_col].isin([1, 4, 5]), asset_col].sum()

    print(f"📊 코드 4,5 총자산 합: {sum_4_5:,.0f}")
    print(f"📊 코드 1,4,5 총자산 합: {sum_1_4_5:,.0f}")

    # 6) Daily 시트에 쓰기
    daily_ws = customer_wb.Worksheets(SHEET_DAILY)
    daily_ws.Range("B14").Value = float(sum_4_5)    # NH 여연금계좌 잔고
    daily_ws.Range("C6").Value = float(sum_1_4_5)   # NH 자문잔고

    print("✅ Daily 시트 B14, C6 업데이트 완료.")

# ===========================
# 5. main 실행부
# ===========================
def main():
    # 1) HTS 폴더에서 두 개 파일 찾기
    first_path, second_path = find_two_hts_files(HTS_FOLDER, HTS_PREFIX)

    # 2) 필요하면 xls → xlsx 변환
    first_xlsx = convert_xls_to_xlsx(first_path)
    second_xlsx = convert_xls_to_xlsx(second_path)

    # 3) parkpark 엑셀 열고 작업
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False  # True로 바꾸면 엑셀 실행되는 거 보이게 할 수 있음

    try:
        print("📘 parkpark 고객 파일 여는 중...")
        wb = excel.Workbooks.Open(CUSTOMER_FILE, False, False, None, PASSWORD)

        # NH_DATA 시트 업데이트
        update_nh_data_sheet(excel, wb, first_xlsx)

        # Daily 시트 업데이트
        update_daily_sheet_from_second(second_xlsx, wb)

        wb.Save()
        print("💾 parkpark 파일 저장 완료.")

    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        excel.Quit()
        print("📁 엑셀 프로세스 종료")


if __name__ == "__main__":
    main()