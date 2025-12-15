import os
import pandas as pd
import win32com.client as win32
import gc

# ===========================
# 1. 기본 설정
# ===========================
DOWNLOAD_DIR = r"C:\Users\pc\Downloads"
T1_PREFIX = "자문결합계좌 실적조회"

CUSTOMER_FILE = r"C:\Users\pc\OneDrive - 주식회사 플레인바닐라\LEEJAEWOOK의 파일 - 플레인바닐라 업무\Customer\고객data\고객data_v101_parkpark.xlsx"
PASSWORD = "nilla17()"
SHEET_DAILY = "Daily"


# ===========================
# 2. 공통 유틸
# ===========================
def convert_xls_to_xlsx(path: str) -> str:
    """ .xls 파일을 Excel로 열어서 .xlsx로 변환 (이미 xlsx면 그대로 반환) """
    base, ext = os.path.splitext(path)
    if ext.lower() != ".xls":
        return path

    if not os.path.exists(path):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {path}")

    print(f"[변환 시작] {path} -> xlsx")
    excel = win32.DispatchEx("Excel.Application")
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


def find_latest_t1_file() -> str:
    """다운로드 폴더에서 '자문결합계좌 실적조회*.xls(x)' 중 가장 최근 파일 반환"""
    candidates = [
        f
        for f in os.listdir(DOWNLOAD_DIR)
        if f.startswith(T1_PREFIX) and f.lower().endswith((".xls", ".xlsx"))
    ]
    if not candidates:
        raise FileNotFoundError(
            f"{DOWNLOAD_DIR} 에 '{T1_PREFIX}*.xls(x)' 파일이 없습니다."
        )

    candidates.sort(
        key=lambda name: os.path.getmtime(os.path.join(DOWNLOAD_DIR, name)),
        reverse=True,
    )
    latest = os.path.join(DOWNLOAD_DIR, candidates[0])
    print(f"📂 최신 T1 파일: {latest}")
    return latest


def parse_numbers_from_t1(path: str):
    """
    T1 파일에서:
      - E4 + E5 합 (원 단위)
      - E6 값 (원 단위)
    을 읽어 반환 (sum_4_5, val_6)
    """
    xlsx = convert_xls_to_xlsx(path)

    print("📖 T1 파일 pandas로 읽는 중...(header=None, 절대셀 접근)")
    # header=None 으로 해서 엑셀의 1행=0, 2행=1, ... 그대로 맞춰 씀
    df = pd.read_excel(xlsx, header=None)

    # E열 = 5번째 열 = 인덱스 4
    def to_number(v):
        s = str(v)
        # 숫자/마이너스/점 빼고 전부 제거 (콤마, 원 등)
        s_clean = "".join(ch for ch in s if ch.isdigit() or ch in "-.")
        try:
            return float(s_clean) if s_clean not in ("", "-", ".", "-.") else 0.0
        except ValueError:
            return 0.0

    e4 = to_number(df.iloc[3, 4])  # 4행(E4)
    e5 = to_number(df.iloc[4, 4])  # 5행(E5)
    e6 = to_number(df.iloc[5, 4])  # 6행(E6)

    sum_4_5 = e4 + e5

    print(f"🔢 E4: {e4:,.0f}")
    print(f"🔢 E5: {e5:,.0f}")
    print(f"💰 E4 + E5 합계(원): {sum_4_5:,.0f}")
    print(f"💰 E6 값(원): {e6:,.0f}")

    return sum_4_5, e6


# ===========================
# 3. parkpark Daily 업데이트
# ===========================
def write_to_daily(sum_4_5_won: float, e6_won: float):
    print("📘 parkpark 파일 열어서 Daily 업데이트 중...")

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    wb = None

    try:
        try:
            excel.ScreenUpdating = False
            excel.DisplayAlerts = False
        except Exception:
            pass

        wb = excel.Workbooks.Open(CUSTOMER_FILE, False, False, None, PASSWORD)
        ws = wb.Worksheets(SHEET_DAILY)

        # ⭐ 억 단위 변환
        b12_value = sum_4_5_won / 100_000_000
        g6_value = e6_won / 100_000_000

        # 소수점 그대로 넣기
        ws.Range("B12").Value = float(b12_value)
        ws.Range("G6").Value = float(g6_value)

        print(f"✏ Daily!B12 = {ws.Range('B12').Value}")
        print(f"✏ Daily!G6  = {ws.Range('G6').Value}")

        wb.Save()
        print("💾 parkpark 저장 완료.")

        wb.Close(SaveChanges=False)
        wb = None

    finally:
        try:
            excel.ScreenUpdating = True
        except Exception:
            pass

        try:
            excel.Quit()
        except Exception:
            pass

        del excel
        gc.collect()
        print("📁 엑셀 종료")

# ===========================
# 4. main
# ===========================
def main():
    latest_t1 = find_latest_t1_file()
    sum_4_5, e6 = parse_numbers_from_t1(latest_t1)
    write_to_daily(sum_4_5, e6)


if __name__ == "__main__":
    main()