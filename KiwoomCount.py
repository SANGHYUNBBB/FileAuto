import os
import pandas as pd
import win32com.client as win32
import gc

# ======================
# 1. 기본 설정
# ======================
DOWNLOAD_DIR = r"C:\Users\pc\Downloads"
LIST_PREFIX = "Excel_List_"

CUSTOMER_FILE = r"C:\Users\pc\OneDrive - 주식회사 플레인바닐라\LEEJAEWOOK의 파일 - 플레인바닐라 업무\Customer\고객data\고객data_v101.xlsx"
# ↑ 앞에서 만든 작업용 파일 쓰는 걸 추천. 원본 쓰고 싶으면 이름만 바꿔줘.
PASSWORD = "nilla17()"
SHEET_DAILY = "Daily"


# ======================
# 2. 공통 유틸
# ======================
def convert_xls_to_xlsx(path: str) -> str:
    """ .xls 를 Excel로 열어서 .xlsx 로 변환 """
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
        wb.SaveAs(xlsx_path, FileFormat=51)
        wb.Close()
    finally:
        excel.Quit()
    print(f"[변환 완료] {path} -> {xlsx_path}")
    return xlsx_path


def norm_col(s: str) -> str:
    """컬럼 이름 정리: 줄바꿈, CR/LF, 공백 제거"""
    s = str(s)
    for token in ["_x000D_", "\r", "\n", " "]:
        s = s.replace(token, "")
    return s.strip()


# ======================
# 3. 최신 Excel_List_ 찾기 + 연금 예탁자산 합계 계산
# ======================
def get_latest_list_file() -> str:
    files = [
        f for f in os.listdir(DOWNLOAD_DIR)
        if f.startswith(LIST_PREFIX) and f.lower().endswith((".xls", ".xlsx"))
    ]
    if not files:
        raise FileNotFoundError(f"{DOWNLOAD_DIR} 에 '{LIST_PREFIX}*.xls(x)' 파일이 없습니다.")

    # 수정시간 기준 최신 파일
    files.sort(
        key=lambda name: os.path.getmtime(os.path.join(DOWNLOAD_DIR, name)),
        reverse=True,
    )
    latest = os.path.join(DOWNLOAD_DIR, files[0])
    print(f"📂 최신 Excel_List 파일: {latest}")
    return latest


def calc_pension_total_eok() -> float:
    """Excel_List_ 최신 파일에서 계좌유형=연금 의 예탁자산 합계를 억 단위로 계산"""
    latest_path = get_latest_list_file()
    latest_xlsx = convert_xls_to_xlsx(latest_path)


    df = pd.read_excel(latest_xlsx)

    original_cols = list(df.columns)
    df.columns = [norm_col(c) for c in df.columns]


    col_type = "계좌유형"
    col_asset = "예탁자산"

    if col_type not in df.columns or col_asset not in df.columns:
        raise KeyError(
            "Excel_List 파일에서 '계좌유형' 또는 '예탁자산' 컬럼을 찾지 못했습니다.\n"
            f"원본 컬럼: {original_cols}\n"
            f"정리 후: {df.columns.tolist()}"
        )

    # 계좌유형에 '연금' 이 들어간 행만 필터
    mask = df[col_type].astype(str).str.contains("연금", na=False)
    df_pension = df.loc[mask, [col_type, col_asset]].copy()
    print(f"📊 '연금' 계좌 행 수: {len(df_pension)}")

    if df_pension.empty:
        print("⚠ 연금 계좌가 없습니다. 0원으로 처리합니다.")
        return 0.0

    # 예탁자산 문자열에서 숫자만 추출 (콤마, '원' 등 제거)
    asset_str = df_pension[col_asset].astype(str)
    asset_clean = asset_str.str.replace(r"[^0-9\-\.]", "", regex=True)
    asset_num = pd.to_numeric(asset_clean, errors="coerce").fillna(0)

    total_won = asset_num.sum()
    print(f"💰 연금 계좌 예탁자산 합계(원): {total_won:,.0f}")

    total_eok = total_won / 100_000_000.0
    print(f"💰 연금 계좌 예탁자산 합계(억): {total_eok}")

    return float(total_eok)


# ======================
# 4. parkpark Daily!B12 업데이트
# ======================
def write_to_daily_b11(value_eok: float):
    import gc
    print("📘 parkpark 파일 열어서 Daily 업데이트 중...")

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False

    wb = None
    try:
        # 화면 업데이트 끄기 (속도 ↑)
        try:
            excel.ScreenUpdating = False
            excel.DisplayAlerts = False
        except Exception:
            pass

        # 🔐 비밀번호 자동 입력하여 엑셀 열기
        wb = excel.Workbooks.Open(CUSTOMER_FILE, False, False, None, PASSWORD)

        print("   → 실제로 연 파일:", wb.FullName)
        print("   → ReadOnly 여부:", wb.ReadOnly)

        ws_daily = wb.Worksheets(SHEET_DAILY)

        # ⭐⭐⭐ 핵심 변경점: B11에만 값 넣기 ⭐⭐⭐
        ws_daily.Range("B11").Value = float(value_eok)
        print("✏ Daily!B11 현재 값(엑셀 내부):", ws_daily.Range("B11").Value)

        # 저장
        wb.Save()


        # 닫기
        wb.Close(SaveChanges=False)
        wb = None


    except Exception as e:
        print("❌ Daily 업데이트 중 오류:", e)
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
            wb = None
        raise

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


# ======================
# 5. main
# ======================
def main():
    total_eok = calc_pension_total_eok()
    write_to_daily_b11(total_eok)


if __name__ == "__main__":
    main()