import os
import re
import pandas as pd
import win32com.client as win32
import gc
import time
import pywintypes

# ===========================
# 1) 설정
# ===========================
DOWNLOAD_DIR = r"C:\Users\pc\Downloads"
SRC_PREFIX = "통합 문서1"  # 삼성증권 파일 이름(접두사)

PARKPARK_FILE = r"C:\Users\pc\OneDrive - 주식회사 플레인바닐라\LEEJAEWOOK의 파일 - 플레인바닐라 업무\Customer\고객data\고객data_v101_parkpark.xlsx"
PASSWORD = "nilla17()"

SHEET_DST = "삼성_DATA"

DST_START_ROW = 6      # B6부터 데이터
DST_START_COL = 2      # B열
DST_REMARK_COL = 1     # A열(비고)

# 삼성증권 파일에서 "B~X" (총 23개 컬럼)
PASTE_COLS = 23

# B열부터의 상대 위치로 계약번호는 E열이므로 (B,C,D,E) = 4번째
CONTRACT_REL_IDX = 3  # 0-based: B=0,C=1,D=2,E=3


# ===========================
# 2) 유틸
# ===========================
def com_call_with_retry(fn, tries=8, delay=0.3, name="COM call"):
    """
    Excel COM 호출이 0x800AC472(바쁨)로 실패할 때 재시도
    """
    last_err = None
    for i in range(tries):
        try:
            return fn()
        except pywintypes.com_error as e:
            last_err = e
            # 엑셀 Busy/Call rejected 류
            if e.args and isinstance(e.args[0], int) and e.args[0] in (-2146777998, -2147418111):
                time.sleep(delay)
                continue
            raise
    raise last_err
def norm_contract(v) -> str:
    """계약번호 정규화 (공백/개행 제거)"""
    if v is None:
        return ""
    s = str(v).strip().replace("\r", "").replace("\n", "")
    return s

def extract_number_from_name(name: str) -> int:
    nums = re.findall(r"\d+", name)
    return int(nums[-1]) if nums else 0

def find_latest_source_file() -> str:
    candidates = [
        f for f in os.listdir(DOWNLOAD_DIR)
        if f.startswith(SRC_PREFIX) and f.lower().endswith((".xls", ".xlsx"))
    ]
    if not candidates:
        raise FileNotFoundError(f"{DOWNLOAD_DIR} 에 '{SRC_PREFIX}*.xls(x)' 파일이 없습니다.")

    candidates.sort(key=lambda n: os.path.getmtime(os.path.join(DOWNLOAD_DIR, n)), reverse=True)
    latest = os.path.join(DOWNLOAD_DIR, candidates[0])
    print(f"📂 최신 삼성증권 파일: {latest}")
    return latest

def convert_xls_to_xlsx(path: str) -> str:
    """xls면 xlsx로 변환. xlsx면 그대로."""
    base, ext = os.path.splitext(path)
    if ext.lower() != ".xls":
        return path

    print(f"[변환 시작] {path} -> xlsx")
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    wb = None
    try:
        wb = excel.Workbooks.Open(path)
        xlsx_path = base + ".xlsx"
        wb.SaveAs(xlsx_path, FileFormat=51)  # xlsx
        wb.Close(False)
        wb = None
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        try:
            excel.Quit()
        except Exception:
            pass
        del excel
        gc.collect()

    print(f"[변환 완료] {path} -> {xlsx_path}")
    return xlsx_path

def no_sci_number(x):
    """
    7.15E+11 → 715000000000 같은 '일반 숫자'로 변환
    """
    if x is None:
        return ""
    try:
        return int(float(x))
    except Exception:
        return x
# ===========================
# 3) 삼성증권 파일 읽기 + 정렬
# ===========================
def to_text_no_sci(x):
    if pd.isna(x):
        return ""
    if isinstance(x, str):
        return x.strip().replace("\r", "").replace("\n", "")
    try:
        return str(int(float(x)))
    except Exception:
        return str(x)
    
def read_and_sort_source(src_path: str):
    src_xlsx = convert_xls_to_xlsx(src_path)

    # 🔥 핵심: 전 컬럼 문자열로 읽기 (지수표기 원천 차단)
    df = pd.read_excel(
        src_xlsx,
        header=0,
        dtype=str
    )

    # B~X만 사용
    df_bx = df.iloc[:, 1:1 + PASTE_COLS].copy()

    # 계약번호(E열) 정규화
    df_bx["__contract__"] = df_bx.iloc[:, CONTRACT_REL_IDX].map(
        lambda x: "" if x is None else str(x).strip()
    )

    # 계약번호 없는 행 제거 + 정렬
    df_bx = (
        df_bx[df_bx["__contract__"] != ""]
        .sort_values(by="__contract__", ascending=True)
        .copy()
    )

    # 붙여넣기용 DF
    values_df = df_bx.drop(columns=["__contract__"]).fillna("").astype(str)
    values_list = values_df.values.tolist()

    # ===========================
    # 🔥 계좌 관련 컬럼 처리
    # ===========================
    TARGET_COLS = {"계좌번호", "수수료출금계좌"}
    target_indexes = []

    for i, col in enumerate(values_df.columns):
        if str(col).strip() in TARGET_COLS:
            target_indexes.append(i)

    for row in values_list:
        for idx in target_indexes:
            s = row[idx].strip()
            if s == "":
                continue

            # 혹시 남아있을 수 있는 지수표기/소수 제거
            if "E+" in s or "e+" in s:
                s = format(int(float(s)), "d")
            if s.endswith(".0"):
                s = s[:-2]

            # ✅ 최종: 무조건 텍스트
            row[idx] = "'" + s
    # ===========================

    sorted_contracts = df_bx["__contract__"].tolist()
    print(f"✅ 삼성증권 원본 데이터 행 수(계약번호 기준): {len(sorted_contracts)}")

    return values_list, sorted_contracts
# ===========================
# 4) parkpark 삼성_DATA 기존 비고 맵 만들기
# ===========================
def build_remark_map(ws):
    """
    삼성_DATA 시트에서
    - A열 비고
    - E열 계약번호 (실제 열 위치: E)
    를 읽어서 {계약번호: 비고} 맵 생성
    """
    xlUp = -4162

    # 계약번호 열은 E열(5)
    last_row = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row
    if last_row < DST_START_ROW:
        print("ℹ 삼성_DATA 기존 데이터가 거의 없습니다. 비고 맵은 빈 상태로 시작합니다.")
        return {}, [], []

    remark_map = {}
    old_contracts = []
    old_remarks = []

    # A~E까지만 읽어도 충분 (비고/계약번호만)
    rng = ws.Range(ws.Cells(DST_START_ROW, 1), ws.Cells(last_row, 5)).Value  # (rows x 5)

    for r in rng:
        remark = r[0]  # A
        contract = norm_contract(r[4])  # E
        if contract == "":
            continue
        remark_map[contract] = "" if remark is None else remark
        old_contracts.append(contract)
        old_remarks.append("" if remark is None else remark)

    print(f"📝 기존 삼성_DATA 비고 보유 계약 수: {len(remark_map)}")
    return remark_map, old_contracts, old_remarks



# ===========================
# 5) parkpark에 쓰기(비고 매칭 포함)
# ===========================
def write_to_parkpark(sorted_rows, sorted_contracts):
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    wb = None

    try:
        try:
            excel.ScreenUpdating = False
            excel.DisplayAlerts = False
            excel.EnableEvents = False
            excel.Calculation = -4135  # xlCalculationManual
        except Exception:
            pass

        print("📘 parkpark 파일 여는 중...")
        wb = excel.Workbooks.Open(PARKPARK_FILE, False, False, None, PASSWORD)
        ws = wb.Worksheets(SHEET_DST)

        # 1) 기존 비고 맵
        remark_map, old_contracts, _ = build_remark_map(ws)

        # 2) 변화 체크(로그)
        new_set = set(sorted_contracts)
        old_set = set(old_contracts)
        removed = sorted(old_set - new_set)
        added = sorted(new_set - old_set)
        print(f"🔍 변경 감지: 해지(사라짐) {len(removed)}명, 신규(추가) {len(added)}명")

        # 3) 붙여넣기 전에 기존 영역 비우기
        row_count = len(sorted_rows)
        if row_count == 0:
            print("⚠ 붙여넣을 데이터가 없습니다. 종료합니다.")
            return

        xlUp = -4162
        last_row = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row
        if last_row < DST_START_ROW:
            last_row = DST_START_ROW

        print("🧹 삼성_DATA 기존 데이터(A~X) 비우는 중...")
        ws.Range(
            ws.Cells(DST_START_ROW, 1),
            ws.Cells(max(last_row, DST_START_ROW + row_count + 200), 24)
        ).ClearContents()

        # 4) 비고(A열) 재구성
        remarks_to_write = [remark_map.get(c, "") for c in sorted_contracts]

        print("📥 비고(A열) 붙여넣기...")
        ws.Range(
            ws.Cells(DST_START_ROW, 1),
            ws.Cells(DST_START_ROW + row_count - 1, 1)
        ).Value = tuple((v,) for v in remarks_to_write)

        # ===========================
        # ✅ 계좌번호: 지수표기 방지 (텍스트 서식 + 값 강제 텍스트)
        # ===========================
        account_rel_idx = None
        for i in range(PASTE_COLS):
            col_name = ws.Cells(1, DST_START_COL + i).Value
            col_name = "" if col_name is None else str(col_name).strip()
            if col_name == "계좌번호":
                account_rel_idx = i
                break

        if account_rel_idx is not None:
            excel_col = DST_START_COL + account_rel_idx

            # 1) 붙여넣기 전에 해당 열을 "텍스트"로 강제
            ws.Range(
                ws.Cells(DST_START_ROW, excel_col),
                ws.Cells(DST_START_ROW + row_count - 1, excel_col)
            ).NumberFormat = "@"

            # 2) 값도 문자열로 강제 (앞에 ' 붙이면 엑셀이 무조건 텍스트로 처리)
            for r in sorted_rows:
                v = r[account_rel_idx]
                s = "" if v is None else str(v).strip()
                r[account_rel_idx] = "'" + s if s else ""
        # ===========================

        # 5) 고객데이터(B~X) 붙여넣기
        print("📥 고객데이터(B~X) 붙여넣기...")
        ws.Range(
            ws.Cells(DST_START_ROW, DST_START_COL),
            ws.Cells(DST_START_ROW + row_count - 1, DST_START_COL + PASTE_COLS - 1)
        ).Value = tuple(tuple(r) for r in sorted_rows)

        # 6) 확인 로그
        print("🔎 확인:")
        print("   - 삼성_DATA!A6(비고) =", ws.Cells(DST_START_ROW, 1).Value)
        print("   - 삼성_DATA!E6(계약번호) =", ws.Cells(DST_START_ROW, 5).Value)

        print("💾 parkpark 저장 중...")
        com_call_with_retry(lambda: wb.Save(), name="wb.Save")
        print("💾 parkpark 저장 완료.")
        time.sleep(0.4)  # Save 직후 Close 충돌 방지
        

        print("📕 워크북 닫는 중...")
        com_call_with_retry(lambda: wb.Close(False), name="wb.Close")
        wb = None
        print("📕 워크북 닫기 완료.")
    finally:
    # 엑셀 환경 복구
        try:
            excel.Calculation = -4105  # xlCalculationAutomatic
        except Exception:
            pass
    try:
        excel.EnableEvents = True
    except Exception:
        pass
    try:
        excel.ScreenUpdating = True
    except Exception:
        pass
    try:
        excel.DisplayAlerts = True
    except Exception:
        pass

    # 남아있으면 닫기 재시도
    try:
        if wb is not None:
            com_call_with_retry(lambda: wb.Close(False), name="finally wb.Close")
    except Exception:
        pass

    # Quit도 재시도
    try:
        com_call_with_retry(lambda: excel.Quit(), name="excel.Quit")
    except Exception:
        pass

    del excel
    gc.collect()
    print("📁 엑셀 종료")

# ===========================
# 6) main
# ===========================
def main():
    src = find_latest_source_file()
    rows, contracts = read_and_sort_source(src)
    write_to_parkpark(rows, contracts)

if __name__ == "__main__":
    main()