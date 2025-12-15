import os
import re
import pandas as pd
import win32com.client as win32
import gc

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


# ===========================
# 3) 삼성증권 파일 읽기 + 정렬
# ===========================
def read_and_sort_source(src_path: str):
    src_xlsx = convert_xls_to_xlsx(src_path)

    # header=0: 첫 행을 컬럼명으로 읽음
    df = pd.read_excel(
    src_xlsx,
    header=0,
    dtype={"계좌번호": str}
)

    # 실제로는 A~X까지 있을 텐데 우리는 B~X만 필요
    # pandas 기준 0-based로 B는 index 1
    if df.shape[1] < 24:
        print(f"⚠ 원본 컬럼 수가 예상보다 적습니다. 현재 컬럼 수={df.shape[1]}. 그래도 가능한 범위로 진행합니다.")

    df_bx = df.iloc[:, 1:1+PASTE_COLS].copy()  # B~X

    # 계약번호(E열)는 B~X 내부에서 4번째(0-based 3)
    contract_series = df_bx.iloc[:, CONTRACT_REL_IDX].map(norm_contract)

    df_bx["__contract__"] = contract_series

    # 완전 빈 행 제거(계약번호 없으면 제거)
    df_bx = df_bx[df_bx["__contract__"] != ""].copy()

    # 계약번호 오름차순 정렬
    df_bx = df_bx.sort_values(by="__contract__", ascending=True)

    # 붙여넣기용 값(2D list)
    values = df_bx.drop(columns=["__contract__"]).astype(object).where(pd.notnull(df_bx), "").drop(columns=["__contract__"], errors="ignore")
    # 위 라인이 복잡해질 수 있어 안전하게 재작성:
    values = df_bx.drop(columns=["__contract__"]).astype(object).where(pd.notnull(df_bx.drop(columns=["__contract__"])), "")

    sorted_contracts = df_bx["__contract__"].tolist()

    print(f"✅ 삼성증권 원본 데이터 행 수(계약번호 기준): {len(sorted_contracts)}")
    return values.values.tolist(), sorted_contracts


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

        # 3) 붙여넣기 전에 기존 B~X 영역 비우기 (A 비고는 우리가 다시 채울 거라 같이 비워도 됨)
        row_count = len(sorted_rows)
        if row_count == 0:
            print("⚠ 붙여넣을 데이터가 없습니다. 종료합니다.")
            return

        # 충분히 큰 범위 비우기(기존 데이터가 더 많았을 수도 있으니 넉넉히)
        xlUp = -4162
        last_row = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row
        if last_row < DST_START_ROW:
            last_row = DST_START_ROW

        print("🧹 삼성_DATA 기존 데이터(A~X) 비우는 중...")
        ws.Range(ws.Cells(DST_START_ROW, 1), ws.Cells(max(last_row, DST_START_ROW + row_count + 200), 24)).ClearContents()
        # 24열 = X

        # 4) 비고(A열) 재구성: 계약번호 기준으로 맵핑
        remarks_to_write = []
        for c in sorted_contracts:
            remarks_to_write.append(remark_map.get(c, ""))

        # 5) 실제 쓰기 (속도: 한 번에 Range 넣기)
        print("📥 비고(A열) 붙여넣기...")
        ws.Range(ws.Cells(DST_START_ROW, 1), ws.Cells(DST_START_ROW + row_count - 1, 1)).Value = tuple((v,) for v in remarks_to_write)

        print("📥 고객데이터(B~X) 붙여넣기...")
        ws.Range(
            ws.Cells(DST_START_ROW, DST_START_COL),
            ws.Cells(DST_START_ROW + row_count - 1, DST_START_COL + PASTE_COLS - 1)
        ).Value = tuple(tuple(r) for r in sorted_rows)

        # 6) 확인 로그
        print("🔎 확인:")
        print("   - 삼성_DATA!A6(비고) =", ws.Cells(DST_START_ROW, 1).Value)
        print("   - 삼성_DATA!E6(계약번호) =", ws.Cells(DST_START_ROW, 5).Value)

        wb.Save()
        print("💾 parkpark 저장 완료.")

        wb.Close(False)
        wb = None
        print("📕 워크북 닫기 완료.")

    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
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
# 6) main
# ===========================
def main():
    src = find_latest_source_file()
    rows, contracts = read_and_sort_source(src)
    write_to_parkpark(rows, contracts)

if __name__ == "__main__":
    main()