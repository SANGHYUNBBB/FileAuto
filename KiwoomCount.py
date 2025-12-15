import os
import pandas as pd
import win32com.client as win32
import gc
from datetime import datetime

# ======================
# 1. 기본 설정
# ======================
DOWNLOAD_DIR = r"C:\Users\pc\Downloads"
LIST_PREFIX = "Excel_List_"

CUSTOMER_FILE = r"C:\Users\pc\OneDrive - 주식회사 플레인바닐라\LEEJAEWOOK의 파일 - 플레인바닐라 업무\Customer\고객data\고객data_v101_parkpark.xlsx"
PASSWORD = "nilla17()"

HEADER_ROW = 5
SHEET_KIWOOM = "키움_DATA_"

DEFAULT_CONTRACT_DATE_STR = "2025.10.10"
DATE_FMT_STR = "%Y.%m.%d"

# ===== 키움_DATA_ 헤더명(5행과 100% 일치해야 함) =====
COL_NO = "NO."
COL_GUBUN = "구분"
COL_PLATFORM = "플랫폼"
COL_NAME = "이름"
COL_ACCT = "계좌(계약)번호"
COL_TYPE = "유형"
COL_CONTRACT = "계약일"
COL_CONTRACT_END = "계약종료일"
COL_BALANCE = "잔고"  # 잔고는 비움

# ===== 증권사 파일 컬럼명(pandas) =====
BROKER_COL_NAME = "이름"
BROKER_COL_ACCT = "계약계좌번호"
BROKER_COL_TYPE = "계좌유형"


# ======================
# 2. 공통 유틸
# ======================
def convert_xls_to_xlsx(path: str) -> str:
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
        wb.SaveAs(xlsx_path, FileFormat=51)  # 51=xlsx
        wb.Close()
    finally:
        excel.Quit()
    print(f"[변환 완료] {path} -> {xlsx_path}")
    return xlsx_path


def norm_col(s: str) -> str:
    s = str(s)
    for token in ["_x000D_", "\r", "\n", " "]:
        s = s.replace(token, "")
    return s.strip()


def get_latest_list_file() -> str:
    files = [
        f for f in os.listdir(DOWNLOAD_DIR)
        if f.startswith(LIST_PREFIX) and f.lower().endswith((".xls", ".xlsx"))
    ]
    if not files:
        raise FileNotFoundError(f"{DOWNLOAD_DIR} 에 '{LIST_PREFIX}*.xls(x)' 파일이 없습니다.")

    files.sort(key=lambda name: os.path.getmtime(os.path.join(DOWNLOAD_DIR, name)), reverse=True)
    latest = os.path.join(DOWNLOAD_DIR, files[0])
    print(f"📂 최신 Excel_List 파일: {latest}")
    return latest


def parse_contract_date(date_str: str) -> datetime:
    return datetime.strptime(date_str, DATE_FMT_STR)


def add_one_year(dt: datetime) -> datetime:
    try:
        return dt.replace(year=dt.year + 1)
    except ValueError:
        return dt.replace(month=2, day=28, year=dt.year + 1)


def get_last_row(ws, col_idx: int) -> int:
    # xlUp = -4162
    return ws.Cells(ws.Rows.Count, col_idx).End(-4162).Row


def cell_text(ws, r: int, c: int) -> str:
    """엑셀 표시값(Text) 기반 문자열"""
    try:
        return str(ws.Cells(r, c).Text or "").strip()
    except Exception:
        return str(ws.Cells(r, c).Value or "").strip()


def norm_digits(s) -> str:
    if s is None:
        return ""
    s = str(s).strip()
    if s.lower() == "nan":
        return ""
    return "".join(ch for ch in s if ch.isdigit())


def map_broker_type_to_customer(t: str) -> str:
    """증권사 계좌유형 -> 우리 유형(비교/저장용)"""
    t = (t or "").strip()
    if t == "위탁종합":
        return "일반"
    return t


def make_customer_key(name, acct, cust_type):
    """우리 키움_DATA_ 비교키: 이름+계좌+유형"""
    return (
        (name or "").strip(),
        norm_digits(acct),
        (cust_type or "").strip(),
    )


def make_broker_key(name, acct, acct_type):
    """증권사 비교키를 '우리 유형' 기준으로 맞춤(위탁종합->일반)"""
    return (
        (name or "").strip(),
        norm_digits(acct),
        map_broker_type_to_customer(acct_type),
    )


def set_cell_value_safe(ws, addr: str, value: str):
    """A1/A2가 병합셀이어도 좌상단에 기록"""
    rng = ws.Range(addr)
    if rng.MergeCells:
        rng.MergeArea.Cells(1, 1).Value = value
    else:
        rng.Value = value


def find_last_kiwoom_row(ws, start_row: int, end_row: int, platform_col: int, name_col: int, keyword="키움"):
    """
    플랫폼 셀에 keyword('키움') 포함 + 이름 존재하는 '마지막 행' 찾기
    """
    for r in range(end_row, start_row - 1, -1):
        platform_txt = str(ws.Cells(r, platform_col).Text or ws.Cells(r, platform_col).Value or "").strip()
        name_txt = str(ws.Cells(r, name_col).Text or ws.Cells(r, name_col).Value or "").strip()
        if name_txt and (keyword in platform_txt):
            return r
    return None


# ======================
# 3. 증권사 파일 로드
# ======================
def load_broker_df() -> pd.DataFrame:
    latest_path = get_latest_list_file()
    latest_xlsx = convert_xls_to_xlsx(latest_path)

    df = pd.read_excel(latest_xlsx)
    df.columns = [norm_col(c) for c in df.columns]
    return df


def build_broker_maps(df: pd.DataFrame):
    """broker_keys(set) + broker_lookup(dict: key->row_series)"""
    missing = [c for c in [BROKER_COL_NAME, BROKER_COL_ACCT, BROKER_COL_TYPE] if c not in df.columns]
    if missing:
        raise KeyError(f"증권사 파일에 필요한 컬럼이 없습니다: {missing}\n현재 컬럼: {df.columns.tolist()}")

    broker_keys = set()
    broker_lookup = {}

    for _, r in df.iterrows():
        k = make_broker_key(
            r.get(BROKER_COL_NAME),
            r.get(BROKER_COL_ACCT),
            r.get(BROKER_COL_TYPE),
        )
        if all(k):
            broker_keys.add(k)
            broker_lookup[k] = r

    return broker_keys, broker_lookup


# ======================
# 4. 키움_DATA_ 업데이트
# ======================
def update_kiwoom_data():
    df_broker = load_broker_df()
    broker_keys, broker_lookup = build_broker_maps(df_broker)

    contract_dt = parse_contract_date(DEFAULT_CONTRACT_DATE_STR)
    end_dt = add_one_year(contract_dt)

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    wb = None

    new_names = []
    canceled_names = []

    try:
        excel.ScreenUpdating = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(CUSTOMER_FILE, False, False, None, PASSWORD)
        ws = wb.Worksheets(SHEET_KIWOOM)

        # ✅ 1) 헤더(5행) 매핑
        max_scan_cols = 80
        header_map = {}
        for c in range(1, max_scan_cols + 1):
            v = ws.Cells(HEADER_ROW, c).Value
            if v is None:
                continue
            txt = str(v).strip()
            if txt:
                header_map[txt] = c

        required = [COL_NO, COL_GUBUN, COL_PLATFORM, COL_NAME, COL_ACCT, COL_TYPE, COL_CONTRACT, COL_CONTRACT_END]
        missing = [c for c in required if c not in header_map]
        if missing:
            raise KeyError(
                f"키움_DATA_ 시트 헤더에서 필요한 컬럼을 못 찾음: {missing}\n"
                f"현재 헤더 일부: {list(header_map.keys())[:40]}"
            )

        # ✅ 2) 시트 전체 마지막행 (NO 기준)
        sheet_last_row = get_last_row(ws, header_map[COL_NO])
        data_start_row = HEADER_ROW + 1
        print(f"✅ 시트 전체 데이터 범위: {data_start_row} ~ {sheet_last_row}")

        # ✅ 3) 키움 플랫폼 구간 마지막 고객 행 찾기 (한경미 같은 마지막 키움 고객)
        last_kiwoom_row = find_last_kiwoom_row(
            ws,
            start_row=data_start_row,
            end_row=sheet_last_row,
            platform_col=header_map[COL_PLATFORM],
            name_col=header_map[COL_NAME],
            keyword="키움"
        )
        if last_kiwoom_row is None:
            raise RuntimeError("키움 플랫폼(키움) 마지막 고객 행을 찾지 못했습니다. 플랫폼/이름 컬럼 값을 확인하세요.")

        print(f"✅ 키움 마지막 고객 행: {last_kiwoom_row} / 이름: {cell_text(ws, last_kiwoom_row, header_map[COL_NAME])}")

        # ✅ 4) 키움 마지막 NO
        last_no_txt = cell_text(ws, last_kiwoom_row, header_map[COL_NO])
        last_no = int(float(last_no_txt)) if last_no_txt else 0
        next_no = last_no + 1

        # ✅ 5) 우리 데이터 전체 key->row (해지 포함해서 '존재'로 취급)
        existing_key_to_row = {}
        for r in range(data_start_row, sheet_last_row + 1):
            name = cell_text(ws, r, header_map[COL_NAME])
            acct = cell_text(ws, r, header_map[COL_ACCT])
            cust_type = cell_text(ws, r, header_map[COL_TYPE])

            k = make_customer_key(name, acct, cust_type)
            if all(k):
                existing_key_to_row[k] = r

        existing_keys = set(existing_key_to_row.keys())

        # ✅ 6) 신규 = broker에는 있고, 우리에는 없는 키
        new_keys = broker_keys - existing_keys

        # ✅ 7) 해지 처리
        # - 이미 해지면 그대로
        # - 기존/신규 중 broker에 없으면 해지로 변경
        for k, row in existing_key_to_row.items():
            gubun = cell_text(ws, row, header_map[COL_GUBUN])

            if gubun == "해지":
                continue

            if gubun in ("기존", "신규") and k not in broker_keys:
                ws.Cells(row, header_map[COL_GUBUN]).Value = "해지"
                canceled_names.append(k[0])

        # ✅ 8) 신규 삽입 위치: 마지막 키움 고객 바로 아래
        insert_row = last_kiwoom_row + 1

        # ✅ 9) 신규 고객은 행 삽입으로 "연달아" 붙이기
        for k in sorted(list(new_keys), key=lambda x: (x[0], x[1], x[2])):
            r = broker_lookup.get(k)
            if r is None:
                continue

            ws.Rows(insert_row).Insert()  # shift down

            # NO 연속
            ws.Cells(insert_row, header_map[COL_NO]).Value = next_no
            next_no += 1

            ws.Cells(insert_row, header_map[COL_GUBUN]).Value = "신규"
            ws.Cells(insert_row, header_map[COL_PLATFORM]).Value = "키움증권"

            ws.Cells(insert_row, header_map[COL_NAME]).Value = k[0]
            ws.Cells(insert_row, header_map[COL_ACCT]).Value = str(r.get(BROKER_COL_ACCT, "") or "").strip()
            ws.Cells(insert_row, header_map[COL_TYPE]).Value = map_broker_type_to_customer(str(r.get(BROKER_COL_TYPE, "") or ""))

            ws.Cells(insert_row, header_map[COL_CONTRACT]).Value = contract_dt.strftime("%Y.%m.%d")
            ws.Cells(insert_row, header_map[COL_CONTRACT_END]).Value = end_dt.strftime("%Y.%m.%d")

            # 잔고 비움
            if COL_BALANCE in header_map:
                ws.Cells(insert_row, header_map[COL_BALANCE]).Value = ""

            new_names.append(k[0])
            insert_row += 1

        # ✅ 10) A1/A2는 키움_DATA_에만 기록 (덮어쓰기)
        set_cell_value_safe(ws, "A1", "\n".join(new_names))
        set_cell_value_safe(ws, "A2", "\n".join(canceled_names))

        wb.Save()
        wb.Close(SaveChanges=False)
        wb = None

        print(f"✅ 신규 추가: {len(new_names)}명 / 해지 처리: {len(canceled_names)}명")
        print("🔎 신규 이름 목록:", new_names)
        print("🔎 해지 이름 목록:", canceled_names)

        return new_names, canceled_names

    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
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


# ======================
# 5. main
# ======================
def main():
    update_kiwoom_data()


if __name__ == "__main__":
    main()
