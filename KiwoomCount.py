import os
import pandas as pd
import win32com.client as win32
import gc
from datetime import datetime

# ======================
# 1. 기본 설정
# ======================
DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
LIST_PREFIX = "Excel_List_"

PASSWORD = "nilla17()"

HEADER_ROW = 5
SHEET_KIWOOM = "키움_DATA_"


DATE_FMT_STR = "%Y.%m.%d"

INVEST_COL_FIXED = 13  # M열 (투자성향 고정)

# ===== 키움_DATA_ 헤더명 =====
COL_NO = "NO."
COL_GUBUN = "구분"
COL_PLATFORM = "플랫폼"
COL_NAME = "이름"
COL_ACCT = "계좌(계약)번호"
COL_TYPE = "유형"
COL_CONTRACT = "계약일"
COL_CONTRACT_END = "계약종료일"
COL_BALANCE = "잔고"
COL_BIRTH = "생년"
COL_PHONE = "전화번호"
COL_EMAIL = "이메일"

# ===== 증권사 파일 컬럼명 =====
BROKER_COL_NAME = "이름"
BROKER_COL_ACCT = "계약계좌번호"
BROKER_COL_TYPE = "계좌유형"
BROKER_COL_BIRTH = "생년월일"
BROKER_COL_INVEST = "투자유형"
BROKER_COL_PHONE = "연락처"
BROKER_COL_EMAIL = "이메일"
BROKER_COL_CONTRACT = "계약일"

# ======================
# 2. 유틸 함수
# ======================
def format_phone_korea(raw):
    """
    전화번호를 010-0000-0000 형식으로 변환
    """
    digits = norm_digits(raw)

    if not digits:
        return ""

    # 앞자리가 0이 아니면 0 보정
    if not digits.startswith("0"):
        digits = "0" + digits

    # 휴대폰 번호 (11자리)만 포맷
    if len(digits) == 11:
        return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"
    elif len(digits) == 10:  # 예외 케이스
        return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
    else:
        # 이상한 길이는 그냥 원본 반환
        return digits
def get_onedrive_path():
    # 회사 OneDrive 우선
    for env in ("OneDriveCommercial", "OneDrive"):
        p = os.environ.get(env)
        if p and os.path.exists(p):
            return p
    raise EnvironmentError("OneDrive 경로를 찾을 수 없습니다.")

ONEDRIVE_ROOT = get_onedrive_path()

CUSTOMER_FILE = os.path.join(
    ONEDRIVE_ROOT,
    "LEEJAEWOOK의 파일 - 플레인바닐라 업무",
    "Customer",
    "고객data",
    "고객data_v101.xlsx",
)
def norm_col(s: str) -> str:
    s = str(s)
    for t in ["_x000D_", "\r", "\n", " "]:
        s = s.replace(t, "")
    return s.strip()


def norm_digits(s) -> str:
    if s is None:
        return ""
    return "".join(ch for ch in str(s) if ch.isdigit())


def clean_cell(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


def parse_contract_date(s: str) -> datetime:
    return datetime.strptime(s, DATE_FMT_STR)


def add_one_year(dt: datetime) -> datetime:
    try:
        return dt.replace(year=dt.year + 1)
    except ValueError:
        return dt.replace(month=2, day=28, year=dt.year + 1)


def cell_text(ws, r, c) -> str:
    try:
        return str(ws.Cells(r, c).Text or "").strip()
    except Exception:
        return str(ws.Cells(r, c).Value or "").strip()


def map_broker_type_to_customer(t: str) -> str:
    return "일반" if (t or "").strip() == "위탁종합" else (t or "").strip()


def make_customer_key(name, acct, cust_type):
    return ((name or "").strip(), norm_digits(acct), (cust_type or "").strip())


def make_broker_key(name, acct, acct_type):
    return ((name or "").strip(), norm_digits(acct), map_broker_type_to_customer(acct_type))


def set_cell_value_safe(ws, addr: str, value: str):
    rng = ws.Range(addr)
    if rng.MergeCells:
        rng.MergeArea.Cells(1, 1).Value = value
    else:
        rng.Value = value


def get_last_row(ws, col_idx: int) -> int:
    return ws.Cells(ws.Rows.Count, col_idx).End(-4162).Row


def find_last_kiwoom_row(ws, start_row, end_row, platform_col, name_col):
    for r in range(end_row, start_row - 1, -1):
        platform = cell_text(ws, r, platform_col)
        name = cell_text(ws, r, name_col)
        if name and "키움" in platform:
            return r
    return None


# ======================
# 3. 증권사 파일 로드
# ======================
def load_broker_df() -> pd.DataFrame:
    files = [
        f for f in os.listdir(DOWNLOAD_DIR)
        if f.startswith(LIST_PREFIX) and f.lower().endswith((".xls", ".xlsx"))
    ]
    if not files:
        raise FileNotFoundError("증권사 파일 없음")

    files.sort(key=lambda f: os.path.getmtime(os.path.join(DOWNLOAD_DIR, f)), reverse=True)
    path = os.path.join(DOWNLOAD_DIR, files[0])

    if path.lower().endswith(".xls"):
        excel = win32.DispatchEx("Excel.Application")
        wb = excel.Workbooks.Open(path)
        new_path = path.replace(".xls", ".xlsx")
        wb.SaveAs(new_path, FileFormat=51)
        wb.Close()
        excel.Quit()
        path = new_path

    df = pd.read_excel(path)
    df.columns = [norm_col(c) for c in df.columns]
    return df


def build_broker_maps(df: pd.DataFrame):
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
# 4. 메인 로직
# ======================
def update_kiwoom_data():
    df_broker = load_broker_df()
    broker_keys, broker_lookup = build_broker_maps(df_broker)

    contract_dt = datetime.today()
    end_dt = add_one_year(contract_dt)

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.ScreenUpdating = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(CUSTOMER_FILE, False, False, None, PASSWORD)
    ws = wb.Worksheets(SHEET_KIWOOM)

    # 헤더 매핑
    header_map = {}
    for c in range(1, 80):
        v = ws.Cells(HEADER_ROW, c).Value
        if v:
            header_map[str(v).strip()] = c

    data_start = HEADER_ROW + 1
    last_row = get_last_row(ws, header_map[COL_NO])

    last_kiwoom_row = find_last_kiwoom_row(
        ws, data_start, last_row, header_map[COL_PLATFORM], header_map[COL_NAME]
    )

    last_no = int(cell_text(ws, last_kiwoom_row, header_map[COL_NO]))
    next_no = last_no + 1

    # 기존 키 생성
    existing = {}
    for r in range(data_start, last_row + 1):
        k = make_customer_key(
            cell_text(ws, r, header_map[COL_NAME]),
            cell_text(ws, r, header_map[COL_ACCT]),
            cell_text(ws, r, header_map[COL_TYPE]),
        )
        if all(k):
            existing[k] = r

    existing_keys = set(existing.keys())
    new_keys = broker_keys - existing_keys

    new_names = []
    canceled_names = []

    # 해지 처리
    for k, r in existing.items():
        gubun = cell_text(ws, r, header_map[COL_GUBUN])
        if gubun != "해지" and k not in broker_keys:
            ws.Cells(r, header_map[COL_GUBUN]).Value = "해지"
            canceled_names.append(k[0])

    insert_row = last_kiwoom_row + 1

    # 신규 추가
    for k in sorted(new_keys):
        r = broker_lookup[k]

        ws.Rows(insert_row).Insert()
        ws.Cells(insert_row, header_map[COL_NO]).Value = next_no
        next_no += 1

        ws.Cells(insert_row, header_map[COL_GUBUN]).Value = "신규"
        ws.Cells(insert_row, header_map[COL_PLATFORM]).Value = "키움증권"
        ws.Cells(insert_row, header_map[COL_NAME]).Value = k[0]
        ws.Cells(insert_row, header_map[COL_ACCT]).Value = r.get(BROKER_COL_ACCT)
        ws.Cells(insert_row, header_map[COL_TYPE]).Value = map_broker_type_to_customer(r.get(BROKER_COL_TYPE))
        broker_contract_raw = r.get(BROKER_COL_CONTRACT)

        if pd.notna(broker_contract_raw):
            # 엑셀 datetime / 문자열 모두 대응
            if isinstance(broker_contract_raw, datetime):
                contract_dt = broker_contract_raw
            else:
                contract_dt = datetime.strptime(str(broker_contract_raw)[:10], "%Y.%m.%d")
        else:
            # 혹시 없으면 오늘 날짜 fallback
            contract_dt = datetime.today()

        end_dt = add_one_year(contract_dt)

        ws.Cells(insert_row, header_map[COL_CONTRACT]).Value = contract_dt.strftime("%Y.%m.%d")
        ws.Cells(insert_row, header_map[COL_CONTRACT_END]).Value = end_dt.strftime("%Y.%m.%d")
        # 생년
        birth = norm_digits(r.get(BROKER_COL_BIRTH))
        if COL_BIRTH in header_map and len(birth) >= 2:
            ws.Cells(insert_row, header_map[COL_BIRTH]).Value = birth[:2]

        # 투자성향 (M열 고정)
        ws.Cells(insert_row, INVEST_COL_FIXED).Value = clean_cell(r.get(BROKER_COL_INVEST))
        # 전화번호 (010-0000-0000 포맷)
        phone = format_phone_korea(r.get(BROKER_COL_PHONE))

        if COL_PHONE in header_map:
            ws.Cells(insert_row, header_map[COL_PHONE]).Value = phone
        # 이메일
        if COL_EMAIL in header_map:
            ws.Cells(insert_row, header_map[COL_EMAIL]).Value = clean_cell(r.get(BROKER_COL_EMAIL))

        if COL_BALANCE in header_map:
            ws.Cells(insert_row, header_map[COL_BALANCE]).Value = ""

        new_names.append(k[0])
        insert_row += 1

    # A1 / A2 기록
    set_cell_value_safe(ws, "A1", "\n".join(new_names))
    set_cell_value_safe(ws, "A2", "\n".join(canceled_names))

    wb.Save()
    wb.Close(SaveChanges=False)
    excel.Quit()
    gc.collect()

    print("신규:", new_names)
    print("해지:", canceled_names)


# ======================
# 5. 실행
# ======================
if __name__ == "__main__":
    update_kiwoom_data()
