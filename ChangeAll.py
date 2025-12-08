import os
import pandas as pd

# ------------------------------
# 1. 기본 설정
# ------------------------------
download_path = r"C:\Users\pc\Downloads"

FILE_PREFIX = "file_066"  # 증권사 엑셀 접두사

# 결과 파일 저장 위치 (원하는 경로로 바꿔도 됨)
OUTPUT_FILE = r"C:\Users\pc\OneDrive - 주식회사 플레인바닐라\LEEJAEWOOK의 파일 - 플레인바닐라 업무\Customer\고객data\FOK_DATA_자동생성.xlsx"


# ------------------------------
# 2. xls → xlsx 변환 함수
# ------------------------------
def convert_xls_to_xlsx(xls_path: str) -> str:
    import win32com.client as win32

    if not os.path.exists(xls_path):
        raise FileNotFoundError(f"xls 파일을 찾을 수 없습니다: {xls_path}")

    excel = win32.Dispatch("Excel.Application")
    try:
        wb = excel.Workbooks.Open(xls_path)
        xlsx_path = os.path.splitext(xls_path)[0] + ".xlsx"
        wb.SaveAs(xlsx_path, FileFormat=51)  # 51 = xlsx
        wb.Close()
    finally:
        excel.Quit()

    print(f"[변환 완료] {xls_path} -> {xlsx_path}")
    return xlsx_path


# ------------------------------
# 3. 최신 증권사 xls 찾기 → xlsx 변환 → pandas로 가공
# ------------------------------
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

# 증권사 데이터 읽기
df = pd.read_excel(latest_xlsx)
print("증권사 컬럼 목록:", list(df.columns))

# '예수금', '평가금액' 컬럼 삭제
drop_cols = ["예수금", "평가금액"]
df = df.drop(columns=drop_cols, errors="ignore")
print("삭제 후 컬럼 목록:", list(df.columns))

# NaN → 빈 문자열
df = df.fillna("")

# ------------------------------
# 4. 결과를 새 엑셀 파일로 저장 (FOK_DATA 시트)
# ------------------------------
# 같은 이름 파일이 있으면 덮어쓰기
os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="FOK_DATA", index=False)

print("✅ 완료: 결과 파일 생성 ->", OUTPUT_FILE)