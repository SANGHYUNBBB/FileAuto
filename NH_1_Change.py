import win32com.client as win32
from datetime import datetime, date

CUSTOMER_FILE = r"C:\Users\pc\OneDrive - 주식회사 플레인바닐라\LEEJAEWOOK의 파일 - 플레인바닐라 업무\Customer\고객data\고객data_v101_parkpark.xlsx"
PASSWORD = "nilla17()"

SHEET_SRC = "NH_DATA"
SHEET_DST = "NH_DATA_1"


def norm(v):
    if v is None:
        return ""
    return str(v).replace("\r", "").replace("\n", "").strip()


def main():
    print("📘 parkpark 고객 파일 여는 중...")
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.ScreenUpdating = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(CUSTOMER_FILE, False, False, None, PASSWORD)

    try:
        ws_src = wb.Worksheets(SHEET_SRC)

        print("📖 NH_DATA UsedRange 읽는 중...")
        used = ws_src.UsedRange
        data = used.Value

        # tuple → list
        rows = [list(r) for r in data]

        # 헤더
        header = [norm(c) for c in rows[0]]
        body = rows[1:]

        # ===== 필수 컬럼 index 찾기 =====
        def find_col(name):
            for i, c in enumerate(header):
                if c == name:
                    return i
            raise RuntimeError(f"'{name}' 컬럼을 찾지 못했습니다. 헤더: {header}")

        idx_code = find_col("상품")
        idx_date = find_col("계약일자")

        # ===== 상품코드 필터링 =====
        filtered = []
        for row in body:
            if all(norm(c) == "" for c in row):  # 빈 행 스킵
                continue

            code = norm(row[idx_code]).replace(".0", "")
            if code in ("1", "4", "5", "001", "004", "005"):
                filtered.append(row)

        print(f"📊 필터링된 행 수: {len(filtered)}")

        # ===== 날짜 정렬 =====
        def get_date(v):
            v = v[idx_date]
            if isinstance(v, (datetime, date)):
                return v
            s = norm(v)
            if s == "":
                return datetime.max
            for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%Y%m%d"):
                try:
                    return datetime.strptime(s, fmt)
                except:
                    pass
            return datetime.max

        filtered.sort(key=get_date)
        print("📅 계약일자 오름차순 정렬 완료.")

        # ===== NH_DATA_1에 행 단위로 붙여넣기 =====
        ws_dst = wb.Worksheets(SHEET_DST)

        print("🧹 NH_DATA_1 비우는 중...")
        ws_dst.Range("A1:AZ50000").ClearContents()

        # 헤더 먼저 넣기
        ws_dst.Range("A1").Resize(1, len(header)).Value = header

        print("📥 행 단위 붙여넣기 시작...")

        for i, row in enumerate(filtered, start=2):
            # 엑셀 셀 갯수 맞추기
            row_fixed = row + [""] * (len(header) - len(row))
            ws_dst.Range(
                ws_dst.Cells(i, 1),
                ws_dst.Cells(i, len(header))
            ).Value = row_fixed

            # 진행상황
            if i % 50 == 0:
                print(f"   → {i-1}행 완료")

        print("🎉 모든 행 복사 완료!")
        print("🔎 NH_DATA_1!A2 =", ws_dst.Cells(2, 1).Value)

        wb.Save()
        print("💾 저장 완료!")

    finally:
        excel.ScreenUpdating = True
        wb.Close(False)
        excel.Quit()
        print("📁 엑셀 종료")


if __name__ == "__main__":
    main()