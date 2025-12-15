import win32com.client as win32
from datetime import datetime, date
import gc

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

    # 속도 옵션
    try:
        excel.ScreenUpdating = False
        excel.DisplayAlerts = False
    except Exception:
        pass

    wb = None
    ws_src = None
    ws_dst = None
    used = None

    try:
        wb = excel.Workbooks.Open(CUSTOMER_FILE, False, False, None, PASSWORD)
        ws_src = wb.Worksheets(SHEET_SRC)

    
        used = ws_src.UsedRange
        data = used.Value

        rows = [list(r) for r in data]

        # 0행: 헤더
        raw_header = rows[0]
        header = [norm(c) for c in raw_header]
        body = rows[1:]

        # ===== 필수 컬럼 index =====
        def find_col(name):
            for i, c in enumerate(header):
                if c == name:
                    return i
            raise RuntimeError(f"'{name}' 컬럼을 찾지 못했습니다. 헤더: {header}")

        idx_code = find_col("상품")
        idx_date = find_col("계약일자")

        # ===== 상품코드 필터링: 1/4/5, 001/004/005 =====
        filtered = []
        for row in body:
            if all(norm(c) == "" for c in row):
                continue

            code = norm(row[idx_code]).replace(".0", "")
            if code in ("1", "4", "5", "001", "004", "005"):
                filtered.append(row)

        if not filtered:
            print("⚠ 필터 결과가 없습니다. 종료.")
            return

        # ===== 계약일자 오름차순 정렬 =====
        def key_date(row):
            v = row[idx_date]
            if isinstance(v, (datetime, date)):
                return v
            s = norm(v)
            if s == "":
                return datetime.max
            for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%Y%m%d"):
                try:
                    return datetime.strptime(s, fmt)
                except ValueError:
                    continue
            return datetime.max

        filtered.sort(key=key_date)


        # ===== NH_DATA_1 작성 =====
        ws_dst = wb.Worksheets(SHEET_DST)

        print("🧹 NH_DATA_1 비우는 중...")
        ws_dst.Range("A1:AZ50000").ClearContents()

        # 헤더 1행 그대로 복사

        col_count = len(raw_header)
        for j, val in enumerate(raw_header, start=1):
            ws_dst.Cells(1, j).Value = val

        # 데이터 행 복사
        print("📥 행 단위 붙여넣기 시작...")
        for i, row in enumerate(filtered, start=2):
            if len(row) < col_count:
                row_fixed = row + [""] * (col_count - len(row))
            else:
                row_fixed = row[:col_count]

            dest = ws_dst.Range(
                ws_dst.Cells(i, 1),
                ws_dst.Cells(i, col_count)
            )
            dest.Value = (tuple(row_fixed),)  # 2차원 튜플로 넣기

            if (i - 1) % 50 == 0:
                print(f"   → {i-1}행 완료")

        print("🎉 모든 행 복사 완료!")


        wb.Save()
        print("💾 저장 완료!")

    finally:
        # COM 객체들 먼저 참조 해제
        try:
            del used
        except Exception:
            pass
        try:
            del ws_src
        except Exception:
            pass
        try:
            del ws_dst
        except Exception:
            pass

        gc.collect()  # 참조 정리

        # 워크북 닫기
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass

        # 엑셀 종료
        try:
            excel.ScreenUpdating = True
        except Exception:
            pass

        try:
            excel.Quit()
        except Exception:
            pass

        del wb
        del excel
        gc.collect()

        print("📁 엑셀 종료 (리소스 정리 완료)")


if __name__ == "__main__":
    main()