
import subprocess
import time

def kill_all_excel():
    """
    실행 중인 모든 Excel 프로세스 강제 종료
    (사용자가 열어둔 Excel 포함 전부 종료)
    """
    subprocess.run(
        ["taskkill", "/F", "/IM", "EXCEL.EXE"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )
    time.sleep(1)

if __name__ == "__main__":
    kill_all_excel()
