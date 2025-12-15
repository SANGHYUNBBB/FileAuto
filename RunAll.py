import os
import sys
import subprocess
from datetime import datetime

SCRIPTS = [
    "FokChange.py",
    "NhChange.py",
    "NH_1_Change.py",
    "KiwoomCount.py",
    "Han.py",
    "SamChange.py",
]

WORKDIR = r"C:\Code"
STOP_ON_ERROR = True

def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def run_one(script: str) -> int:
    script_path = os.path.join(WORKDIR, script)

    if not os.path.exists(script_path):
        print(f"❌ NOT FOUND: {script_path}")
        return 2

    print(f"\n▶ START: {script}  ({now()})")

    # ✅ 핵심: UTF-8 강제 + 이모지 출력 에러 방지
    # -X utf8 : 파이썬 UTF-8 모드
    # PYTHONIOENCODING=utf-8 : stdout/stderr 강제
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"

    cmd = [
        sys.executable,
        "-X", "utf8",
        script_path,
    ]

    # 출력은 "그대로 콘솔에 흘려보내기" (프린트 진행내역 보임)
    p = subprocess.Popen(
        cmd,
        cwd=WORKDIR,
        env=env,
        text=True,
        encoding="utf-8",
        errors="replace",   # 혹시라도 인코딩 문제 생기면 깨진 글자 대신 대체문자
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,  # 에러도 같은 스트림으로 합쳐서 순서대로 보이게
        bufsize=1,
    )

    # 실시간 출력
    assert p.stdout is not None
    for line in p.stdout:
        print(line, end="")

    p.wait()

    if p.returncode == 0:
        print(f"✅ OK  : {script}  ({now()})")
    else:
        print(f"\n---- ERROR (code={p.returncode}) ----")
        print(f"❌ FAIL: {script}")

    return p.returncode

def main():
    print("=== RunAll START ===")
    print(f"WORKDIR={WORKDIR}")
    print(f"STOP_ON_ERROR={STOP_ON_ERROR}")

    results = []
    for s in SCRIPTS:
        code = run_one(s)
        results.append((s, code))
        if code != 0 and STOP_ON_ERROR:
            break

    print("\n=== RunAll SUMMARY ===")
    for s, code in results:
        status = "OK " if code == 0 else "FAIL"
        print(f"- {status} | {s} | code={code}")

if __name__ == "__main__":
    main()