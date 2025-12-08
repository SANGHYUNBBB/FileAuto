import shutil
import os
import win32com.client.gencache

# 1) gen_py 캐시 폴더 삭제 시도
gen_py_dir = os.path.join(os.environ["LOCALAPPDATA"], "Temp", "gen_py")
print("gen_py 폴더:", gen_py_dir)

if os.path.exists(gen_py_dir):
    print("🔄 gen_py 폴더 삭제 중...")
    shutil.rmtree(gen_py_dir, ignore_errors=True)
else:
    print("gen_py 폴더가 없습니다. 건너뜁니다.")

# 2) pywin32 타입 라이브러리 캐시 재생성
print("♻ gencache 재생성 중...")
win32com.client.gencache.Rebuild()
print("✅ gencache 재생성 완료")