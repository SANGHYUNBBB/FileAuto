import os
from pathlib import Path

def get_fixed_customer_path():
    """
    고객data_v101.xlsx 파일의 고정된 경로를 반환합니다.
    다양한 OneDrive 경로 구조를 지원합니다.
    """
    import getpass
    
    # 현재 사용자 이름 가져오기
    current_user = getpass.getuser()
    
    # 가능한 경로 패턴들 (우선순위 순)
    base_onedrive = f"C:\\Users\\{current_user}\\OneDrive - 주식회사 플레인바닐라"
    
    possible_paths = [
        # 표준 경로 (xmfos 등 대부분 사용자)
        os.path.join(base_onedrive, "플레인바닐라 업무", "Customer", "고객data", "고객data_v101.xlsx"),
        # 현재 pc 사용자의 경로 (LEEJAEWOOK의 파일 포함)
        os.path.join(base_onedrive, "LEEJAEWOOK의 파일 - 플레인바닐라 업무", "Customer", "고객data", "고객data_v101.xlsx"),
    ]
    
    # 가능한 경로들을 확인
    for path in possible_paths:
        if os.path.exists(path):
            return path
    
    # 어떤 경로도 찾지 못한 경우
    error_msg = f"고객data 파일을 찾을 수 없습니다.\n사용자: {current_user}\n확인한 경로들:\n"
    for i, path in enumerate(possible_paths, 1):
        error_msg += f"  {i}. {path}\n"
    error_msg += "OneDrive 경로와 파일 구조를 확인하세요."
    
    raise FileNotFoundError(error_msg)

def find_customer_file():
    """이전 버전과의 호환성을 위한 함수 - 이제는 고정된 경로를 사용합니다."""
    return get_fixed_customer_path()