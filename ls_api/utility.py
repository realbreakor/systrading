from ls_api.config import BASE_DIR
import pickle
import os
import datetime


async def save_state(state, name):
    tmp_path = os.path.join(BASE_DIR, f"{name}_temp.pkl")
    final_path = os.path.join(BASE_DIR, f"{name}.pkl")
    with open(tmp_path, "wb") as f:
        pickle.dump(state, f)
    os.replace(tmp_path, final_path)


async def load_state(name):
    file_path = os.path.join(BASE_DIR, f"{name}.pkl")
    if os.path.exists(file_path):
        try:
            with open(f"{name}.pkl", "rb") as f:
                state = pickle.load(f)
            now = datetime.datetime.now()
            start = state["inTime"]
            end = state["outTime"]
            if start <= now <= end:
                print("현재 세션 → 상태 복구")
                return state
            else:
                print("세션 종료 → 새로 시작")
                return None
        except Exception as e:
            print("로드 실패:", e)
            return None


async def mailing(title, text, whom, send=0):
    pass