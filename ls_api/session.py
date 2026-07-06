import win32com.client
import pythoncom
import asyncio
import datetime

from ls_api.config import ID, MOUI_ID_PSWD, CERT


class XASessionEventHandler:
    login_state = 0

    def OnLogin(self, code, msg):
        if code == "0000":
            print("Login success")
            XASessionEventHandler.login_state = 1
        else:
            print("Login failed")


_inst_session = None


async def xingLogin():
    global _inst_session
    pythoncom.CoInitialize()
    _inst_session = win32com.client.DispatchWithEvents("XA_Session.XASession", XASessionEventHandler)
    _inst_session.ConnectServer("demo.ls-sec.co.kr", 20001)
    _inst_session.Login(ID, MOUI_ID_PSWD, '', 0, 0)
    ctime = datetime.datetime.now()
    while XASessionEventHandler.login_state == 0:
        pythoncom.PumpWaitingMessages()
        await asyncio.sleep(0.005)
        if datetime.datetime.now() > ctime + datetime.timedelta(seconds=20):
            break
    if XASessionEventHandler.login_state == 0:
        raise TimeoutError("xingAPI 로그인 실패 또는 응답 없음 (20초 초과)")
    XASessionEventHandler.login_state = 0
    return 'Login success'


async def xingAccount(digit):
    global _inst_session
    if _inst_session is None:
        raise RuntimeError("로그인이 먼저 필요합니다. xingLogin()을 호출하세요.")
    num_account = _inst_session.GetAccountListCount()
    for i in range(num_account):
        num = _inst_session.GetAccountList(i)
        if num[-4:] == digit:
            print(num)
            return num
    raise ValueError(f"계좌번호 끝 4자리 '{digit}'에 해당하는 계좌를 찾을 수 없습니다.")
