import win32com.client
import pythoncom
import asyncio
import datetime

from ls_api.config import MOUI_PSWD


class XAQueryEventHandlerT0441:
    query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandlerT0441.query_state = 1


class XAQueryEventHandlerCFOBQ10500:
    query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandlerCFOBQ10500.query_state = 1


class XAQueryEventHandlerT0434:
    query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandlerT0434.query_state = 1


# 각 TR별 inst 캐시 (처음 호출 시 한 번만 생성)
_inst_t0441 = None
_inst_cfobq10500 = None
_inst_t0434 = None


async def _wait(handler_cls, timeout=10):
    ctime = datetime.datetime.now()
    while handler_cls.query_state == 0:
        pythoncom.PumpWaitingMessages()
        await asyncio.sleep(0.005)
        if datetime.datetime.now() > ctime + datetime.timedelta(seconds=timeout):
            break
    if handler_cls.query_state == 0:
        raise TimeoutError(f"xingAPI 응답 없음: {handler_cls.__name__} ({timeout}초 초과)")
    handler_cls.query_state = 0


async def xingJango(acnt):
    global _inst_cfobq10500
    if _inst_cfobq10500 is None:
        _inst_cfobq10500 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerCFOBQ10500)
        _inst_cfobq10500.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\CFOBQ10500.res"
    _inst_cfobq10500.SetFieldData("CFOBQ10500InBlock1", "RecCnt", 0, 1)
    _inst_cfobq10500.SetFieldData("CFOBQ10500InBlock1", "AcntNo", 0, acnt)
    _inst_cfobq10500.SetFieldData("CFOBQ10500InBlock1", "Pwd", 0, MOUI_PSWD)
    _inst_cfobq10500.Request(0)
    await _wait(XAQueryEventHandlerCFOBQ10500)
    return _inst_cfobq10500.GetFieldData("CFOBQ10500OutBlock2", "MnyOrdAbleAmt", 0)


async def xingCheckAccount(acnt, codeName):
    global _inst_t0441
    if _inst_t0441 is None:
        _inst_t0441 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerT0441)
        _inst_t0441.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\t0441.res"
    _inst_t0441.SetFieldData("t0441InBlock", "accno", 0, acnt)
    _inst_t0441.SetFieldData("t0441InBlock", "passwd", 0, MOUI_PSWD)
    _inst_t0441.SetFieldData("t0441InBlock", "cts_expcode", 0, codeName)
    _inst_t0441.Request(0)
    await _wait(XAQueryEventHandlerT0441)
    count = _inst_t0441.GetBlockCount("t0441OutBlock1")
    if count == 0:
        _numOfContract = 0
        _enteredDirection = '0'
        _tdtsunik = _inst_t0441.GetFieldData("t0441OutBlock1", "tdtsunik", 0)
        _medosu = _inst_t0441.GetFieldData("t0441OutBlock1", "medosu", 0)
        _avgPrice = _inst_t0441.GetFieldData("t0441OutBlock1", "pamt", 0)
    else:
        _numOfContract = int(_inst_t0441.GetFieldData("t0441OutBlock1", "jqty", 0))
        _enteredDirection = _inst_t0441.GetFieldData("t0441OutBlock1", "medocd", 0)
        _tdtsunik = _inst_t0441.GetFieldData("t0441OutBlock", "tdtsunik", 0)
        _medosu = _inst_t0441.GetFieldData("t0441OutBlock1", "medosu", 0)
        _avgPrice = _inst_t0441.GetFieldData("t0441OutBlock1", "pamt", 0)
    return _numOfContract, _enteredDirection, _tdtsunik, _medosu, _avgPrice


async def xingResult(acnt):
    global _inst_t0441
    if _inst_t0441 is None:
        _inst_t0441 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerT0441)
        _inst_t0441.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\t0441.res"
    _inst_t0441.SetFieldData("t0441InBlock", "accno", 0, acnt)
    _inst_t0441.SetFieldData("t0441InBlock", "passwd", 0, MOUI_PSWD)
    _inst_t0441.Request(0)
    await _wait(XAQueryEventHandlerT0441)
    return _inst_t0441.GetFieldData("t0441OutBlock", "tdtsunik", 0)


async def xingCheckOrder(acnt, codeName, option=False):
    global _inst_t0434
    if _inst_t0434 is None:
        _inst_t0434 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerT0434)
        _inst_t0434.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\t0434.res"
    _inst_t0434.SetFieldData("t0434InBlock", "accno", 0, acnt)
    _inst_t0434.SetFieldData("t0434InBlock", "passwd", 0, MOUI_PSWD)
    _inst_t0434.SetFieldData("t0434InBlock", "expcode", 0, codeName)
    _inst_t0434.SetFieldData("t0434InBlock", "chegb", 0, '2')
    _inst_t0434.SetFieldData("t0434InBlock", "sortgb", 0, '1')
    _inst_t0434.Request(0)
    await _wait(XAQueryEventHandlerT0434)
    count = _inst_t0434.GetBlockCount("t0434OutBlock1")
    _uncheckedTrade = 0
    ordrem = 0
    ordno = 0
    for i in range(count):
        _uncheckedTrade += int(_inst_t0434.GetFieldData("t0434OutBlock1", "qty", i))
        ordrem += int(_inst_t0434.GetFieldData("t0434OutBlock1", "ordrem", i))
        ordno = int(_inst_t0434.GetFieldData("t0434OutBlock1", "ordno", i))
    if option == True:
        return ordrem, ordno
    else:
        return _uncheckedTrade


