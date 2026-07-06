import win32com.client
import pythoncom
import asyncio
import datetime
import time


class XAQueryEventHandlerT2111:
    query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandlerT2111.query_state = 1


class XAQueryEventHandlerT2112:
    query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandlerT2112.query_state = 1


class XAQueryEventHandlerT8426:
    query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandlerT8426.query_state = 1


class XAQueryEventHandlerT9943:
    query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandlerT9943.query_state = 1


class XAQueryEventHandlerT8435:
    query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandlerT8435.query_state = 1


class XAQueryEventHandlerT9944:
    query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandlerT9944.query_state = 1


class XAQueryEventHandlerMMDAQ91200:
    query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandlerMMDAQ91200.query_state = 1


# 각 TR별 inst 캐시 (처음 호출 시 한 번만 생성)
_inst_t2111 = None
_inst_t2112 = None
_inst_t8426 = None
_inst_t9943 = None
_inst_t8435 = None
_inst_t9944 = None
_inst_mmdaq91200 = None


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


async def xingCurrentPrice(codeName, option=None):
    global _inst_t2111
    if _inst_t2111 is None:
        _inst_t2111 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerT2111)
        _inst_t2111.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\t2111.res"
    _inst_t2111.SetFieldData("t2111InBlock", "focode", 0, codeName)
    _inst_t2111.Request(0)
    await _wait(XAQueryEventHandlerT2111)
    if option == 'yeprice':
        yeprice = _inst_t2111.GetFieldData("t2111OutBlock", "yeprice", 0)
        return yeprice if yeprice != '' else 0
    _currentPrice = _inst_t2111.GetFieldData("t2111OutBlock", "price", 0)
    _openPrice = _inst_t2111.GetFieldData("t2111OutBlock", "open", 0)
    _lastMonth = _inst_t2111.GetFieldData("t2111OutBlock", "lastmonth", 0)
    return _currentPrice, _openPrice, _lastMonth


async def xingBidAsk(codeName):
    global _inst_t2112
    if _inst_t2112 is None:
        _inst_t2112 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerT2112)
        _inst_t2112.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\t2112.res"
    _inst_t2112.SetFieldData("t2112InBlock", "shcode", 0, codeName)
    _inst_t2112.Request(0)
    await _wait(XAQueryEventHandlerT2112)
    ask_1 = _inst_t2112.GetFieldData("t2112OutBlock", "offerho1", 0)
    ask_vol_1 = _inst_t2112.GetFieldData("t2112OutBlock", "offerrem1", 0)
    ask_2 = _inst_t2112.GetFieldData("t2112OutBlock", "offerho2", 0)
    ask_vol_2 = _inst_t2112.GetFieldData("t2112OutBlock", "offerrem2", 0)
    ask_3 = _inst_t2112.GetFieldData("t2112OutBlock", "offerho3", 0)
    ask_vol_3 = _inst_t2112.GetFieldData("t2112OutBlock", "offerrem3", 0)
    ask_4 = _inst_t2112.GetFieldData("t2112OutBlock", "offerho4", 0)
    ask_vol_4 = _inst_t2112.GetFieldData("t2112OutBlock", "offerrem4", 0)
    ask_5 = _inst_t2112.GetFieldData("t2112OutBlock", "offerho5", 0)
    ask_vol_5 = _inst_t2112.GetFieldData("t2112OutBlock", "offerrem5", 0)
    return [(ask_1, ask_vol_1), (ask_2, ask_vol_2), (ask_3, ask_vol_3), (ask_4, ask_vol_4), (ask_5, ask_vol_5)]


async def xingUsdCode():
    global _inst_t8426
    if _inst_t8426 is None:
        _inst_t8426 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerT8426)
        _inst_t8426.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\t8426.res"
    _inst_t8426.SetFieldData("t8426InBlock", "dummy", 0, '')
    _inst_t8426.Request(0)
    await _wait(XAQueryEventHandlerT8426)
    count = _inst_t8426.GetBlockCount("t8426OutBlock")
    for i in range(count):
        _hname = _inst_t8426.GetFieldData("t8426OutBlock", "hname", i)
        _shcode = _inst_t8426.GetFieldData("t8426OutBlock", "shcode", i)
        _expcode = _inst_t8426.GetFieldData("t8426OutBlock", "expcode", i)
        if _hname[:4] == '미국달러':
            _next_code = _inst_t8426.GetFieldData("t8426OutBlock", "shcode", i + 1)
            _next_expcode = _inst_t8426.GetFieldData("t8426OutBlock", "expcode", i + 1)
            return [_shcode, _expcode, _next_code, _next_expcode]


async def xingCode():
    global _inst_t9943
    if _inst_t9943 is None:
        _inst_t9943 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerT9943)
        _inst_t9943.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\t9943.res"
    _inst_t9943.SetFieldData("t9943InBlock", "gubun", 0, '')
    _inst_t9943.Request(0)
    await _wait(XAQueryEventHandlerT9943)
    _code = _inst_t9943.GetFieldData("t9943OutBlock", "shcode", 0)
    _expcode = _inst_t9943.GetFieldData("t9943OutBlock", "expcode", 0)
    await asyncio.sleep(1)
    return _code, _expcode


async def xingMiniCode(next=False):
    global _inst_t8435
    if _inst_t8435 is None:
        _inst_t8435 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerT8435)
        _inst_t8435.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\t8435.res"
    _inst_t8435.SetFieldData("t8435InBlock", "gubun", 0, 'MF')
    _inst_t8435.Request(0)
    await _wait(XAQueryEventHandlerT8435)
    idx = 1 if next else 0
    _code = _inst_t8435.GetFieldData("t8435OutBlock", "shcode", idx)
    _exp = _inst_t8435.GetFieldData("t8435OutBlock", "expcode", idx)
    await asyncio.sleep(1)
    return _code, _exp


async def xingOptCode(jango_clear=None):
    global _inst_t9944
    if _inst_t9944 is None:
        _inst_t9944 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerT9944)
        _inst_t9944.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\t9944.res"
    _inst_t9944.SetFieldData("t9944InBlock", "dummy", 0, '')
    _inst_t9944.Request(0)
    await _wait(XAQueryEventHandlerT9944, timeout=5)

    count = _inst_t9944.GetBlockCount("t9944OutBlock")
    optCode = _inst_t9944.GetFieldData("t9944OutBlock", "shcode", 0)
    month = optCode[3:5]

    currentPrice, openPrice, lastMonth = await xingCurrentPrice(optCode)
    if datetime.datetime.now().date().strftime('%Y%m%d') == lastMonth:
        lst = []
        for i in range(count):
            optCode = _inst_t9944.GetFieldData("t9944OutBlock", "shcode", i)
            if optCode[3:5] not in lst:
                lst.append(optCode[3:5])
        month = lst[1]

    names = []
    for i in range(count):
        optCode = _inst_t9944.GetFieldData("t9944OutBlock", "shcode", i)
        if optCode[:3] not in names:
            names.append(optCode[:3])
    call_name = names[0]
    put_name = names[2] if len(names) == 4 else names[1]

    if jango_clear == True:
        lst = []
        for i in range(count):
            optCode = _inst_t9944.GetFieldData("t9944OutBlock", "shcode", i)
            if optCode[3:5] == month:
                lst.append(optCode)
        return lst

    return call_name + month, put_name + month


async def xingGetMgn():
    global _inst_mmdaq91200
    if _inst_mmdaq91200 is None:
        _inst_mmdaq91200 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerMMDAQ91200)
        _inst_mmdaq91200.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\MMDAQ91200.res"
    _inst_mmdaq91200.SetFieldData("MMDAQ91200InBlock1", "RecCnt", 0, 1)
    _inst_mmdaq91200.SetFieldData("MMDAQ91200InBlock1", "IsuLgclssCode", 0, '')
    _inst_mmdaq91200.SetFieldData("MMDAQ91200InBlock1", "IsuMdclssCode", 0, '')
    _inst_mmdaq91200.Request(0)
    await _wait(XAQueryEventHandlerMMDAQ91200)
    count = _inst_mmdaq91200.GetBlockCount("MMDAQ91200OutBlock2")
    for i in range(count):
        name = _inst_mmdaq91200.GetFieldData("MMDAQ91200OutBlock2", "ShtnHanglIsuNm", 0)
        rate_Mgn = _inst_mmdaq91200.GetFieldData("MMDAQ91200OutBlock2", "CsgnMgnrt", 0)
    return name, rate_Mgn
