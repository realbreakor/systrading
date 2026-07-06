import win32com.client
import pythoncom
import asyncio
import datetime
import logging

from ls_api.config import MOUI_PSWD


class XAQueryEventHandlerT8455:
    query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandlerT8455.query_state = 1


class XAQueryEventHandlerT8456:
    query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandlerT8456.query_state = 1


class XAQueryEventHandlerCCENT00100:
    query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandlerCCENT00100.query_state = 1


class XAQueryEventHandlerCCENQ90200:
    query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandlerCCENQ90200.query_state = 1


class XAQueryEventHandlerCCENQ30100:
    query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandlerCCENQ30100.query_state = 1


class XAQueryEventHandlerCCENT00300:
    query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandlerCCENT00300.query_state = 1


# 각 TR별 inst 캐시 (처음 호출 시 한 번만 생성)
_inst_t8455 = None
_inst_t8456 = None
_inst_ccent00100 = None
_inst_ccenq90200 = None
_inst_ccenq30100 = None
_inst_ccent00300 = None


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


async def n_find_code(kind):
    global _inst_t8455
    if _inst_t8455 is None:
        _inst_t8455 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerT8455)
        _inst_t8455.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\t8455.res"
    _inst_t8455.SetFieldData("t8455InBlock", "gubun", 0, kind)
    _inst_t8455.Request(0)
    await _wait(XAQueryEventHandlerT8455, timeout=5)
    fcode = _inst_t8455.GetFieldData("t8455OutBlock", "shcode", 0)
    await asyncio.sleep(1)
    return fcode


async def n_xingCurrentPrice(codes):
    global _inst_t8456
    if _inst_t8456 is None:
        _inst_t8456 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerT8456)
        _inst_t8456.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\t8456.res"
    _inst_t8456.SetFieldData("t8456InBlock", "focode", 0, codes)
    _inst_t8456.Request(0)
    await _wait(XAQueryEventHandlerT8456, timeout=5)
    openPrice = _inst_t8456.GetFieldData("t8456OutBlock", "open", 0)
    currentPrice = _inst_t8456.GetFieldData("t8456OutBlock", "price", 0)
    return openPrice, currentPrice


async def n_xingFutOrder(acnt, direct, numOfOrder, codeName, orderType="03", orderPrice=""):
    global _inst_ccent00100
    log_msg = f"계좌번호: {acnt} | 방향: {direct} | 주문계약수: {numOfOrder} | 종목코드: {codeName} | 주문타입: {orderType} | 주문가: {orderPrice}"
    logging.info(log_msg)
    if _inst_ccent00100 is None:
        _inst_ccent00100 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerCCENT00100)
        _inst_ccent00100.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\CCENT00100.res"
    _inst_ccent00100.SetFieldData("CCENT00100InBlock1", "AcntNo", 0, acnt)
    _inst_ccent00100.SetFieldData("CCENT00100InBlock1", "Pwd", 0, MOUI_PSWD)
    _inst_ccent00100.SetFieldData("CCENT00100InBlock1", "FnoIsuNo", 0, codeName)
    _inst_ccent00100.SetFieldData("CCENT00100InBlock1", "BnsTpCode", 0, direct)
    _inst_ccent00100.SetFieldData("CCENT00100InBlock1", "FnoOrdprcPtnCode", 0, orderType)
    _inst_ccent00100.SetFieldData("CCENT00100InBlock1", "FnoOrdPrc", 0, orderPrice)
    _inst_ccent00100.SetFieldData("CCENT00100InBlock1", "OrdQty", 0, numOfOrder)
    _inst_ccent00100.Request(0)
    await _wait(XAQueryEventHandlerCCENT00100, timeout=5)
    print('Order done')


async def n_xingCheckAccount(account, codeName):
    global _inst_ccenq90200
    if _inst_ccenq90200 is None:
        _inst_ccenq90200 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerCCENQ90200)
        _inst_ccenq90200.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\CCENQ90200.res"
    _inst_ccenq90200.SetFieldData("CCENQ90200InBlock1", "RecCnt", 0, 1)
    _inst_ccenq90200.SetFieldData("CCENQ90200InBlock1", "AcntNo", 0, account)
    _inst_ccenq90200.SetFieldData("CCENQ90200InBlock1", "InptPwd", 0, MOUI_PSWD)
    _inst_ccenq90200.SetFieldData("CCENQ90200InBlock1", "BalEvalTp", 0, '0')
    _inst_ccenq90200.SetFieldData("CCENQ90200InBlock1", "FutsPrcEvalTp", 0, '1')
    _inst_ccenq90200.Request(0)
    await _wait(XAQueryEventHandlerCCENQ90200)

    numOfContract = 0
    avgPrice = 0
    enteredDirection = '0'
    medosu = ''

    count = _inst_ccenq90200.GetBlockCount("CCENQ90200OutBlock3")
    for i in range(count):
        code = _inst_ccenq90200.GetFieldData("CCENQ90200OutBlock3", "FnoIsuNo", 0)
        if code == codeName:
            numOfContract = int(_inst_ccenq90200.GetFieldData("CCENQ90200OutBlock3", "UnsttQty", 0))
            medosu = _inst_ccenq90200.GetFieldData("CCENQ90200OutBlock3", "BnsTpNm", 0)
            avgPrice = float(_inst_ccenq90200.GetFieldData("CCENQ90200OutBlock3", "FnoAvrPrc", 0))
            enteredDirection = _inst_ccenq90200.GetFieldData("CCENQ90200OutBlock3", "BnsTpCode", 0)

    return numOfContract, avgPrice, enteredDirection, medosu


async def n_xingCheckOrder(account, code, lst=False):
    global _inst_ccenq30100
    now = datetime.datetime.now().date()
    tom = now + datetime.timedelta(days=1)
    now = now.strftime('%Y%m%d')
    tom = tom.strftime('%Y%m%d')
    if _inst_ccenq30100 is None:
        _inst_ccenq30100 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerCCENQ30100)
        _inst_ccenq30100.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\CCENQ30100.res"
    _inst_ccenq30100.SetFieldData("CCENQ30100InBlock1", "RecCnt", 0, 1)
    _inst_ccenq30100.SetFieldData("CCENQ30100InBlock1", "AcntNo", 0, account)
    _inst_ccenq30100.SetFieldData("CCENQ30100InBlock1", "InptPwd", 0, MOUI_PSWD)
    _inst_ccenq30100.SetFieldData("CCENQ30100InBlock1", "QrySrtDt", 0, now)
    _inst_ccenq30100.SetFieldData("CCENQ30100InBlock1", "QryEndDt", 0, tom)
    _inst_ccenq30100.SetFieldData("CCENQ30100InBlock1", "FnoClssCode", 0, '11')
    _inst_ccenq30100.SetFieldData("CCENQ30100InBlock1", "PrdgrpCode", 0, '00')
    _inst_ccenq30100.SetFieldData("CCENQ30100InBlock1", "PrdtExecTpCode", 0, '2')
    _inst_ccenq30100.SetFieldData("CCENQ30100InBlock1", "StnlnSeqTp", 0, '3')
    _inst_ccenq30100.SetFieldData("CCENQ30100InBlock1", "MktTpCode", 0, '0')
    _inst_ccenq30100.SetFieldData("CCENQ30100InBlock1", "CommdaCode", 0, '99')
    _inst_ccenq30100.SetFieldData("CCENQ30100InBlock1", "FnoIsuNo", 0, code)
    _inst_ccenq30100.SetFieldData("CCENQ30100InBlock1", "FnoTrdPtnCode", 0, '03')
    _inst_ccenq30100.SetFieldData("CCENQ30100InBlock1", "SrtOrdNo2", 0, 0)
    _inst_ccenq30100.Request(0)
    await _wait(XAQueryEventHandlerCCENQ30100)

    count = _inst_ccenq30100.GetBlockCount("CCENQ30100OutBlock3")

    if lst == True:
        result = []
        for i in range(count):
            ordrem = int(_inst_ccenq30100.GetFieldData("CCENQ30100OutBlock3", "UnercQty", i))
            ordno = int(_inst_ccenq30100.GetFieldData("CCENQ30100OutBlock3", "OrdNo", i))
            result.append((ordno, ordrem))
        return result
    else:
        unchecked = 0
        for i in range(count):
            try:
                unchecked += int(_inst_ccenq30100.GetFieldData("CCENQ30100OutBlock3", "UnercQty", i))
            except ValueError:
                unchecked = 0
        return unchecked


async def n_xingCancelOrder(acnt, codeName, ordNo, qty):
    global _inst_ccent00300
    if _inst_ccent00300 is None:
        _inst_ccent00300 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerCCENT00300)
        _inst_ccent00300.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\CCENT00300.res"
    _inst_ccent00300.SetFieldData("CCENT00300InBlock1", "AcntNo", 0, acnt)
    _inst_ccent00300.SetFieldData("CCENT00300InBlock1", "Pwd", 0, MOUI_PSWD)
    _inst_ccent00300.SetFieldData("CCENT00300InBlock1", "FnoIsuNo", 0, codeName)
    _inst_ccent00300.SetFieldData("CCENT00300InBlock1", "OrgOrdNo", 0, ordNo)
    _inst_ccent00300.SetFieldData("CCENT00300InBlock1", "CancQty", 0, qty)
    _inst_ccent00300.Request(0)
    await _wait(XAQueryEventHandlerCCENT00300)
    print('Cancel done')