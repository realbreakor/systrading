import win32com.client
import pythoncom
import asyncio
import datetime
import random
import logging

from ls_api.config import MOUI_PSWD
from ls_api.account import xingCheckAccount, xingCheckOrder


class XAQueryEventHandlerCFOAT00100:
    query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandlerCFOAT00100.query_state = 1


class XAQueryEventHandlerCFOAT00300:
    query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandlerCFOAT00300.query_state = 1


# 각 TR별 inst 캐시 (처음 호출 시 한 번만 생성)
_inst_cfoat00100 = None
_inst_cfoat00300 = None


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


async def xingFutOrder(acnt, direct, numOfOrder, codeName, orderType="03", orderPrice=""):
    global _inst_cfoat00100
    log_msg = f"계좌번호: {acnt} | 방향: {direct} | 주문계약수: {numOfOrder} | 종목코드: {codeName} | 주문타입: {orderType} | 주문가: {orderPrice}"
    logging.info(log_msg)
    if _inst_cfoat00100 is None:
        _inst_cfoat00100 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerCFOAT00100)
        _inst_cfoat00100.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\CFOAT00100.res"
    _inst_cfoat00100.SetFieldData("CFOAT00100InBlock1", "AcntNo", 0, acnt)
    _inst_cfoat00100.SetFieldData("CFOAT00100InBlock1", "Pwd", 0, MOUI_PSWD)
    _inst_cfoat00100.SetFieldData("CFOAT00100InBlock1", "FnoIsuNo", 0, codeName)
    _inst_cfoat00100.SetFieldData("CFOAT00100InBlock1", "BnsTpCode", 0, direct)
    _inst_cfoat00100.SetFieldData("CFOAT00100InBlock1", "FnoOrdprcPtnCode", 0, orderType)
    _inst_cfoat00100.SetFieldData("CFOAT00100InBlock1", "FnoOrdPrc", 0, orderPrice)
    _inst_cfoat00100.SetFieldData("CFOAT00100InBlock1", "OrdQty", 0, numOfOrder)
    _inst_cfoat00100.Request(0)
    await _wait(XAQueryEventHandlerCFOAT00100)
    print('Order done')


async def xingCancelOrder(acnt, codeName, ordNo, qty):
    global _inst_cfoat00300
    if _inst_cfoat00300 is None:
        _inst_cfoat00300 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerCFOAT00300)
        _inst_cfoat00300.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\CFOAT00300.res"
    _inst_cfoat00300.SetFieldData("CFOAT00300InBlock1", "AcntNo", 0, acnt)
    _inst_cfoat00300.SetFieldData("CFOAT00300InBlock1", "Pwd", 0, MOUI_PSWD)
    _inst_cfoat00300.SetFieldData("CFOAT00300InBlock1", "FnoIsuNo", 0, codeName)
    _inst_cfoat00300.SetFieldData("CFOAT00300InBlock1", "OrgOrdNo", 0, ordNo)
    _inst_cfoat00300.SetFieldData("CFOAT00300InBlock1", "CancQty", 0, qty)
    _inst_cfoat00300.Request(0)
    await _wait(XAQueryEventHandlerCFOAT00300)
    print('Cancel done')
