import win32com.client
import datetime


class XReal_JIF:
    def __init__(self):
        super().__init__()
        self.jangubun = ''
        self.jstatus = ''
        print('JIF')

    def OnReceiveRealData(self, tr_code):
        self.jangubun = self.GetFieldData("OutBlock", "jangubun")
        self.jstatus = self.GetFieldData("OutBlock", "jstatus")
        print(self.jangubun, self.jstatus, datetime.datetime.now())

    def start(self):
        self.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\JIF.res"
        self.SetFieldData('InBlock', 'jangubun', '0')
        self.AdviseRealData()

    def end(self):
        self.UnadviseRealData()

    @classmethod
    def get_instance(cls):
        return win32com.client.DispatchWithEvents("XA_DataSet.XAReal", cls)


class XReal_FX9:
    def __init__(self):
        super().__init__()
        self.upstep = ''
        self.dnstep = ''
        self.uplmtprice = ''
        self.dnlmtprice = ''
        self.futcode = ''

    def OnReceiveRealData(self, tr_code):
        self.upstep = self.GetFieldData("OutBlock", "upstep")
        self.dnstep = self.GetFieldData("OutBlock", "dnstep")
        self.uplmtprice = self.GetFieldData("OutBlock", "uplmtprice")
        self.dnlmtprice = self.GetFieldData("OutBlock", "dnlmtprice")
        self.futcode = self.GetFieldData("OutBlock", "futcode")
        print(f'upstep: {self.upstep} | dnstep: {self.dnstep} | uplmtprice: {self.uplmtprice} | dnlmtprice: {self.dnlmtprice} | futcode: {self.futcode}')

    def start(self, code):
        self.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\FX9.res"
        self.SetFieldData('InBlock', 'futcode', code)
        self.AdviseRealData()

    def end(self):
        self.UnadviseRealData()

    @classmethod
    def get_instance(cls):
        return win32com.client.DispatchWithEvents("XA_DataSet.XAReal", cls)


class XReal_FC9:
    def __init__(self):
        super().__init__()
        self.open = ''
        self.price = ''
        self.jgubun = ''
        self.futcode = ''

    def OnReceiveRealData(self, tr_code):
        self.open = self.GetFieldData("OutBlock", "open")
        self.price = self.GetFieldData("OutBlock", "price")
        self.jgubun = self.GetFieldData("OutBlock", "jgubun")
        self.futcode = self.GetFieldData("OutBlock", "futcode")

    def start(self, code):
        self.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\FC9.res"
        self.SetFieldData('InBlock', 'futcode', code)
        self.AdviseRealData()

    def end(self):
        self.UnadviseRealData()

    @classmethod
    def get_instance(cls):
        return win32com.client.DispatchWithEvents("XA_DataSet.XAReal", cls)


class XReal_OC0:
    def __init__(self):
        super().__init__()
        self.open = ''
        self.price = ''
        self.jgubun = ''
        self.optcode = ''

    def OnReceiveRealData(self, tr_code):
        self.open = self.GetFieldData("OutBlock", "open")
        self.price = self.GetFieldData("OutBlock", "price")
        self.jgubun = self.GetFieldData("OutBlock", "jgubun")
        self.optcode = self.GetFieldData("OutBlock", "optcode")

    def start(self, code):
        self.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\OC0.res"
        self.SetFieldData('InBlock', 'optcode', code)
        self.AdviseRealData()

    def add_item(self, code):
        self.SetFieldData("InBlock", "optcode", code)
        self.AdviseRealData()

    def remove_item(self, code):
        self.UnadviseRealDataWithKey(code)

    def end(self):
        print('OUT')
        self.UnadviseRealData()

    @classmethod
    def get_instance(cls):
        return win32com.client.DispatchWithEvents("XA_DataSet.XAReal", cls)


class XReal_C01:
    def __init__(self):
        super().__init__()
        self.accno = ''
        self.cheprice = 0
        self.chevol = 0
        self.chetime = ''
        self.dosugb = ''
        self.order_list = {}

    def OnReceiveRealData(self, tr_code):
        self.accno = self.GetFieldData("OutBlock", "accno")
        self.cheprice = float(self.GetFieldData("OutBlock", "cheprice"))
        self.chevol = int(self.GetFieldData("OutBlock", "chevol"))
        self.chetime = self.GetFieldData("OutBlock", "chetime")
        self.dosugb = self.GetFieldData("OutBlock", "dosugb")

        if self.accno not in self.order_list:
            self.order_list[self.accno] = {'매수': 0, '매도': 0, '가격': 0}

        self.order_list[self.accno]['가격'] = self.cheprice

        if self.dosugb == '2':
            self.order_list[self.accno]['매수'] += self.chevol
        elif self.dosugb == '1':
            self.order_list[self.accno]['매도'] += self.chevol

    def start(self):
        self.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\C01.res"
        self.AdviseRealData()

    def end(self):
        self.UnadviseRealData()

    def set_info(self, account):
        del self.order_list[account]

    @classmethod
    def get_instance(cls):
        return win32com.client.DispatchWithEvents("XA_DataSet.XAReal", cls)


class XReal_FH9:
    def __init__(self):
        super().__init__()
        self.totoffercnt = 0
        self.totbidcnt = 0
        self.futcode = ''

    def OnReceiveRealData(self, tr_code):
        self.totoffercnt = int(self.GetFieldData("OutBlock", "totoffercnt"))
        self.totbidcnt = int(self.GetFieldData("OutBlock", "totbidcnt"))
        self.futcode = self.GetFieldData("OutBlock", "futcode")

    def start(self, code):
        self.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\FH9.res"
        self.SetFieldData('InBlock', 'futcode', code)
        self.AdviseRealData()

    def end(self):
        self.UnadviseRealData()

    @classmethod
    def get_instance(cls):
        return win32com.client.DispatchWithEvents("XA_DataSet.XAReal", cls)


class XReal_YF9:
    def __init__(self):
        super().__init__()
        self.ychetime = ''
        self.yeprice = 0
        self.expct_ccls_q = 0

    def OnReceiveRealData(self, tr_code):
        self.ychetime = self.GetFieldData("OutBlock", "ychetime")
        self.yeprice = self.GetFieldData("OutBlock", "yeprice")
        self.expct_ccls_q = self.GetFieldData("OutBlock", "expct_ccls_q")

    def start(self, code):
        self.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\YF9.res"
        self.SetFieldData('InBlock', 'futcode', code)
        self.AdviseRealData()

    def end(self):
        self.UnadviseRealData()

    @classmethod
    def get_instance(cls):
        return win32com.client.DispatchWithEvents("XA_DataSet.XAReal", cls)


class XReal_YOC:
    def __init__(self):
        super().__init__()
        self.ychetime = ''
        self.yeprice = 0
        self.optcode = ''
        self.expct_ccls_q = 0

    def OnReceiveRealData(self, tr_code):
        self.ychetime = self.GetFieldData("OutBlock", "ychetime")
        self.yeprice = self.GetFieldData("OutBlock", "yeprice")
        self.optcode = self.GetFieldData("OutBlock", "optcode")
        self.expct_ccls_q = self.GetFieldData("OutBlock", "expct_ccls_q")

    def start(self, code):
        self.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\YOC.res"
        self.SetFieldData('InBlock', 'optcode', code)
        self.AdviseRealData()

    def add_item(self, code):
        self.SetFieldData("InBlock", "optcode", code)
        self.AdviseRealData()

    def remove_item(self, code):
        self.UnadviseRealDataWithKey(code)

    def end(self):
        self.UnadviseRealData()

    @classmethod
    def get_instance(cls):
        return win32com.client.DispatchWithEvents("XA_DataSet.XAReal", cls)


# ========================= 야간 선물 =========================

class XReal_DC0:
    def __init__(self):
        super().__init__()
        self.open = ''
        self.price = ''
        self.jgubun = ''

    def OnReceiveRealData(self, tr_code):
        self.open = self.GetFieldData("OutBlock", "open")
        self.price = self.GetFieldData("OutBlock", "price")
        self.jgubun = self.GetFieldData("OutBlock", "jgubun")

    def start(self, code):
        self.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\DC0.res"
        self.SetFieldData('InBlock', 'futcode', code)
        self.AdviseRealData()

    def end(self):
        print('OUT')
        self.UnadviseRealData()

    @classmethod
    def get_instance(cls):
        return win32com.client.DispatchWithEvents("XA_DataSet.XAReal", cls)


class XReal_C02:
    def __init__(self):
        super().__init__()
        self.accno = ''
        self.cheprice = 0
        self.chevol = 0
        self.chetime = ''
        self.dosugb = ''
        self.order_list = {}

    def OnReceiveRealData(self, tr_code):
        self.accno = self.GetFieldData("OutBlock", "accno")
        self.cheprice = float(self.GetFieldData("OutBlock", "cheprice"))
        self.chevol = int(self.GetFieldData("OutBlock", "chevol"))
        self.chetime = self.GetFieldData("OutBlock", "chetime")
        self.dosugb = self.GetFieldData("OutBlock", "dosugb")

        if self.accno not in self.order_list:
            self.order_list[self.accno] = {'매수': 0, '매도': 0, '가격': 0}

        self.order_list[self.accno]['가격'] = self.cheprice

        if self.dosugb == '2':
            self.order_list[self.accno]['매수'] += self.chevol
        elif self.dosugb == '1':
            self.order_list[self.accno]['매도'] += self.chevol

        print(self.order_list)

    def start(self):
        self.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\C02.res"
        self.AdviseRealData()

    def end(self):
        self.UnadviseRealData()

    def set_info(self, account):
        del self.order_list[account]

    @classmethod
    def get_instance(cls):
        return win32com.client.DispatchWithEvents("XA_DataSet.XAReal", cls)


class XReal_DH0:
    def __init__(self):
        super().__init__()
        self.totoffercnt = 0
        self.totbidcnt = 0

    def OnReceiveRealData(self, tr_code):
        self.totoffercnt = int(self.GetFieldData("OutBlock", "totoffercnt"))
        self.totbidcnt = int(self.GetFieldData("OutBlock", "totbidcnt"))

    def start(self, code):
        self.ResFileName = "C:\\LS_SEC\\xingAPI\\Res\\DH0.res"
        self.SetFieldData('InBlock', 'futcode', code)
        self.AdviseRealData()

    def end(self):
        self.UnadviseRealData()

    @classmethod
    def get_instance(cls):
        return win32com.client.DispatchWithEvents("XA_DataSet.XAReal", cls)
