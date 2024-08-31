from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        print("Kiwoom 클래스")

        ######## eventloop 모듈
        self.login_event_loop = None
        self.detail_account_info_event_loop = None
        self.detail_account_info_event_loop2 = None
        ############################

        ######## 변수모음
        self.account_num = None
        ############################


        self.get_ocx_instance()
        self.event_slots()

        #로그인
        self.signal_login_connect()

        # 계좌비밀번호 설정 호출(함수이용)
        #self.dynamicCall("KOA_Functions(String, String)", "ShowAccountWindow", "")

        self.get_account_info()
        self.detail_account_info() #예수금 가져오는것

        # 계좌평가 잔고내역 요청
        self.detail_acount_mystock()

    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)
        self.OnReceiveTrData.connect(self.trdata_slot)

    def login_slot(self, err_code):
        print(errors(err_code))

        self.login_event_loop.exit()

    def signal_login_connect(self):
        self.dynamicCall("CommConnect()")

        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(String)", "ACCNO")

        self.account_num = account_list.split(":")[0][:-1]
        print("나의 보유 계좌번호는 %s" % self.account_num)

    def detail_account_info(self):
        print("예수금 요청하는 부분 %s" %self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        self.dynamicCall("CommRqData(String, String, int, String)", "예수금상세현황요청","opw00001", "0", "2000")

        self.detail_account_info_event_loop = QEventLoop()
        self.detail_account_info_event_loop.exec_()

    def detail_acount_mystock(self, sPrevNext="0"):
        print("계좌평가 잔고내역 요청")
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")

        self.dynamicCall("CommRqData(String, String, int, String)", "계좌평가잔고내역요청", "opw00018", sPrevNext, "2000")

        self.detail_account_info_event_loop2 = QEventLoop()
        self.detail_account_info_event_loop2.exec_()

    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        '''
        tr 요청을 받는 구역 슬롯
        :param sCrNo: 스크린번호
        :param sRQName: 요청시 지은 이름
        :param sTrCode: 요청id , tr코드
        :param sRecordName: 사용 안함
        :param sPrevNext: 다음페이지가 있는지...
        :return:
        '''

        if sRQName == "예수금상세현황요청":
            deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "예수금")

            print("예수금 %s" % int(deposit))

            ok_deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "출금가능금액")
            print("출금가능금액 %s" % int(ok_deposit))


            self.detail_account_info_event_loop.exit()

        if sRQName == "계좌평가잔고내역요청":
            total_buy_money = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "총매입금액")
            total_buy_money_result = int(total_buy_money)

            print("총매입금액 %s" % total_buy_money_result)

            total_profit_loss_rate = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "총수익률(%)")
            total_profit_loss_rate_result = float(total_profit_loss_rate)

            print("총수익률(%%): %s" % total_profit_loss_rate_result)

            self.detail_account_info_event_loop2.exit()