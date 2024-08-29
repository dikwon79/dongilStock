from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        print("Kiwoom 클래스")

        ######## eventloop 모듈
        self.login_event_loop = None
        ############################

        ######## 변수모음
        self.account_num = None
        ############################

        self.get_ocx_instance()
        self.event_slots()
        self.signal_login_connect()
        self.get_account_info()
        self.detail_account_info() #예수금 가져오는것

    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)

    def login_slot(self, err_code):
        print(errors(err_code))

        self.login_event_loop.exit()

    def signal_login_connect(self):
        self.dynamicCall("CommConnect()")

        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(String)", "ACCNO")

        self.account_num = account_list.split(":")[0]
        print("나의 보유 계좌번호는 %s" % self.account_num)

    def detail_account_info(self):
        print("예수금 가져오는 파트")
