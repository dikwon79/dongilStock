import sys
import pandas as pd
import datetime
import numpy as np
from find_theme.stock_data import *
from collections import deque
from queue import Queue

from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QAxContainer import QAxWidget
from PyQt5 import QtGui, uic
from PyQt5.QtCore import Qt, QSettings, QTimer, QCoreApplication, QAbstractTableModel


form_class = uic.loadUiType("main.ui")[0]



class PandasModel(QAbstractTableModel):
    def __init__(self, data):
        super().__init__()
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                return str(self._data.iloc[index.row(), index.column()])
        return None

    def headerData(self, section, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._data.columns[section]

        if orientation == Qt.Vertical and role == Qt.DisplayRole:
            return self._data.index[section]
        return None

    def setData(self, index, value, role):
        # Always return False to disable editing
        return False

    def flags(self, index):
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable


class KiwoomAPI(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.show()
        self.scrnum = 5000
        self.using_condition_name = ""

        self.condition_name_to_condition_idx_dict = dict()
        self.registed_condition_df = pd.DataFrame(columns=["화면번호", "조건식이름"])
        self.registed_conditions_list = []

        self.account_info_df = pd.DataFrame(
            columns=[
                "종목명",
                "테마",
                "매매가능수량",
                "보유수량",
                "매입가",
                "현재가",
                "수익률",
            ]
        )
        self.accountTableView.resizeColumnsToContents()
        self.accountTableView.resizeRowsToContents()

        self.watchListTableView.resizeColumnsToContents()
        self.watchListTableView.resizeRowsToContents()
        self.is_updated_realtime_watchlist = False
        self.stock_code_to_price_info_dict = dict()

        try:
            self.realtime_watchList_df = pd.read_pickle("./realtime_watchlist_df.pkl")
        except FileNotFoundError:
            self.realtime_watchList_df = pd.DataFrame(columns=["종목명","현재가","평균단가","목표가","손절가","수익률","매수기반조건식","보유수량", "매수주문완료여부"])

        self.realtime_registered_codes= set()
        self.conditionInputButton.clicked.connect(self.condition_in)
        self.conditionOutButton.clicked.connect(self.condition_out)
        self.settings = QSettings('My company', 'myApp')
        self.load_settings()
        #self.setWindowIcon(QtGui.QIcon('icon.ico'))

        self.max_send_per_sec: int = 4  # 조당 TR 호출 최대 4번
        self.max_send_per_minute: int = 55  # 분당 TR 호출 죄대 55번
        self.max_send_per_hour: int = 950  # 시간당 TR 호줄 최대 950번
        self.last_tr_send_times = deque(maxlen=self.max_send_per_hour)

        self.tr_req_queue = Queue()
        self.orders_queue = Queue()


        self.unfinished_order_num_to_info_dict = dict()

        self.account_num = None
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self._set_signal_slots()
        self._login()


        # Create an instance of QTimer
        self.timer1 = QTimer()
        self.timer2 = QTimer()
        self.timer3 = QTimer()
        self.timer4 = QTimer()
        self.timer5 = QTimer()
        self.timer6 = QTimer()
        self.timer7 = QTimer()
        self.timer8 = QTimer()

        # Connect the timeout signal to the update_pandas_models method
        self.timer1.timeout.connect(self.update_pandas_models)
        self.timer2.timeout.connect(self.send_tr_request)
        self.timer3.timeout.connect(self.send_orders)
        self.timer4.timeout.connect(self.request_get_account_balance)
        self.timer5.timeout.connect(self.request_current_order_info)
        self.timer6.timeout.connect(self.save_pickle)
        self.timer7.timeout.connect(self.check_unfinished_orders)
        self.timer8.timeout.connect(self.check_outliers)


        #cell 클릭 이벤트
        self.watchListTableView.clicked.connect(self.on_cell_clicked)

    def on_cell_clicked(self, index):

        row = index.row()
        column = index.column()

        # QTableView의 모델에 접근
        model = self.account_info_df.model()
        item = model.data(index)

        print(f"로우:{row}, 컬럼:{column}, 아이템 :{item}")
        # 비동기 작업 스레드 시작
        #self.worker = Worker(row, column, item.text())
        #self.worker.finished.connect(self.on_task_finished)
        #self.worker.start()

    def on_task_finished(self, result: str):
        pass
        # 스레드 작업이 완료되면 GUI 업데이트
        #QMessageBox.information(self, "Task Finished", result)

    def check_outliers(self):
        pop_list = []
        for row in self.realtime_watchList_df.itertuples():
            stock_code = getattr(row, "Index")
            목표가 = getattr(row, "목표가")
            손절가 = getattr(row, "손절가")
            if np.isnan(float(목표가)) or np.isnan(float(손절가)):
                pop_list.append(stock_code)
        for stock_code in pop_list:
            print(f"종목코드: {stock_code}, outlier! Pop!")
            self.realtime_watchList_df.drop(stock_code, inplace=True)

    def check_unfinished_orders(self):
        pop_list = []
        for order_num, stock_info_dict in self.unfinished_order_num_to_info_dict.items():
            주문번호 = order_num
            종목코드 = stock_info_dict['종목코드']
            주문체결시간 = stock_info_dict['주문체결시간']
            미체결수량 = stock_info_dict['미체결수량']
            주문가격 = stock_info_dict['주문가격']

            order_time = datetime.datetime.now().replace(
                hour=int(주문체결시간[:-4]),
                minute=int(주문체결시간[-4:-2]),
                second=int(주문체결시간[-2:]),
            )

            정정주문가격 = self.stock_code_to_price_info_dict.get(종목코드, None)
            if not 정정주문가격:
                print(f"종목코드 : {종목코드}, 최우선매수호가X 주문실패!")
                return
            if self.now_time - order_time >= datetime.timedelta(seconds=10):
                print(f"종목코드: {종목코드}, 미체결수량: {미체결수량}, 주문번호: {주문번호}, 지정가 매도 정정 주문!")
                # 시장가 매도 정정
                self.orders_queue.put(
                    [
                        "매도정정주문",
                        self._get_screen_num(),
                        self.account_num,
                        6,
                        종목코드,
                        미체결수량,
                        정정주문가격,
                        "00",
                        주문번호,
                    ],
                )
                pop_list.append(주문번호)

        for order_num in pop_list:
            self.unfinished_order_num_to_info_dict.pop(order_num)

    def load_settings(self):
        self.resize(self.settings.value("size", self.size()))

        self.move(self.settings.value("pos", self.pos()))
        self.buyAmountLineEdit.setText(self.settings.value('buyAmountLineEdit', defaultValue="100000", type=str))
        self.goalReturnLineEdit.setText(self.settings.value('goalReturnLineEdit',defaultValue="2.5", type=str))
        self.stopLossLineEdit.setText(self.settings.value('stopLossLineEdit', defaultValue="-2.5",type=str))

    def save_pickle(self):
        self.realtime_watchList_df.to_pickle("./realtime_watchlist_df.pkl")
        self.save_settings()

    def request_current_order_info(self):
        self.tr_req_queue.put([self.get_current_order_info()])

    def get_current_order_info(self):
        self.set_input_value("계좌번호", self.account_num)
        self.set_input_value("체결구분", "1")
        self.set_input_value("매매구분", "0")
        self.comm_rq_data("opt10075_req", "opt10075", 0, self._get_screen_num())


    def request_get_account_balance(self):
        self.tr_req_queue.put([self.get_account_balance])

    def send_tr_request(self):
        self.now_time = datetime.datetime.now()

        if self.is_check_tr_req_condition() and not self.tr_req_queue.empty():
            # 큐에서 함수 및 인수 가져오기
            request_func, *func_args = self.tr_req_queue.get()

            # None 체크 추가
            if request_func is None:
                print("Error: request_func is None")
            else:
                print(f"Executing TR request function: {request_func}")

                # 함수가 있으면 호출
                request_func(*func_args) if func_args else request_func()

            self.last_tr_send_times.append(self.now_time)

    def get_account_info(self):
        account_nums = str(self.kiwoom.dynamicCall("GetLoginInfo(QString)", ["ACCNO"]).rstrip(';'))

        print(f"계좌번호 리스트: {account_nums}")
        self.account_num = account_nums.split(';')[0]
        print(f"사용 계좌 번호: {self.account_num}")

    def get_account_balance(self):
        if self.is_check_tr_req_condition():

            self.set_input_value("계좌번호", self.account_num)
            self.set_input_value("비밀번호", "0000")
            self.set_input_value("비밀번호입력매체구분", "00")
            self.comm_rq_data("opw00018_req", "opw00018", 0, self._get_screen_num())

    def comm_rq_data(self, rqname, trcode, next, screen_no):

        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", rqname, trcode, next, screen_no)

    def update_pandas_models(self):
        pd_model = PandasModel(self.registed_condition_df)
        self.registeredTableView.setModel(pd_model)
        pd_model2 = PandasModel(self.realtime_watchList_df)
        self.watchListTableView.setModel(pd_model2)
        pd_model3 = PandasModel(self.account_info_df)
        self.accountTableView.setModel(pd_model3)

    def receive_tr_data(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext, nDataLength, sErrorCode, sMessage, sSplmMsg):
        # SScrNo:화면번호 SROName: 사용자 구분명 sTrCode: TR이름, sRecordName:레코드 이름 sPrevNext: 연속조회 유무를  판단한는 값 0
        #  연속(추가조회)데미터,없음, 2:연속(주가조회) 데미터 있음
        # 조회요청 응답을 받거나 조회데이터를 수신했올 때 호출됩니다.
        # 조회데이터는 이 이벤트에서 GetCommData()함수를 이용해서 얻어올수 있습니다.
        print(
            f"Received TR data sScrNo: {sScrNo}, sRQName: {sRQName}, "
            f"sTrCode: {sTrCode}, sRecordName: {sRecordName}, sPrevNext: {sPrevNext}, "
            f"nDataLength: {nDataLength}, sErrorCode: {sErrorCode}, sMessage: {sMessage}, sSplmMsg: {sSplmMsg}")

        if sRQName == "opw00018_req":
            self.on_opw00018_req(sTrCode, sRQName)
        elif sRQName == "opt10075_req":
            self.on_opt10075_req(sTrCode, sRQName)

    def on_opt10075_req(self, sTrCode, sRQName):
        cnt = self._get_repeat_cnt(sTrCode, sRQName)
        for i in range(cnt):
            주문번호 = self.get_comm_data(sTrCode, sRQName, i, "주문번호").strip()
            미체결수량 = int(self.get_comm_data(sTrCode, sRQName, i, "미체결수량"))
            주문가격 = int(self.get_comm_data(sTrCode, sRQName, i, "주문가격"))
            종목코드 = self.get_comm_data(sTrCode, sRQName, i, "종목코드").strip()
            주문구분 = self.get_comm_data(sTrCode, sRQName, i, "주문구분").replace("+", "").replace("-", "").strip()
            시간 = self.get_comm_data(sTrCode, sRQName, i, "시간").strip()
            order_time = datetime.datetime.now().replace(
                hour= int(시간[:-4]),
                minute= int(시간[-4:-2]),
                second= int(시간[-2:]),
            )

            정정주문가격 = self.get_sell_price(주문가격 * 0.93)
            if 주문구분 in ("매도","매도정정") and self.now_time - order_time >= datetime.timedelta(seconds=10):
                print(f"종목코드: {종목코드}, 미체결수량: {미체결수량}, 주문번호: {주문번호}, 시장가 매도 정정 주문!")
                # 시장가 매도 정정
                self.orders_queue.put(
                    [
                        "매도정정주문",
                        self._get_screen_num(),
                        self.account_num,
                        6,
                        종목코드,
                        미체결수량,
                        정정주문가격,
                        "03",
                        주문번호,
                    ],
                )


    def get_chejandata(self, nFid):
        ret = self.kiwoom.dynamicCall("GetChejanData(int)", nFid)
        return ret

    def receive_chejandata(self, sGubun, nItemCnt, sFIdList):
        # sGubun: 체결구분 접수와 체결시 '0'값, 국내주식 잔고전달은 '1'값, 파생잔고 전달은 '4'
        if sGubun == "0":
            종목코드 = self.get_chejandata(9001).replace("A", "").strip()
            종목명 = self.get_chejandata(302).strip()
            주문체결시간 = self.get_chejandata(908).strip()
            주문수량 = 0 if len(self.get_chejandata(900)) == 0 else int(self.get_chejandata(900))
            주문가격 = 0 if len(self.get_chejandata(901)) == 0 else int(self.get_chejandata(901))
            체결수량 = 0 if len(self.get_chejandata(911)) == 0 else int(self.get_chejandata(911))
            체결가격 = 0 if len(self.get_chejandata(910)) == 0 else int(self.get_chejandata(910))
            미체결수량 = 0 if len(self.get_chejandata(902)) == 0 else int(self.get_chejandata(902))
            주문구분 = self.get_chejandata(905).replace("+", "").replace("-", "").strip()
            매매구분 = self.get_chejandata(906).strip()
            단위체결가 = 0 if len(self.get_chejandata(914)) == 0 else int(self.get_chejandata(914))
            단위체결량 = 0 if len(self.get_chejandata(915)) == 0 else int(self.get_chejandata(915))
            원주문번호 = self.get_chejandata(904).strip()
            주문번호 = self.get_chejandata(9203).strip()
            print(f"Received cheiandata! 주문체결시간: {주문체결시간}, 종목코드:{종목코드}, "
                  f"종목명: {종목명}, 주문수량: {주문수량}, 주문가격: {주문가격}, 체결수량: {체결수량}, 체결가격: {체결가격}, "
                  f"주문구분: {주문구분}, 미체결수량: {미체결수량}, 매매구분: {매매구분}, 단위체결가: {단위체결가}, "
                  f"단위체결량: {단위체결량}, 주문번호: {주문번호}, 원주문번호: {원주문번호}")
            if 매매구분 == "매수" and 체결수량 > 0:
                self.realtime_watchList_df.loc[종목코드, "보유수량"] = 체결수량

            if 주문구분 in ("매도", "매도정정"):
                self.unfinished_order_num_to_info_dict[주문번호] = dict(
                    종목코드 = 종목코드,
                    미체결수량 = 미체결수량,
                    주문가격 = 주문가격,
                    주문체결시간 = 주문체결시간,
                )
                if 미체결수량 == 0:
                    self.unfinished_order_num_to_info_dict.pop(주문번호)


        if sGubun == 1:
            print("잔고동보")

    def save_settings(self):
        # Write window size and position to config file
        self.settings.setValue("size", self.size())
        self.settings.setValue("pos", self.pos())
        self.settings.setValue('buyAmountLineEdit', self.buyAmountLineEdit.text())
        self.settings.setValue('goalReturnLineEdit', self.goalReturnLineEdit.text())
        self.settings.setValue('stopLossLineEdit', self.stopLossLineEdit.text())

    def _get_repeat_cnt(self, trcode, rqname):
        ret = self.kiwoom.dynamicCall("GetRepeatCnt(QString, QString", trcode, rqname)
        return ret


    def on_opw00018_req(self, sTrCode, sRQName):
        현재평가잔고 = int(self.get_comm_data(sTrCode, sRQName, 0, "총평가금액"))
        print(f"현재평가잔고: {현재평가잔고: ,}원")
        self.currentBalanceLabel.setText(f"현재 평가 잔고: {현재평가잔고: ,}원")
        cnt = self._get_repeat_cnt(sTrCode, sRQName)
        self.account_info_df =pd.DataFrame(
            columns=[
                "종목명",
                "테마",
                "매매가능수량",
                "현재가",
                "보유수량",
                "매입가",
                "수익률",
            ]
        )
        current_account_code_list =[]
        for i in range(cnt):
            종목코드 = self.get_comm_data(sTrCode, sRQName, i, "종목번호").replace("A", "").strip()

            current_account_code_list.append(종목코드)
            종목명 = self.get_comm_data(sTrCode, sRQName, i, "종목명").strip()
            테마 = stockcode(종목코드)
            매매가능수량 = int(self.get_comm_data(sTrCode, sRQName, i, "매매가능수량"))
            보유수량 = int(self.get_comm_data(sTrCode, sRQName, i, "보유수량"))
            현재가 = int(self.get_comm_data(sTrCode, sRQName, i, "현재가"))
            매입가 = int(self.get_comm_data(sTrCode, sRQName, i, "매입가"))
            수익률 = float(self.get_comm_data(sTrCode, sRQName, i, "수익률(%)"))
            print(
                f"종목코드: {종목코드}, 테마:{테마}, 매매가능수량: {매매가능수량}, 보유수량: {보유수량}, 매입가: {매입가}, 수익률: {수익률}")

            if 종목코드 in self.realtime_watchList_df.index.to_list():
                self.realtime_watchList_df.loc[종목코드, "종목명"] = 종목명
                self.realtime_watchList_df.loc[종목코드, "평균단가"] = 매입가
                self.realtime_watchList_df.loc[종목코드, "보유수량"] = 보유수량

            self.account_info_df.loc[종목코드] = {
                "종목명": 종목명,
                "테마":테마,
                "매매가능수량": 매매가능수량,
                "현재가": 현재가,
                "보유수량": 보유수량,
                "매입가": 매입가,
                "수익률": 수익률,

            }


        if not self.is_updated_realtime_watchlist:
            for 종목코드 in current_account_code_list:
                self.register_code_to_realtime_list(종목코드)

            self.is_updated_realtime_watchlist = True
            realtime_tracking_code_list = self.realtime_watchList_df.index.to_list()
            for stock_code in realtime_tracking_code_list:
                if stock_code not in current_account_code_list:
                    self.realtime_watchList_df.drop(stock_code, inplace=True)
                    print(f"종목코드: {stock_code} self.realtime_watchlist_df 에서 drop!")

        self.accountTableView.resizeColumnsToContents()  # 열을 콘텐츠에 맞춰 자동 조정
        self.accountTableView.resizeRowsToContents()  # 행을 콘텐츠에 맞춰 자동 조정

    def set_input_value(self, id, value):
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", id, value)

    def get_comm_data(self, strTrCode, strRecordName, nIndex, strItemName):
        ret = self.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", strTrCode, strRecordName, nIndex,
                                      strItemName)

        return ret.strip()



    def condition_in(self):
        condition_name = self.conditionComboBox.currentText()
        condition_idx = self.condition_name_to_condition_idx_dict.get(condition_name, None)
        if not condition_idx:
            print("잘못된 조건 검색식 다시 선택하세요")
            return
        else:
            print(f"{condition_name} 실시간 조건 등록 요청!")
            self.send_condition(self._get_screen_num(), condition_name, condition_idx, 1)



    def condition_out(self):
        condition_name = self.conditionComboBox.currentText()
        condition_idx = self.condition_name_to_condition_idx_dict.get(condition_name, None)
        if not condition_idx:
            print("잘못된 조건 검색식 다시 선택하세요")
            return
        elif condition_idx in self.registed_condition_df.index:
            print(f"{condition_name} 실시간 조건 등록 편출!")
           # screen_num = self.registed_condition_df.loc[condition_idx,"화면번호"]
           # self.send_condition_stop(screen_num, condition_name, condition_idx)

            self.registed_condition_df.drop(condition_idx, inplace=True)
        else:
            print(f"조건식 편출 실패!")
            return
        print(self.registed_condition_df)


    def _set_signal_slots(self):
        self.kiwoom.OnEventConnect.connect(self._event_connect)
        self.kiwoom.OnReceiveRealData.connect(self._receive_realdata)
        self.kiwoom.OnReceiveConditionVer.connect(self._receive_condition)
        self.kiwoom.OnReceiveRealCondition.connect(self._receive_real_condition)
        self.kiwoom.OnReceiveTrData.connect(self.receive_tr_data)
        self.kiwoom.OnReceiveChejanData.connect(self.receive_chejandata)

        self.kiwoom.OnReceiveMsg.connect(self.receive_msg)

    def receive_msg(self, sScrNo, sRQName, sTrCode, sMsg):

        print(f"Received MSG! 화면번호: {sScrNo}, 사용자 구분명: {sRQName},TR이름: {sTrCode}, 메체지: {sMsg}")

    def _login(self):
        ret = self.kiwoom.dynamicCall("CommConnect()")
        if ret == 0:
            print("로그인 창 열기 성공!")

    def _event_connect(self, err_code):
        if err_code == 0:
            print("login성공")
            self._after_login()
        else:
            raise Exception("로그인 실패")

    def _after_login(self):
        self.get_account_info()
        self.kiwoom.dynamicCall("GetConditionLoad()")


        # Start the timer with an interval of 300 milliseconds (0.3 seconds)
        self.timer1.start(300)
        self.timer2.start(10)  # 0.01초마다 체크
        self.timer3.start(10)  # 5초마다 체크
        self.timer4.start(5000)  # 5초마다 체크
        self.timer5.start(60000)  # 60초마다 체크
        self.timer6.start(30000)  # 30초마다 저장
        self.timer7.start(100)  # 0.1초마다 저장
        self.timer8.start(1000)  # 1초마다 저장

    def _receive_condition(self):
        condition_info = self.kiwoom.dynamicCall("GetConditionNameList()").split(';')
        for condition_name_idx_str in condition_info:
            if len(condition_name_idx_str) == 0:
                continue
            condition_idx, condition_name = condition_name_idx_str.split('^')
            self.condition_name_to_condition_idx_dict[condition_name] = condition_idx
            #print(condition_idx, condition_name)

        self.conditionComboBox.addItems(self.condition_name_to_condition_idx_dict.keys())

    def _get_screen_num(self):
        self.scrnum += 1
        if self.scrnum > 5190:
            self.scrnum = 5000
        return str(self.scrnum)

    def send_condition(self, scr_num, condition_name, condition_idx, n_search):

        # nSearch: 조회구분, 0:조건검색, 1:실시간 조건검색
        result = self.kiwoom.dynamicCall(
            "SendCondition(QString, QString, int, int)",
             scr_num, condition_name, condition_idx, n_search
        )
        if result == 1:
            print(f"{condition_name} 조건검색 등록")
            self.registed_condition_df.loc[condition_idx] = {"화면번호":scr_num, "조건식이름":condition_name}
            self.registed_conditions_list.append(condition_name)

        elif result !=1 and condition_name in self.registed_conditions_list:
            print(f"{condition_name} 조건식 이미 등록 완료")
            self.registed_condition_df.loc[condition_idx] = {"화면번호": scr_num, "조건식이름": condition_name}
        else:
            print(f"{condition_name} 조건검색 등록 실패!")

    def send_condition_stop(self, scr_num, condition_name, condition_idx):
        print(f"{condition_name} 조건검색 실시간 해제!")
        #nSearch: 조회구분, 0:조건검색, 1:실시간 조건검색
        self.kiwoom.dynamicCall(
            "SendConditionStop(QString, QString, int)",
             scr_num, condition_name, condition_idx
        )


    def _receive_real_condition(self, strCode, strType, strConditionName, strConditionIndex):

        # strType: 이벤트 종류, "I":종목편입, "D", 종목미탈
        # strConditionName: 조건식 미듬
        # strConditionIndex: 조건명 인덱스
        print(f"Received real condition, {strCode}, {strType}, {strConditionName}, {strConditionIndex}")

        if strConditionIndex.zfill(3) not in self.registed_condition_df.index.to_list():
            print(f"조건명: {strConditionName}, 편입 조건식에 해당 안됨 Pass")
            return

        if strType == "I" and strCode not in self.realtime_watchList_df.index:
            if strCode not in self.realtime_registered_codes:
                self.register_code_to_realtime_list(strCode) #실시간 체결등록
            name = self.kiwoom.dynamicCall("GetMasterCodeName(QString",[strCode])
            # current_price = int(self.kiwoom.dynamicCall("GetMasterLastPrice(QString", [strCode]))
            # goal_price = current_price * (1 + float(self.goalReturnLineEdit.text())/ 100)
            # stoploss_price = current_price * (1 + float(self.stopLossLineEdit.text()) / 100)

            self.realtime_watchList_df.loc[strCode] = {
                '종목명': name,
                '현재가': None,
                '평균단가': None,
                '목표가': None,
                '손절가': None,
                '수익률': None,
                '매수기반조건식': strConditionName,
                '보유수량': 0,
                '매수주문완료여부': False
            }



    def get_comm_realdata(self, strCode, nFid):
        return self.kiwoom.dynamicCall("GetCommRealData(QString, int)", strCode, nFid)

    def _receive_realdata(self, sJongmokCode, sRealType, sRealData):
        if sRealType == "주식체결":
            self.now_time = datetime.datetime.now()
            now_price = int(self.get_comm_realdata(sRealType, 10).replace('-', '')) #현재가
            최우선매수호가 = int(self.get_comm_realdata(sRealType, 28).replace('-', ''))  # 최우선 매수호가
            self.stock_code_to_price_info_dict[sJongmokCode] = 최우선매수호가
            if sJongmokCode in self.realtime_watchList_df.index.to_list():
                if not self.realtime_watchList_df.loc[sJongmokCode, "매수주문완료여부"]:
                    goal_price = now_price * (1 + float(self.goalReturnLineEdit.text()) / 100)
                    stoploss_price = now_price * (1 + float(self.stopLossLineEdit.text()) / 100)
                    self.realtime_watchList_df.loc[sJongmokCode, "목표가"] = goal_price
                    self.realtime_watchList_df.loc[sJongmokCode, "손절가"] = stoploss_price
                    order_amount = int(self.buyAmountLineEdit.text()) // now_price
                    if order_amount < 1:
                        print(f"종목코드: {sJongmokCode}, 주문수량 부족으로 매수 진행 X")
                        return
                    # 매수 주문 진행
                    self.orders_queue.put(
                        [
                            "시장가매수주문",
                            self._get_screen_num(),
                            self.account_num,
                            1,
                            sJongmokCode,
                            order_amount,
                            "",
                            "03",
                            "",
                        ],
                    )
                    self.realtime_watchList_df.loc[sJongmokCode, "매수주문완료여부"] = True

                self.realtime_watchList_df.loc[sJongmokCode, '현재가'] = now_price

                mean_buy_price = self.realtime_watchList_df.loc[sJongmokCode,'평균단가']

                if mean_buy_price is not None:
                    self.realtime_watchList_df.loc[sJongmokCode, '수익률'] = round(
                        (now_price - mean_buy_price) / mean_buy_price * 100 - 0.87,
                        2,
                    )



                보유수량 = int(self.realtime_watchList_df.loc[sJongmokCode, '보유수량'])
                if 보유수량 > 0 and now_price < self.realtime_watchList_df.loc[sJongmokCode, '손절가']:
                    print(f"종목코드: {sJongmokCode} 매도 진행!(손절)")
                    주문가격 = self.stock_code_to_price_info_dict.get(sJongmokCode, None)
                    if not 주문가격:
                        print(f"종목코드 : {sJongmokCode}, 최우선매수호가X 주문실패!")
                        return
                    self.orders_queue.put(
                        [
                            "매도주문",
                            self._get_screen_num(),
                            self.account_num,
                            2,
                            sJongmokCode,
                            self.realtime_watchList_df.loc[sJongmokCode, '보유수량'],
                            주문가격,
                            "00",
                            "",
                        ],
                    )
                    self.realtime_watchList_df.drop(sJongmokCode, inplace=True)
                elif 보유수량>0 and now_price > self.realtime_watchList_df.loc[sJongmokCode, '목표가']:
                    #지정가 -> 10초 이상 미체결 발생시 시장가 정정주문
                    print(f"종목코드: {sJongmokCode} 매도 진행!(익절)")

                    self.orders_queue.put(
                        [
                            "매도주문",
                            self._get_screen_num(),
                            self.account_num,
                            2,
                            sJongmokCode,
                            self.realtime_watchList_df.loc[sJongmokCode, '보유수량'],
                            now_price,
                            "00",
                            "",
                        ],
                    )
                    self.realtime_watchList_df.drop(sJongmokCode, inplace=True)




    def send_orders(self):
        self.now_time = datetime.datetime.now()

        if self.is_check_tr_req_condition() and not self.orders_queue.empty():
            sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo = self.orders_queue.get()
            ret = self.send_order(sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo)
            if ret == 0:
                print(f"{sRQName} 주문 접수 성공!")
            self.last_tr_send_times.append(self.now_time)

    def send_order(self, sRQName, sScreenNo, sAccNo,nOrderType, sCode, nQty, nPrice, sHogaGb, s0rgOrderNo):
        print("Sending order")
        return self.kiwoom.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                                   [sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, s0rgOrderNo])
        # sRQName: 사용자 구분명 -> OnReceiveTrData에서 받을 미듬으로!
        # sAccNo: 18자리!
        # nOrderType: 주문유형 1:신규매수, 2:신규매도 3:매수취소, 4:매도취소, 5:매수정정, 6:매도정정
        # sCode: 종목코드, nQty: 수량, nPrice: 가격,
        # sHogaGb: 거래구분(혹믄 호가구분)믄 마래 참고
        # sOrgOrderNo: 원주문번호입니다. 신규주문에는 공백, 정정(취소)주문할 원주문번호를 입력합니다.
        # 80: 지정가, 03: 시장가, 05: 조건부지정가, 06: 죄유리지정가, 07: 죄우선지정가, 10: 지정가IOC
        # 13: 시장가IOC, 16: 최유리IOC, 28: 지정가FOK, 23: 시장가FOK, 26: 최유리FOK, 61: 장전시간의종가
        # 62: 시간외단일가매매, 81: 장후시간의종가
        # 예시, SendOrder("주식주문", get_screen_num2(), cbo계좌.Text.Trim(), 1, 종목코드, 수량, 현재가, "00", "");





    def set_real(self, scrNum, strCodeList, strFidList, strRealType):
        self.kiwoom.dynamicCall("SetRealReg(QString, QString, QString, QString)", scrNum, strCodeList, strFidList,
                                strRealType)

    def register_code_to_realtime_list(self, code):
        fid_list = "10;12;20;28"

        if len(code) != 0:
            self.realtime_registered_codes.add(code)
            self.set_real(self._get_screen_num(), code, fid_list, "1")
            print(f"{code}, 실시간 등록 완료!")

    def is_check_tr_req_condition(self):
        now_time = datetime.datetime.now()

        if len(self.last_tr_send_times) >= self.max_send_per_sec and \
            now_time - self.last_tr_send_times[-self.max_send_per_sec] < datetime.timedelta(milliseconds=1000):
            print(f"초 단위 TR 요정 제한! Wait for time to send!")
            return False
        elif len(self.last_tr_send_times) >= self.max_send_per_minute and \
            self.now_time - self.last_tr_send_times[-self.max_send_per_minute] < datetime.timedelta(minutes=1):
            print(f"분 단위 TR 요청 제한! Wait for time to send!")
            return False
        elif len(self.last_tr_send_times) >= self.max_send_per_hour and \
            now_time - self.last_tr_send_times[-self.max_send_per_hour] < datetime.timedelta(minutes=60):
            print(f"시간 단위 TR 요청 제한! Wait for time to send!")
            return False
        else:
            return True

    @staticmethod
    def get_sell_price(now_price):
        now_price = int(now_price)

        if now_price < 2000:
            return now_price
        elif 5000 > now_price >= 2000:
            return now_price - now_price % 5
        elif now_price >= 5000 and now_price < 20000:
            return now_price - now_price % 10
        elif now_price >= 20000 and now_price < 50000:
            return now_price - now_price % 50
        elif now_price >= 50000 and now_price < 200000:
            return now_price - now_price % 100
        elif now_price >= 200000 and now_price < 500000:
            return now_price - now_price % 500
        else:
            return now_price - now_price % 1000

sys._excepthook = sys.excepthook

def my_exception_hook(exctype,value,traceback):
    # Print the error and traceback
    print(exctype, value, traceback)
    # Call the normal Exception hook after
    sys.excepthook(exctype, value, traceback)
    sys.exit(1)

# Set the exception hook to our wrapping function
sys.excepthook = my_exception_hook

if __name__ == "__main__":
    app = QApplication(sys.argv)
    kiwoom_api = KiwoomAPI()
    sys.exit(app.exec_())