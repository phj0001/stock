# -*- coding: utf-8 -*-

import sys
import os
import pickle
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from . err_Code import *
from PyQt5.QtTest import *
from . kiwoom_fid import *

class Kiwoom(QAxWidget):

    def __init__(self):
        super().__init__()

        self.realType = RealType()

        # eventloop 모음
        self.loop = QEventLoop()
        self.acct_event_loop = QEventLoop()

        # 스크린 번호
        self.scr_real = "1000"
        self.scr_acct = "2000"
        self.scr_st_anal = "3000"
        self.scr_stock = "4000"
        self.scr_order = "5000"


        # 변수모음
        self.acct_info = {}
        self.acct_num = [] # 전체 계좌번호 저장 리스트
        self.acct_cnt = None
        self.deposit = None    #예수금
        self.withdraw = None   #출금가능금액
        self.invest_amount = None   #투자액
        self.invest_percent = 0.5  #투자비율
        self.acct_eval = None  # 계좌손익
        self.acct_yield = None # 계좌수익률
        self.msg = ""

        self.pr_list = []  # 종목분석용 데이터
        self.my_stock_dict = {}  # 보유종록 목록
        self.no_deal_order_dict = {} #미체결 종목 목록
        self.portfolio_dict = {}
        self.real_deal_dict = {} # 실시간 거래 종목
        self.allcode_dict = {} # 종목 코드 저장

        # 실행 함수
        # self.create_kiwoom_instance()
        # self.event_slots()
        # self.login()
        # self.get_account_info()
        # self.get_deposit_info()
        # self.my_stock_info()  # 보유종목 평가현황
        # self.no_deal_order()  # 미체결 종목 현황
        # self.get_code() #종목코드 내려받기 -> 데이터 요청 & 받기 -> 분석하기
        # self.menu()
        # self.read_data() #파일에 저장된 데이터 불러오기
        # self.scr_setting() #스크린번호 할당
        # self.real_trading_reg()
        # self.real_event_slots()


    # 키움 오픈 API 불러오기
    def create_kiwoom_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")
        self.event_slots()
        self.login()
        self.get_account_info()
        if int(self.acct_cnt) > 1:
            for acct in self.acct_num:
                if acct[-2:] == '11':
                    # print("계좌: %s" % acct)
                    self.get_deposit_info(acct)
                    self.my_stock_info(acct)
                    info = {}
                    info['투자액'] = self.invest_amount
                    info['손익'] = self.acct_eval
                    info['계좌수익률'] = self.acct_yield
                    info['예수금'] = self.deposit
                    info['출금가능액'] = self.withdraw
                    self.acct_info[acct] = info
                    QTest.qWait(1200)

        print("보유종목: %s개" % len(self.my_stock_dict))
        for key, value in self.my_stock_dict.items():
            print(value)

    def login(self):
        self.dynamicCall("CommConnect()")
        self.loop.exec_()   #로그인 루프 실행

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)
        self.OnReceiveTrData.connect(self.tr_slot)

    def real_trading_reg(self):
        self.dynamicCall("SetRealReg(QString, QString, QString, QString)", self.scr_real, "",
                         self.realType.REALTYPE['장시작시간']['장운영시간'], "0")

    def real_event_slots(self):
        self.OnReceiveRealData.connect(self.realdata_slot)

    def login_slot(self, err_code):
        if err_code == 0:
            self.msg = "로그인 성공"
            # print(msg)
        else:
            self.msg = "로그인 실패: " + errors(err_code)[1]
            # print(msg & errors(err_code)[1])
            sys.exit(0)
        self.loop.exit()  #로그인 루프 종료

    def menu(self):
        sel = ""
        while True:
            print('[작업선택]')
            print("1. RSI 분석")
            print("2. 매수대상 종목 검색")
            print("3. 계좌평가 현황")
            print("4. 계좌잔고 조회")
            print("Q. 프로그램 종료")
            sel = input("=> ")

            if sel == "Q" or sel =="q":
                sys.exit(0)
            if sel == "1":
                st_nm = input("테스트할 종목 입력: ")
                result = self.find_code(st_nm)
                market = result[0]
                st_code = result[1]
                filelocation = "files/" + st_nm + ".pkl"
                if os.path.isfile(filelocation):
                    with open(filelocation, 'rb') as f:
                        self.pr_list = pickle.load(f)
                    print(self.pr_list)

                    # f = open(filelocation, 'rt')
                    # lines = csv.reader(f)
                    # for line in lines:
                    #     print(line)
                    # for line in lines:
                    #     lists = []
                    #     line = line.lstrip('[').rstrip(']\n')
                    #     print(type(line))
                    #     lists.append(line)
                    #     self.pr_list.append(lists)
                    #     print(self.pr_list)
                    #     break
                    # self.pr_list.append(data)
                    # print(data)
                else:
                    self.pr_list.clear()
                    print("%s 데이터 내려받기" % st_nm)
                    self.day_pr_rq(st_code)
                    self.cal_rsi(st_code)
            elif sel == "2":
                break
            elif sel == "3":
                break
            elif sel == "4":
                break

    def get_account_info(self):
        # print("[계좌정보]")
        self.acct_cnt = self.dynamicCall("GetLoginInfo(String)", "ACCOUNT_CNT")
        # print("보유계좌 수: %s" % self.acct_cnt)
        acct_list = self.dynamicCall("GetLoginInfo(String)", "ACCNO")
        self.acct_num = acct_list.split(';')
        self.acct_num.remove("")
        # self.acct_no = self.acct_num[0]
        # print("계좌번호: %s " % self.acct_num[0])  #계좌: 8157071611
        # user_id = self.dynamicCall("GetLoginInfo(String)", "USER_ID")
        # print("아이디: %s" % user_id)



    def get_deposit_info(self, acct):
        self.dynamicCall("SetInputValue(String, String)","계좌번호", acct)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        self.dynamicCall("CommRqData(String, String, int, String)", "예수금상세현황", "opw00001", "0", self.scr_acct)

        self.acct_event_loop.exec_()

    def my_stock_info(self, acct, sPrevNext = "0"): #주식 잔고 요청
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", acct)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "1")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "계좌평가잔고내역", "opw00018", sPrevNext, self.scr_acct)

        if not self.acct_event_loop.isRunning():
            self.acct_event_loop.exec_()

    def no_deal_order(self, sPrevNext = '0'):   # 미체결 종목 조회
        # print("[미체결 현황 조회]")
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.acct_num[0])
        self.dynamicCall("SetInputValue(String, String)", "전체종목구분", "0")
        self.dynamicCall("SetInputValue(String, String)", "매매구분", "0")
        self.dynamicCall("SetInputValue(String, String)", "체결구분", "1")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "미체결요청", "opt10075", sPrevNext,
                         self.scr_acct)
        self.acct_event_loop.exec()

    def scr_disconnect(self, scr_no):
        self.dynamicCall("DisconnectRealData(QString)", scr_no)

    # 종목코드 업데이트 및 일봉 자료 요청함수 day_pr_rq 호출 -> init에서 실행
    def get_code(self, market_code=None):
        market_code = 0
        code_dict = {}
        code_list = self.dynamicCall("GetCodeListByMarket(QString)", market_code)
        code_list = code_list.split(";")[:-1]
        print("코스피 종목코드 %s개 받음" % len(code_list))

        for idx, code in enumerate(code_list):
            st_nm = self.dynamicCall("GetMasterCodeName(QString)", code)
            code_dict[st_nm] = code

        self.allcode_dict[market_code] = code_dict.copy()

        market_code = 10
        code_dict.clear()
        code_list = self.dynamicCall("GetCodeListByMarket(QString)", market_code)
        code_list = code_list.split(";")[:-1]
        print("코스닥 종목코드 %s개 받음" % len(code_list))

        for idx, code in enumerate(code_list):
            st_nm = self.dynamicCall("GetMasterCodeName(QString)", code)
            code_dict[st_nm] = code

        self.allcode_dict[market_code] = code_dict.copy()

    # def get_code_list(self, market_code):  #종목코드 반환  -> get_code 함수에서 호출
    #   code_list = self.dynamicCall("GetCodeListByMarket(QString)", market_code)
    #   code_list = code_list.split(";")[:-1]
    #   return code_list

    def find_code(self, StockName):
        if StockName in self.allcode_dict[0]:
            st_code = self.allcode_dict[0][StockName]
            print("코스피종목, 코드: %s " % st_code)
            a = (0, st_code)
            return a
        elif StockName in self.allcode_dict[10]:
            st_code = self.allcode_dict[10][StockName]
            print("코스닥종목, 코드: %s" % st_code)
            a = (0, st_code)
            return a
        else:
            print("%s의 종목코드를 찾을 수 없습니다" % StockName)

    def tr_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        '''
        TR 요청을 받는 곳
        :param sScrNo: 스크린번호
        :param sRQName: 요청 이름
        :param sTrCode: TR 코드
        :param sRecordName: 사용안함
        :param sPrevNext: 다음 페이지 여부 확인
        :return:
        '''

        if sRQName == "예수금상세현황":
            money = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, "0", "예수금")
            self.deposit = format(int(money), ',')
            money = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, "0", "d+2출금가능금액")
            self.withdraw = format(int(money), ',')

            # if self.deposit != None:
            #     print("예수금: %s" % self.deposit)
            # else:
            #     print("예수금 받기 실패")
            #
            # if self.withdraw != None:
            #     print("출금가능액: %s" % self.withdraw)
            # else:
            #     print("출금가능액 받기 실패")
            #
            self.scr_disconnect(self.scr_acct)
            self.acct_event_loop.exit()

        elif sRQName == "계좌평가잔고내역":  # tr_slot

            # if (self.invest_amount == None or self.acct_yield == None
            #         or self.acct_eval == None or self.acct_yield == None):
            print("[계좌평가 현황]")
            invest_amount = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "총매입금액")
            self.invest_amount = format(int(invest_amount), ',')
            # print("투자액 %s" % self.invest_amount)

            acct_eval = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총평가손익금액")
            self.acct_eval = format(int(acct_eval), ',')
            # print("계좌손익 %s" % self.acct_eval)

            acct_yield = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "총수익률(%)")
            self.acct_yield = format(float(acct_yield)/100, ',')
            # print("계좌수익률(%%): %s" % self.acct_yield)

            cnt = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            for i in range(cnt):
                stock_code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목번호")
                stock_code = stock_code.strip()[1:]

                st_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                st_nm = st_nm.strip()

                stock_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                  "보유수량")
                stock_quantity = int(stock_quantity.strip())

                buy_pr = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입가")
                buy_pr = int(buy_pr.strip())

                stock_yield = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                 "수익률(%)")
                stock_yield = float(stock_yield.strip())
                stock_yield = stock_yield/100

                current_pr = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
                current_pr = int(current_pr.strip())

                buy_amount = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입금액")
                buy_amount = int(buy_amount.strip())

                if not stock_code in self.my_stock_dict:
                    self.my_stock_dict[stock_code] = {}

                self.my_stock_dict[stock_code].update({'종목명': st_nm})
                self.my_stock_dict[stock_code].update({'보유수량': stock_quantity})
                self.my_stock_dict[stock_code].update({'매입가': buy_pr})
                self.my_stock_dict[stock_code].update({'수익율': stock_yield})
                self.my_stock_dict[stock_code].update({'현재가': current_pr})
                self.my_stock_dict[stock_code].update({'매입금액': buy_amount})

            if sPrevNext == "2":
                self.my_stock_info(sPrevNext="2")
            else:
                self.scr_disconnect(self.scr_acct)
                self.acct_event_loop.exit()

        elif sRQName == "미체결요청":   # tr_slot
            cnt = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            print("미체결 종목 수 %s" % cnt)
            for i in range(cnt):
                stock_code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                              "종목코드")

                st_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                st_nm = st_nm.strip()
                order_kind = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문구분")
                order_vol = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문수량")
                order_pr = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문가격")
                no_deal_vol = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "미체결수량")
                deal_vol = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "체결량")
                deal_pr = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "체결가")
                order_status = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문상태")  # 접수, 확인, 체결
                order_no = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문번호")

                stock_code = stock_code.strip()
                order_kind = order_kind.strip().lstrip('+').lstrip('-')
                order_vol = int(order_vol.strip())
                order_pr = int(order_pr.strip())
                no_deal_vol = int(no_deal_vol.strip())
                deal_vol = int(deal_vol.strip())
                deal_pr = int(deal_pr.strip())
                order_status = order_status.strip()
                order_no = int(order_no.strip())

                if order_no in self.no_deal_order_dict:
                    pass
                else:
                    self.no_deal_order_dict[order_no] = {}

                ad = self.no_deal_order_dict[order_no]
                ad.update({"종목코드":stock_code})
                ad.update({"종목명": st_nm})
                ad.update({"주문구분": order_kind})
                ad.update({"주문수량": order_vol})
                ad.update({"주문가격": order_pr})
                ad.update({"미체결수량": no_deal_vol})
                ad.update({"체결량": deal_vol})
                ad.update({"주문상태": order_status})
                ad.update({"체결량": deal_vol})

            for key, value in self.no_deal_order_dict.items():
                print(value)
            self.scr_disconnect(self.scr_acct)
            self.acct_event_loop.exit()

        elif sRQName == "주식일봉차트조회":    #tr_slot> 일봉데이터 수신, 데이터 분석
            code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, "0", "종목코드")
            code = code.strip()
            cnt = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            print("%s개 수신" % cnt)

            # 전체 데이터 내려받기
            for i in range(cnt):
                data = []
                current_pr = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
                start_pr = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "시가")
                high_pr = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "고가")
                low_pr = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "저가")
                deal_vol = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "거래량")
                deal_amount = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "거래대금")
                date = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "일자")

                data.append("")
                data.append(current_pr.strip())
                data.append(start_pr.strip())
                data.append(high_pr.strip())
                data.append(low_pr.strip())
                data.append(deal_vol.strip())
                data.append(deal_amount.strip())
                data.append(date.strip())
                data.append("")

                self.pr_list.append(data.copy())

            if sPrevNext =="2":
                self.day_pr_rq(code = code, sPrevNext = sPrevNext)
            else:
                st_nm = self.dynamicCall("GetMasterCodeName(QString)", code)
                file_save = "files/" + st_nm + ".pkl"
                with open(file_save, 'wb') as f:
                    pickle.dump(self.pr_list, f, pickle.HIGHEST_PROTOCOL)
                self.scr_disconnect(self.scr_st_anal)
                self.acct_event_loop.exit()

    # 가격데이터 요청하는 함수 -> 데이터가 600개 이상이면 tr에서 재호출
    def day_pr_rq(self, code=None, date=None, sPrevNext='0'):
        QTest.qWait(3600)
        self.dynamicCall("SetInputValue(Qstring, QString)", "종목코드", code)
        self.dynamicCall("SetInputValue(Qstring, QString)", "수정주가구분", '1')
        if date != None:
            self.dynamicCall("SetInputValue(Qstring, QString)", "기준일자", date)

        self.dynamicCall("CommRqData(QSring, QString, int, QString)", "주식일봉차트조회", "opt10081", sPrevNext, self.scr_st_anal)

        if not self.acct_event_loop.isRunning():
            self.acct_event_loop.exec_()

    # 매수종목 대상 분석, menu()에서 호출
    def stock_anal(self, code):
        pass_condition = False

        if self.pr_list == None or len(self.pr_list) < 120:
            pass
        else:
            price_sum = 0
            for value in self.pr_list[:120]:
                price_sum += int(value[1])
            ma120_price = price_sum / 120  # 120일 이동 평균선의 가격 계산

            # pr_list 순서:  0 "", 1 현재가, 2 거래량, 3 거래대금,  4 기준일자, 5 시가,  6 고가, 7 저가, 8 ""
            is_stock_price_bottom = False
            today_price = None

            # 과거 20일 간의 일봉 데이터를 조회하면서 120일 이동 평균선보다 주가가 아래에 위치하는지 확인
            if int(self.pr_list[0][7]) <= ma120_price and int(self.pr_list[0][6]) >= ma120_price:
                is_stock_price_bottom = True
                today_price = int(self.pr_list[0][6])

            prev_price = None
            if is_stock_price_bottom:
                ma120_price_prev = 0
                is_stock_price_prev_top = False
                idx = 1

                while True:
                    if len(self.pr_list[idx:]) < 120:
                        break

                    price_sum = 0
                    for value in self.pr_list[idx:idx + 120]:
                        price_sum += int(value[1])
                    ma120_price_prev = price_sum / 120

                    if ma120_price_prev <= int(self.pr_list[idx][6]) and idx <= 20:
                        break

                    if int(self.pr_list[idx][7] > ma120_price_prev and idx > 20):
                        is_stock_price_prev_top = True
                        prev_price = int(self.pr_list[idx][7])
                        break
                    idx += 1

                if is_stock_price_prev_top:
                    if ma120_price > ma120_price_prev and today_price > prev_price:
                        pass_condition = True

            if pass_condition:
                stock_name = self.dynamicCall("GetMasterCodeName(QString", code)
                print("%s 분석결과 충족" % stock_name)
                f = open("files/pass_stock.txt", "a", encoding="UTF8")
                f.write(f"{code}\t{stock_name}\t{self.pr_list[0][1]}\n")
                f.close()
            else:
                stock_name = self.dynamicCall("GetMasterCodeName(QString", code)
                print("%s 분석결과 미충족" % stock_name)

    #검색조건을 충족한 종목을 저장한 파일에서 종목 읽어오기
    def cal_rsi(self, stock_code):
        pass

    # 검색 조건을 통과한 종목을 저장한 파일을 읽어 포트 딕션에 저장
    def read_data(self):
        if os.path.exists("files/pass_stock.txt"):
            f = open("files/pass_stock.txt","r", encoding="utf8")
            lines = f.readlines()
            for line in lines:
                if line != "":
                    ls = line.split("\t")
                    stock_code = ls[0]
                    stock_name = ls[1]
                    stock_price = int(ls[2].split("\n")[0])
                    stock_price = abs(stock_price)
                    self.portfolio_dict.update({stock_code:{"종목명":stock_name, "현재가":stock_price}})
            f.close()

    def scr_setting(self):
        stock_code_lists = []

        #계좌평가잔고내역에 있는 종목들
        for code in self.my_stock_dict.keys():
            if code not in stock_code_lists:
                stock_code_lists.append(code)

        #미체결에 있는 종목
        for order_no in self.no_deal_order_dict.keys():
            code = self.no_deal_order_dict[order_no]["종목코드"]

            if code not in stock_code_lists:
                stock_code_lists.append(code)

        #포트폴리오에 있는 종목들들
        for code in self.portfolio_dict.keys():
            if code not in stock_code_lists:
                stock_code_lists.append(code)

        #스크린 번호 할당
        cnt = 0
        for code in stock_code_lists:
            real_scr = int(self.scr_stock)
            trade_scr = int(self.scr_order)

            if (cnt % 50) == 0:
                real_scr += 1
                self.scr_stock = str(real_scr)

            if (cnt % 50)  == 0:
                trade_scr += 1
                self.scr_order = str(trade_scr)

            if code in self.portfolio_dict.keys():
                self.portfolio_dict[code].update({"종목스크린": str(self.scr_stock)})
                self.portfolio_dict[code].update({"주문스크린": str(self.scr_order)})

            elif code not in self.portfolio_dict.keys():
                self.portfolio_dict.update({code: {"주문스크린": str(self.scr_order), "종목스크린": str(self.scr_stock)}})

            cnt += 1

        print(self.portfolio_dict)


    def realdata_slot(self, sCode, sRealType, sRealData):
        if sRealType == "장시작시간":
            fid = self.realType.REALTYPE[sRealType]["장운영구분"]
            value = self.dynamicCall("GetCommData(QString, int)", sCode, fid)

            if value == "0":
                print("장 시작 전")
            elif value =='3':
                print("장 시작")
            elif value == '2':
                print("동시호가 시간")
            elif value == '4':
                print("장 종료")
