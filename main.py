# -*- coding: utf-8 -*-

import sys
import PyQt5
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5 import uic
from kiwoom.kiwoom import *

BtUi = 'ui/backtesting.ui'
# msg = 'ui/msgBox.ui'

class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        uic.loadUi(BtUi, self)
        self.login_pbtn1.clicked.connect(self.kiwoom_login)

    def kiwoom_login(self):
        self.k = Kiwoom()
        self.k.create_kiwoom_instance()
        # self.statusBar(self.k.msg)
        # self.statusmsg(self.k.msg)
        # msg = "보유계좌수 : " + str(self.k.account_cnt)
        # self.statusmsg(msg)
        acct1 = self.k.acct_num[0]
        acct2 = self.k.acct_num[1]
        self.acct1.addItem(acct1)
        self.acct2.addItem(acct2)
        self.investamount1.setText(self.k.acct_info[acct1]["투자액"])
        self.investamount2.setText(self.k.acct_info[acct2]["투자액"])
        self.acct_rate1.setText(self.k.acct_info[acct1]["계좌수익률"])
        self.acct_rate2.setText(self.k.acct_info[acct2]["계좌수익률"])
        self.acct_eval1.setText(self.k.acct_info[acct1]["손익"])
        self.acct_eval2.setText(self.k.acct_info[acct2]["손익"])
        self.deposit1.setText(self.k.acct_info[acct1]["예수금"])
        self.deposit2.setText(self.k.acct_info[acct2]["예수금"])
        self.withdraw1.setText(self.k.acct_info[acct1]["출금가능액"])
        self.withdraw2.setText(self.k.acct_info[acct2]["출금가능액"])

    def statusmsg(self, msg):
        msgBox = QMessageBox()
        msgBox.setWindowTitle('알림창')
        msgBox.setText("{}".format(msg))
        msgBox.addButton('확인', QMessageBox.YesRole)
        ret = msgBox.exec_()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainwindow = MainWindow()
    mainwindow.show()
    app.exec_()