from kiwoom.kiwoom import *

import sys
from PyQt5.QtWidgets import *


class Ui_class():
    def __init__(self):
        print("Ui_class 입니다")

        self.app = QApplication(sys.argv)

        self.kiwoom = Kiwoom()

        self.app.exec_()  # 프로그램 종료되지 않게 유지
