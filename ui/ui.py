from kiwoom.kiwoom import *

import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic
class Ui:
    def __init__(self):

        print("ui class")

        self.app = QApplication(sys.argv)

        self.kiwoom = Kiwoom()

        self.app.exec_()