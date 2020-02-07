# -*- coding: utf-8 -*-
from PyQt5.QtWidgets import QMessageBox, QInputDialog, QWidget

from Matcher import Matcher
from design import *


class MainWidgets(QtWidgets.QMainWindow):
    def __init__(self):
        QtWidgets.QWidget.__init__(self)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.buttonpress)

    @staticmethod
    def showComplete():
        msg = QMessageBox()
        msg.setText("Complete!")
        msg.exec_()

    def buttonpress(self):
        a = self.ui.lineEdit.text()
        b = self.ui.lineEdit_2.text()
        c = self.ui.lineEdit_3.text()
        d = self.ui.lineEdit_4.text()
        matcher = Matcher()
        Matcher.setdirectory(a)
        Matcher.setreportsdirectory(b)
        Matcher.setdrugsfilepath(c)
        Matcher.setgarbagefilepath(d)
        matcher.rename_drugs()
        self.showComplete()
