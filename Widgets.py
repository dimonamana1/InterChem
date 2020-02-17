# -*- coding: utf-8 -*-
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QInputDialog, QWidget, QMessageBox


class Widgets(QWidget):

    def __init__(self):
        super().__init__()
        self.setWindowTitle('Renamer')
        self.setWindowIcon(QIcon('12.ico'))

    def showDialog1(self, key, value):  # 1st step response dialog
        text, ok = QInputDialog.getText(self, 'Input Dialog', key + " = " + value)
        if ok:
            return str(text)

    def showDialog2(self, drug1, drug2, drug3):  # 2nd step response dialog
        text, ok = QInputDialog.getText(self, 'Input Dialog',
                                        "Choose one: " + "\n" + "1: " + drug1 + "\n" + " 2: " + drug2 + "\n" + "3: " +
                                        drug3)
        if ok:
            return str(text)

    def showDialog(self):  # 3rd step response dialog
        text, ok = QInputDialog.getText(self, 'Input Dialog', "Insert our own:")
        if ok:
            return str(text)

    @staticmethod
    def showError(filenm):  # error message
        msg = QMessageBox()
        msg.setWindowTitle('Renamer')
        msg.setWindowIcon(QIcon('12.ico'))
        msg.setText(filenm + " not renamed")
        msg.exec_()

    @staticmethod
    def showFNF():  # error message
        msg = QMessageBox()
        msg.setWindowTitle('Renamer')
        msg.setWindowIcon(QIcon('12.ico'))
        msg.setText("File or directory does not exist, please try again")
        msg.exec_()
