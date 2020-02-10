# -*- coding: utf-8 -*-
from PyQt5.QtWidgets import QInputDialog, QWidget, QMessageBox


class Widgets(QWidget):

    def __init__(self):
        super().__init__()

    def showDialog1(self, key, value):
        text, ok = QInputDialog.getText(self, 'Input Dialog', key + " = " + value)
        if ok:
            return str(text)

    def showDialog2(self, drug1, drug2, drug3):
        text, ok = QInputDialog.getText(self, 'Input Dialog',
                                        "Choose one: " + "\n" + "1: " + drug1 + "\n" + " 2: " + drug2 + "\n" + "3: " +
                                        drug3)
        if ok:
            return str(text)

    def showDialog(self):
        text, ok = QInputDialog.getText(self, 'Input Dialog', "Insert our own:")
        if ok:
            return str(text)

    @staticmethod
    def showError(filenm):
        msg = QMessageBox()
        msg.setText(filenm + "not renamed")
        msg.exec_()
