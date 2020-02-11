# -*- coding: utf-8 -*-
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QMessageBox, QFileDialog
from design import *


class MainWidgets(QtWidgets.QMainWindow):
    def __init__(self):
        QtWidgets.QWidget.__init__(self)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setWindowTitle('Renamer')
        self.setWindowIcon(QIcon('12.jpg'))
        self.mas = ["C:/Users/4r4r5/Desktop/reports/",  # paths list
                    "C:/Users/4r4r5/Desktop/reports_result/",
                    "C:/Users/4r4r5/Desktop/Препараты.csv", "C:/Users/4r4r5/Desktop/allbd.csv"]
        self.ui.pushButton.clicked.connect(self.__buttonpress)
        # shitcode = not OOP
        self.ui.pushButton_2.clicked.connect(self.__getDirectory1)  # shitcode, need to refactor
        self.ui.pushButton_3.clicked.connect(self.__getDirectory2)  # shitcode, need to refactor
        self.ui.pushButton_4.clicked.connect(self.__getFileName1)  # shitcode, need to refactor
        self.ui.pushButton_5.clicked.connect(self.__getFileName2)  # shitcode, need to refactor

    @staticmethod
    def showComplete():  # complete message
        msg = QMessageBox()
        msg.setWindowTitle('Renamer')
        msg.setWindowIcon(QIcon('12.jpg'))
        msg.setText("Complete!")
        msg.exec_()

    def __getDirectory1(self):  # shitcode, need to refactor
        dirlist = QFileDialog.getExistingDirectory(self, "Выбрать папку", ".")
        self.mas[0] = dirlist + '/'
        self.ui.label.setText(str(self.mas[0]))

    def __getDirectory2(self):  # shitcode, need to refactor
        dirlist = QFileDialog.getExistingDirectory(self, "Выбрать папку", ".")
        self.mas[1] = dirlist + '/'
        self.ui.label_2.setText(str(self.mas[1]))

    def __getFileName1(self):  # shitcode, need to refactor
        filename, filetype = QFileDialog.getOpenFileName(self,
                                                         "Выбрать файл",
                                                         ".",
                                                         "CSV Files(*.csv)")
        self.mas[2] = filename
        self.ui.label_3.setText(str(self.mas[2]))

    def __getFileName2(self):  # shitcode, need to refactor

        filename, filetype = QFileDialog.getOpenFileName(self,
                                                         "Выбрать файл",
                                                         ".",
                                                         "CSV Files(*.csv)")
        self.mas[3] = filename
        self.ui.label_4.setText(str(self.mas[3]))

    def __buttonpress(self):  # rename drugs handler
        try:
            from Matcher import Matcher
            matcher = Matcher(self.mas[0], self.mas[1], self.mas[2], self.mas[3])
            try:
                matcher.rename_drugs()
                self.showComplete()
            except IOError:
                from Widgets import Widgets
                Widgets.showFNF()
        except Exception as e:
            print(e)
