# -*- coding: utf-8 -*-
import sys
from design import *
import MainWidget
app = QtWidgets.QApplication(sys.argv)
myapp = MainWidget.MainWidgets()
myapp.show()
sys.exit(app.exec_())