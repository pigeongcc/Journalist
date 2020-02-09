import sys
from PyQt5 import QtWidgets, QtGui
import Ui_design_1

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    ui = Ui_design_1.Ui_Dialog()
    sys.exit(app.exec_())