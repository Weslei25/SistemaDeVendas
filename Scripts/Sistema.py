import sys, os
from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *


class MyApp(QWidget):

    def __init__(self):
        super().__init__()
        self.setFocusPolicy(QtCore.Qt.ClickFocus)
        self.initUI()

    def initUI(self):
        vbox = QVBoxLayout()
        vbox.addStretch(2)
        btn = QPushButton("Test")
        btn.setToolTip("This tooltip")
        vbox.addWidget(btn)
        vbox.addStretch(1)

        self.setLayout(vbox)
        self.setGeometry(300, 300, 300, 200)
        self.show()

    def focusOutEvent(self, event):
        self.setFocus(True)
        self.activateWindow()
        self.raise_()
        self.show()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())