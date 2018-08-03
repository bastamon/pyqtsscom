# -*-coding:utf-8 -*-
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog
from pyqtsscom import Ui_Form
import analysisV2
import re
import win32api
import win32gui
import sys
import os


class mywindow(QtWidgets.QWidget, Ui_Form):
    def __init__(self):
        super(mywindow, self).__init__()
        self.setupUi(self)
        # self.filepath.setText("SaveWindows2018_7_31_7-57-57.TXT")
        # filepathstr = self.filepath.text()
        # print(filepathstr)

        self.analyzeButton.clicked.connect(self.analyzeentry)
        self.openxlsx.clicked.connect(self.openxls)
        self.browseButton.clicked.connect(self.setBrowerPath)

    def openxls(self):
        filepathstr = self.getfilepath()
        pattern = r'(.+?)\.'
        pathname = "".join(re.findall(pattern, filepathstr, flags=re.IGNORECASE)) + '.xlsx'
        win32api.ShellExecute(0, 'open', pathname, '', '', 1)

    def getfilepath(self):
        return self.filepath.text()

    def analyzeentry(self):
        if self.proId.text() == '':
            assert self.proId.text() == '', 'need protId'
            # analysisV2.analyze(self.getfilepath(), self.proId.text())
        else:
            self.proId.setText('05')
        analysisV2.analyze(self.getfilepath(), self.proId.text())
        self.analyzeButton.setEnabled(False)
        self.openxlsx.setEnabled(True)

    def setBrowerPath(self):
        fileName1, filetype = QFileDialog.getOpenFileName(self, "选取文件", os.getcwd(),
                                                          "Text Files (*.txt);;All Files (*)")
        self.filepath.setText(fileName1)
        filepathstr = self.getfilepath()
        pattern = r'(.+?)\.'
        pathname = "".join(re.findall(pattern, filepathstr, flags=re.IGNORECASE)) + '.db'
        if os.path.exists(pathname):
            self.analyzeButton.setEnabled(False)
            self.infolabel.setStyleSheet("color: rgb(255, 0, 127);")
            self.infolabel.setText("文件已解析完毕，可直接打开")
            self.openxlsx.setEnabled(True)
        else:
            self.analyzeButton.setEnabled(True)


if __name__ == '__main__':
    ct = win32api.GetConsoleTitle()
    hd = win32gui.FindWindow(0,ct)
    win32gui.ShowWindow(hd,0)
    app = QtWidgets.QApplication(sys.argv)
    myshow = mywindow()
    myshow.show()
    sys.exit(app.exec_())
