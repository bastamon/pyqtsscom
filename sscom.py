# -*-coding:utf-8 -*-
from PyQt5 import QtWidgets
from pyqtsscom import Ui_Form
import analysisV2
import re
import win32api
import sys




class mywindow(QtWidgets.QWidget, Ui_Form):
    def __init__(self):
        super(mywindow, self).__init__()
        self.setupUi(self)
        # self.filepath.setText("SaveWindows2018_7_31_7-57-57.TXT")
        filepathstr = self.filepath.text()
        print(filepathstr)

        self.analyzeButton.clicked.connect(self.analyzeentry)
        # self.browseButton.clicked.connect()
        # self.filename.textChanged.connect()
        self.openxlsx.clicked.connect(self.openxls)


    def openxls(self):
        filepathstr = self.getfilepath()
        pattern = r'(.+?)\.'
        pathname = "".join(re.findall(pattern, filepathstr.encode('utf-8'), flags=re.IGNORECASE))+'.xlsx'
        win32api.ShellExecute(0, 'open', pathname, '', '', 1)


    def getfilepath(self):
        return self.filepath.text()


    def analyzeentry(self):
        if self.proId.text() == '':
            assert self.proId.text() == '', 'need protId'
            analysisV2.analyze(self.getfilepath(), self.proId.text())
        else:
            self.proId.setText('05')
            analysisV2.analyze(self.getfilepath(), self.proId.text())










if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    myshow = mywindow()
    myshow.show()
    sys.exit(app.exec_())