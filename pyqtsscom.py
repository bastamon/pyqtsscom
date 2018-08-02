# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'pyqtsscom.ui'
#
# Created by: PyQt5 UI code generator 5.6
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(320, 240)
        self.browseButton = QtWidgets.QPushButton(Form)
        self.browseButton.setGeometry(QtCore.QRect(260, 80, 41, 23))
        self.browseButton.setObjectName("browseButton")
        self.analyzeButton = QtWidgets.QPushButton(Form)
        self.analyzeButton.setGeometry(QtCore.QRect(50, 190, 75, 23))
        self.analyzeButton.setObjectName("analyzeButton")
        self.openxlsx = QtWidgets.QPushButton(Form)
        self.openxlsx.setGeometry(QtCore.QRect(180, 190, 75, 23))
        self.openxlsx.setObjectName("openxlsx")
        self.file_label = QtWidgets.QLabel(Form)
        self.file_label.setGeometry(QtCore.QRect(20, 80, 31, 16))
        self.file_label.setObjectName("file_label")
        self.filepath = QtWidgets.QLineEdit(Form)
        self.filepath.setGeometry(QtCore.QRect(60, 80, 181, 21))
        self.filepath.setObjectName("filepath")
        self.proId = QtWidgets.QLineEdit(Form)
        self.proId.setGeometry(QtCore.QRect(130, 140, 113, 20))
        self.proId.setObjectName("proId")
        self.ptoId_label = QtWidgets.QLabel(Form)
        self.ptoId_label.setGeometry(QtCore.QRect(30, 140, 71, 16))
        self.ptoId_label.setObjectName("ptoId_label")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.browseButton.setText(_translate("Form", "浏览"))
        self.analyzeButton.setText(_translate("Form", "统计"))
        self.openxlsx.setText(_translate("Form", "显示结果"))
        self.file_label.setText(_translate("Form", "路径:"))
        self.ptoId_label.setText(_translate("Form", "协议识别码:"))

