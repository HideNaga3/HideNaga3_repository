# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'progress_msg_ui.ui'
#
# Created by: PyQt5 UI code generator 5.15.10
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_ProgressMsgDialog(object):
    def setupUi(self, ProgressMsgDialog):
        ProgressMsgDialog.setObjectName("ProgressMsgDialog")
        ProgressMsgDialog.setEnabled(True)
        ProgressMsgDialog.resize(574, 141)
        font = QtGui.QFont()
        font.setFamily("メイリオ")
        ProgressMsgDialog.setFont(font)
        self.progress_msg_label = QtWidgets.QLabel(ProgressMsgDialog)
        self.progress_msg_label.setGeometry(QtCore.QRect(10, 30, 551, 101))
        self.progress_msg_label.setAlignment(QtCore.Qt.AlignCenter)
        self.progress_msg_label.setObjectName("progress_msg_label")
        self.label = QtWidgets.QLabel(ProgressMsgDialog)
        self.label.setGeometry(QtCore.QRect(10, 10, 81, 16))
        self.label.setObjectName("label")

        self.retranslateUi(ProgressMsgDialog)
        QtCore.QMetaObject.connectSlotsByName(ProgressMsgDialog)

    def retranslateUi(self, ProgressMsgDialog):
        _translate = QtCore.QCoreApplication.translate
        ProgressMsgDialog.setWindowTitle(_translate("ProgressMsgDialog", "ProgressBar"))
        self.progress_msg_label.setText(_translate("ProgressMsgDialog", "Progress"))
        self.label.setText(_translate("ProgressMsgDialog", "処理状況"))
