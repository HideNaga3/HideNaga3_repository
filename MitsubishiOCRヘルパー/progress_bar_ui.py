# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'progress_bar_ui.ui'
#
# Created by: PyQt5 UI code generator 5.15.9
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_ProgressBarDialog(object):
    def setupUi(self, ProgressBarDialog):
        ProgressBarDialog.setObjectName("ProgressBarDialog")
        ProgressBarDialog.setEnabled(True)
        ProgressBarDialog.resize(351, 183)
        self.progressBar = QtWidgets.QProgressBar(ProgressBarDialog)
        self.progressBar.setGeometry(QtCore.QRect(30, 60, 301, 51))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")

        self.retranslateUi(ProgressBarDialog)
        QtCore.QMetaObject.connectSlotsByName(ProgressBarDialog)

    def retranslateUi(self, ProgressBarDialog):
        _translate = QtCore.QCoreApplication.translate
        ProgressBarDialog.setWindowTitle(_translate("ProgressBarDialog", "ProgressBar"))
