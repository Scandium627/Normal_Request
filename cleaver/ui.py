# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'main1.ui'
##
## Created by: Qt User Interface Compiler version 5.15.2
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################



from PyQt5.Qt import *
class Ui_Form(object):
    def setupUi(self, Form):
        if not Form.objectName():
            Form.setObjectName(u"Form")
        Form.resize(819, 435)
        self.label_4 = QLabel(Form)
        self.label_4.setObjectName(u"label_4")
        self.label_4.setGeometry(QRect(10, 140, 101, 31))
        self.tableView = QTableView(Form)
        self.tableView.setObjectName(u"tableView")
        self.tableView.setGeometry(QRect(10, 171, 381, 221))
        self.label_3 = QLabel(Form)
        self.label_3.setObjectName(u"label_3")
        self.label_3.setGeometry(QRect(420, 140, 101, 31))
        self.tableView_2 = QTableView(Form)
        self.tableView_2.setObjectName(u"tableView_2")
        self.tableView_2.setGeometry(QRect(420, 170, 381, 221))
        self.textBrowser = QTextBrowser(Form)
        self.textBrowser.setObjectName(u"textBrowser")
        self.textBrowser.setGeometry(QRect(10, 80, 331, 31))
        self.pushButton = QPushButton(Form)
        self.pushButton.setObjectName(u"pushButton")
        self.pushButton.setGeometry(QRect(10, 50, 131, 21))
        self.pushButton_3 = QPushButton(Form)
        self.pushButton_3.setObjectName(u"pushButton_3")
        self.pushButton_3.setGeometry(QRect(710, 410, 75, 23))
        self.label_6 = QLabel(Form)
        self.label_6.setObjectName(u"label_6")
        self.label_6.setGeometry(QRect(10, 10, 261, 41))
        self.pushButton_2 = QPushButton(Form)
        self.pushButton_2.setObjectName(u"pushButton_2")
        self.pushButton_2.setGeometry(QRect(620, 410, 75, 23))

        self.retranslateUi(Form)

        QMetaObject.connectSlotsByName(Form)
    # setupUi

    def retranslateUi(self, Form):
        Form.setWindowTitle(QCoreApplication.translate("Form", u"Form", None))
        self.label_4.setText(QCoreApplication.translate("Form", u"<html><head/><body><p><span style=\" font-size:11pt; font-weight:600; color:#fa9a24;\">\u5373\u5c06\u5230\u671fCASE</span></p></body></html>", None))
        self.label_3.setText(QCoreApplication.translate("Form", u"<html><head/><body><p><span style=\" font-size:11pt; font-weight:600; color:#f40004;\">\u8d85\u671fCASE</span></p></body></html>", None))
        self.pushButton.setText(QCoreApplication.translate("Form", u"\u9009\u62e9\u76d1\u63a7Excel", None))
        self.pushButton_3.setText(QCoreApplication.translate("Form", u"Close", None))
        self.label_6.setText(QCoreApplication.translate("Form", u"<html><head/><body><p><span style=\" font-size:14pt; font-weight:600; color:#d72cf5;\">Cleaver_Sister Monitor</span></p></body></html>", None))
        self.pushButton_2.setText(QCoreApplication.translate("Form", u"Refresh", None))
    # retranslateUi

