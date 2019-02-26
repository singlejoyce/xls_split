# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'mainui.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!
import os
import sys
import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import *

__author__ = "joyce"


class Example(QWidget):
    def __init__(self):
        super().__init__()
        self.cwd = os.getcwd()  # 获取当前程序文件位置
        # self.cwd = "D:\\project\\LoveDance_N1\\data\\xlsx"
        self.xls_dir_list = []
        self.selectButton = QtWidgets.QPushButton(self)
        self.startButton = QtWidgets.QPushButton(self)
        self.sheetname = QtWidgets.QTextEdit(self)
        self.xls_dir = QtWidgets.QTextEdit(self)
        self.sheetlabel = QtWidgets.QLabel(self)
        self.xlsdirlabel = QtWidgets.QLabel(self)
        self.initUI()

    def initUI(self):
        self.setGeometry(100, 100, 1066, 784)
        self.setWindowTitle('xls拆分导出txt工具')
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(14)
        self.xls_dir.setFont(font)
        self.xls_dir.setGeometry(QtCore.QRect(150, 30, 691, 351))
        self.xls_dir.setObjectName("xls_dir")
        self.sheetname.setFont(font)
        self.sheetname.setGeometry(QtCore.QRect(150, 410, 691, 321))
        self.sheetname.setObjectName("sheetname")
        self.sheetlabel.setGeometry(QtCore.QRect(30, 410, 111, 31))
        self.sheetlabel.setFont(font)
        self.sheetlabel.setObjectName("sheetlabel")
        self.xlsdirlabel.setGeometry(QtCore.QRect(40, 20, 121, 51))
        self.xlsdirlabel.setFont(font)
        self.xlsdirlabel.setObjectName("xlsdirlabel")
        self.selectButton.setGeometry(QtCore.QRect(860, 40, 171, 51))
        self.selectButton.setFont(font)
        self.selectButton.setObjectName("selectButton")
        self.selectButton.clicked.connect(self.btn_chooseMutiFile)
        self.startButton.setFont(font)
        self.startButton.setGeometry(QtCore.QRect(860, 520, 171, 51))
        self.startButton.setObjectName("startButton")
        self.startButton.clicked.connect(self.start)
        self.retranslateUi()
        QtCore.QMetaObject.connectSlotsByName(self)
        self.show()

    def retranslateUi(self):
        _translate = QtCore.QCoreApplication.translate
        self.startButton.setText(_translate("Dialog", "开始"))
        self.selectButton.setText(_translate("Dialog", "选择xls文件"))
        self.sheetlabel.setText(_translate("Dialog", "sheetname："))
        self.xlsdirlabel.setText(_translate("Dialog", "xls文件地址："))

    def closeEvent(self, event):
        reply = QMessageBox.question(self, '提示', '确认退出吗？',
                                     QMessageBox.Ok | QMessageBox.Cancel, QMessageBox.Cancel)
        if reply == QMessageBox.Ok:
            event.accept()
        elif reply == QMessageBox.Cancel:
            event.ignore()

    def process(self, xls_dir):
        xls = pd.ExcelFile(xls_dir)
        sheet_names = xls.sheet_names
        sheet_names_show = os.path.basename(xls_dir) + "表名显示为：\n" + str(sheet_names) + "\n"
        self.sheetname.append(sheet_names_show)
        txt_save_path = os.getcwd() + '\\reslut\\' + os.path.basename(xls_dir) + "\\"
        if not os.path.isdir(txt_save_path):
            os.makedirs(txt_save_path)
        for sheet_name in sheet_names:
            txt_dir = txt_save_path + sheet_name + ".txt"
            # tempsheet.to_excel(reslutdir + sheet_name + ".xlsx", index=False)  # write data to excel
            with open(txt_dir, 'w', encoding='utf-8') as f:
                tempsheet = pd.read_excel(xls_dir, sheet_name=sheet_name)
                f.write(tempsheet.to_string())

    def start(self):
        sheet_names_show = ""
        if isinstance(self.xls_dir_list, list):
            for xls_dir in self.xls_dir_list:
                self.process(xls_dir)
        elif isinstance(self.xls_dir_list, str):
            self.process(self.xls_dir_list)
        # 弹出完成提示框
        QMessageBox.information(self, "提示", self.tr("xls拆分完成!"), QMessageBox.Ok)

    def btn_chooseFile(self):
        self.xls_dir_list, filetype = QFileDialog.getOpenFileName(self,
                                                                  "选取文件",
                                                                  self.cwd,  # 起始路径
                                                                  "Excel Files(*.xls *.xlsx)")  # 设置文件扩展名过滤,用双分号间隔
        if self.xls_dir_list == "":
            print("\n取消选择")
            return
        self.xls_dir.setPlainText(self.xls_dir_list)

    def btn_chooseMutiFile(self):
        self.xls_dir_list, filetype = QFileDialog.getOpenFileNames(self,
                                                                   "多文件选择",
                                                                   self.cwd,  # 起始路径
                                                                   "Excel Files(*.xls *.xlsx)")
        if len(self.xls_dir_list) == 0:
            print("\n取消选择")
            return

        for file in self.xls_dir_list:
            self.xls_dir.append(file)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())
