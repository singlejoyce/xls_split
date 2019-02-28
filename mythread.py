# -*- coding:utf-8 -*-
import os

import pandas as pd
from PyQt5 import QtCore
import time


__author__ = "joyce"


class MyThread(QtCore.QThread):
    def __init__(self, parent=None):
        super(MyThread, self).__init__(parent)
        self.finish_state = False

    update_progressBar_signal = QtCore.pyqtSignal(str)

    def run(self):
        step = 0
        while True:
            step = step + 5
            time.sleep(1)
            # print("self.finish_state = ", self.finish_state)
            if not self.finish_state:
                self.update_progressBar_signal.emit(str(step))
            else:
                break


class WorkThread(QtCore.QThread):
    def __init__(self, xls_dir_list):
        super(WorkThread, self).__init__()
        self.xls_dir_list = xls_dir_list

    finish_state_signal = QtCore.pyqtSignal(str)

    def run(self):
        sheet_names_show = ""
        for xls_dir in self.xls_dir_list:
            xls = pd.ExcelFile(xls_dir)
            sheet_names = xls.sheet_names
            txt_save_path = os.getcwd() + '\\reslut\\' + os.path.basename(xls_dir) + "\\"
            if not os.path.isdir(txt_save_path):
                os.makedirs(txt_save_path)
            for sheet_name in sheet_names:
                txt_dir = txt_save_path + sheet_name + ".txt"
                tempsheet = pd.read_excel(xls_dir, sheet_name=sheet_name)
                # 处理格式问题，强制将所有的float64格式转换为int型
                tempsheet.fillna(0, inplace=True)
                for column in tempsheet.columns:
                    if tempsheet.dtypes[column].name == "float64":
                        tempsheet[column] = tempsheet[column].map(int)
                # 保存文件至.txt
                with open(txt_dir, 'w', encoding='utf-8') as f:
                    f.write(tempsheet.to_string())
            sheet_names_show = sheet_names_show + os.path.basename(xls_dir) + "表名显示为：\n" + str(
                sheet_names) + "\n"

        self.finish_state_signal.emit(sheet_names_show)  # 处理完毕后发出信号





