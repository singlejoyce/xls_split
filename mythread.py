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
            time.sleep(5)
            # print("self.finish_state = ", self.finish_state)
            if not self.finish_state:
                self.update_progressBar_signal.emit(str(step))
            else:
                break


def export_txt(result_path, xls_dir, pos):
    xls = pd.ExcelFile(xls_dir)
    sheet_names = xls.sheet_names
    txt_save_path = os.path.join(result_path, pos, os.path.basename(xls_dir))
    if not os.path.isdir(txt_save_path):
        os.makedirs(txt_save_path)
    for sheet_name in sheet_names:
        txt_dir = os.path.join(txt_save_path, sheet_name + ".txt")
        tempsheet = pd.read_excel(xls_dir, sheet_name=sheet_name)
        # 处理格式问题，强制将所有的float64格式转换为int型
        tempsheet.fillna(0, inplace=True)
        for column in tempsheet.columns:
            if tempsheet.dtypes[column].name == "float64":
                tempsheet[column] = tempsheet[column].map(int)
        tempsheet.reset_index(drop=True, inplace=True)
        tempsheet.to_csv(txt_dir, sep='\t', index=False)


class WorkThread(QtCore.QThread):
    def __init__(self, xls_dir_list):
        super(WorkThread, self).__init__()
        self.xls_dir_list = xls_dir_list

    finish_state_signal = QtCore.pyqtSignal(str)

    def run(self):
        sheet_names_show = ""
        result_path = os.getcwd() + '\\reslut\\' + time.strftime('%Y%m%d-%H%M%S', time.localtime(time.time()))
        for xls_dir in self.xls_dir_list:
            export_txt(result_path, xls_dir, "")
            # sheet_names_show = sheet_names_show + os.path.basename(xls_dir) + "表名显示为：\n" + str(
            #     sheet_names) + "\n"

        self.finish_state_signal.emit("finished!")  # 处理完毕后发出信号


class WorkThread2(QtCore.QThread):
    def __init__(self, xls_dir_list_l, xls_dir_list_r):
        super(WorkThread2, self).__init__()
        self.xls_dir_list_l = xls_dir_list_l
        self.xls_dir_list_r = xls_dir_list_r

    finish_state_signal = QtCore.pyqtSignal(str)

    def run(self):
        result_path = os.getcwd() + '\\reslut\\' + time.strftime('%Y%m%d-%H%M%S', time.localtime(time.time()))
        for xls_dir in self.xls_dir_list_l:
            export_txt(result_path, xls_dir, "Left")

        for xls_dir in self.xls_dir_list_r:
            export_txt(result_path, xls_dir, "Right")

        self.finish_state_signal.emit("finished!")  # 处理完毕后发出信号


