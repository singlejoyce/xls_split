import os
import shutil
import time
import webbrowser
import win32com.client

# 对比报告、拆分表格存放的位置
temp_path = r'D:\temp'
# 需配置eyond Compare 4路径
BComparePath = r'"C:\Program Files\Beyond Compare 4\BCompare.exe"'
# 左边的对比版本
pathL = r'D:\project\LoveDance_N1\data_n1'
# 右边的对比版本
pathR = r'D:\project\LoveDance_P1\data_p1'

comparescript = {
    'data': os.path.abspath("scripts\\datacompare.txt"),
}

TestPath = os.path.join(temp_path, time.strftime('%Y%m%d-%H%M%S', time.localtime(time.time())))
tempL = os.path.join(TestPath, os.path.basename(pathL))
tempR = os.path.join(TestPath, os.path.basename(pathR))


# 遍历目录下的所有文件，self.files属性中保存了所有文件相对路径的集合
class dirinfo(object):
    def __init__(self, path):
        self.path = path
        self.files = set()
        self.getfiles()

    def getfiles(self, path=''):
        list = os.listdir(os.path.join(self.path, path))
        for i in list:
            if not i.startswith("."):  # 排除隐藏文件
                if os.path.isdir(os.path.join(self.path, path, i)):
                    self.getfiles(os.path.join(path, i))
                else:
                    self.files.add(os.path.join(path, i))


# 返回path1独有文件，path2独有文件，两边有差异的文件。
def getdiff(path1, path2):
    l = dirinfo(path1).files
    r = dirinfo(path2).files
    df = [x for x in l & r if open(os.path.join(path1, x), 'rb').read() != open(os.path.join(path2, x), 'rb').read()]
    return l - r, r - l, df


def copyfile(path, outpath, name):
    if not os.path.exists(outpath):
        os.makedirs(outpath)
    shutil.copy(path, os.path.join(outpath, name))


class Excel(object):
    def __init__(self):
        self.excel = win32com.client.Dispatch('Excel.Application')

    def split(self, file, outpath):
        xlbook = self.excel.Workbooks.Open(file)
        for sheet in xlbook.Sheets:
            if sheet.Name == "商城":
                continue
            newbook = self.excel.Workbooks.Add()
            sheet.Copy(newbook.Sheets.Item(1))

            if not os.path.exists(outpath):
                os.makedirs(outpath)
            filename = os.path.join(outpath, os.path.splitext(os.path.basename(file))[0] + "_" + sheet.name + '.xlsx')
            print(filename)
            newbook.SaveAs(filename)
            newbook.Close()
        xlbook.Close()


def show_data_diff(path1, path2, showlog=False):
    data_pathL = os.path.join(path1, 'xlsx')
    if not os.path.exists(data_pathL):
        data_pathL = path1
    data_pathR = os.path.join(path2, 'xlsx')
    if not os.path.exists(data_pathR):
        data_pathR = path2
    if data_pathL and data_pathR:
        onlyl, onlyr, diff = getdiff(data_pathL, data_pathR)
        if onlyl:
            for f in onlyl:
                copyfile(os.path.join(data_pathL, f), tempL, os.path.basename(f))
        if onlyr:
            for f in onlyr:
                copyfile(os.path.join(data_pathR, f), tempR, os.path.basename(f))
        if diff:
            table = Excel()
            for f in diff:
                if ".xlsx" in f:
                    table.split(os.path.join(data_pathL, f), tempL)
                    table.split(os.path.join(data_pathR, f), tempR)
                else:
                    copyfile(os.path.join(data_pathL, f), tempL, os.path.basename(f))
                    copyfile(os.path.join(data_pathR, f), tempR, os.path.basename(f))
            logfile = os.path.join(TestPath, "data_diff.html")
            cmd = ' '.join([BComparePath,
                            r'/silent',
                            "@" + comparescript['data'],
                            tempL,
                            tempR,
                            os.path.join(TestPath, logfile)])
            print(cmd)
            os.system(cmd)
            if showlog:
                webbrowser.open(os.path.join(TestPath, logfile))


if __name__ == "__main__":
    show_data_diff(pathL, pathR, True)

