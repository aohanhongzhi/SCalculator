# coding=utf-8
import math
import os
import pandas as pd
import sys

import xlwt
from PyQt5.QtCore import QTimer, QThread, pyqtSignal
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import *

debug = False

version = '18.3-20190428'


class WorkThread(QThread):
    # 定义一个信号
    trigger = pyqtSignal(int, float, str)

    def __int__(self):
        # 初始化函数，默认
        super(WorkThread, self).__init__()

    # 重点是以0.5为单位计算的，这个只针对网点的，不针对客户
    def weightCal(self, a):
        # 检查价格表规范
        if debug == True:
            print("传入修正的重量是:%s" % (a))
        # 实际测试发现这个检测并不是很靠谱！
        if isinstance(a, float):
            if debug == True:
                print("是float类型的数字")
            b = math.ceil(a)
            if a < 5:
                if b - 0.5 >= a:
                    b = b - 0.5
            # 三元运算符  b=b-0.5>a?b-0.5:b
            if debug == True:
                print("重量是：%s" % (b))
            return b
        else:
            if debug == True:
                print("%s不是一个float类型数字" % (a))
                # QMessageBox.information(self, "错误", "导入的账单表不符合规范，第3列不是重量！")
            pass

    def run(self):
        print('新线程启动')
        print(fname)
        if fname is None:
            self.trigger.emit(1, 1, "请先导入价格！")
            return
        billData = pd.read_excel(fname)
        # data = pd.read_csv(fname)
        # print(billData)
        billArrays = billData.values
        # 读取具体的坐标值
        # print(format(billData.iloc[1, 1]))
        global saveDf, customDf
        saveDf = pd.DataFrame(columns=('网点名称', '单号', '重量', '目的地', '应收费用'))
        customDf = pd.DataFrame(columns=('网点名称', '单号', '重量', '目的地', '应收费用'))
        # 获取重量
        i = 0
        for i in range(0, billData.shape[0]):
            self.trigger.emit(i, billData.shape[0], 'a')
            # self.statusBar().showMessage('正在计算中...' + str(i + 1))
            # 取重量
            # print("%s\t%s"%(billData.iloc[i,1],billData.iloc[i,2]))
            # print(str(billData.iloc[i,2]))
            # 重量区分
            weight = billArrays[i, 2]
            try:
                weightNew = self.weightCal(weight)
            except Exception as result:
                print("发生错误，表的第%d行%s不是数字" % (i + 1, billArrays[i, 2]))
                self.trigger.emit(i, billData.shape[0], "表的第%d行%s不是数字" % (i + 1, billArrays[i, 2]))
                return

            try:
                directoryNetValue = priceNetDictionary[billArrays[i, 3].strip()]
                directoryCustomerValue = priceCustomerDictionary[billArrays[i,3].strip()]

                if debug == True:
                    print('directoryValue', directoryNetValue)
            except:
                #     QMessageBox.information(self, "错误",
                #                             "1.价格表中没有找到：%s" % billArrays[i, 3].strip())
                self.trigger.emit(i, billData.shape[0], "1.价格表中没有找到：%s" % billArrays[i, 3].strip())
                return
            # print("价格天花板：", self.weightCal(billArrays[i, 2]))

            if weightNew != None:

                #网点算法
                consquenceNet=-1
                try:
                    for j in range(0,len(directoryNetValue),4):
                        if(weight>directoryNetValue[j] and weight<=directoryNetValue[j+1]):
                            consquenceNet= weightNew * directoryNetValue[j+2]+directoryNetValue[j+3]
                            break
                    pass
                except:

                    self.trigger.emit(i, billData.shape[0], "价格表中没有找到：%s" % billArrays[i, 3].strip())
                    return
                saveDf.loc[i] = [billData.iloc[i, 0], billData.iloc[i, 1], billData.iloc[i, 2], billData.iloc[i, 3],
                                 consquenceNet]

                #客户算法
                consquenceCustomer=-1
                try:
                    for j in range(0, len(directoryCustomerValue), 4):
                        if (weight > directoryCustomerValue[j] and weight <= directoryCustomerValue[j + 1]):
                            weight=weight if weight > 0.5 else 0.5
                            consquenceCustomer= weight * directoryCustomerValue[j+2]+directoryCustomerValue[j+3]
                            #保留三位小数
                            consquenceCustomer='%.3f' % consquenceCustomer
                            '''
                            if directoryCustomerValue[j+3]=="固定":
                                consquenceCustomer = directoryCustomerValue[j + 2]
                            else:
                                consquenceCustomer = weightNew*directoryCustomerValue[j+2]+directoryCustomerValue[j+3]
                            '''
                            break
                            

                    pass
                except:
                    self.trigger.emit(i, billData.shape[0], "客户导出算法错误" % billArrays[i, 3].strip())
                    return
                customDf.loc[i] = [billData.iloc[i, 0], billData.iloc[i, 1], billData.iloc[i, 2],
                                   billData.iloc[i, 3],
                                   consquenceCustomer]
                '''
                # 网点
                try:

                    if weight <= 1 and weight >= 0:
                        consquence = weightNew * directoryValue[0] + directoryValue[1]  # 第一个[]是字典的键，第二个[]是字典值，元组的第一个值
                    elif weight > 1 and weight <= 3:
                        consquence = weightNew * directoryValue[2] + directoryValue[3]
                    elif weight > 3 and weight <= 5:
                        consquence = weightNew * directoryValue[4] + directoryValue[5]
                    elif weightNew > 5 and weightNew <= 10:
                        consquence = weightNew * directoryValue[6] + directoryValue[7]
                    elif weightNew > 10:
                        consquence = weightNew * directoryValue[8] + directoryValue[9]
                    else:
                        self.trigger.emit(i, billData.shape[0], "表的第%d行%s异常，小于0？" % (i + 1, weight))
                        print("报错!")
                    if debug == True:
                        print("%s:%s" % (billData.iloc[i, 3], consquence))

                except:

                    self.trigger.emit(i, billData.shape[0], "价格表中没有找到：%s" % billArrays[i, 3].strip())
                    return
                saveDf.loc[i] = [billData.iloc[i, 0], billData.iloc[i, 1], billData.iloc[i, 2], billData.iloc[i, 3],
                                 consquence]
                # 客户算法
                if weight <= 0.5 and weight >= 0:
                    customConsquence = directoryValue[10]
                elif weight <= 1 and weight > 0.5:
                    customConsquence = directoryValue[11]
                elif weight <= 2 and weight > 1:
                    customConsquence = directoryValue[12]
                elif weight <= 3 and weight > 2:
                    customConsquence = directoryValue[13]
                elif weight <= 4 and weight > 3:
                    customConsquence = directoryValue[14]
                elif weight > 4 and weight <= 5:
                    customConsquence = directoryValue[15]
                elif weight > 5 and weight <= 10:
                    # 5千克到10千克就是多少钱一千克，还加一个附加额，可以理解为首重
                    customConsquence = weightNew * directoryValue[16] + directoryValue[17]
                elif weight > 10:
                    # 10千克以上就是多少钱一千克，还加一个附加额
                    customConsquence = weightNew * directoryValue[18] + directoryValue[19]
                else:
                    pass
                    # self.trigger.emit(i, billData.shape[0], "数字异常,可能小于0")
                customDf.loc[i] = [billData.iloc[i, 0], billData.iloc[i, 1], billData.iloc[i, 2],
                                   billData.iloc[i, 3],
                                   customConsquence]

                # print(i + 1, "\t ", billData.shape[0], "\t", float((i + 1) / billData.shape[0]))
                # self.pbar.setValue(100 * float((i + 1) / billData.shape[0]))
        if debug == True:
            print(saveDf)
            '''


class customerSaveThread(QThread):
    # 定义一个信号
    trigger = pyqtSignal(int, str, str)

    def __int__(self):
        # 初始化函数，默认
        super(customerSaveThread, self).__init__()

    def run(self):
        print("原值保存的线程响应")
        if customerSaveFilePath[0] == "":
            print("没有保存文件")
            return
        print(customerSaveFilePath)
        saveFile = pd.ExcelWriter(customerSaveFilePath[0] + ".xls")
        customDf.to_excel(saveFile, index=None)
        saveFile.save()
        self.trigger.emit(2, "状态栏", "原值账单导出完成！")


class NetpointSaveThread(QThread):
    trigger = pyqtSignal(int, str, str)

    def __int__(self):
        # 初始化函数，默认
        super(NetpointSaveThread, self).__init__()

    def run(self):
        print("网点保存的线程响应")
        saveFile = pd.ExcelWriter(netSaveFilePath[0] + ".xlsx")
        # 解决导出文件多了一列index
        # df.style.applymap(color_negative_red).to_excel(saveFile)
        # saveDf.to_excel(saveFile, index=None)
        style = "font:colour_index red;"
        red_style = xlwt.easyxf(style)
        # 指定某一列？
        saveDf.style.applymap(self.color_red, subset=['目的地']).to_excel(saveFile, index=None)
        saveFile.save()

        self.trigger.emit(2, "状态栏", "网点账单导出完成！")

    # 声明为静态方法
    @staticmethod
    def color_red(val):
        # color = 'red' if val < 0 else 'black'
        # if isinstance(val, float):
        #     # color = 'red' if val < 1 else 'black'
        #     return 'color: %s' % color
        if debug == True:
            print(val)
        destination = str(val).strip()
        if destination == '青海' or destination == '甘肃省' or destination == '海南省' or destination == '宁夏' or destination == '新疆' or destination == '西藏':
            return 'background-color:red'
        return 'color:black'


class Example3(QMainWindow):
    # 定义全局变量
    global fname, priceData, priceNetDictionary,priceCustomerDictionary0, saveDf, customDf, customerSaveFilePath, netSaveFilePath, TimeConsuming
    fname = None
    priceData = None

    def __init__(self):
        super().__init__()
        self.initUI()

    def bill(self):
        print("导入货单的按钮响应")
        self.billPath.setText("您还没有导入账单")
        self.statusBar().showMessage('请先导入价格！', 10000)
        self.pbar.setValue(0)
        global fname, priceData, priceDictionary

        if priceData is None:
            print("None")
            QMessageBox.information(self, "警告", "请先导入价格！")
            return
        if debug == True:
            print(format(priceData.iloc[1, 1]))
        # 文本打开操作要与新线程执行操作分开
        fname, _ = QFileDialog.getOpenFileName(self, "Open", "", "(*.xls *.xlsx)")
        if fname == "":
            if debug == True:
                print("没有选择文件")
            return

        self.billPath.setText("路径：" + fname)
        if debug == True:
            print(fname)

        # df = pd.read_excel("data.xlsx", usecols=[0, 5])  # 指定读取第1列和第6列
        # 当然也可以用"A,F"代替[0,5]

        # 初始化一个定时器
        self.timer = QTimer(self)
        # 定义时间超时连接start_app
        self.timer.timeout.connect(self.start)
        # 定义时间任务是一次性任务
        self.timer.setSingleShot(True)
        # 启动时间任务
        self.timer.start()

        # 实例化一个线程
        self.work = WorkThread()
        # 多线程的信号触发连接到UpText
        self.work.trigger.connect(self.display)
        # 不加定时器，直接启动会卡死在窗口,卡死是因为当初多线程写的并不正确。
        # self.work.start()

    def display(self, i, total, msg):
        self.pbar.setValue(100 * float((i + 1) / total))
        self.lcd.display(int(i + 1))
        if i + 1 != total:
            self.statusBar().showMessage('正在计算中...' + str(i + 1))
        elif i + 1 == total:
            self.statusBar().showMessage('计算完成，一共' + str(i + 1) + '条数据')
            self.lcdtimer.stop()
        if msg != 'a':
            QMessageBox.information(self, "错误", msg)

    def start(self):
        # time.sleep(2)
        # self.textBrowser.append('test1')
        # 启动另一个线程
        self.lcdtimer.start(1000)
        self.work.start()

    def price(self):
        print("导入价格的按钮响应")
        self.pricePath.setText("您还没有导入价格表")
        fPriceName, _ = QFileDialog.getOpenFileName(self, "Open", "", "(*.xls *.xlsx)")

        if fPriceName == "":
            print("没有选择文件")
            return
        # 检查文件是否为excel文件
        print(fPriceName)
        self.pricePath.setText("路径：" + fPriceName)
        # df = pd.read_excel("data.xlsx", usecols=[0, 5])  # 指定读取第1列和第6列
        # 当然也可以用"A,F"代替[0,5]
        global priceData
        priceData = pd.read_excel(fPriceName)

        Arrays = priceData.values
        if debug == True:
            print(Arrays)
        # data = pd.read_csv(fname)
        # print(priceData)

        # 读取具体的坐标值(行，列)
        #
        a = format(priceData.iloc[0, 0])
        print(format(priceData.iloc[0, 1]))
        print(format(priceData.iloc[1, 0]))
        print(format(priceData.iloc[2, 2]))
        print(format(priceData.iloc[3, 0]))
        print(format(priceData.iloc[3, 1]))
        print(format(priceData.iloc[3, 2]))

        # 这边生成字典（不可修改）
        print("行数", int(priceData.shape[0]));
        global priceNetDictionary,priceCustomerDictionary
        priceNetDictionary = {}
        priceCustomerDictionary = {}
        try:

            # 下面是价格
            for i in range(3, priceData.shape[0]):
                # priceDictionary[priceData.iloc[i,0]]=[priceData.iloc[i,1],priceData.iloc[i,2],priceData.iloc[i,3],priceData.iloc[i,4],priceData.iloc[i,5],priceData.iloc[i,6],priceData.iloc[i,7],priceData.iloc[i,8],priceData.iloc[i,9],priceData.iloc[i,10],priceData.iloc[i,11],priceData.iloc[i,12],priceData.iloc[i,13]]
                # priceDictionary[priceData.iloc[i, 0]] = priceData.iloc[i, 1]
                # 前面两个（1,2）是网点的，（3,4,5）是客户

                # 我想着这个地方应该是，获取一行的长度，也就是列数，之后再遍历下。
                priceNetDictionaryValue = []
                priceCustomerDictionaryValue=[]
                for j in range(1, Arrays.shape[1], 2):

                    if (Arrays[0, j] == "网点"):
                        priceNetDictionaryValue.append(float(Arrays[1, j]))
                        priceNetDictionaryValue.append(float(Arrays[2, j]))
                        priceNetDictionaryValue.append(float(Arrays[i, j]))
                        priceNetDictionaryValue.append(Arrays[i, j + 1])

                    elif (Arrays[0, j] == "原值"):
                        priceCustomerDictionaryValue.append(float(Arrays[1,j]))
                        priceCustomerDictionaryValue.append(float(Arrays[2,j]))
                        priceCustomerDictionaryValue.append(float(Arrays[i,j]))
                        priceCustomerDictionaryValue.append(Arrays[i, j+1])
                        pass

                priceNetDictionary[Arrays[i, 0]] = priceNetDictionaryValue
                priceCustomerDictionary[Arrays[i,0]]=priceCustomerDictionaryValue
                # priceDictionary[Arrays[i, 0]] = [Arrays[i, 1], Arrays[i, 2], Arrays[i, 3], Arrays[i, 4], Arrays[i, 5],
                #                                  Arrays[i, 6],Arrays[i,7],Arrays[i,8],Arrays[i,9],Arrays[i,10],Arrays[i,11],Arrays[i,12],Arrays[i,13],Arrays[i,14],Arrays[i,15],Arrays[i,16],Arrays[i,17]]
        except:
            QMessageBox.information(self, "错误",
                                    "价格表不符合规范！")
        self.statusBar().showMessage('价格表导入完成，一共有' + str(priceData.shape[0]) + '个地区价格')

    # 客户保存的响应按钮
    def custSave(self):

        if priceData is None:
            QMessageBox.information(self, "警告", "没有文件可保存！")
            # self.statusBar().showMessage('请先导入价格！', 10000)
            return

        global saveDf, fname, customDf
        if fname is None:
            return
        print(fname)
        global customerSaveFilePath
        customerSaveFilePath = QFileDialog.getSaveFileName(self, "save file", os.path.splitext(fname)[0] + "原值",
                                                           "(*.xls *.xlsx)")
        self.statusBar().showMessage('原值账单正在导出，请稍后！', 10000)
        # 初始化一个定时器
        self.timer = QTimer(self)
        # 定义时间超时连接start_app
        self.timer.timeout.connect(self.start2)
        # 定义时间任务是一次性任务
        self.timer.setSingleShot(True)
        # 启动时间任务
        self.timer.start()

        # 实例化一个线程
        self.work2 = customerSaveThread()
        # 多线程的信号触发连接到UpText
        self.work2.trigger.connect(self.display2)
        # 不加定时器，直接启动会卡死在窗口
        # self.work2.start()

    def start2(self):
        # time.sleep(2)
        # 启动另一个线程
        self.work2.start()

    def display2(self, int1, str1, str2):
        if int1 == 1:
            pass
        elif int1 == 2:
            self.statusBar().showMessage('客户账单导出完成！', 10000)
        pass

    # 网点导出响应按钮
    def save(self):
        print("导出网点账单按钮的响应")
        self.statusBar().showMessage('请先导入价格！', 10000)
        if priceData is None:
            QMessageBox.information(self, "警告", "没有文件可保存！")
            self.statusBar().showMessage('请先导入价格！', 10000)
            return

        global saveDf, fname, customDf
        if fname is None:
            return
        print(fname)

        print("文件名：", os.path.splitext(fname)[0])
        global netSaveFilePath
        netSaveFilePath = QFileDialog.getSaveFileName(self, "save file", os.path.splitext(fname)[0] + "网点",
                                                      "(*.xls *.xlsx)")
        if netSaveFilePath[0] == "":
            print("没有保存文件")
            return
        print(netSaveFilePath)
        self.statusBar().showMessage('网点账单正在导出，请稍后！', 10000)
        # 初始化一个定时器
        self.timer = QTimer(self)
        # 定义时间超时连接start_app
        self.timer.timeout.connect(self.start3)
        # 定义时间任务是一次性任务
        self.timer.setSingleShot(True)
        # 启动时间任务
        self.timer.start()

        # 实例化一个线程
        self.work3 = NetpointSaveThread()
        # 多线程的信号触发连接到UpText
        self.work3.trigger.connect(self.display3)
        # 不加定时器，直接启动会卡死在窗口
        # self.work2.start()

    def start3(self):
        # time.sleep(2)
        # 启动另一个线程
        self.work3.start()

    def display3(self, int3, str3_1, str3_2):
        if int3 == 2:
            self.statusBar().showMessage('网点账单导出完成！', 10000)

    def about(self):
        QMessageBox.about(self, "关于", "傲寒弘毅信息科技有限公司开发\n版权所有，未经授权，严禁使用！\n当前版本:%s\n联系：aohanhongzhi@qq.com" % (version))

    def initUI(self):
        # 禁止最大化按钮
        # self.setWindowFlags(QtCore.Qt.WindowMinimizeButtonHint)
        # 禁止拉伸窗口大小
        # self.setFixedSize(self.width(), self.height());

        self.setWindowIcon(QIcon("iconCal.png"))

        # exitAction = QAction(QIcon('exit.png'), '退出吧宝宝', self)
        billAction = QAction("导入货单", self)
        billAction.setShortcut('Ctrl+Q')
        # billAction.triggered.connect(qApp.quit)

        billAction.triggered.connect(self.bill)
        priceAction = QAction("导入价格", self)
        priceAction.triggered.connect(self.price)

        calculatorAction = QAction("原值导出", self)
        calculatorAction.triggered.connect(self.custSave)

        saveAction = QAction("网点导出", self)
        saveAction.triggered.connect(self.save)

        aboutAction = QAction('关于', self)
        aboutAction.triggered.connect(self.about)

        self.toolbar = self.addToolBar('Exit123')
        self.toolbar.addAction(priceAction)
        self.toolbar.addAction(billAction)
        self.toolbar.addAction(calculatorAction)
        self.toolbar.addAction(saveAction)
        self.toolbar.addAction(aboutAction)

        # 界面上创建一个表格
        # model=DataFrameModel()

        # 文件路径显示
        self.pricePath = QTextEdit(self)
        self.pricePath.setText("您还没有导入价格表")
        self.pricePath.setGeometry(50, 50, 350, 50)
        self.pricePath.setReadOnly(True)

        self.billPath = QTextEdit(self)
        self.billPath.setText("您还没有导入账单")
        self.billPath.setGeometry(50, 110, 350, 50)

        # 创建一个百分比显示
        self.pbar = QProgressBar(self)
        self.pbar.setGeometry(50, 170, 350, 50)
        s = QWidget(self.pbar)
        self.setCentralWidget(s)
        self.pbar.setValue(0)

        # LCD

        self.lcd = QLCDNumber(self)
        self.lcd.setGeometry(50, 230, 350, 50)
        # 计时器
        self.lcdTime = QLCDNumber(self)
        self.lcdTime.setGeometry(50, 290, 350, 50)
        # 在类中定义一个定时器,并在构造函数中设置启动及其信号和槽
        self.lcdtimer = QTimer(self)
        # 设置计时间隔并启动(1000ms == 1s)

        # 计时结束调用timeout_slot()方法,注意不要加（）
        self.lcdtimer.timeout.connect(self.timeout_slot)

        global TimeConsuming
        TimeConsuming = 0
        text = "%d:%02d" % (TimeConsuming, TimeConsuming)
        self.lcdTime.display(text)

        self.setGeometry(300, 300, 450, 450)
        self.setWindowTitle('苏宁对账系统')
        self.show()
        self.move_center()

    def timeout_slot(self):
        global TimeConsuming
        TimeConsuming += 1
        text = "%d:%02d" % (TimeConsuming / 60, TimeConsuming % 60)
        text = self.lcdTime.display(text)
        pass

    def move_center(self):
        screen = QDesktopWidget().screenGeometry()
        form = self.geometry()
        x_move_step = (screen.width() - form.width()) / 2
        y_move_step = (screen.height() - form.height()) / 2
        self.move(x_move_step, y_move_step)

    def closeEvent(self, event):

        reply = QMessageBox.question(self, '退出程序',
                                     "确定退出程序?\n版本：%s" % (version), QMessageBox.No |
                                     QMessageBox.Yes, QMessageBox.No)

        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()


def main():
    app = QApplication(sys.argv)
    example = Example3()
    sys.exit(app.exec_())
    pass


if __name__ == '__main__':
    main()

'''

5:添加0-5单位为0.5,5+单位为1
6：修复进度条不到百分之百问题
7：修改价格新方案
8：添加友好提示，异常捕捉。
9：添加客户版本与网点版本
10：添加退出调试
11:修正，客户超过5千克以上取整
12:加入多线程，防止程序卡死，UI线程与工作线程分开
14:加入路径显示，但是还是会卡死（不是必然事件）
15-1:检测表的否为数字
15-2：加入特定区域的标记处理，颜色标记
16：升级多线程，可以处理超大量的数据，不会卡死界面
17:网点与客户保存都加入了多线程计算，此外加入计时器
17.1:修复客户导出崩溃
17.2：添加软件图标
17.3:客户版的由原来的0-3，改成0-1和1-3就可以了
17.4:客户版的由原来的，5以上修改成，5-10，10以上，都是有附加值。
17.5:网点版全部加上附加值，首重，续重不减。
17.6：修改导入价格的数组的写法，方便后续修改不用再次增删数组到字典了。
17.7:维护客户版价格表，因为上次重新修改价格字典的写法，因此这次不需要再考虑价格字典的写法了
18.0:重新设计客户的价格表，算法重新修改，客户可以自主随意设计价格表
18.1:加入原值计算,0.5以内取0.5.以上取原值
18.2:导入一次价格表
18.3:原值计算保留三位小数
'''
