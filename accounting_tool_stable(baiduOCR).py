'''
    Pyside2文档：https://doc.qt.io/qtforpython-5/
    Pyside2简易教程：https://www.byhy.net/tut/py/gui/qt_01/
    百度智能云获取项目ak与sk
    参考：https://blog.csdn.net/qq_34673086/article/details/107845615
'''

import os, openpyxl
import requests,base64
from functools import partial
from PySide2.QtCore import QDate, Qt
from PySide2.QtUiTools import QUiLoader
from PySide2.QtWidgets import QApplication, QMessageBox, QTableWidgetItem, QFileDialog

def get_token(ak,sk):
    host = 'https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id={}&client_secret={}'.format(ak,sk)
    response = requests.get(host)
    if response:
        return response.json()['access_token']
    else:
        return None

class books:
    def __init__(self,ak,sk):
        # 加载界面
        self.ui = QUiLoader().load('./ui/accounting_tool_stable.ui')
        self.ui.resize(1200,700)
        # 初始化百度OCR设置：通用文字识别（高精度版）
        self.request_url = "https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic"
        self.access_token = get_token(ak,sk)
        if self.access_token is None:
            raise RuntimeError('获取access_token失败')
        self.request_url = self.request_url + "?access_token=" + self.access_token
        self.headers = {'content-type': 'application/x-www-form-urlencoded'}
        # 获取系数
        self.stores = []
        self.seaFoods = []
        self.others = []
        self.pricesOut1 = {}  # 1-15日售价
        self.pricesOut2 = {}  # 16-31日售价
        self.pricesIn = {}  # 进价
        self.getParam()
        # 词汇处理
        # 1同义词纠错 出货单词汇：系统词汇
        self.synonym = {"商品100": "商品1"}
        # 2停用词
        self.stop_word = ['品名', '品种', '申购', '备注']
        # 税率
        self.rate = 1.13
        # 设置变量与控件
        # tab1
        self.ui.date.setCalendarPopup(True)
        self.ui.date.setDate(QDate.currentDate())
        self.Date = QDate.currentDate().toString('yyyy-MM-dd')
        self.dataInPath = os.path.join('data', 'in')
        self.searchAndShow(0, None)
        self.ui.classes.addItems(self.seaFoods + self.others)
        self.ui.classSearch.addItems(self.seaFoods)
        self.ui.price.setValue(self.pricesIn[self.seaFoods[0]])
        # tab2
        self.ui.date2.setCalendarPopup(True)
        self.ui.date2.setDate(QDate.currentDate())
        self.Date2 = QDate.currentDate().toString('yyyy-MM-dd')
        self.dataOutPath = os.path.join('data', 'out')
        self.searchAndShow(1, None)
        self.ui.store.addItems(self.stores)
        self.ui.classes2.addItems(self.seaFoods)
        # tab3
        self.ui.date3.setCalendarPopup(True)
        self.ui.date3.setDate(QDate.currentDate())
        self.Date3 = QDate.currentDate().toString('yyyy-MM-dd')
        self.checkShow()
        # tab4
        self.ui.storeSearch.addItems(self.stores)
        self.ui.year.setValue(QDate.currentDate().year())
        self.ui.month.setValue(QDate.currentDate().month())
        # 关联事件
        self.ui.insert.clicked.connect(partial(self.insert, 0))
        self.ui.sub.clicked.connect(partial(self.delete, 0))
        self.ui.date.dateChanged.connect(partial(self.searchAndShow, 0))  # 信号自动传日期(位于最后一个参数)
        self.ui.save.clicked.connect(partial(self.save, 0))
        self.ui.insert2.clicked.connect(partial(self.insert, 1))
        self.ui.sub2.clicked.connect(partial(self.delete, 1))
        self.ui.date2.dateChanged.connect(partial(self.searchAndShow, 1))
        self.ui.save2.clicked.connect(partial(self.save, 1))
        self.ui.date3.dateChanged.connect(self.changeDate)
        self.ui.search.clicked.connect(self.checkShow)
        self.ui.search2.clicked.connect(self.statistics)
        self.ui.selectClass.clicked.connect(self.selectClass)
        self.ui.classes.currentIndexChanged.connect(self.changePriceInOrNum)  # 进价显示基准
        self.ui.selectStore.clicked.connect(self.statistics2)
        self.ui.chooseFiles.clicked.connect(self.chooseFiles)
        self.ui.batchImport.clicked.connect(self.batchImport)
        self.ui.tableWeight2.cellChanged.connect(partial(self.calculateTotal, 1))
        self.ui.tableWeight.cellChanged.connect(partial(self.calculateTotal, 0))

    def chooseFiles(self):
        file_name = QFileDialog.getExistingDirectory(self.ui, "选择当日进货单文件夹")
        self.ui.filesName.setText(file_name)

    def batchImport(self):
        image_path = self.ui.filesName.text()
        if image_path and os.path.exists(image_path):
            for file in os.listdir(image_path):
                if not file.endswith(".jpg") and not file.endswith(".png") and not file.endswith(".jpeg"):
                    continue
                try:
                    # 获取识别结果:识别效果与图片倒正有关
                    img_path = os.path.join(image_path, file)
                    f = open(img_path, 'rb')
                    img = base64.b64encode(f.read())
                    params = {"image": img}
                    response = requests.post(self.request_url, data=params, headers=self.headers)
                    results = response.json()
                    if results["words_result_num"] == 0:
                        print(file + "未读取到内容")
                        continue
                    # 取出所有识别到的词
                    txts = [result["words"] for result in results["words_result"]]
                    # print(txts)
                    # 获取词汇处理后有用数据:store_name sea_Foods
                    store_name = ""
                    got_store = False
                    sea_Foods = []
                    for txt in txts:
                        if txt in self.stop_word or self.is_number(txt): # 停用词及纯数字停用
                            continue
                        if txt in self.synonym: # 近义词转换
                            txt = self.synonym[txt]
                        if txt in self.seaFoods: # 品名查询
                            sea_Foods.append(txt)
                            continue
                        if not got_store: # 未匹配到店名时，匹配店名
                            for store in self.stores:
                                if store in txt:
                                    store_name = store
                                    got_store = True
                    # print(store_name,sea_Foods)
                    if store_name != "":
                        print(file + "图片读取成功")
                        # 插入表格
                        for sea_Food in sea_Foods:
                            rowcount = self.ui.tableWeight2.rowCount()
                            self.ui.tableWeight2.insertRow(rowcount)
                            item = QTableWidgetItem(store_name)
                            item.setFlags(Qt.ItemIsEnabled)
                            self.ui.tableWeight2.setItem(rowcount, 0, item)
                            item = QTableWidgetItem(sea_Food)
                            item.setFlags(Qt.ItemIsEnabled)
                            self.ui.tableWeight2.setItem(rowcount, 1, item)
                            day = self.ui.date2.date().day()
                            p = self.pricesOut1[sea_Food] if day >= 1 and day <= 15 else self.pricesOut2[sea_Food]
                            item = QTableWidgetItem(str(p))
                            item.setFlags(Qt.ItemIsEnabled)
                            self.ui.tableWeight2.setItem(rowcount, 3, item)
                    else:
                        print(file + "店铺信息未识别")
                except Exception as e:
                    print(e)
                    print(file + "图片读取失败")
                    continue
        else:
            QMessageBox.warning(
                self.ui,
                '警告',
                '文件夹路径有误！')

    # 重量改变，总价自动改变
    def is_number(self, str):
        try:
            float(str)
        except:
            return False
        else:
            return True

    def calculateTotal(self, flag, row, column):
        if flag == 0:
            w_column, p_column, t_column, taxes = 1, 2, 3, 1
            tableWeight = self.ui.tableWeight
        else:
            w_column, p_column, t_column, taxes = 2, 3, 4, self.rate
            tableWeight = self.ui.tableWeight2
        if column == w_column:
            change_weight = tableWeight.item(row, column).text()
            if change_weight is not None:
                if self.is_number(change_weight):
                    item_price = tableWeight.item(row, p_column)
                    if item_price is not None:
                        item = QTableWidgetItem(str(round(float(change_weight) * float(item_price.text()) * taxes, 2)))
                        item.setFlags(Qt.ItemIsEnabled)
                        tableWeight.setItem(row, t_column, item)
                else:
                    QMessageBox.warning(
                        self.ui,
                        '警告',
                        '输入数据有误，请检查并重新输入！')

    def changePriceInOrNum(self, num):
        if num >= len(self.seaFoods):
            self.ui.weight.setValue(1)
            self.ui.price.setValue(0)
        else:
            self.ui.price.setValue(self.pricesIn[self.seaFoods[num]])

    def getParam(self):
        storesPath = os.path.join('param', 'stores.xlsx')
        wb = openpyxl.load_workbook(storesPath)
        ws = wb[wb.sheetnames[0]]
        for row in ws.rows:
            self.stores.append(row[0].value)
        foodsPath = os.path.join('param', 'foods.xlsx')
        wb = openpyxl.load_workbook(foodsPath)
        ws = wb[wb.sheetnames[0]]
        rowNum = 0
        for row in ws.rows:
            if rowNum == 0:
                rowNum += 1
                continue
            self.seaFoods.append(row[0].value)
            self.pricesOut1[row[0].value] = row[1].value  # 获取1-15售价
            self.pricesOut2[row[0].value] = row[2].value  # 获取16-31售价
            self.pricesIn[row[0].value] = row[3].value  # 获取进价基准
        othersPath = os.path.join('param', 'others.xlsx')
        wb = openpyxl.load_workbook(othersPath)
        ws = wb[wb.sheetnames[0]]
        for row in ws.rows:
            self.others.append(row[0].value)

    def insert(self, flag):
        if flag == 0:
            rowcount = self.ui.tableWeight.rowCount()
            self.ui.tableWeight.insertRow(rowcount)
            item = QTableWidgetItem(self.ui.classes.currentText())
            item.setFlags(Qt.ItemIsEnabled)
            self.ui.tableWeight.setItem(rowcount, 0, item)
            self.ui.tableWeight.setItem(rowcount, 1, QTableWidgetItem(str(self.ui.weight.value())))
            item = QTableWidgetItem(str(self.ui.price.value()))
            item.setFlags(Qt.ItemIsEnabled)
            self.ui.tableWeight.setItem(rowcount, 2, item)
            item = QTableWidgetItem(str(round(self.ui.weight.value() * self.ui.price.value(), 2)))
            item.setFlags(Qt.ItemIsEnabled)
            self.ui.tableWeight.setItem(rowcount, 3, item)
            self.ui.tableWeight.sortItems(0)  # 插入后，根据品名排序，方便发现错误
            self.ui.tableWeight.scrollToItem(item)
        else:
            rowcount = self.ui.tableWeight2.rowCount()
            self.ui.tableWeight2.insertRow(rowcount)
            item = QTableWidgetItem(self.ui.store.currentText())
            item.setFlags(Qt.ItemIsEnabled)
            self.ui.tableWeight2.setItem(rowcount, 0, item)
            item = QTableWidgetItem(self.ui.classes2.currentText())
            item.setFlags(Qt.ItemIsEnabled)
            self.ui.tableWeight2.setItem(rowcount, 1, item)
            self.ui.tableWeight2.setItem(rowcount, 2, QTableWidgetItem(str(self.ui.weight2.value())))
            day = self.ui.date2.date().day()
            if day >= 1 and day <= 15:
                p = self.pricesOut1[self.ui.classes2.currentText()]
            else:
                p = self.pricesOut2[self.ui.classes2.currentText()]
            item = QTableWidgetItem(str(p))
            item.setFlags(Qt.ItemIsEnabled)
            self.ui.tableWeight2.setItem(rowcount, 3, item)
            item = QTableWidgetItem(str(round(self.ui.weight2.value() * p * self.rate, 2)))
            item.setFlags(Qt.ItemIsEnabled)
            self.ui.tableWeight2.setItem(rowcount, 4, item)
            self.ui.tableWeight2.sortItems(0)  # 插入后，根据店名排序，方便查看
            self.ui.tableWeight2.scrollToItem(item)

    def delete(self, flag):
        if flag == 0:
            tableWeight = self.ui.tableWeight
        else:
            tableWeight = self.ui.tableWeight2
        currentrow = tableWeight.currentRow()
        if currentrow == -1:
            QMessageBox.warning(
                self.ui,
                '警告',
                '未选中数据，不可删除！')
        tableWeight.removeRow(currentrow)

    def searchAndShow(self, flag, date):
        if flag == 0:
            if date is not None:
                self.Date = date.toString('yyyy-MM-dd')
            tableWeight = self.ui.tableWeight
            dataPath = self.dataInPath
            Date = self.Date
            totalPrice = self.ui.totalPrice
        else:
            if date is not None:
                self.Date2 = date.toString('yyyy-MM-dd')
            tableWeight = self.ui.tableWeight2
            dataPath = self.dataOutPath
            Date = self.Date2
            totalPrice = self.ui.totalPrice2
        tableWeight.setRowCount(0)
        wbName = os.path.join(dataPath, Date + '.xlsx')
        total = 0.0
        if os.path.exists(wbName):
            wb = openpyxl.load_workbook(wbName)
            ws = wb[wb.sheetnames[0]]
            rowNum = 0
            for row in ws.rows:
                if rowNum == 0:
                    rowNum += 1
                    continue
                insertPosition = rowNum - 1
                line = [cell.value for cell in row]
                tableWeight.insertRow(insertPosition)
                for i in range(len(line) - 2):
                    tableWeight.setItem(insertPosition, i, QTableWidgetItem(line[i] if isinstance(line[i], str) else str(line[i])))
                for i in [1, 2]:
                    item = QTableWidgetItem(str(line[len(line) - i]))
                    item.setFlags(Qt.ItemIsEnabled)
                    tableWeight.setItem(insertPosition, len(line) - i, item)
                total += float(line[len(line) - 1])  # 统计每日总金额
                rowNum += 1
        # 输出总金额
        totalPrice.clear()
        totalPrice.setText('总金额：' + str(total))

    def save(self, flag):
        if flag == 0:
            tableWeight = self.ui.tableWeight
            dataPath = self.dataInPath
            Date = self.Date
            title = ['品种', '重量(斤)', '进价', '总价']
        else:
            tableWeight = self.ui.tableWeight2
            dataPath = self.dataOutPath
            Date = self.Date2
            title = ['店名', '品种', '重量(公斤)', '售价', '总价']
        choice = QMessageBox.question(self.ui, '确认', '确定要保存数据吗？\n注:这将修改原有数据')
        if choice == QMessageBox.Yes:
            rowcount = tableWeight.rowCount()
            columncount = tableWeight.columnCount()
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(title)
            for row in range(rowcount):
                rowData = []
                for column in range(columncount):
                    item = tableWeight.item(row, column)
                    if item is not None and item.text() != "":
                        rowData.append(item.text())
                    else:
                        rowData.append(0.0)
                ws.append(rowData)
            wb.save(os.path.join(dataPath, Date + '.xlsx'))
            QMessageBox.warning(self.ui, '提示', '保存成功')
        self.searchAndShow(flag, None)

    def changeDate(self, date):
        self.Date3 = date.toString('yyyy-MM-dd')

    def checkShow(self):
        self.ui.tableWeight3.setRowCount(0)
        inPath = os.path.join(self.dataInPath, self.Date3 + '.xlsx')
        outPath = os.path.join(self.dataOutPath, self.Date3 + '.xlsx')
        if os.path.exists(inPath) and os.path.exists(outPath):
            # 获取处理in数据
            wb1 = openpyxl.load_workbook(inPath)
            ws1 = wb1[wb1.sheetnames[0]]
            inData = {}
            rowNum = 0
            for row in ws1.rows:
                if rowNum == 0:
                    rowNum += 1
                    continue
                inData[row[0].value] = float(row[1].value)
            # 获取处理out数据
            wb2 = openpyxl.load_workbook(outPath)
            ws2 = wb2[wb2.sheetnames[0]]
            outData = {}
            rowNum = 0
            for row in ws2.rows:
                if rowNum == 0:
                    rowNum += 1
                    continue
                if row[1].value in outData:
                    outData[row[1].value] = round(float(row[2].value) * 2 + outData[row[1].value], 2)
                else:
                    outData[row[1].value] = float(row[2].value) * 2
            # 显示
            for i in range(len(self.seaFoods)):
                foodName = self.seaFoods[i]
                self.ui.tableWeight3.insertRow(i)
                self.ui.tableWeight3.setItem(i, 0, QTableWidgetItem(foodName))
                if foodName in inData:
                    self.ui.tableWeight3.setItem(i, 1, QTableWidgetItem(str(inData[foodName])))
                if foodName in outData:
                    self.ui.tableWeight3.setItem(i, 2, QTableWidgetItem(str(outData[foodName])))
                if foodName in inData and foodName in outData:
                    self.ui.tableWeight3.setItem(i, 3, QTableWidgetItem(
                        str(round(outData[foodName] - inData[foodName], 2)) + ' 正确' if inData[foodName] <= outData[
                            foodName] else str(round(outData[foodName] - inData[foodName], 2)) + ' 错误'))

    def statistics(self):
        year = int(self.ui.year.value())
        month = int(self.ui.month.value())
        excels = os.listdir(self.dataInPath)
        fileNames = []
        for excel in excels:
            if excel[0] == '.':
                continue
            s = excel.split('-')
            y = int(s[0])
            m = int(s[1])
            if y == year and m == month:
                fileNames.append(excel)
        dayPrice = {}
        total = 0.0
        for file in fileNames:
            dayTotal = 0.0
            filePath = os.path.join(self.dataInPath, file)
            wb = openpyxl.load_workbook(filePath)
            ws = wb[wb.sheetnames[0]]
            rowNum = 0
            for row in ws.rows:
                if rowNum == 0:
                    rowNum += 1
                    continue
                line = [cell.value for cell in row]
                dayTotal += float(line[3])
                total += float(line[3])
            dayPrice[file.split('.')[0]] = dayTotal
        self.ui.text.clear()
        self.ui.text.append(str(year) + "年" + " " + str(month) + "月 进货统计")
        for day in dayPrice:
            self.ui.text.append(day + ':' + str(dayPrice[day]) + '元')
        self.ui.text.append('总支出金额：' + str(total) + '元')
        self.ui.text.ensureCursorVisible()

    def statistics2(self):
        store = self.ui.storeSearch.currentText()
        year = int(self.ui.year.value())
        month = int(self.ui.month.value())
        excels = os.listdir(self.dataOutPath)
        fileNames = []
        for excel in excels:
            if excel[0] == '.':
                continue
            s = excel.split('-')
            y = int(s[0])
            m = int(s[1])
            if y == year and m == month:
                fileNames.append(excel)
        storeData = {}
        Total = 0.0
        total = 0.0
        for file in fileNames:
            dayTotal = 0.0
            filePath = os.path.join(self.dataOutPath, file)
            wb = openpyxl.load_workbook(filePath)
            ws = wb[wb.sheetnames[0]]
            rowNum = 0
            for row in ws.rows:
                if rowNum == 0:
                    rowNum += 1
                    continue
                line = [cell.value for cell in row]
                Total += float(line[4])
                if line[0] == store:
                    dayTotal += float(line[4])
                    total += float(line[4])
            storeData[file.split('.')[0]] = dayTotal
        self.ui.text2.clear()
        self.ui.text2.append(str(year) + "年" + " " + str(month) + "月 出货统计")
        self.ui.text2.append("全店铺总金额：" + str(Total) + "\n")
        self.ui.text2.append(store + "店 详细信息")
        for data in storeData:
            self.ui.text2.append(data + ':' + str(storeData[data]) + '元')
        self.ui.text2.append(store + '总金额：' + str(total) + '元')
        self.ui.text2.ensureCursorVisible()

    def selectClass(self):
        self.ui.classText.clear()
        filePath = os.path.join(self.dataOutPath, self.Date3 + '.xlsx')
        if os.path.exists(filePath):
            wb = openpyxl.load_workbook(filePath)
            ws = wb[wb.sheetnames[0]]
            rowNum = 0
            for row in ws.rows:
                if rowNum == 0:
                    rowNum += 1
                    continue
                line = [cell.value for cell in row]
                if line[1] == self.ui.classSearch.currentText():
                    self.ui.classText.append(line[0] + ':' + str(float(line[2]) * 2) + '斤')

if __name__ == "__main__":
    # 此处填入个人百度智能云项目ak与sk
    ak = '百度智能云OCR项目ak'
    sk = '百度智能云OCR项目sk'
    app = QApplication([])
    books = books()
    books.ui.show()
    app.exec_()
