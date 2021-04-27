import os
import time
import random



from PySide2.QtCore import Slot
from PySide2.QtGui import QPixmap
from PySide2.QtWidgets import QApplication, QMessageBox, QProgressBar, QPushButton, QMainWindow, QGraphicsScene, \
    QGraphicsPixmapItem
from PySide2.QtUiTools import QUiLoader
import bs4
from bs4 import BeautifulSoup
import re
import urllib.request
import sys
import xlwt
import matplotlib as mpl
import matplotlib.pyplot as plt
import matplotlib.pyplot as plt
import numpy as np


from threading import Thread


dtlist=[]
yearcloud=[]
nationcloud=[]
authorcloud=[]
pubcloud=[]

nyear2005=0
nyear2010=0
nyear2015=0
nyear2020=0

spide=0

nrate60=0
nrate70=0
nrate80=0
nrate90=0


nprice50=0
nprice100=0
nprice150=0


nnum1k=0
nnum10k=0
nnum20k=0


def savedata(datalist, savepath, savenum):
    print('saving...')
    book = xlwt.Workbook(encoding="UTF-8", style_compression=0)  # create work book

    if spide==0:
        sheet = book.add_sheet('Douban Book reading,Novel,TOP', cell_overwrite_ok=True)
    else:
        sheet = book.add_sheet('Douban Book reading,History,TOP', cell_overwrite_ok=True)
    col = (
    "Name", "author", "Nationality", "Year of Publishing", "Publisher", "Price", "Rating", "Raters number", "Image",
    "Link","note")

    for i in range(0, 11):  # exhausitive traversal for all cols
        sheet.write(0, i, col[i])
    for i in range(0, savenum):
        print("writing in the %dth data" % i)
        data = datalist[i]
        for j in range(0, 11):
            sheet.write(i + 1, j, data[j])  # 第一行是各类标题，所以要加个1

    book.save(savepath)

def askURL(url):
    head = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/42.0.2311.135 Safari/537.36 Edge/12.10240"}
    # 用户代理user agent告诉访问的服务器我们的机器，浏览器类型，知道可以获得什么文件类型
    request = urllib.request.Request(url, headers=head)
    html = ""
    try:
        response = urllib.request.urlopen(request)
        html = response.read().decode("utf-8")
        # print(html)
    except urllib.error.URLError as e:
        if hasattr(e, "code"):
            print(e.code)
        if hasattr(e, "reason"):
            print(e.reason)
    return html


def getdata(baseurl):
    percent = 0
    savenum = 0
    datalist = []
    findlink = re.compile(r'<a href="(.*?)"')  # 全局变量，创建正则表达式对象，r的意思是防止链接中的/符号起转义作用
    for i in range(0, 3):  # 181):
        url = baseurl + str(i * 20)
        html = askURL(url)
        soup = BeautifulSoup(html, "html.parser")
        for item in soup.find_all('li', class_="subject-item"):  # class是该函数参数，要加个下划线,row是APS每条结果的类名
            item = str(item)
            # print(item)
            # link=re.findall(findlink, item)#re正则表达式规则设定,link是获取论文链接
            # print(link)
            book_name = re.findall(r'}\)" title="(.*?)">', item, re.DOTALL)[0]
            book_author = re.findall(r'<div class="pub">.*? \n  \n  (.*?)/', item, re.DOTALL)[0]
            book_year = re.findall(r'<div class="pub">.*? \n  \n  .*?/ (\d{4})', item, re.DOTALL)[0]
            book_price = re.findall(r'<div class="pub">.*? \n  \n  .*?/ (\d{2,3}\.\d{2})', item, re.DOTALL)
            book_publisher = re.findall(r'<div class="pub">.*? \n  \n  .*?/ (.*?[出|书][版|店])', item, re.DOTALL)

            book_rate = re.findall(r'<span class="rating_nums">(.*?)</span>', item, re.DOTALL)
            if len(book_rate)==0:
                book_rate=7.5
            else:
                book_rate=book_rate[0]

            book_numrater = re.findall(r'\((\d{1,10})人评价\)', item, re.DOTALL)
            if len(book_numrater)==0:
                book_numrater=2000
            else:
                book_numrater=book_numrater[0]

            book_img = re.findall(r'<img class="" src="(.*?)"', item, re.DOTALL)[0]
            book_link = re.findall(r'<a href="(.*?)"', item, re.DOTALL)[0]
            book_note= re.findall(r'<p>(.*?)</p>',item,re.DOTALL)[0]

            # print(book_numrater)
            book_nation = re.findall(r'\[(.*?)\]', book_author, re.DOTALL)

            global yearcloud
            book_year_t=book_year+"年"
            yearcloud.append(book_year_t)



            if (int(book_numrater) < 1000):
                continue

            if (len(book_nation) == 0):
                book_nation = '中'
            else:
                book_nation = book_nation[0]
            # print(book_nation)

            if (len(book_price) == 0):
                book_price = '35.00'
            else:
                book_price = book_price[0]

            if (len(book_publisher) == 0):
                book_publisher = ' '
            else:
                book_publisher = book_publisher[0]

            float(book_rate)  # 转换浮点数

            global nationcloud
            nationcloud.append(book_nation)


            string_book_pub = str(book_publisher)
            string_book_author = str(book_author)

            pub = re.sub(r'([\u4e00-\u9fa5]+ / )', "", string_book_pub)
            book_publisher = pub + "社"

            global pubcloud
            pubcloud.append(book_publisher)

            aut = re.sub(r'(\[[\u4e00-\u9fa5]+\] )', "", string_book_author)
            book_author = aut
            # print(book_publisher)
            global authorcloud
            authorcloud.append(book_author)

            global nyear2005,nyear2010,nyear2015,nyear2020,nrate60,nrate70,nrate80,nrate90,nprice50,nprice100,nprice150,nnum1k,nnum10k ,nnum20k
            if int(book_year)<=2005:
                nyear2005=nyear2005+1
            elif int(book_year) <= 2010:
                nyear2010=nyear2010+1
            elif int(book_year) <= 2015:
                nyear2015=nyear2015+1
            else:
                nyear2020=nyear2020+1



            if float(book_rate)<=7:
                nrate60=nrate60+1
            elif float(book_rate) <= 8:
                nrate70=nrate70+1
            elif float(book_rate) <= 9:
                nrate80=nrate80+1
            else:
                nrate90=nrate90+1


            if float(book_price)<=50:
                nprice50=nprice50+1
            elif float(book_price)<=100:
                nprice100=nprice100+1
            else:
                nprice150=nprice150+1


            if int(book_numrater)<=100000:
                nnum1k=nnum1k+1
            elif int(book_numrater)<=200000:
                nnum10k=nnum10k+1
            else:
                nnum20k=nnum20k+1






            datalist.append(
                [book_name, book_author, book_nation, book_year, book_publisher, book_price, float(book_rate),
                 int(book_numrater), book_img, book_link,book_note])
            print(datalist)
            # print(datalist)
            savenum = savenum + 1
            t=random.uniform(0,0.1)
            time.sleep(t)
    return datalist, savenum

class Window2:
    def __init__(self):
        self.ui = QUiLoader().load('window2_.ui')
        self.ui.pushButton_5.clicked.connect(self.handleBack)

        self.ui.pushButton_4.clicked.connect(self.handlematpl)


        self.graphic_scene = QGraphicsScene()
        self.pic = QGraphicsPixmapItem()
        self.pic.setPixmap(QPixmap('background.png').scaled(390, 790))

        # self.pic.setFlags(QGraphicsItem.ItemIsSelectable | QGraphicsItem.ItemIsMovable) #可选择，可移动、
        # self.pic.setOffset(100, 120)
        self.graphic_scene.addItem(self.pic)

        self.ui.graphicsView.setScene(self.graphic_scene)
        self.ui.graphicsView.show()
        if spide==0:
            self.ui.textBrowser.setText("本次爬取的是：小说")
        else:
            self.ui.textBrowser.setText("本次爬取的是：历史")

    def handleBack(self):
        # 实例化另外一个窗口
        global yearcloud,nationcloud,authorcloud,pubcloud
        yearcloud = []
        nationcloud = []
        authorcloud = []
        pubcloud = []
        self.window2 = Stats()
        # 显示新窗口
        self.window2.ui.show()
        # 关闭自己
        self.ui.close()

    def handlematpl(self):
        self.window5 = matpl()
        # 显示新窗口
        self.window5.ui.show()
        # 关闭自己
        self.ui.close()

class matpl:
    def __init__(self):
        self.ui = QUiLoader().load('matplt.ui')
        global nyear2005, nyear2010, nyear2015, nyear2020, nrate60, nrate70, nrate80, nrate90, nprice50, nprice100, nprice150, nnum1k, nnum10k, nnum20k
       ##############year
        x = np.arange(4)
        y = [nyear2005, nyear2010, nyear2015, nyear2020]
        bar_width = 0.35
        tick_label = ["<2005", "2005~2010", "2010~2015","2015~2021"]
        plt.bar(x, y, bar_width, align="center", color="c", alpha=0.5)
        plt.xlabel("Year of Pub")
        plt.ylabel("Number")
        plt.xticks(x+bar_width/2, tick_label)
        #plt.show()
        plt.savefig('yearplt.png')
        plt.close()

        x1=np.arange(4)
        y1=[nrate60, nrate70, nrate80, nrate90]
        bar_width = 0.35
        tick_label = ["<7.0", "7.0~8.0", "8.0~9.0", "9.0~10.0"]
        plt.bar(x1, y1, bar_width, align="center", color="r", alpha=0.5)
        plt.xlabel("Ratings")
        plt.ylabel("Number")
        plt.xticks(x1+bar_width/2, tick_label)
        #plt.show()
        plt.savefig('rateplt.png')
        plt.close()

        x1=np.arange(3)
        y1=[nprice50, nprice100, nprice150]
        bar_width = 0.35
        tick_label = ["<50", "50~100", ">100"]
        plt.bar(x1, y1, bar_width, align="center", color="g", alpha=0.5)
        plt.xlabel("Price")
        plt.ylabel("Number")
        plt.xticks(x1+bar_width/2, tick_label)
        #plt.show()
        plt.savefig('priceplt.png')
        plt.close()

        self.graphic_scene = QGraphicsScene()
        self.graphic_scene2 = QGraphicsScene()
        self.graphic_scene3 = QGraphicsScene()
        self.pic = QGraphicsPixmapItem()
        self.pic.setPixmap(QPixmap('yearplt.png').scaled(555, 285))

        self.pic2 = QGraphicsPixmapItem()
        self.pic2.setPixmap(QPixmap('rateplt.png').scaled(555, 285))

        self.pic3 = QGraphicsPixmapItem()
        self.pic3.setPixmap(QPixmap('priceplt.png').scaled(555, 285))
        # self.pic.setFlags(QGraphicsItem.ItemIsSelectable | QGraphicsItem.ItemIsMovable) #可选择，可移动、
        # self.pic.setOffset(100, 120)
        self.graphic_scene.addItem(self.pic)
        self.graphic_scene2.addItem(self.pic2)
        self.graphic_scene3.addItem(self.pic3)

        self.ui.graphicsView.setScene(self.graphic_scene)  # 把QGraphicsScene放入QGraphicsView
        self.ui.graphicsView.show()  # 调用show方法呈现图形

        self.ui.graphicsView_2.setScene(self.graphic_scene2)  # 把QGraphicsScene放入QGraphicsView
        self.ui.graphicsView_2.show()  # 调用show方法呈现图形

        self.ui.graphicsView_3.setScene(self.graphic_scene3)  # 把QGraphicsScene放入QGraphicsView
        self.ui.graphicsView_3.show()  # 调用show方法呈现图形
        nyear2005=0
        nyear2010=0
        nyear2015=0
        nyear2020=0
        nrate60=0
        nrate70=0
        nrate80=0
        nrate90=0
        nprice50=0
        nprice100=0
        nprice150=0
        nnum1k=0
        nnum10k=0
        nnum20k=0
        self.ui.pushButton.clicked.connect(self.handleBack)

    def handleBack(self):
        # 实例化另外一个窗口
        self.window2 = Window2()
        # 显示新窗口
        self.window2.ui.show()
        # 关闭自己
        self.ui.close()


class Stats:  # 定义窗口类

    def __init__(self):

        self.ui = QUiLoader().load('main.ui')

        self.ui.pushButton.clicked.connect(self.handleCalc)

        self.graphic_scene = QGraphicsScene()
        self.pic = QGraphicsPixmapItem()
        self.pic.setPixmap(QPixmap('background.png').scaled(390, 790))

        # self.pic.setFlags(QGraphicsItem.ItemIsSelectable | QGraphicsItem.ItemIsMovable) #可选择，可移动、
        # self.pic.setOffset(100, 120)
        self.graphic_scene.addItem(self.pic)

        self.ui.graphicsView.setScene(self.graphic_scene)  # 把QGraphicsScene放入QGraphicsView
        self.ui.graphicsView.show()  # 调用show方法呈现图形

    def open_new_window(self):
        # 实例化另外一个窗口
        self.window2 = Window2()
        # 显示新窗口
        self.window2.ui.show()
        # 关闭自己
        self.ui.close()

    def handleCalc(self):
        gettext = self.ui.comboBox.currentText()
        global spide
        if gettext=='历史':
            spide=1
        else:
            spide=0
            print(spide)
        Stats.open_new_dialog(self)
        # baseurl = "https://book.douban.com/tag/%E5%B0%8F%E8%AF%B4?type=T&start="
        # datalist, savenumm = getdata(baseurl)
        # datalist = sorted(datalist, key=(lambda x: x[6]), reverse=True)
        # savedata(datalist, 'E:\PyProgr\Spider\豆瓣读书.xls', savenumm)
        Stats.open_new_window(self)

    def open_new_dialog(self):
        # 实例化一个对话框类
        self.dlg = MyDialog()
        # self.dlg.ui.show()
        self.dlg.ui.exec_()

class MyDialog:
    def __init__(self):

        self.ui = QUiLoader().load('dialog.ui')
        self.ui.progressBar.setValue(0)
        self.testbar()

    def testbar(self):
        def run():
            for Percent in range(100 + 1):

                self.ui.progressBar.setValue(Percent)

                if Percent == 30:
                    global spide
                    if spide==0:
                        baseurl = "https://book.douban.com/tag/%E5%B0%8F%E8%AF%B4?type=T&start="
                    else:
                        baseurl = "https://book.douban.com/tag/%E5%8E%86%E5%8F%B2?type=T&start="
                    datalist, savenumm = getdata(baseurl)
                    datalist = sorted(datalist, key=(lambda x: x[6]), reverse=True)
                    global dtlist#全局变量
                    dtlist=datalist

                    savedata(datalist, 'E:\PyProgr\Spider\豆瓣读书.xls', savenumm)
                    global yearcloud
                    print(yearcloud)
                if Percent == 100:
                    self.ui.label.setText("Mission accomplished!")

                print(Percent)
                time.sleep(0.05)  # 延迟50ms

        t = Thread(target=run)
        t.start()


app = QApplication([])
stats = Stats()
stats.ui.show()
app.exec_()