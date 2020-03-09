from fbs_runtime.application_context.PyQt5 import ApplicationContext
from PyPDF2 import PdfFileWriter, PdfFileReader
from PyQt5.QtWidgets import (QWidget, QLabel, QLineEdit,
                             QFileDialog,
                             QVBoxLayout,
                             QGroupBox,
                             QGridLayout,
                             QProgressBar,
                             QCheckBox,
                             QMessageBox, QPushButton)
from PyQt5 import QtGui, QtCore
from PyQt5.QtCore import  pyqtSignal, QSize
from PyQt5.QtGui import QColor

import sys
import pathlib
from pathlib import Path
import os
from time import strftime, gmtime
from pathlib import Path
import json






class ClickableLineEdit(QLineEdit):
    clicked = pyqtSignal()  # signal when the text entry is left clicked

    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self.clicked.emit()
        else:
            super().mousePressEvent(event)




class ClickableLineEdit(QLineEdit):
    clicked = pyqtSignal()  # signal when the text entry is left clicked

    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self.clicked.emit()
        else:
            super().mousePressEvent(event)

class MainWindow(QWidget):
    def __init__(self, appctxt, config, configPath):
        super().__init__()
        self.lastOpenFoldPath = str(Path.home())
        self.outputFileNameValue = ""
        self.outputFoldPath = config['lastPath'] or str(Path.home())
        self.splitWayValue = config['byWay']
        self.config = config
        self.appctxt = appctxt
        self.configPath = configPath
        self.initUI()

    def initUI(self):
        self.setAutoFillBackground(True)
        pattern = self.palette()
        pattern.setColor(QtGui.QPalette.Background, QColor("#EFFCFC"))
        self.setPalette(pattern)

        mainLayout = QVBoxLayout()
        uploadButton = QPushButton()
        uploadButton.setObjectName("upload")
        uploadButton.setCursor(QtCore.Qt.PointingHandCursor)
        uploadButton.clicked.connect(self.handleOpenPDF)
        icon = QtGui.QIcon()
        iconPath = appctxt.get_resource("images/pdf.png")
        icon.addPixmap(QtGui.QPixmap(iconPath))
        uploadButton.setIcon(icon)
        uploadButton.setIconSize(QSize(30, 30))



        note = QLabel(
            '如何分页，默认为1，也可以自定义， Exp * [2] 表示每隔2页分隔成一个PDF *  [2，3] 表示第一个文档取前2页，第二个文档取接下去的3页，3. * [1-2，3-5] 表示 1～2页分割成一个文档， 3～5页分割成一个文档')
        note.setObjectName('note')
        note.setWordWrap(True)

        self.outputFileName = QLineEdit()
        self.outputFileName.setPlaceholderText('输出文件名称')
        self.outputFileName.setDisabled(True)



        # self.compressCheckBox = QCheckBox()
        # self.compressCheckBox.setText("输出文件是否打包压缩")
        self.categoryByTime = QCheckBox()
        self.categoryByTime.setText("输出文件夹按时间分类")

        state = 0
        if self.config['byTime'] == True:
            state = 2
        else:
            state = 0
        self.categoryByTime.setCheckState(state)
        self.categoryByTime.stateChanged.connect(self.byTimeChecked)
        self.categoryByTime.setDisabled(True)


        self.outputFold = ClickableLineEdit()
        self.outputFold.setText(self.outputFoldPath)
        self.outputFold.setPlaceholderText("输出文件夹位置")
        self.outputFold.clicked.connect(self.selectOutputFold)
        self.outputFold.textChanged.connect(self.textChangedOutputFold)
        self.outputFold.setDisabled(True)

        doButton = QPushButton("执行")
        doButton.setObjectName("do")
        doButton.setCursor(QtCore.Qt.PointingHandCursor)
        doButton.clicked.connect(self.handleSplitPDF)
        doButton.setDisabled(True)
        self.doButton = doButton
        self.splitWay = QLineEdit()
        self.splitWay.setText(self.splitWayValue)
        self.splitWay.textChanged.connect(self.splitWayChange)
        self.splitWay.setDisabled(True)

        self.fileName = QLabel('文件名称')
        self.pageNumber = QLabel("文件页数")
        self.fileSize = QLabel("文件大小")
        self.updateTime = QLabel("文件最近更新时间")

        self.setLayout(mainLayout)
        self.setGeometry(300, 200, 500, 600)

        self.setWindowTitle('PDF 分割工具')
        self.progressStatus = QProgressBar()
        self.progressStatus.setObjectName("progressBar")
        self.progressStatus.setGeometry(0,0, 500, 20)
        self.progressStatus.setMaximum(100)
        self.progressStatus.setMinimum(0)
        self.progressStatus.setTextVisible(False)
        self.progressStatus.setVisible(False)



        # self.progressStatus.setHidden(True)








        headerGridGroup = QGroupBox()
        bodyGridGroup = QGroupBox()
        gridHeaderLayout = QGridLayout()
        gridHeaderLayout.addWidget(uploadButton, 1, 1, 4, 1)  # row 1 col 2
        gridHeaderLayout.addWidget(self.fileName, 1, 2, 1, 1)
        gridHeaderLayout.addWidget(self.pageNumber, 2, 2, 1, 1)
        gridHeaderLayout.addWidget(self.fileSize, 3, 2, 1, 1)
        gridHeaderLayout.addWidget(self.updateTime, 4, 2, 1, 1)

        gridBodyLayout = QGridLayout()
        gridBodyLayout.addWidget(self.outputFileName, 2, 1, 1, 2)  # row 2 col 1
        gridBodyLayout.addWidget(self.outputFold, 3, 1, 1, 2)

        gridBodyLayout.addWidget(self.categoryByTime, 4,1,1,2)




        gridBodyLayout.addWidget(note, 7, 1, 1, 2)
        gridBodyLayout.addWidget(self.splitWay, 8, 1, 1, 2)
        gridBodyLayout.addWidget(doButton, 9, 2, 1, 1)
        gridBodyLayout.addWidget(self.progressStatus, 10, 1, 1, 2)




        headerGridGroup.setLayout(gridHeaderLayout)
        bodyGridGroup.setLayout(gridBodyLayout)

        mainLayout.addWidget(headerGridGroup)
        mainLayout.addWidget(bodyGridGroup)


    def byTimeChecked(self):
        self.config['byTime'] = self.categoryByTime.isChecked()
        self.updateConfigFile()


    def textChangedOutputFold(self):
        self.config['lastPath'] = self.outputFold.text()
        self.updateConfigFile()

    def splitWayChange(self):
        self.config['byWay'] = self.splitWay.text()
        self.updateConfigFile()


    def selectOutputFold(self):
        foldName = QFileDialog.getExistingDirectory(self, "选择存放文件夹")
        self.outputFoldPath = str(foldName)
        self.outputFold.setText(foldName)

    def handleOpenPDF(self):
        lastOpenFoldPath = self.lastOpenFoldPath
        fname = QFileDialog.getOpenFileName(self, '打开PDF文件', lastOpenFoldPath, "pdf(*.pdf)")

        if fname[0]:
            fileName = os.path.splitext(os.path.basename(fname[0]))[0]

            fileSizeByte = os.path.getsize(fname[0])
            fileSize = "{:.2f}".format(fileSizeByte / 1024 / 1024) + 'M'
            updateTime = strftime("%Y-%m-%d %H:%M:%S", gmtime(os.path.getmtime(fname[0])))
            self.inputPDF = PdfFileReader(open(fname[0], "rb"))
            pageNumber = str(self.inputPDF.numPages) + '页'

            self.lastOpenFoldPath = str(pathlib.Path(fname[0]).parent.absolute())

            self.outputFileNameValue = fileName



            self.fileName.setText('文件名：' + self.outputFileNameValue)
            self.pageNumber.setText('PDF总页数：' + pageNumber)
            self.fileSize.setText('文件大小：' + fileSize)
            self.updateTime.setText('最近更新时间：' + updateTime)

            outputName = os.path.splitext(fileName)[0]
            self.outputFileName.setText(outputName)
            self.outputFileName.setDisabled(False)
            self.outputFold.setDisabled(False)
            self.doButton.setDisabled(False)
            self.outputFileName.setDisabled(False)
            self.categoryByTime.setDisabled(False)
            self.splitWay.setDisabled(False)


    # 弹框展示错误信息
    def showDialog(self, message):
        msgBox = QMessageBox()
        msgBox.setIcon(QMessageBox.Information)
        msgBox.setText(message)
        msgBox.exec()


    # 信息验证
    def checkInfomationValidate(self):
        fileName = self.outputFileName.text()
        outputPath = self.outputFold.text()
        splitWay = self.splitWay.text()

        if not len(fileName):
            return {
                'status': False,
                "msg": "请填写需要输出的文件的名称"
            }

        if not len(outputPath):
            return {
                'status': False,
                'msg': '请选择需要输出的文件的存放地址'
            }

        if not len(splitWay):
            return {
                'status': False,
                'msg': '请指定分割方法'
            }

        return {
            'status':True,
            'msg': ''
        }



    def updateConfigFile(self):
        with open(self.configPath, 'w') as f:
            json.dump(self.config, f);
    def handleSetSplitWay(self):
        pass

    def handleSplitPDF(self):

        self.checkInfomationValidate()
        self.doButton.setVisible(False)
        self.progressStatus.setVisible(True)
        fileName = self.outputFileName.text()
        outputPath = self.outputFold.text()
        splitWay = self.splitWay.text()
        # compress = self.compressCheckBox.checkState()
        categoryByTime = self.categoryByTime.checkState()
        # print(compress, categoryByTime)
        funcWays = []
        try:
            ways = splitWay.split(",")

            for i in ways:
                stripText =  i.strip()
                if '-' in splitWay:
                    p = [  int(i.strip())  for  i in stripText.split('-')]
                else:
                    p = int(i.strip())
                funcWays.append(p)
        except:
           self.showDialog("指定文件分页方式有误！")

        inputPDF = self.inputPDF
        totalNumber = inputPDF.numPages
        pages = []


        def doPage( page, size, total):
            t =  [*range(page, page + size, 1 )]
            t = list(filter(lambda x: x <= total, t))
            return t





        if len(funcWays) == 1 and  (not isinstance(funcWays[0], list)):
            #固定分页
            tmp = [ *range(1, totalNumber + 1, funcWays[0])]
            pages = [ doPage(i, funcWays[0], totalNumber) for i in tmp  ]

        else:
            #自定义分页

            if isinstance(funcWays[0], list):
                # [[1,2], [3,4]

                tmp = []
                for i in funcWays:
                    t = [*range(i[0], i[1] + 1, 1)]
                    tmp.append(t)
                pages = tmp




            else:

                tmp = []
                start = 1
                for i in funcWays:
                    end = start + i
                    if end > totalNumber:
                        end = totalNumber + 1
                    p = [*range(start,  end, 1)]
                    start = p[-1] + 1
                    tmp.append(p)
                pages = tmp


        dirname = os.path.join(self.outputFoldPath, self.outputFileNameValue)


        if categoryByTime == 2:
            timeDir = strftime("%Y-%m-%d", gmtime())
            dirname = os.path.join(self.outputFoldPath,timeDir, self.outputFileNameValue)
            print(self.outputFoldPath,timeDir, self.outputFileNameValue)

        if not os.path.exists(dirname):
            os.makedirs(dirname)





        print("last pages", pages)
        for i in pages:
            output = PdfFileWriter()
            for p in i:
                if p <= totalNumber:
                    output.addPage(inputPDF.getPage(p-1))
                    self.progressStatus.setValue(p % totalNumber * 100)
            stri = []
            if len(i) == 1:
                stri = [str(d) for d  in i]
            else:
                stri =[str(i[0]), str(i[-1])]
            name = self.outputFileNameValue + '-'.join(stri) + '.pdf'
            with open(os.path.join(dirname, name), "wb") as outputStream:
                output.write(outputStream)

        self.doButton.setVisible(True)
        self.progressStatus.setVisible(False)

if __name__ == '__main__':
    appctxt = ApplicationContext()
    stylesheet = appctxt.get_resource('styles.qss')
    configPath = appctxt.get_resource("config.json")
    with open(configPath) as f:
        config = json.load(f)
        print(config)
    appctxt.app.setStyleSheet(open(stylesheet).read())
    window = MainWindow(appctxt, config, configPath)
    window.show()
    exit_code = appctxt.app.exec_()
    sys.exit(exit_code)