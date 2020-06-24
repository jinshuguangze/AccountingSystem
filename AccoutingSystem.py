import sys
from threading import Timer

import pandas as pd
from PyQt5 import Qt, QtCore, QtGui, QtWidgets
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *

from AccoutingSystemUI import *
from ButtonIconFinish_jpg import img


class AccoutingSystem(QMainWindow, Ui_mainWindows):
    openedFile1 = ''
    openedFile2 = ''

    def __init__(self, parent=None):
        super(AccoutingSystem, self).__init__(parent)
        self.setupUi(self)
        self.sheetInput1.clicked.connect(self.sheetInput1_click)
        self.sheetInput2.clicked.connect(self.sheetInput2_click)
        self.generateOutput.clicked.connect(self.generateOutput_click)

    def sheetInput1_click(self):
        filePath, filetype = QFileDialog.getOpenFileName(self, '选择员工表格', os.path.join(
            os.path.expanduser("~"), 'Desktop'), 'Excel files(*.xlsx , *.xls)')
        self.openedFile1 = filePath
        icon = QtGui.QIcon()
        icon.addPixmap(Image.open(BytesIO(base64.b64decode(img))).toqpixmap())
        self.sheetInput1.setIcon(icon)
        self.sheetInput1.setIconSize(QtCore.QSize(350, 350))

    def sheetInput2_click(self):
        filePath, filetype = QFileDialog.getOpenFileName(self, '选择金额表格', os.path.join(
            os.path.expanduser("~"), 'Desktop'), 'Excel files(*.xlsx , *.xls)')
        self.openedFile2 = filePath
        icon = QtGui.QIcon()
        icon.addPixmap(Image.open(BytesIO(base64.b64decode(img))).toqpixmap())
        self.sheetInput2.setIcon(icon)
        self.sheetInput2.setIconSize(QtCore.QSize(350, 350))

    def generateOutput_click(self):
        if(self.openedFile1 != '' and self.openedFile2 != ''):
            filePath, filetype = QFileDialog.getSaveFileName(None, '保存生成表格', os.path.join(
                os.path.expanduser("~"), 'Desktop')+'/生成表格.xlsx', 'Excel files(*.xlsx , *.xls)')
            if(filePath != ''):
                df1 = pd.read_excel(self.openedFile1)
                df2 = pd.read_excel(self.openedFile2)

                inName = df1.loc[:, '轮入支局'].values.tolist()
                dfo1 = pd.DataFrame({'轮入支局': inName})

                name = df1.loc[:, '员工姓名'].values.tolist()
                dfo2 = pd.DataFrame({'员工姓名': name})

                attribute = df1.loc[:, '员工属性'].values.tolist()
                dfo3 = pd.DataFrame({'员工属性': attribute})

                outName = df1.loc[:, '轮出支局'].values.tolist()
                inNameCheck = df2.loc[:, '轮入支局'].values.tolist()
                attributeCheck = df2.loc[:, '员工属性'].values.tolist()
                moneyCheck = df2.loc[:, '核算金额'].values.tolist()

                n = -1
                flag = [0]*df1.shape[0]
                moneytotal = [0]*df1.shape[0]
                for i in inName:
                    n = n+1
                    m = -1
                    for j in inNameCheck:
                        m = m+1
                        if(i == j and attribute[n] == attributeCheck[m]):
                            moneytotal[n] = moneytotal[n]+moneyCheck[m]
                            flag[n] = flag[n]+1
                            break

                    m = -1
                    for j in inNameCheck:
                        m = m+1
                        if(outName[n] == j and attribute[n] == attributeCheck[m]):
                            moneytotal[n] = moneytotal[n]+moneyCheck[m]
                            flag[n] = flag[n]+2
                            break
                    moneytotal[n] = moneytotal[n]/2

                flag = list(map(str, flag))
                moneytotal = list(map(str, moneytotal))
                n = -1
                for i in moneytotal:
                    n = n+1
                    if(flag[n] == '0'):
                        moneytotal[n] = i+'(轮入和论出支局均未匹配成功！)'
                    elif(flag[n] == '1'):
                        moneytotal[n] = i+'(论出支局未匹配成功！)'
                    elif(flag[n] == '2'):
                        moneytotal[n] = i+'(论入支局未匹配成功！)'

                dfo4 = pd.DataFrame({'核发金额': moneytotal})

                writer = pd.ExcelWriter(filePath)
                dfo1.to_excel(writer, sheet_name='总表', startcol=0, index=False)
                dfo2.to_excel(writer, sheet_name='总表', startcol=1, index=False)
                dfo3.to_excel(writer, sheet_name='总表', startcol=2, index=False)
                dfo4.to_excel(writer, sheet_name='总表', startcol=3, index=False)
                writer.save()

                inNameList = []
                [inNameList.append(i)
                 for i in inNameCheck if not i in inNameList]
                for i in inNameList:
                    shot = []
                    for j in range(0, df1.shape[0]):
                        if(i == inName[j]):
                            shot.append(j)

                    secName = []
                    secAttribute = []
                    secMoneytotal = []
                    secTotal = 0
                    for j in range(0, len(shot)):
                        secName.append(name[shot[j]])
                        secAttribute.append(attribute[shot[j]])
                        secMoneytotal.append(moneytotal[shot[j]])
                        if(moneytotal[shot[j]].find('(') != -1):
                            secTotal = secTotal + \
                                float(moneytotal[shot[j]].split('(')[0])
                        else:
                            secTotal = secTotal+float(moneytotal[shot[j]])

                    secInName = [i]*len(shot)
                    secInName.append('合计：')
                    secMoneytotal.append(str(secTotal))
                    dfos1 = pd.DataFrame({'轮入支局': secInName})
                    dfos2 = pd.DataFrame({'员工姓名': secName})
                    dfos3 = pd.DataFrame({'员工属性': secAttribute})
                    dfos4 = pd.DataFrame({'核发金额': secMoneytotal})
                    dfos1.to_excel(writer, sheet_name=i,
                                   startcol=0, index=False)
                    dfos2.to_excel(writer, sheet_name=i,
                                   startcol=1, index=False)
                    dfos3.to_excel(writer, sheet_name=i,
                                   startcol=2, index=False)
                    dfos4.to_excel(writer, sheet_name=i,
                                   startcol=3, index=False)
                    writer.save()
                writer.close()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    accoutingSystem = AccoutingSystem()
    accoutingSystem.show()
    sys.exit(app.exec_())
