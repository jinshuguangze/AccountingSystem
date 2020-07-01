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
        if(filePath != ''):
            icon = QtGui.QIcon()
            icon.addPixmap(Image.open(
                BytesIO(base64.b64decode(img))).toqpixmap())
            self.sheetInput1.setIcon(icon)
            self.sheetInput1.setIconSize(QtCore.QSize(350, 350))

    def sheetInput2_click(self):
        filePath, filetype = QFileDialog.getOpenFileName(self, '选择金额表格', os.path.join(
            os.path.expanduser("~"), 'Desktop'), 'Excel files(*.xlsx , *.xls)')
        self.openedFile2 = filePath
        if(filePath != ''):
            icon = QtGui.QIcon()
            icon.addPixmap(Image.open(
                BytesIO(base64.b64decode(img))).toqpixmap())
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
                dfo1_1 = pd.DataFrame({'轮入支局': inName})

                name = df1.loc[:, '员工姓名'].values.tolist()
                dfo2 = pd.DataFrame({'员工姓名': name})

                attribute = df1.loc[:, '用工性质'].values.tolist()
                dfo3 = pd.DataFrame({'用工性质': attribute})

                outName = df1.loc[:, '轮出支局'].values.tolist()
                dfo1_2 = pd.DataFrame({'轮出支局': outName})

                inNameCheck = df2.loc[:, '轮入支局'].values.tolist()
                attributeCheck = df2.loc[:, '用工性质'].values.tolist()
                moneyCheck = df2.loc[:, '核算金额'].values.tolist()

                n = -1
                flag = [0]*df1.shape[0]
                moneytotal1 = [0]*df1.shape[0]
                moneytotal2 = [0]*df1.shape[0]
                moneytotal3 = [0]*df1.shape[0]
                for i in inName:
                    n = n+1
                    m = -1
                    for j in inNameCheck:
                        m = m+1
                        if(i == j and attribute[n] == attributeCheck[m]):
                            moneytotal1[n] = moneyCheck[m]
                            flag[n] = flag[n]+1
                            break

                    m = -1
                    for j in inNameCheck:
                        m = m+1
                        if(outName[n] == j and attribute[n] == attributeCheck[m]):
                            moneytotal2[n] = moneyCheck[m]
                            flag[n] = flag[n]+2
                            break
                    moneytotal1[n] = moneytotal1[n]/2
                    moneytotal2[n] = moneytotal2[n]/2
                    moneytotal3[n] = moneytotal1[n] + moneytotal2[n]

                flag = list(map(str, flag))
                moneytotal1 = list(map(str, moneytotal1))
                moneytotal2 = list(map(str, moneytotal2))
                moneytotal3 = list(map(str, moneytotal3))
                n = -1
                for i in moneytotal3:
                    n = n+1
                    if(flag[n] == '0'):
                        moneytotal1[n] = moneytotal1[n]+'(轮入支局未匹配成功！)'
                        moneytotal2[n] = moneytotal2[n]+'(轮出支局未匹配成功！)'
                        moneytotal3[n] = i+'(轮入和轮出支局均未匹配成功！)'
                    elif(flag[n] == '1'):
                        moneytotal1[n] = moneytotal1[n]+'(轮入支局未匹配成功！)'
                        moneytotal3[n] = i+'(轮入支局未匹配成功！)'
                    elif(flag[n] == '2'):
                        moneytotal2[n] = moneytotal2[n]+'(轮出支局未匹配成功！)'
                        moneytotal3[n] = i+'(轮出支局未匹配成功！)'

                dfo4 = pd.DataFrame({'轮入核发金额': moneytotal1})
                dfo5 = pd.DataFrame({'轮出核发金额': moneytotal2})
                dfo6 = pd.DataFrame({'核发总金额': moneytotal3})

                writer = pd.ExcelWriter(filePath)
                dfo1_1.to_excel(writer, sheet_name='员工明细',
                                startcol=0, index=False)
                dfo1_2.to_excel(writer, sheet_name='员工明细',
                                startcol=1, index=False)
                dfo2.to_excel(writer, sheet_name='员工明细',
                              startcol=2, index=False)
                dfo3.to_excel(writer, sheet_name='员工明细',
                              startcol=3, index=False)
                dfo4.to_excel(writer, sheet_name='员工明细',
                              startcol=4, index=False)
                dfo5.to_excel(writer, sheet_name='员工明细',
                              startcol=5, index=False)
                dfo6.to_excel(writer, sheet_name='员工明细',
                              startcol=6, index=False)
                writer.save()

                inNameList = []
                [inNameList.append(i)
                 for i in inNameCheck if not i in inNameList]
                secTotal = [0]*len(inNameList)
                Ain = [0]*len(inNameList)
                Bin = [0]*len(inNameList)
                Cin = [0]*len(inNameList)
                Din = [0]*len(inNameList)
                Aout = [0]*len(inNameList)
                Bout = [0]*len(inNameList)
                Cout = [0]*len(inNameList)
                Dout = [0]*len(inNameList)
                n = -1
                for i in inNameList:
                    n = n+1
                    shotin = []
                    shotout = []
                    for j in range(0, df1.shape[0]):
                        if(i == inName[j]):
                            shotin.append(j)
                        if(i == outName[j]):
                            shotout.append(j)

                    for j in range(0, len(shotin)):
                        if(attribute[shotin[j]] == 'A'):
                            Ain[n] = Ain[n]+1
                        elif(attribute[shotin[j]] == 'B'):
                            Bin[n] = Bin[n]+1
                        elif(attribute[shotin[j]] == 'C'):
                            Cin[n] = Cin[n]+1
                        else:
                            Din[n] = Din[n]+1
                        if(moneytotal1[shotin[j]].find('(') != -1):
                            secTotal[n] = secTotal[n] + \
                                float(moneytotal1[shotin[j]].split('(')[0])
                        else:
                            secTotal[n] = secTotal[n] + \
                                float(moneytotal1[shotin[j]])

                    for j in range(0, len(shotout)):
                        if(attribute[shotout[j]] == 'A'):
                            Aout[n] = Aout[n]+1
                        elif(attribute[shotout[j]] == 'B'):
                            Bout[n] = Bout[n]+1
                        elif(attribute[shotout[j]] == 'C'):
                            Cout[n] = Cout[n]+1
                        else:
                            Dout[n] = Dout[n]+1
                        if(moneytotal2[shotout[j]].find('(') != -1):
                            secTotal[n] = secTotal[n] + \
                                float(moneytotal2[shotout[j]].split('(')[0])
                        else:
                            secTotal[n] = secTotal[n] + \
                                float(moneytotal2[shotout[j]])

                dfos1 = pd.DataFrame({'支局名称': list(map(str, inNameList))})
                dfos2 = pd.DataFrame({'核发总金额': list(map(str, secTotal))})
                dfos3 = pd.DataFrame({'原A类员工人数': list(map(str, Aout))})
                dfos4 = pd.DataFrame({'原B类员工人数': list(map(str, Bout))})
                dfos5 = pd.DataFrame({'原C类员工及劳务承揽人数': list(map(str, Cout))})
                dfos6 = pd.DataFrame({'原其他未知类别人数': list(map(str, Dout))})
                dfos7 = pd.DataFrame({'现A类员工人数': list(map(str, Ain))})
                dfos8 = pd.DataFrame({'现B类员工人数': list(map(str, Bin))})
                dfos9 = pd.DataFrame({'现C类员工及劳务承揽人数': list(map(str, Cin))})
                dfos10 = pd.DataFrame({'现其他未知类别人数': list(map(str, Din))})
                dfos1.to_excel(writer, sheet_name='支局明细',
                               startcol=0, index=False)
                dfos2.to_excel(writer, sheet_name='支局明细',
                               startcol=1, index=False)
                dfos3.to_excel(writer, sheet_name='支局明细',
                               startcol=2, index=False)
                dfos4.to_excel(writer, sheet_name='支局明细',
                               startcol=3, index=False)
                dfos5.to_excel(writer, sheet_name='支局明细',
                               startcol=4, index=False)
                dfos6.to_excel(writer, sheet_name='支局明细',
                               startcol=5, index=False)
                dfos7.to_excel(writer, sheet_name='支局明细',
                               startcol=6, index=False)
                dfos8.to_excel(writer, sheet_name='支局明细',
                               startcol=7, index=False)
                dfos9.to_excel(writer, sheet_name='支局明细',
                               startcol=8, index=False)
                dfos10.to_excel(writer, sheet_name='支局明细',
                                startcol=9, index=False)

                writer.save()
                writer.close()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    accoutingSystem = AccoutingSystem()
    accoutingSystem.show()
    sys.exit(app.exec_())
