# создание базы в экселе и заполнение ее новыми строками
# -*- coding: utf-8 -*-
import clr
import System
from System import Array

import operator

clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
from System.Runtime.InteropServices import Marshal

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Drawing import *
from System.Windows.Forms import *

def indForDelete(l):
    IndUnit = l.index("AG_Spc_ЕдИзм")
    IndPos = l.index("AG_Spc_Позиция")
    IndMat = l.index("AG_Spc_Материал")
    IndCount = l.index("AG_Spc_Количество")
    IndLevel = l.index("AG_Spc_Уровень")
    IndToSpec = l.index("AG_Spc_Занесение в спецификацию")
    return [IndPos, IndMat, IndUnit, IndCount, IndLevel, IndToSpec]

def searchData(zippedLists):
    searchData = []
    for i in zippedLists:
        _=""
        for ii in i:
            if ii != None:
                _+= ii
        searchData.append(_)
    return searchData

def run():
    viewSchedule = doc.ActiveView
    userName = uiapp.Application.Username

    if viewSchedule.ViewType != ViewType.Schedule:
        MessageBox.Show("Откройте спецификацию", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information)
    else:
        sectionData = viewSchedule.GetTableData().GetSectionData(SectionType.Body)
        numberOfRows = sectionData.NumberOfRows
        numberOfColumns = sectionData.NumberOfColumns
        headData = [viewSchedule.GetCellText(SectionType.Body, 0, column) for column in range(numberOfColumns)]
        #спека ревита, очищенная от ненужных столбцов
        revData = [[viewSchedule.GetCellText(SectionType.Body, row, column) for column in range(numberOfColumns) if column not in indForDelete(headData)] for row in range(numberOfRows)]
        revData[0].append("Имя пользователя")
        revData[1].append("")
        for i in revData[2:]:
            i.append(userName)
        numberOfColumns = len(revData[0])
    return revData

revData = run()

'''
excel = Excel.ApplicationClass()

excel.Visible = False
path = r"C:\Users\a-suchkova\Desktop\TEST\БАЗА.xlsx"

workbooks = excel.Workbooks
workbook = workbooks.Open(path)
worksheet = workbook.Worksheets[1]
usedRange = worksheet.UsedRange                                                                                                                    
literals = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

# если таблица пустая, то заполняем ее полностью данными из экспликации
xlData = []
try:
    xlData = list(usedRange.Value2)
except:
    xlrange = worksheet.Range["A1:"+literals[numberOfColumns-1]+str(numberOfRows)]
    a = Array.CreateInstance(str, numberOfRows, numberOfColumns)
    for row in range(numberOfRows):
        for column in range(numberOfColumns):
            a[row, column] = revData[row][column]
    xlrange.Value2 = a

if xlData: 
    # поисковая таблица из эксель
    xlhead = list(usedRange.Rows(1).Value2)
    xlIndFT = xlhead.index("Семейство и типоразмер")
    xlIndSpcThi = xlhead.index("AG_Spc_Толщина Угол")
    xlIndSpcSize = xlhead.index("AG_Spc_Размер")
    xlIndCode = xlhead.index("AG_Spc_Код категории")

    FamTypeDataXl = list(usedRange.Columns(xlIndFT+1).Value2)
    ThiDataXl = list(usedRange.Columns(xlIndSpcThi+1).Value2)
    SizeDataXl = list(usedRange.Columns(xlIndSpcSize+1).Value2)

    searchDataXl = searchData(zip(FamTypeDataXl, ThiDataXl, SizeDataXl))

    # поисковая таблица из спеки ревита
    revIndFT = revData[0].index("Семейство и типоразмер")
    revIndSpcThi = revData[0].index("AG_Spc_Толщина Угол")
    revIndSpcSize = revData[0].index("AG_Spc_Размер")

    FamTypeDataRev = [i[revIndFT] for i in revData]
    ThiDataRev = [i[revIndSpcThi] for i in revData]
    SizeDataRev = [i[revIndSpcSize] for i in revData]

    searchDataRev = searchData(zip(FamTypeDataRev, ThiDataRev, SizeDataRev))

    # поиск разницы между searchDataXl и searchDataRev
    notInBase = filter(lambda x: x not in searchDataXl, searchDataRev)

    # делим список из экселя по рядам
    xlData2 = [xlData[i:i+len(xlhead)] for i in range(0, len(xlData), len(xlhead))]

    # добавляем в список для эксель недостающие данные
    if notInBase:
        newData = []
        for i in notInBase:
            ind = searchDataRev.index(i)
            revData[ind].append(userName)
            newData.append(revData[ind])

        numberOfRows = len(xlData2)
        numberOfColumns = len(xlData2[0])
        
        xlrange = worksheet.Range["A"+str(numberOfRows+1)+":"+literals[numberOfColumns-1]+str(numberOfRows+len(newData))]
        a = Array.CreateInstance(str, len(newData), numberOfColumns)
        for row in range(len(newData)):
            for column in range(numberOfColumns):
                if newData[row][column] != None:
                    a[row, column] = newData[row][column].ToString()
        xlrange.Value2 = a
        MessageBox.Show("Таблица обновлена", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information)
        workbook.Close(True)
        
    else:
        MessageBox.Show("Новых данных для базы нет", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information)
        workbook.Close()
        


workbooks.Close()
excel.Quit()

Marshal.ReleaseComObject(usedRange) 
Marshal.ReleaseComObject(worksheet)
Marshal.ReleaseComObject(workbook)
Marshal.ReleaseComObject(workbooks)
Marshal.ReleaseComObject(excel)	

#MessageBox.Show(str("lol"), "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information)
'''
