# создание базы в экселе и заполнение ее новыми строками
# -*- coding: utf-8 -*-
import clr

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

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

# Запуск главной функции (для дебага и возврата основного массива данных в переменной)
revData = run()
