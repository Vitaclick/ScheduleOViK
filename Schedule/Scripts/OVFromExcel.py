# заполнение параметров элементов из базы эксель
# -*- coding: utf-8 -*-

import clr
import System
from System import Array

import sys
sys.path.append("C:/Program Files (x86)/IronPython 2.7/Lib")

clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
from System.Runtime.InteropServices import Marshal

from operator import itemgetter
from itertools import groupby

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Drawing import *
from System.Windows.Forms import *

# -------------------------функции

def setValue(excelValue, revitValue):
    if excelValue != None and excelValue != "" and revitValue.AsString() != excelValue:
        revitValue.Set(excelValue.ToString())

def setParameters(elems):
    if elems:
        for e in elems:
            spcName = e.LookupParameter("AG_Spc_Наименование")
            scpType = e.LookupParameter("AG_Spc_Тип")
            spcArt = e.LookupParameter("AG_Spc_Артикул")
            spcProd = e.LookupParameter("AG_Spc_Изготовитель")
            spcMass = e.LookupParameter("AG_Spc_Масса")

            famName = e.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
            spcThi = e.LookupParameter("AG_Spc_Толщина Угол").AsString()
            spcSize = e.LookupParameter("AG_Spc_Размер").AsString()
            code = ""
            for i in [famName, spcThi, spcSize]:
                if i != None:
                    code += i
            if code in dictXl:
                setValue(dictXl[code][0], spcName)
                setValue(dictXl[code][1], scpType)
                setValue(dictXl[code][2], spcArt)
                setValue(dictXl[code][3], spcProd)
                setValue(dictXl[code][4], spcMass)
            
def toList(elems):
    if isinstance(elems, list):
        return elems
    else:
        return [elems]

excel = Excel.ApplicationClass()
excel.Visible = False
path = r"C:\Users\a-suchkova\Desktop\TEST\БАЗА.xlsx"
workbooks = excel.Workbooks
workbook = workbooks.Open(path)
worksheet = workbook.Worksheets[1]
usedRange = worksheet.UsedRange

# заголовок с именами параметров
xlhead = list(usedRange.Rows(1).Value2)
indFT = xlhead.index("Семейство и типоразмер")
# индексы параметров
indSpcSize = xlhead.index("AG_Spc_Размер")
indSpcThi = xlhead.index("AG_Spc_Толщина Угол")

indSpcName = xlhead.index("AG_Spc_Наименование")
indSpcType = xlhead.index("AG_Spc_Тип")
indSpcArt = xlhead.index("AG_Spc_Артикул")
indSpcProd = xlhead.index("AG_Spc_Изготовитель")
indSpcMass = xlhead.index("AG_Spc_Масса")

colFTxl = list(usedRange.Columns(indFT+1).Value2)
colThixl = list(usedRange.Columns(indSpcThi+1).Value2)
colSizexl = list(usedRange.Columns(indSpcSize+1).Value2)

keys = []
for tup in zip(colFTxl, colThixl, colSizexl):
    code = ""
    for i in tup:
        if i != None:
            code += i
    keys.append(code)

colNamexl = list(usedRange.Columns(indSpcName+1).Value2)
colTypexl = list(usedRange.Columns(indSpcType+1).Value2)
colArtxl = list(usedRange.Columns(indSpcArt+1).Value2)
colProdxl = list(usedRange.Columns(indSpcProd+1).Value2)
colMassxl = list(usedRange.Columns(indSpcMass+1).Value2)

values = zip(colNamexl, colTypexl, colArtxl, colProdxl, colMassxl)

dictXl = dict(zip(keys, values))

# воздуховоды
ducts = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType().ToElements()
# гибкий воздуховод
flexDuct = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FlexDuctCurves).WhereElementIsNotElementType().ToElements()
# соед детали воздуховодов
fitings = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctFitting).WhereElementIsNotElementType().ToElements()
# арматура
accessory = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctAccessory).WhereElementIsNotElementType().ToElements()
# воздухораспределители
terminal = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType().ToElements()
# изоляция
isol = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctInsulations).WhereElementIsNotElementType().ToElements()
# оборудование
equipment = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().ToElements()
# трубы
pipes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements()
# соед детали труб
pipeFitings = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeFitting).WhereElementIsNotElementType().ToElements()
# арматура труб
pipeAccessory = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType().ToElements()
# гибкий трубопровод
flexPipe = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FlexPipeCurves).WhereElementIsNotElementType().ToElements()
# сантехнические приборы
plumbing = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PlumbingFixtures).WhereElementIsNotElementType().ToElements()
# спринклеры
sprinklers = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sprinklers).WhereElementIsNotElementType().ToElements()
# изоляция труб
pipeIsol = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeInsulations).WhereElementIsNotElementType().ToElements()

t = Transaction(doc, "SetParameters")
t.Start()
setParameters(ducts)
setParameters(flexDuct)
setParameters(fitings)
setParameters(accessory)
setParameters(terminal)
setParameters(isol)
setParameters(equipment)
setParameters(pipes)
setParameters(pipeFitings)
setParameters(pipeAccessory)
setParameters(flexPipe)
setParameters(plumbing)
setParameters(sprinklers)
setParameters(pipeIsol)
t.Commit()

workbook.Close()
workbooks.Close()
excel.Quit()

Marshal.ReleaseComObject(usedRange)
Marshal.ReleaseComObject(worksheet)
Marshal.ReleaseComObject(workbook)
Marshal.ReleaseComObject(workbooks)
Marshal.ReleaseComObject(excel)	

MessageBox.Show("ОК", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information)



