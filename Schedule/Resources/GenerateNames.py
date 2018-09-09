# -*- coding: utf-8 -*-
# генерация наименований для соединительных деталей, труб, воздуховодов

import clr
import System
import math

import sys
sys.path.append("C:/Program Files (x86)/IronPython 2.7/Lib")

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Drawing import *
from System.Windows.Forms import *

# генерация имени
def generateName(elems):
    if elems:
        for e in elems:
            thi = e.LookupParameter("AG_Spc_Толщина Угол").AsString()
            if e.Category.Id.IntegerValue == int(BuiltInCategory.OST_DuctFitting) or e.Category.Id.IntegerValue == int(BuiltInCategory.OST_PipeFitting):
                typeDetail = e.MEPModel.PartType
                if typeDetail == PartType.Elbow:
                    angle = thi.split(" ")[1]
                    thi = thi.split(" ")[0]
                    catName = "Отвод"
                    catName = catName + " " + angle
                elif typeDetail == PartType.Transition:
                    catName = "Переход"
                elif typeDetail == PartType.Tee:
                    catName = "Тройник"
                elif typeDetail == PartType.SpudAdjustable or typeDetail == PartType.TapAdjustable:
                    catName = "Врезка"
                elif typeDetail == PartType.Union:
                    catName = "Соединение"
                elif typeDetail == PartType.Cross:
                    catName = "Крестовина"
                elif typeDetail == PartType.Cap:
                    catName = "Заглушка"
                else:
                    catName = ""
            elif e.Category.Id.IntegerValue == int(BuiltInCategory.OST_DuctCurves):
                catName = "Воздуховод"
            elif e.Category.Id.IntegerValue == int(BuiltInCategory.OST_FlexDuctCurves):
                catName = "Гибкий воздуховод"
            elif e.Category.Id.IntegerValue == int(BuiltInCategory.OST_FlexPipeCurves):
                catName = "Гибкий трубопровод"
            else:
                catName = "Трубопровод"

            spcSize = e.LookupParameter("AG_Spc_Размер").AsString()
            spcMaterial = doc.GetElement(e.GetTypeId()).LookupParameter("AG_Spc_Материал").AsString()
            if spcMaterial != None and spcMaterial.strip() != "":
                mat = spcMaterial + ", "
            else:
                mat = ""
            # продумать алгоритм расчета толщины гофрированных труб
            if (e.Category.Id.IntegerValue == int(BuiltInCategory.OST_DuctFitting) or
                e.Category.Id.IntegerValue == int(BuiltInCategory.OST_DuctCurves) or 
                e.Category.Id.IntegerValue == int(BuiltInCategory.OST_FlexDuctCurves)):
                discribe = mat + "толщ. "+ thi +" мм, " + spcSize
            elif (e.Category.Id.IntegerValue == int(BuiltInCategory.OST_PipeCurves) or 
                e.Category.Id.IntegerValue == int(BuiltInCategory.OST_PipeFitting)):
                discribe = mat + spcSize + "x" + thi
            elif e.Category.Id.IntegerValue == int(BuiltInCategory.OST_FlexPipeCurves):
                discribe = mat + spcSize

            expectedName = catName+" "+ discribe
            spcName = e.LookupParameter("AG_Spc_Наименование")
            spcName.Set(expectedName)

# воздуховоды
ducts = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType().ToElements()
# гибкий воздуховод
flexDuct = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FlexDuctCurves).WhereElementIsNotElementType().ToElements()
# соед детали воздуховодов
fitings = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctFitting).WhereElementIsNotElementType().ToElements()
# трубы
pipes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements()
# соед детали труб
pipeFitings = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeFitting).WhereElementIsNotElementType().ToElements()
# гибкий трубопровод
flexPipe = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FlexPipeCurves).WhereElementIsNotElementType().ToElements()

t = Transaction(doc, "GenerateNames")
t.Start()

generateName(ducts)
generateName(flexDuct)
generateName(fitings)
generateName(pipes)
generateName(pipeFitings)
generateName(flexPipe)

t.Commit()

MessageBox.Show("ОК", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information)

