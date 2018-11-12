# -*- coding: utf-8 -*-
# заполнение непользовательских параметров в элементах (Система, ТолщинаУгол, Размер, Количество, ЕдИзм, Уровень, Код категории, Наименование (генерит для воздуховодов и фасонки предварительно))
import clr
import System

import sys
sys.path.append("C:/Program Files (x86)/IronPython 2.7/Lib")

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *


def cleanParameters(l):
    if l:
        for e in l:
            spcName = e.LookupParameter("AG_Spc_Наименование")
            spcType = e.LookupParameter("AG_Spc_Тип")
            spcArt = e.LookupParameter("AG_Spc_Артикул")
            spcProd = e.LookupParameter("AG_Spc_Изготовитель")
            spcMass = e.LookupParameter("AG_Spc_Масса")
            spcPos = e.LookupParameter("AG_Spc_Позиция")

            spcName.Set("")
            spcType.Set("")
            spcArt.Set("")
            spcProd.Set("")
            spcMass.Set("")
            spcPos.Set("")

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

t = Transaction(doc, "CleanParameters")
t.Start()
cleanParameters(ducts)
cleanParameters(flexDuct)
cleanParameters(fitings)
cleanParameters(accessory)
cleanParameters(terminal)
cleanParameters(isol)
cleanParameters(equipment)
cleanParameters(pipes)
cleanParameters(pipeFitings)
cleanParameters(pipeAccessory)
cleanParameters(flexPipe)
cleanParameters(plumbing)
cleanParameters(sprinklers)
cleanParameters(pipeIsol)
t.Commit()


