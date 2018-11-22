# заполнение параметров элементов из базы эксель
# -*- coding: utf-8 -*-

import clr
import sys

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

NONELIST = (None, '', ' ')
parameters = [
    '',
    'AG_Spc_Наименование',
    '','',
    'AG_Spc_Тип',
    'AG_Spc_Артикул',
    'AG_Spc_Изготовитель',
    'AG_Spc_Масса',
    'AG_Spc_Примечание',
    '','','']

def Uniq(e):
    famName = e.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
    spcThi = e.LookupParameter("AG_Spc_Толщина Угол").AsString()
    spcSize = e.LookupParameter("AG_Spc_Размер").AsString()
    code = ''
    for i in [famName, spcThi, spcSize]:
        if i != None:
            code += i
    return code

def setParameters(elems):
    for e in elems:
        key = Uniq(e)
        if key in keysFromSpreadsheet:
            rowIndex = keysFromSpreadsheet.index(key)
            dataRow = data[rowIndex]
            for k, cell in enumerate(dataRow):
                if cell not in NONELIST and parameters[k] != '':
                    p = e.LookupParameter(parameters[k])
                    p.Set(cell.ToString())


# import Spreadsheet data
data = sys.dataFromSpreadsheet

keysFromSpreadsheet = []
for row in data:
    code = ''
    unc = (row[0], row[2], row[3])
    for i in unc:
        if i:
            code += i
    keysFromSpreadsheet.append(code)

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
# обобщенные модели
generic = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()

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
setParameters(generic)
t.Commit()


