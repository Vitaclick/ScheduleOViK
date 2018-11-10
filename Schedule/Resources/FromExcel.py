# заполнение параметров элементов из базы эксель
# -*- coding: utf-8 -*-

import clr
import sys

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *


# функции
def setValue(excelValue, revitValue):
    if excelValue != None and excelValue != "" and revitValue.AsString() != excelValue:
        revitValue.Set(excelValue.ToString())

def setParameters(elems, dictXl):
    if elems:
        for e in elems:
            spcName = e.LookupParameter("AG_Spc_Наименование")
            scpType = e.LookupParameter("AG_Spc_Тип")
            spcArt = e.LookupParameter("AG_Spc_Артикул")
            spcProd = e.LookupParameter("AG_Spc_Изготовитель")
            spcMass = e.LookupParameter("AG_Spc_Масса")
            spcRemark = e.LookupParameter("AG_Spc_Примечание")
            spcMTO = e.LookupParameter("AG_Spc_МТО")

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
                setValue(dictXl[code][5], spcRemark)
                setValue(dictXl[code][6], spcMTO)
            
def toList(elems):
    if isinstance(elems, list):
        return elems
    else:
        return [elems]

def run():
    # import Spreadsheet data
    data = sys.dataFromSpreadsheet
    # transposed data from Spreadsheet
    transposedDataFromSpreadsheet = [[data[j][i] for j in range(len(data))] for i in range(len(data[0]))]

    colFTxl = transposedDataFromSpreadsheet[0]
    colThixl = transposedDataFromSpreadsheet[2]
    colSizexl = transposedDataFromSpreadsheet[3]

    keys = []
    for tup in zip(colFTxl, colThixl, colSizexl):
        code = ""
        for i in tup:
            if i != None:
                code += i
        keys.append(code)

    colNamexl = transposedDataFromSpreadsheet[1]
    colTypexl = transposedDataFromSpreadsheet[4]
    colArtxl = transposedDataFromSpreadsheet[5]
    colProdxl = transposedDataFromSpreadsheet[6]
    colMassxl = transposedDataFromSpreadsheet[7]
    colRemarkxl = transposedDataFromSpreadsheet[8]
    colMTOxl = transposedDataFromSpreadsheet[9]

    values = zip(colNamexl, colTypexl, colArtxl, colProdxl, colMassxl, colRemarkxl, colMTOxl)

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
    # обобщенные модели
    generic = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()

    t = Transaction(doc, "SetParameters")
    t.Start()
    setParameters(ducts, dictXl)
    setParameters(flexDuct, dictXl)
    setParameters(fitings, dictXl)
    setParameters(accessory, dictXl)
    setParameters(terminal, dictXl)
    setParameters(isol, dictXl)
    setParameters(equipment, dictXl)
    setParameters(pipes, dictXl)
    setParameters(pipeFitings, dictXl)
    setParameters(pipeAccessory, dictXl)
    setParameters(flexPipe, dictXl)
    setParameters(plumbing, dictXl)
    setParameters(sprinklers, dictXl)
    setParameters(pipeIsol, dictXl)
    setParameters(generic, dictXl)
    t.Commit()
run()
