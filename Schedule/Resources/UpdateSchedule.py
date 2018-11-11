# -*- coding: utf-8 -*-
# заполнение непользовательских параметров в элементах (Система, ТолщинаУгол, Размер, Количество, ЕдИзм, Уровень, Код категории, Наименование (генерит для воздуховодов и фасонки предварительно))
# продумать алгоритм расчета толщины гофрированных труб

import clr
import System
import math

import sys
sys.path.append("C:/Program Files (x86)/IronPython 2.7/Lib")

from operator import itemgetter
from itertools import groupby

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Drawing import *
from System.Windows.Forms import *

antismoke = ['ВПВ', 'ППВ', 'ДУ', 'КДУ', 'ПВ']
# функция определения настоящего уровня элемента
def findLevel(z, levelsSort):
    n = 0
    while n < len(levels)-1:
        if z <= levelsSort[0].Elevation:
            return levelsSort[0].get_Parameter(BuiltInParameter.LEVEL_ELEV).AsValueString()
            break
        elif levelsSort[n].Elevation <= z < levelsSort[n+1].Elevation:
            return levelsSort[n].get_Parameter(BuiltInParameter.LEVEL_ELEV).AsValueString()
            break
        elif z >= levelsSort[len(levelsSort)-1].Elevation:
            return levelsSort[len(levelsSort)-1].get_Parameter(BuiltInParameter.LEVEL_ELEV).AsValueString()
            break
        n += 1
def itemLevel(e):
    z = e.Location.Point.Z
    lvl = findLevel(z, levelsSort)
    spcLevel = e.LookupParameter("AG_Spc_Уровень")
    spcLevel.Set(str(lvl))
def runLevel(e):
    pipePoints = e.Location.Curve.Tessellate()
    if pipePoints[0].Z < pipePoints[1].Z:
        downZ = pipePoints[0].Z
        upZ = pipePoints[1].Z
    elif pipePoints[0].Z > pipePoints[1].Z:
        downZ = pipePoints[1].Z
        upZ = pipePoints[0].Z
    else:
        downZ = pipePoints[0].Z
        upZ = pipePoints[1].Z
    z = downZ + (upZ - downZ)/2
    lvl = findLevel(z, levelsSort)
    spcLevel = e.LookupParameter("AG_Spc_Уровень")
    spcLevel.Set(str(lvl))

def parSys(e):
    system = e.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
    spcSystem = e.LookupParameter("AG_Spc_Система")
    if system != None:
        spcSystem.Set(system)
def parSize(e):
    size = e.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString()
    if e.Category.Id.IntegerValue == int(BuiltInCategory.OST_DuctFitting) or e.Category.Id.IntegerValue == int(BuiltInCategory.OST_PipeFitting):
        if (e.MEPModel.PartType == PartType.Elbow or 
            e.MEPModel.PartType == PartType.Union or
            e.MEPModel.PartType == PartType.SpudAdjustable or
            e.MEPModel.PartType == PartType.TapAdjustable):
            size = size.split("-")[0]
    if e.Category.Id.IntegerValue == int(BuiltInCategory.OST_DuctAccessory) or e.Category.Id.IntegerValue == int(BuiltInCategory.OST_PipeAccessory):
        size = size.split("-")[0]
    spcSize = e.LookupParameter("AG_Spc_Размер")
    spcSize.Set(size)
def parIsolQuant(e):
    square = e.get_Parameter(BuiltInParameter.RBS_CURVE_SURFACE_AREA).AsDouble()
    squareM = UnitUtils.ConvertFromInternalUnits(square, DisplayUnitType.DUT_SQUARE_METERS)
    spcQuant = e.LookupParameter("AG_Spc_Количество")
    spcQuant.Set(squareM)
def parIsolQuantPipe(e):
    length = e.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()
    len_m = UnitUtils.ConvertFromInternalUnits(length, DisplayUnitType.DUT_METERS)
    spcQuant = e.LookupParameter("AG_Spc_Количество")
    spcQuant.Set(len_m)
def parItemQuant(e):
    spcQuant = e.LookupParameter("AG_Spc_Количество")
    spcQuant.Set(1)
def parDuctQuant(e):
    quant = e.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()
    quantM = round(UnitUtils.ConvertFromInternalUnits(quant, DisplayUnitType.DUT_METERS),2)
    spcQuant = e.LookupParameter("AG_Spc_Количество")
    spcQuant.Set(quantM)
def parUnit(e,val):
    spcUnit = e.LookupParameter("AG_Spc_ЕдИзм")
    spcUnit.Set(val)
def tByConnector(sysName, con, typeIsol):
    shape = con.Shape
    if shape == ConnectorProfileType.Rectangular:
        h = con.Height
        w = con.Width
        maxS = max(h, w)
        maxMM = UnitUtils.ConvertFromInternalUnits(maxS, DisplayUnitType.DUT_MILLIMETERS)
        if typeIsol != None and "EI" in typeIsol:
            if maxMM <= 2000.0:
                t = "0.9"
            else:
                t = "1.2"
        else:
            if sysName != None and any(sysName.__contains__(i) for i in antismoke):
                if maxMM <= 2000.0:
                    t = "0.9"
                else:
                    t = "1.2"
            else:
                if maxMM <= 250.0:
                    t = "0.5"
                elif maxMM <= 1000:
                    t = "0.7"
                else:
                    t = "0.9"
    elif shape == ConnectorProfileType.Round:
        d = con.Radius * 2
        dMM = UnitUtils.ConvertFromInternalUnits(d, DisplayUnitType.DUT_MILLIMETERS)
        if typeIsol != None and "EI" in typeIsol:
            if dMM <= 2000.0:
                t = "0.9"
            else:
                t = "1.2"
        else:
            if sysName != None and any(sysName.__contains__(i) for i in antismoke):
                if dMM <= 800.0:
                    t = "0.9"
                elif dMM <= 1250.0:
                    t = "1.0"
                elif dMM <= 1600.0:
                    t = "1.2"
                else:
                    t = "1.4"
            else:
                if dMM <= 200.0:
                    t = "0.5"
                elif dMM <= 450.0:
                    t = "0.6"
                elif dMM <= 800.0:
                    t = "0.7"
                elif dMM <= 1250.0:
                    t = "1.0"
                elif dMM <= 1600.0:
                    t = "1.2"
                else:
                    t = "1.4"
    return t
# заполнить параметр Толщина Угол для воздуховодов
def setThiDucts(e):
    sysName = e.LookupParameter("AG_Spc_Система").AsString()
    spcThiAngle = e.LookupParameter("AG_Spc_Толщина Угол")
    if e.Category.Id.IntegerValue == int(BuiltInCategory.OST_DuctCurves):
        cc = [i for i in e.ConnectorManager.Connectors]
        con = cc[0]
        typeIsol = e.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_TYPE).AsString()
        thi = tByConnector(sysName, con, typeIsol)
    elif e.Category.Id.IntegerValue == int(BuiltInCategory.OST_FlexDuctCurves):
        thi = "0.15"
    elif e.Category.Id.IntegerValue == int(BuiltInCategory.OST_PipeCurves):
        dOut = e.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble()
        dIn = e.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM).AsDouble()
        t = (dOut-dIn)/2
        thi = str(UnitUtils.ConvertFromInternalUnits(t, DisplayUnitType.DUT_MILLIMETERS))
    spcThiAngle.Set(thi)

# заполнить параметр Толщина Угол для фитингов
def setThiItems(e):
    sysName = e.LookupParameter("AG_Spc_Система").AsString()
    spcThiAngle = e.LookupParameter("AG_Spc_Толщина Угол")
    thi = "-"
    if e.Category.Id.IntegerValue == int(BuiltInCategory.OST_DuctFitting):
        cc = [i for i in e.MEPModel.ConnectorManager.Connectors]
        maxConSquare = 0.0
        for con in cc:
            if con.Shape == ConnectorProfileType.Rectangular:
                conSquare = con.Height * con.Width
            elif con.Shape == ConnectorProfileType.Round:
                conSquare = math.pi * con.Radius ** 2
            if conSquare > maxConSquare:
                outCon = con
        typeIsol = e.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_TYPE).AsString()
        thi = tByConnector(sysName, outCon, typeIsol)
    elif e.Category.Id.IntegerValue == int(BuiltInCategory.OST_PipeFitting):
        cc = [i for i in e.MEPModel.ConnectorManager.Connectors]
        
        for con in cc:
            allRefs = [i for i in con.AllRefs]
            for i in allRefs:
                own = i.Owner
                if own.Category.Id.IntegerValue == int(BuiltInCategory.OST_PipeCurves):
                    thi = own.LookupParameter("AG_Spc_Толщина Угол").AsString()
                    break
    else:
        thi = "-"
    
    typeDetail = e.MEPModel.PartType
    if typeDetail == PartType.Elbow:
        cSet = [i for i in e.MEPModel.ConnectorManager.Connectors]
        con = cSet[0]
        angle = str(int(round(math.degrees(con.Angle)/5.0))*5) + "°"
        thiAngle = thi + " " + angle
    else:
        thiAngle = thi
    spcThiAngle.Set(thiAngle)

def setCategoryCode(e, num):
    spcCode = e.LookupParameter("AG_Spc_Код категории")
    spcCode.Set(num)

def sizePipeIsol(e):
    s = e.LookupParameter("AG_Spc_Размер")
    t = e.LookupParameter("AG_Spc_Толщина Угол")
    vs = e.get_Parameter(BuiltInParameter.RBS_PIPE_CALCULATED_SIZE).AsString()
    if "-" in vs:
        vs = vs.split("-")[0]
    vt = e.get_Parameter(BuiltInParameter.RBS_INSULATION_THICKNESS_FOR_PIPE).AsValueString()
    s.Set(vs)
    t.Set(vt)
def setStock(e, value):
    pStock = e.LookupParameter('AG_Spc_Запас')
    pStock.Set(value)

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
# уровни
levels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
levelsSort = sorted(levels, key = lambda i: i.Elevation)


t = Transaction(doc, "SetParameters")
t.Start()

# оборудование
if equipment:
    for e in equipment:
        parSys(e)
        itemLevel(e)
        parItemQuant(e)
        parUnit(e, "компл.")
        setCategoryCode(e, 10)
        setStock(e, 1.0)

# воздухораспределители
if terminal:
    for e in terminal:
        parSys(e)
        parSize(e)
        parItemQuant(e)
        parUnit(e, "шт.")
        itemLevel(e)
        setCategoryCode(e, 20)
        setStock(e, 1.0)

# арматура
if accessory:
    for e in accessory:
        parSys(e)
        parSize(e)
        parItemQuant(e)
        parUnit(e, "шт.")
        itemLevel(e)
        setCategoryCode(e, 30)
        setStock(e, 1.0)

# фитинги
if fitings:
    for e in fitings:
        parSys(e)
        parSize(e)
        parItemQuant(e)
        parUnit(e, "шт.")
        itemLevel(e)
        setThiItems(e)
        setCategoryCode(e, 40)
        setStock(e, 1.0)

# воздуховоды
if ducts:
    for e in ducts:
        parSys(e)
        parSize(e)
        parDuctQuant(e)
        parUnit(e, "м")
        runLevel(e)
        setThiDucts(e)
        setCategoryCode(e, 50)
        setStock(e, 1.15)

# гибкие воздуховоды
if flexDuct:
    for e in flexDuct:
        parSys(e)
        parSize(e)
        parDuctQuant(e)
        parUnit(e, "м")
        runLevel(e)
        setThiDucts(e)
        setCategoryCode(e, 51)
        setStock(e, 1.15)

# трубы
if pipes:
    for e in pipes:
        parSys(e)
        parSize(e)
        parDuctQuant(e)
        parUnit(e, "м")
        runLevel(e)
        setThiDucts(e)
        setCategoryCode(e, 60)
        setStock(e, 1.15)

# гибкие трубы
if flexPipe:
    for e in flexPipe:
        parSys(e)
        parSize(e)
        parDuctQuant(e)
        parUnit(e, "м")
        runLevel(e)
        setCategoryCode(e, 61)
        setStock(e, 1.15)

# соед детали труб
if pipeFitings:
    for e in pipeFitings:
        parSys(e)
        parSize(e)
        parItemQuant(e)
        parUnit(e, "шт.")
        itemLevel(e)
        setThiItems(e)
        setCategoryCode(e, 70)
        setStock(e, 1.0)

# арматура труб
if pipeAccessory:
    for e in pipeAccessory:
        parSys(e)
        parSize(e)
        parItemQuant(e)
        parUnit(e, "шт.")
        itemLevel(e)
        setCategoryCode(e, 80)
        setStock(e, 1.0)

# сантехника
if plumbing:
    for e in plumbing:
        parSys(e)
        itemLevel(e)
        parItemQuant(e)
        parUnit(e, "шт.")
        setCategoryCode(e, 90)
        setStock(e, 1.0)

# сприклеры
if sprinklers:
    for e in sprinklers:
        parSys(e)
        itemLevel(e)
        parItemQuant(e)
        parUnit(e, "шт.")
        setCategoryCode(e, 100)
        setStock(e, 1.0)

# изоляция воздуховодов
if isol:
    for e in isol:
        parSys(e)
        parIsolQuant(e)
        parUnit(e, "м²")
        runLevel(e)
        setCategoryCode(e, 110)
        setStock(e, 1.45)

# изоляция труб
if pipeIsol:
    for e in pipeIsol:
        parSys(e)
        parIsolQuantPipe(e)
        parUnit(e, "м")
        runLevel(e)
        sizePipeIsol(e)
        setCategoryCode(e, 120)
        setStock(e, 1.25)

# обобщенные модели
if generic:
    for e in generic:
        setCategoryCode(e, 130)
        setStock(e, 1.0)


t.Commit()


MessageBox.Show("ОК", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information)

