# -*- coding: utf-8 -*-
# длина заглушки - параметр Length или Длина воздуховодов у ADSK - по экземпляру!

import clr
import System
import math

import sys
sys.path.append("C:/Program Files (x86)/IronPython 2.7/Lib")

from operator import itemgetter
from itertools import groupby

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

antismoke = ["ВПВ", "ППВ", "ДУ", "КДУ", "ПВ"]
# функция определения настоящего уровня элемента


# генерация имени
def generateName(e):
    if e.Category.Id.IntegerValue == int(BuiltInCategory.OST_DuctFitting) or e.Category.Id.IntegerValue == int(BuiltInCategory.OST_PipeFitting):
        typeDetail = e.MEPModel.PartType
        if typeDetail == PartType.Elbow:
            cSet = [i for i in e.MEPModel.ConnectorManager.Connectors]
            con = cSet[0]
            angle = str(int(round(math.degrees(con.Angle)/5.0))*5) + "°"
            catName = "Отвод "+angle + " / Bend "+angle
        elif typeDetail == PartType.Transition:
            size = e.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString()
            size = size.replace("-", " / ")
            catName = "Переход / Reduction " + size
        elif typeDetail == PartType.Tee:
            size = e.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString()
            size = size.replace("-", " / ")
            catName = "Тройник / Tee " + size
        elif typeDetail == PartType.SpudAdjustable or typeDetail == PartType.TapAdjustable:
            catName = "Врезка / Inset"
        elif typeDetail == PartType.Union:
            catName = "Соединение / Union"
        elif typeDetail == PartType.Cross:
            catName = "Крестовина / Cross"
        elif typeDetail == PartType.Cap:
            catName = "Заглушка / Plug"
        else:
            catName = ""
    elif e.Category.Id.IntegerValue == int(BuiltInCategory.OST_DuctCurves):
        catName = "Воздуховод прямой участок / Duct straight part"
    elif e.Category.Id.IntegerValue == int(BuiltInCategory.OST_FlexDuctCurves):
        catName = "Гибкий воздуховод / Flex duct"
    elif e.Category.Id.IntegerValue == int(BuiltInCategory.OST_FlexPipeCurves):
        catName = "Гибкий трубопровод / Flex pipe"
    else:
        catName = "Трубопровод / Pipe"

    spcName = e.LookupParameter("PL_Part")
    spcName.Set(catName)

def parSys(e):
    system = e.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
    spcSystem = e.LookupParameter("AG_Spc_Система")
    if system != None:
        spcSystem.Set(system)


def parSizeFromText(e, wp, hp, dp):
    size = e.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString()
    if 'x' in size:
        a = size.split('-')[0].split('x')
        wp.Set(a[0])
        hp.Set(a[1])
    elif 'ø' in size:
        a = size.split('-')[0].replace('ø', '')
        dp.Set(a)
def lenSizeFromMultipleParameters(e, lp):
    lParNames = ('Length', 'Длина воздуховода', 'L')
    for p in lParNames:
        try:
            l = e.LookupParameter(p).AsValueString()
            lp.Set(l)
        except: pass
def roundTo10(num):
    digit = int(num)
    ost = digit % 10
    if ost >= 5:
        digit += 10 - ost
    else:
        digit -= ost
    return str(digit)
def parSize(e):
    wp = e.LookupParameter("PL_Width")
    hp = e.LookupParameter("PL_Height")
    dp = e.LookupParameter("PL_Diameter")
    lp = e.LookupParameter("PL_Length")
    dp = e.LookupParameter("PL_Diameter")

    if e.Category.Id.IntegerValue == int(BuiltInCategory.OST_DuctCurves) or e.Category.Id.IntegerValue == int(BuiltInCategory.OST_FlexDuctCurves):
        try:
            w = e.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).AsValueString()
            h = e.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsValueString()
            wp.Set(roundTo10(w))
            hp.Set(roundTo10(h))
        except:
            d = e.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM).AsValueString()
            dp.Set(roundTo10(d))
        l = e.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsValueString()
        lp.Set(roundTo10(l))
    elif e.Category.Id.IntegerValue == int(BuiltInCategory.OST_DuctFitting):
        typeDetail = e.MEPModel.PartType
        if typeDetail == PartType.Elbow:
            parSizeFromText(e, wp, hp, dp)
        elif typeDetail == PartType.SpudAdjustable or typeDetail == PartType.TapAdjustable:
            parSizeFromText(e, wp, hp, dp)
            l = e.get_Parameter(BuiltInParameter.RBS_FAMILY_CONTENT_TAKEOFF_LENGTH).AsDouble()
            l = str(int(UnitUtils.ConvertFromInternalUnits(l, DisplayUnitType.DUT_MILLIMETERS)))
            lp.Set(l)
        elif typeDetail == PartType.Cap:
            parSizeFromText(e, wp, hp, dp)
            lenSizeFromMultipleParameters(e, lp)
        elif typeDetail == PartType.Transition:
            lenSizeFromMultipleParameters(e, lp)




def tByConnector(sysName, con, typeIsol):
    shape = con.Shape
    if shape == ConnectorProfileType.Rectangular:
        h = con.Height
        w = con.Width
        maxS = max(h, w)
        maxMM = UnitUtils.ConvertFromInternalUnits(maxS, DisplayUnitType.DUT_MILLIMETERS)
        if typeIsol != None and "EI" in typeIsol:
            if maxMM <= 2000.0:
                t = 0.9
            else:
                t = 1.2
        else:
            if sysName != None and any(sysName.__contains__(i) for i in antismoke):
                if maxMM <= 2000.0:
                    t = 0.9
                else:
                    t = 1.2
            else:
                if maxMM <= 250.0:
                    t = 0.5
                elif maxMM <= 1000:
                    t = 0.7
                else:
                    t = 0.9
    elif shape == ConnectorProfileType.Round:
        d = con.Radius * 2
        dMM = UnitUtils.ConvertFromInternalUnits(d, DisplayUnitType.DUT_MILLIMETERS)
        if typeIsol != None and "EI" in typeIsol:
            if dMM <= 2000.0:
                t = 0.9
            else:
                t = 1.2
        else:
            if sysName != None and any(sysName.__contains__(i) for i in antismoke):
                if dMM <= 800.0:
                    t = 0.9
                elif dMM <= 1250.0:
                    t = 1.0
                elif dMM <= 1600.0:
                    t = 1.2
                else:
                    t = 1.4
            else:
                if dMM <= 200.0:
                    t = 0.5
                elif dMM <= 450.0:
                    t = 0.6
                elif dMM <= 800.0:
                    t = 0.7
                elif dMM <= 1250.0:
                    t = 1.0
                elif dMM <= 1600.0:
                    t = 1.2
                else:
                    t = 1.4
    return t
# заполнить параметр Толщина Угол для воздуховодов
def setThiDucts(e):
    sysName = e.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
    spcThiAngle = e.LookupParameter("AG_Thickness")
    if e.Category.Id.IntegerValue == int(BuiltInCategory.OST_DuctCurves):
        cc = [i for i in e.ConnectorManager.Connectors]
        con = cc[0]
        typeIsol = e.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_TYPE).AsString()
        thi = tByConnector(sysName, con, typeIsol)
    elif e.Category.Id.IntegerValue == int(BuiltInCategory.OST_FlexDuctCurves):
        thi = 0.15
    spcThiAngle.Set(thi)

# заполнить параметр Толщина Угол для фитингов
def setThiItems(e):
    sysName = e.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
    spcThi = e.LookupParameter("AG_Thickness")
    thi = 0
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
                    thi = own.LookupParameter("AG_Thickness").AsString()
                    break
    else:
        thi = 0
    spcThi.Set(thi)



def setCategoryCode(e, num):
    spcCode = e.LookupParameter("AG_Spc_Код категории")
    spcCode.Set(num)

def squareDuct(e):
    sp = e.LookupParameter("PL_Area S")
    s = e.get_Parameter(BuiltInParameter.RBS_CURVE_SURFACE_AREA).AsDouble()
    s_m = round(UnitUtils.ConvertFromInternalUnits(s, DisplayUnitType.DUT_SQUARE_METERS),1)
    sp.Set(s_m)

def squareFittigs(e):
    sp = e.LookupParameter("PL_Area S")
    geo1 = e.get_Geometry(Options())
    enum1 = geo1.GetEnumerator()
    enum1.MoveNext()
    geo2 = enum1.Current.GetInstanceGeometry()
    solids = [g for g in geo2 if g.GetType() == Solid]
    connectors = e.MEPModel.ConnectorManager.Connectors

    consArea = 0
    for c in connectors:
        if c.Shape == ConnectorProfileType.Rectangular:
            consArea += c.Height*c.Width
        elif c.Shape == ConnectorProfileType.Round:
            consArea += math.pi*(c.Radius**2)

    area = max([i.SurfaceArea for i in solids]) - consArea
    area_m = round(UnitUtils.ConvertFromInternalUnits(area, DisplayUnitType.DUT_SQUARE_METERS),1)
    sp.Set(area_m)

def groupCode(e):
    gp = e.LookupParameter('PL_GroupCode')

    a1 = e.LookupParameter('PL_Part').AsString()
    a2 = e.LookupParameter('PL_Diameter').AsString()
    a3 = e.LookupParameter('PL_Width').AsString()
    a4 = e.LookupParameter('PL_Height').AsString()
    a5 = e.LookupParameter('PL_Length').AsString()
    a6 = e.LookupParameter('AG_Thickness').AsValueString()

    A = filter(lambda x: x != None, [a1, a2, a3, a4, a5, a6])

    code = '-'.join(A)
    gp.Set(code)


# воздуховоды
ducts = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType().ToElements()
# гибкий воздуховод
flexDuct = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FlexDuctCurves).WhereElementIsNotElementType().ToElements()
# соед детали воздуховодов
fitings = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctFitting).WhereElementIsNotElementType().ToElements()



t = Transaction(doc, "SetParameters")
t.Start()

# фитинги
if fitings:
    for e in fitings:
        parSys(e)
        parSize(e)
        setThiItems(e)
        setCategoryCode(e, 40)
        generateName(e)
        squareFittigs(e)
        groupCode(e)

# воздуховоды
if ducts:
    for e in ducts:
        parSys(e)
        parSize(e)
        setThiDucts(e)
        setCategoryCode(e, 50)
        generateName(e)
        squareDuct(e)
        groupCode(e)

# гибкие воздуховоды
if flexDuct:
    for e in flexDuct:
        parSys(e)
        parSize(e)
        setThiDucts(e)
        setCategoryCode(e, 51)
        generateName(e)
        groupCode(e)

t.Commit()

