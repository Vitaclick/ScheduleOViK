# -*- coding: utf-8 -*-
import clr

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

userName = uiapp.Application.Username
revData = []
keysFromModel = []
def getCodesFromElements(elems):
    if elems:
        elemsWithCatCode = []
        for e in elems:
            parCatCode = e.LookupParameter('AG_Spc_Код категории')
            if parCatCode != None:
                parCatCodeValue = parCatCode.AsDouble()
                if parCatCodeValue > 0.0:
                    elemsWithCatCode.append(e)
        for e in elemsWithCatCode:
            parCatCodeValue = e.LookupParameter('AG_Spc_Код категории').AsDouble()
            famName = e.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
            spcThi = e.LookupParameter("AG_Spc_Толщина Угол").AsString()
            spcSize = e.LookupParameter("AG_Spc_Размер").AsString()
            row = [famName, '', spcThi, spcSize, '', '', '', '', '', '', parCatCodeValue, userName]
            keysFromModel.append(row)


ducts = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType().ToElements()
flexDuct = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FlexDuctCurves).WhereElementIsNotElementType().ToElements()
fitings = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctFitting).WhereElementIsNotElementType().ToElements()
accessory = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctAccessory).WhereElementIsNotElementType().ToElements()
terminal = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType().ToElements()
isol = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctInsulations).WhereElementIsNotElementType().ToElements()
equipment = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().ToElements()
pipes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements()
pipeFitings = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeFitting).WhereElementIsNotElementType().ToElements()
pipeAccessory = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType().ToElements()
flexPipe = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FlexPipeCurves).WhereElementIsNotElementType().ToElements()
plumbing = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PlumbingFixtures).WhereElementIsNotElementType().ToElements()
sprinklers = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sprinklers).WhereElementIsNotElementType().ToElements()
pipeIsol = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeInsulations).WhereElementIsNotElementType().ToElements()
generic = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()

getCodesFromElements(ducts)
getCodesFromElements(flexDuct)
getCodesFromElements(fitings)
getCodesFromElements(accessory)
getCodesFromElements(terminal)
getCodesFromElements(isol)
getCodesFromElements(equipment)
getCodesFromElements(pipes)
getCodesFromElements(pipeFitings)
getCodesFromElements(pipeAccessory)
getCodesFromElements(flexPipe)
getCodesFromElements(plumbing)
getCodesFromElements(sprinklers)
getCodesFromElements(pipeIsol)
getCodesFromElements(generic)

for i in keysFromModel:
    if i not in revData:
        revData.append(i)
