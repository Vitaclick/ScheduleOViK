# -*- coding: utf-8 -*-
import clr
import System
import math

import sys
sys.path.append("C:/Program Files (x86)/IronPython 2.7/Lib")

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

rvtLinks = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType().ToElements()
rvtLinkDocs = [i.GetLinkDocument() for i in rvtLinks]
keysFromModels = []

def addCodesFromElemsToList(elems):
  if elems:
    for e in elems:
      famName = e.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
      spcThi = e.LookupParameter("AG_Spc_Толщина Угол").AsString()
      spcSize = e.LookupParameter("AG_Spc_Размер").AsString()
      code = ''
      for i in [famName, spcThi, spcSize]:
        if i != None:
          code += i
      keysFromModels.append(code)

for linkdoc in rvtLinkDocs:
  ducts = FilteredElementCollector(linkdoc).OfCategory(BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType().ToElements()
  flexDuct = FilteredElementCollector(linkdoc).OfCategory(BuiltInCategory.OST_FlexDuctCurves).WhereElementIsNotElementType().ToElements()
  fitings = FilteredElementCollector(linkdoc).OfCategory(BuiltInCategory.OST_DuctFitting).WhereElementIsNotElementType().ToElements()
  accessory = FilteredElementCollector(linkdoc).OfCategory(BuiltInCategory.OST_DuctAccessory).WhereElementIsNotElementType().ToElements()
  terminal = FilteredElementCollector(linkdoc).OfCategory(BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType().ToElements()
  isol = FilteredElementCollector(linkdoc).OfCategory(BuiltInCategory.OST_DuctInsulations).WhereElementIsNotElementType().ToElements()
  equipment = FilteredElementCollector(linkdoc).OfCategory(BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().ToElements()
  pipes = FilteredElementCollector(linkdoc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements()
  pipeFitings = FilteredElementCollector(linkdoc).OfCategory(BuiltInCategory.OST_PipeFitting).WhereElementIsNotElementType().ToElements()
  pipeAccessory = FilteredElementCollector(linkdoc).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType().ToElements()
  flexPipe = FilteredElementCollector(linkdoc).OfCategory(BuiltInCategory.OST_FlexPipeCurves).WhereElementIsNotElementType().ToElements()
  plumbing = FilteredElementCollector(linkdoc).OfCategory(BuiltInCategory.OST_PlumbingFixtures).WhereElementIsNotElementType().ToElements()
  sprinklers = FilteredElementCollector(linkdoc).OfCategory(BuiltInCategory.OST_Sprinklers).WhereElementIsNotElementType().ToElements()
  pipeIsol = FilteredElementCollector(linkdoc).OfCategory(BuiltInCategory.OST_PipeInsulations).WhereElementIsNotElementType().ToElements()
  generic = FilteredElementCollector(linkdoc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()

  addCodesFromElemsToList(ducts)
  addCodesFromElemsToList(flexDuct)
  addCodesFromElemsToList(fitings)
  addCodesFromElemsToList(accessory)
  addCodesFromElemsToList(terminal)
  addCodesFromElemsToList(isol)
  addCodesFromElemsToList(equipment)
  addCodesFromElemsToList(pipes)
  addCodesFromElemsToList(pipeFitings)
  addCodesFromElemsToList(pipeAccessory)
  addCodesFromElemsToList(flexPipe)
  addCodesFromElemsToList(plumbing)
  addCodesFromElemsToList(sprinklers)
  addCodesFromElemsToList(pipeIsol)
  addCodesFromElemsToList(generic)

uniqkeysFromModels = list(set(keysFromModels))

_keysAtSheet = [i for i in keysAtSheet]

status = []
for i in _keysAtSheet:
    if i not in uniqkeysFromModels:
        status.append(['удалите'])
    else:
        status.append([''])
status[0] = ['Статус'] # head row
status[1] = ['']
