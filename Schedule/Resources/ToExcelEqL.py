# -*- coding: utf-8 -*-
import clr

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

def keyEquipment(elems, catcode):
  if elems:
    return [[e.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString(), catcode] for e in elems]

def keyOther(elems, catcode):
  if elems:
    return [[e.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString() + ' | ' + e.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString(), catcode] for e in elems]

# воздухораспределители
terminal = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType().ToElements()
# арматура
ductAccessory = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctAccessory).WhereElementIsNotElementType().ToElements()
# арматура труб
pipeAccessory = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType().ToElements()
# оборудование
equipment = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().ToElements()

_revData = keyEquipment(equipment, 10) + keyOther(terminal, 20) + keyOther(ductAccessory, 30) + keyOther(pipeAccessory, 40)

revData = []
for i in _revData:
  if i not in revData:
    revData.append(i)

