# заполнение параметров элементов из базы эксель
# -*- coding: utf-8 -*-

import clr
import sys

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

NONELIST = (None, '', ' ')

textParameters = [
  '','',
  'EL_Pos',
  'EL_Rev',
  'EL_Work Type',
  'EL_Flow Diagramm',
  'EL_Maker Type',
  '','','','','','',
  'EL_Mechanical EQP Data Sheet',
  'EL_Motoric Actuation',
  'EL_Equipment number electrical parts',
  'EL_HVAC Control Equipment List',
  'EL_Remarks',
  'EL_Air out off AHU',
  '','','','','','','',
  'EL_AHU-Layout and HX-Calculation',
  'EL_Purpose',
  'EL_Type',
  'EL_Fusible Link Actuation',
  'EL_Activation Temperature',
  'EL_Open Limit Switch',
  'EL_Closed Limit Switch',
  'EL_Bird Screen',
  'EL_Primary Filter',
  'EL_electrical heated',
  'EL_Position',
  'EL_Material Information',
  '', '',
  'EL_Maker Type Motoric Actuation',
  'EL_Medium glycol',
  'EL_Medium temperature range',
  'EL_Nominal Diameter',
  'EL_Connection',
  'EL_Working Area',
  'EL_Voltage Level and Phase',
  '',
  'EL_Type way',
  '']

numberParameters = [
  '','','','','','','',
  'EL_Air Flow',
  'EL_Safety margin',
  'EL_Length',
  'EL_Width',
  'EL_Height',
  'EL_Weight',
  '','','','','','',
  'EL_Elt power for AHU motors',
  'EL_Elt power for Emergency AHU motors',
  'EL_Cooling',
  'EL_Elt power for Cooling',
  'EL_Heating',
  'EL_Humidity',
  'EL_Heating for back heating after humidification',
  '','','','','','','','','','','','',
  'EL_Face Velocity',
  'EL_Max pressure loss',
  '','','','','','','',
  'EL_Elt power for Pump motors',
  '',
  'EL_Kv-value']



def getKeyFromElement(e):
  if e.Category.Id.IntegerValue != int(BuiltInCategory.OST_MechanicalEquipment):
    key = e.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString() + ' | ' + e.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString()
  else:
    key = e.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
  return key

def setParameter(k):
  if textParameters[k] != '':
    p = e.LookupParameter(textParameters[k])
    if p.AsString() != cell:
      p.Set(cell.ToString())
  elif numberParameters[k] != '':
    p = e.LookupParameter(numberParameters[k])
    if p.AsDouble() != cell:
      p.Set(float(cell))


# воздухораспределители
terminal = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType().ToElements()
# арматура
ductAccessory = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctAccessory).WhereElementIsNotElementType().ToElements()
# арматура труб
pipeAccessory = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType().ToElements()
# оборудование
equipment = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().ToElements()

elems = list(equipment) + list(terminal) + list(ductAccessory) + list(pipeAccessory)

# import Spreadsheet data
data = sys.dataFromSpreadsheet
keysFromSpreadsheet = [i[0] for i in data]

t = Transaction(doc, "SetParameters")
t.Start()
for e in elems:
    key = getKeyFromElement(e)
    if key in keysFromSpreadsheet:
        rowIndex = keysFromSpreadsheet.index(key)
        dataRow = data[rowIndex]
        for k, cell in enumerate(dataRow):
            if cell not in NONELIST:
                setParameter(k)
t.Commit()


