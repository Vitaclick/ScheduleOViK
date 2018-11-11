# -*- coding: utf-8 -*-
import clr
import sys

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

NoneList = (None, '', ' ')
projInfo = doc.ProjectInformation
code = projInfo.LookupParameter('CODE').AsString()

parameters = [
  '','',
  'EQ_Description',
  'EQ_Material ID',
  'EQ_HVAC Panel',
  'EQ_Place',
  'EQ_Reference drawing',
  'EQ_ATEX',
  'EQ_IPxx',
  'EQ_Fabricant',
  'EQ_Phases',
  'EQ_Voltage',
  'EQ_Frequency',
  'EQ_User Type',
  'EQ_Absorbed Power',
  'EQ_Installed Power',
  'EQ_Nominal Current',
  'EQ_Syncr Speed',
  'EQ_Efficiency',
  'EQ_cos fi',
  'EQ_Operation Duty',
  'EQ_Factor',
  'EQ_Length',
  'EQ_Width',
  'EQ_Height',
  'EQ_Weight',
  'CL_TAG',
  'CL_Type',
  'CL_Cores',
  'CL_Armour',
  'CL_Outer Diameter',
  'CL_Length',
  '',
  'S_Description',
  'S_Type',
  'S_Range',
  'S_Setpoint',
  'S_Alarm',
  '']

def run():
    def Uniq(e):
        groupcode = e.LookupParameter('GroupTag').AsString()
        partition = e.LookupParameter('partition').AsString()
        designation = e.LookupParameter('Signal_designation').AsString()
        eqtag = e.LookupParameter('EQ_TAG').AsString()
        tag = '-'.join([code, partition, designation, eqtag])
        uniq = groupcode + ' | ' + tag
        return uniq

    def setParameter(k):
        if parameters[k] != '':
            p = e.LookupParameter(parameters[k])
            if p.StorageType == StorageType.String:
                if p.AsString() != cell.ToString():
                    p.Set(cell.ToString())
            elif p.StorageType == StorageType.Double:
                if p.AsDouble() != float(cell):
                    p.Set(float(cell))

    elems = FilteredElementCollector(doc).OfClass(FamilyInstance).OfCategory(BuiltInCategory.OST_DetailComponents).WhereElementIsNotElementType().ToElements()
    elemsWithElType = []
    for e in elems:
        parElementType = e.Symbol.LookupParameter('Element Type')
        parGroupTag = e.LookupParameter('GroupTag')
        if parElementType != None and parGroupTag != None:
            if parElementType.AsString() not in NoneList and parGroupTag.AsString() not in NoneList:
                elemsWithElType.append(e)

    # import Spreadsheet data
    data = sys.dataFromSpreadsheet
    keysFromSpreadsheet = [i[0] for i in data]

    t = Transaction(doc, "SetParameters")
    t.Start()
    for e in elemsWithElType:
        key = Uniq(e)
        if key in keysFromSpreadsheet:
            rowIndex = keysFromSpreadsheet.index(key)
            dataRow = data[rowIndex]
            for k, cell in enumerate(dataRow):
                if cell not in NoneList:
                    setParameter(k)
    t.Commit()

run()

