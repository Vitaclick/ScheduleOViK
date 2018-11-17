# -*- coding: utf-8 -*-
import clr
import sys
sys.path.append("C:/Program Files (x86)/IronPython 2.7/Lib")

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from itertools import groupby

def run():
    NoneList = (None, '', ' ')
    projInfo = doc.ProjectInformation
    code = projInfo.LookupParameter('CODE').AsString()

    def UniqAndTag(e):
        groupcode = e.LookupParameter('GroupTag').AsString()
        partition = e.LookupParameter('partition').AsString()
        designation = e.LookupParameter('Signal_designation').AsString()
        eqtag = e.LookupParameter('EQ_TAG').AsString()
        tag = '-'.join([code, partition, designation, eqtag])
        uniq = groupcode + ' | ' + tag
        return [uniq, tag]

    elems = FilteredElementCollector(doc).OfClass(FamilyInstance).OfCategory(BuiltInCategory.OST_DetailComponents).WhereElementIsNotElementType().ToElements()
    elemsWithElType = []
    for e in elems:
        parElementType = e.Symbol.LookupParameter('Element Type')
        parGroupTag = e.LookupParameter('GroupTag')
        if parElementType != None and parGroupTag != None:
            if parElementType.AsString() not in NoneList and parGroupTag.AsString() not in NoneList:
                elemsWithElType.append(e)
    uniqCodeElems = []
    uniqCodes = []
    for e in elemsWithElType:
        unq = e.LookupParameter('partition').AsString() + e.LookupParameter('Signal_designation').AsString() + e.LookupParameter('EQ_TAG').AsString()
        if unq not in uniqCodes:
            uniqCodes.append(unq)
            uniqCodeElems.append(e)
    sortElemsByGroupTag = sorted(uniqCodeElems, key = lambda e: e.LookupParameter('GroupTag').AsString())
    groupElemsByGroupTag = [[x for x in g] for k,g in groupby(sortElemsByGroupTag, lambda e: e.LookupParameter('GroupTag').AsString())]

    groupByElemType = []
    revData = []
    for group in groupElemsByGroupTag:
        equipment = list(filter(lambda e: 'EQ' in e.Symbol.LookupParameter('Element Type').AsString(), group))
        signal = list(filter(lambda e: 'S' in e.Symbol.LookupParameter('Element Type').AsString(), group))
        if equipment:
            for e in equipment:
                uniqandtag = UniqAndTag(e)
                revData.append([uniqandtag[0], None, uniqandtag[1]])
        if signal:
            for e in signal:
                uniqandtag = UniqAndTag(e)
                revData.append([uniqandtag[0]] + [None]*32 + [uniqandtag[1]])
    
    uniqsInModel = list(set([UniqAndTag(e)[0] for e in elemsWithElType]))

    _keysAtSheet = [i for i in keysAtSheet[0]]
    status = []
    for i in _keysAtSheet:
        if i not in uniqsInModel:
            status.append(['Удалено из модели'])
        else:
            status.append([''])
    status[0] = ['Статус'] # head row
    return [revData, status]

data = run()
revData = data[0]
status = data[1]