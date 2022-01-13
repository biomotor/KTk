#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pythoncom
from win32com.client import Dispatch, gencache

#  Полный путь к библиотеке
FPath = 'C:\Program Files\ASCON\KOMPAS-3D v19\Libs\PLib.kle'

#  Подключим описание интерфейсов API5
kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
kompas_object = kompas6_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))

#  Подключим описание интерфейсов API7
kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
application = kompas_api7_module.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID, pythoncom.IID_IDispatch))

#  Получим текущий активный документ
kompas_document = application.ActiveDocument

#  Если ни один документ не открыт выход
if kompas_document == None: quit()

#  Получим тип документа
DocType = kompas_object.ksGetDocumentType(0)

#  Если документ не деталь или сборка выход
if DocType not in [5, 6]: quit()

#  Получим указатель на интерфейс текущего документа трехмерной модели
iDocument3D = kompas_object.ActiveDocument3D()

#  Получим указатель на интерфейс менеджера выделенных объектов
SlcMan = iDocument3D.GetSelectionMng()

#  Выход из программы при отсутствии выделенного объекта
if not SlcMan.GetObjectByIndex(0): quit()

#  Получим количество выделенных объектов
Count = SlcMan.GetCount()

#  Получим указатель на интерфейс библиотеки моделей
IModelLibrary = kompas_object.GetModelLibrary()

#  Выберем объект из библиотеки
FileName = IModelLibrary.ChoiceModelFromLib(FPath, 0)[0]

#  Меняем все выделенные библиотечные объекты
Part = {}
for n in range(0, Count):
    Part[n] = SlcMan.GetObjectByIndex(n)
    Part[n].fileName = FileName
for n in range(0, Count):
    Part[n].Update()
