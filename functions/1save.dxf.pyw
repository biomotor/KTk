#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pythoncom
from win32com.client import Dispatch, gencache

#  Подключим описание интерфейсов API5
kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
kompas_object = kompas6_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))

#  Подключим описание интерфейсов API7
kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
application = kompas_api7_module.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID, pythoncom.IID_IDispatch))

#  Получим активный документ
kompas_document = application.ActiveDocument

#  Если ни один документ не открыт выход
if kompas_document == None: quit()

#  Получим путь и имя файла
Path = (kompas_document.Path)
Name = (kompas_document.Name[:-4])
FPath = Path+Name+".dxf"

#  Получим тип документа
DocType = kompas_object.ksGetDocumentType(0)

#  Получим интерфейс параметров сохранения документа и сохраним
#  1.Чертеж 3.Фрагмент 4.Cпецификация 5.Деталь 6.Сборка 7.Текстовый документ
if DocType in [1, 3, 4]:
    kompas_document.SaveAs(FPath)

#  Если прочее выход
else: quit()
