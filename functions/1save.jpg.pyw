#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pythoncom
from win32com.client import Dispatch, gencache

#  Выберем расширение сохраняемого файла
Extension = ".jpg"

#  Выберем разрешение сохраняемого файла
Resolution = 600

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

# Получим коллекцию открытых документов в приложении
Documents = application.Documents

#  Получим путь и имя файла
Path = (kompas_document.Path)
Name = (kompas_document.Name[:-4])
FPath = Path+Name+Extension

#  Получим тип документа
DocType = kompas_object.ksGetDocumentType(0)

#  Определим параметры сохранения. Получим интерфейс параметров сохранения документа и сохраним
def Save(kdoc):
    rasterPar = kdoc.RasterFormatParam()
    rasterPar.format = 2
    rasterPar.greyScale = 1 
    rasterPar.extResolution = Resolution  # Разрешение
    kdoc.SaveAsToRasterFormat(FPath,rasterPar)

#  Передадим в функцию Save указатель на интерфейс текущего документа
if   DocType == 1:       #  Чертеж
    Save(kompas_object.ActiveDocument2D())
elif DocType == 4:       #  Спецификация
    Save(kompas_object.SpcActiveDocument())
elif DocType == 7:       #  Текстовый документ
    Save(kompas_object.ActiveDocumentTxt())
elif DocType in [5, 6]:  #  Деталь или Сборка
    Resolution = 0
    Save(kompas_object.ActiveDocument3D())

#  Если прочее выход
else: quit()
