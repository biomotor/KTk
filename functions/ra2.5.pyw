#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pythoncom
from win32com.client import Dispatch, gencache

#  Введем желаемое значение шероховатости
Text = "Ra 2,5"

#  Подключим константы API Компас
kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants

#  Подключим описание интерфейсов API5
kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
kompas_object = kompas6_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))

#  Подключим описание интерфейсов API7
kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
application = kompas_api7_module.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID, pythoncom.IID_IDispatch))

Documents = application.Documents

#  Получим активный документ
kompas_document = application.ActiveDocument
kompas_document_2d = kompas_api7_module.IKompasDocument2D(kompas_document)
iDocument2D = kompas_object.ActiveDocument2D()

#  Произведем запись шероховатости
iDrawingDocument = kompas_document._oleobj_.QueryInterface(kompas_api7_module.IDrawingDocument.CLSID, pythoncom.IID_IDispatch)
iDrawingDocument = kompas_api7_module.IDrawingDocument(iDrawingDocument)
iSpecRough = iDrawingDocument.SpecRough
iSpecRough.SignType = kompas6_constants.ksNoProcessingType
iSpecRough.Text = Text
iSpecRough.Distance = 2
iSpecRough.AddSign = True
iSpecRough.Update()
