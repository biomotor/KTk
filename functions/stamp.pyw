#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pythoncom
import LDefin2D
import datetime
from win32com.client import Dispatch, gencache

#  Подключим константы API Компас
kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants

#  Подключим описание интерфейсов API5
kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
kompas_object = kompas6_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))

#  Подключим описание интерфейсов API7
kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
application = kompas_api7_module.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID, pythoncom.IID_IDispatch))

# Получим коллекцию открытых документов в приложении
Documents = application.Documents

#  Получим активный документ
kompas_document = application.ActiveDocument

#  Если ни один документ не открыт выход
if kompas_document == None: quit()

#  Получим тип документа
DocType = kompas_object.ksGetDocumentType(0)

#  Получим интерфейс основной надписи документа
if DocType==1:  #  Чертеж
    kompas_document_2d = kompas_api7_module.IKompasDocument2D(kompas_document)
    iDocument2D = kompas_object.ActiveDocument2D()
    iStamp = iDocument2D.GetStamp()
elif DocType==4:  #  Спецификация
    kompas_document_spc = kompas_api7_module.ISpecificationDocument(kompas_document)
    iDocumentSpc = kompas_object.SpcActiveDocument()
    iStamp = iDocumentSpc.GetStamp()
elif DocType==7:  #  Текстовый документ
    kompas_document_txt = kompas_api7_module.ITextDocument(kompas_document)
    iDocumentTxt = kompas_object.ActiveDocumentTxt()
    iStamp = iDocumentTxt.GetStamp()

#  Если документ не Чертеж, Спецификация или Текстовый документ выход
else: quit()

#  Произведем запись значения графы 'Разработал'
iStamp.ksOpenStamp()
iStamp.ksColumnNumber(110)

iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
iTextLineParam.Init()
iTextLineParam.style = 32768
iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
iTextItemParam.Init()
iTextItemParam.iSNumb = 0
iTextItemParam.s = "Чаленко С.Н."
iTextItemParam.type = 0
iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
iTextItemFont.Init()
iTextItemFont.bitVector = 4096
iTextItemFont.color = 0
iTextItemFont.fontName = "GOST type A"
iTextItemFont.height = 3.5
iTextItemFont.ksu = 1
iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
iTextLineParam.SetTextItemArr(iTextItemArray)

#  Произведем запись значения графы 'Проверил'
iStamp.ksTextLine(iTextLineParam)
iStamp.ksColumnNumber(111)

iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
iTextLineParam.Init()
iTextLineParam.style = 32768
iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
iTextItemParam.Init()
iTextItemParam.iSNumb = 0
iTextItemParam.s = "Шишов А.А."
iTextItemParam.type = 0
iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
iTextItemFont.Init()
iTextItemFont.bitVector = 4096
iTextItemFont.color = 0
iTextItemFont.fontName = "GOST type A"
iTextItemFont.height = 3.5
iTextItemFont.ksu = 1
iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
iTextLineParam.SetTextItemArr(iTextItemArray)

#  Произведем запись значения графы 'Утвердил'
iStamp.ksTextLine(iTextLineParam)
iStamp.ksColumnNumber(115)

iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
iTextLineParam.Init()
iTextLineParam.style = 32768
iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
iTextItemParam.Init()
iTextItemParam.iSNumb = 0
iTextItemParam.s = "Кусуров А.Ю."
iTextItemParam.type = 0
iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
iTextItemFont.Init()
iTextItemFont.bitVector = 4096
iTextItemFont.color = 0
iTextItemFont.fontName = "GOST type A"
iTextItemFont.height = 3.5
iTextItemFont.ksu = 1
iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
iTextLineParam.SetTextItemArr(iTextItemArray)

#  Произведем запись значения 'Подпись' графы 'Разработал'
iStamp.ksTextLine(iTextLineParam)
iStamp.ksColumnNumber(120)

iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
iTextLineParam.Init()
iTextLineParam.style = 32768
iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
iTextItemParam.Init()
iTextItemParam.iSNumb = 0
iTextItemParam.s = "Чаленко."
iTextItemParam.type = 0
iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
iTextItemFont.Init()
iTextItemFont.bitVector = 4096
iTextItemFont.color = 0
iTextItemFont.fontName = "Denistina"
iTextItemFont.height = 3.5
iTextItemFont.ksu = 1
iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
iTextItemParam.Init()
iTextItemParam.iSNumb = 0
iTextItemParam.s = ""
iTextItemParam.type = 0
iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
iTextItemFont.Init()
iTextItemFont.bitVector = 0
iTextItemFont.color = 0
iTextItemFont.fontName = "GOST type A"
iTextItemFont.height = 3.5
iTextItemFont.ksu = 1
iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
iTextLineParam.SetTextItemArr(iTextItemArray)

#  Произведем запись значения 'Дата' графы 'Разработал'
iStamp.ksTextLine(iTextLineParam)
iStamp.ksColumnNumber(130)

iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
iTextLineParam.Init()
iTextLineParam.style = 32768
iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
iTextItemParam.Init()
iTextItemParam.iSNumb = 0
iTextItemParam.s = datetime.date.today().strftime('%d.%m.%y')
iTextItemParam.type = 0
iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
iTextItemFont.Init()
iTextItemFont.bitVector = 4096
iTextItemFont.color = 0
iTextItemFont.fontName = "GOST type A"
iTextItemFont.height = 3.5
iTextItemFont.ksu = 1
iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
iTextLineParam.SetTextItemArr(iTextItemArray)

#  Произведем запись значения 'Дата' графы 'Проверил'
iStamp.ksTextLine(iTextLineParam)
iStamp.ksColumnNumber(131)

iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
iTextLineParam.Init()
iTextLineParam.style = 32768
iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
iTextItemParam.Init()
iTextItemParam.iSNumb = 0
iTextItemParam.s = datetime.date.today().strftime('%d.%m.%y')
iTextItemParam.type = 0
iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
iTextItemFont.Init()
iTextItemFont.bitVector = 4096
iTextItemFont.color = 0
iTextItemFont.fontName = "GOST type A"
iTextItemFont.height = 3.5
iTextItemFont.ksu = 1
iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
iTextLineParam.SetTextItemArr(iTextItemArray)

#  Произведем запись значения 'Дата' графы 'Утвердил'
iStamp.ksTextLine(iTextLineParam)
iStamp.ksColumnNumber(135)

iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
iTextLineParam.Init()
iTextLineParam.style = 32768
iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
iTextItemParam.Init()
iTextItemParam.iSNumb = 0
iTextItemParam.s = datetime.date.today().strftime('%d.%m.%y')
iTextItemParam.type = 0
iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
iTextItemFont.Init()
iTextItemFont.bitVector = 4096
iTextItemFont.color = 0
iTextItemFont.fontName = "GOST type A"
iTextItemFont.height = 3.5
iTextItemFont.ksu = 1
iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
iTextLineParam.SetTextItemArr(iTextItemArray)

#  Произведем запись значения 'Фирма'
iStamp.ksTextLine(iTextLineParam)
iStamp.ksColumnNumber(9)

iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
iTextLineParam.Init()
iTextLineParam.style = 32769
iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
iTextItemParam.Init()
iTextItemParam.iSNumb = 0
iTextItemParam.s = "ООО «Технопром»"
iTextItemParam.type = 0
iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
iTextItemFont.Init()
iTextItemFont.bitVector = 4096
iTextItemFont.color = 0
iTextItemFont.fontName = "Bauhaus-Heavy"
iTextItemFont.height = 4
iTextItemFont.ksu = 1
iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
iTextLineParam.SetTextItemArr(iTextItemArray)

iStamp.ksTextLine(iTextLineParam)
iStamp.ksCloseStamp()
