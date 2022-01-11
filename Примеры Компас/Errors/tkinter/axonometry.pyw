#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pythoncom
import subprocess
from tkinter import Tk
from tkinter.messagebox import showerror, Dialog
from win32com.client import Dispatch, gencache

#  Проверим запущен ли КОМПАС
def checkKompasRun():
    cmd = 'tasklist /FI "IMAGENAME eq %s" /FI "STATUS eq running"' % 'KOMPAS.Exe'
    status = subprocess.Popen(cmd, stdout=subprocess.PIPE).stdout.read()
    return 'KOMPAS.Exe' in str(status)

if checkKompasRun() == True:

    #  Подключим описание интерфейсов API5
    kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
    kompas_object = kompas6_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))

    #  Подключим описание интерфейсов API7
    kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    application = kompas_api7_module.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID, pythoncom.IID_IDispatch))

    #  Получим текущий активный документ
    kompas_document = application.ActiveDocument
    if kompas_document == None:
        Tk().withdraw()
        showwarning("Ошибка:", "   Не открыт ни один документ!   \n   Чтобы сохранить документ вначале его нужно открыть!")
        quit()

    #  Получим тип документа
    DocType = kompas_object.ksGetDocumentType(0)

    if DocType == (5 or 6):  #  Деталь или Сборка

        #  Получим указатель на интерфейс текущего документа трехмерной модели
        kompas_document_3d = kompas_api7_module.IKompasDocument3D(kompas_document)
        iDocument3D = kompas_object.ActiveDocument3D()

        #  Произведем изменение ориентации
        ksViewProjectionCollection = iDocument3D.GetViewProjectionCollection()
        ProjectionType = 1      # Номер типа проекции: Z-Аксонометрия - 1
        ksViewProjectionCollection.viewProjectionScheme = ProjectionType

        #  Установим ориентацию модели "Изометрия"
        #  ...

    else:
        Tk().withdraw()
        showerror("Ошибка:", "   Тип документа не соответствует!   \n   Пожалуйста, переключитесь на деталь или сборку!   ")
        quit()

else:
    Tk().withdraw()
    showerror("Ошибка:", "   Не запущено ни одного окна КОМПАС !   \n   Пожалуйста, запустите КОМПАС   \n   и откройте деталь или сборку!   ")
    quit()

