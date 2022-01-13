#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pythoncom
import textwrap
from win32com.client import Dispatch, gencache

#  Введем необходимый цвет
Color = "#000000"

#  Подключим описание интерфейсов API5
KAPI = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
iKompasObject = KAPI.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(KAPI.KompasObject.CLSID, pythoncom.IID_IDispatch))
iDocument3D = iKompasObject.ActiveDocument3D()

#  Произведем преобразование HEX кода в OLE
def hex_to_rgb(value):
    if value[0] == '#':
        value = value[1:]
    len_value = len(value)
    if len_value not in [3, 6]:
        raise ValueError('Incorect a value hex {}'.format(value))
    if len_value == 3:
        value = ''.join(i * 2 for i in value)
    return tuple(int(i, 16) for i in textwrap.wrap(value, 2))
rgb = hex_to_rgb(Color)
Color = rgb[0] + (rgb[1] * 256) + (rgb[2] * 256 * 256)

#  Производим замену цвета выделенных объектов
SlcMan = iDocument3D.GetSelectionMng()
Count = SlcMan.GetCount()
n = 0
Part0 = SlcMan.GetObjectByIndex (0)
ColorPart0 = Part0.ColorParam()
color0 = ColorPart0.color
for n in range(0,Count,1):
    Part = SlcMan.GetObjectByIndex (n)
    ColorPart = Part.ColorParam()
    color = ColorPart.color
    ColorPart.color = Color
    Part.Update()
SlcMan.UnselectAll ()
