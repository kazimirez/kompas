import pythoncom
from win32com.client import Dispatch, gencache

import LDefin2D
import MiscellaneousHelpers as MH
from win32comext.mapi.mapiutil import SetPropertyValue

KompasAPI5 = gencache.EnsureModule('{0422828C-F174-495E-AC5D-D31014DBBE87}', 0,
                                   1, 0)
KompasAPI7 = gencache.EnsureModule('{69AC2981-37C0-4379-84FD-5DD2F3C0A520}', 0,
                                   1, 0)
KompasConst = gencache.EnsureModule('{75C9F5D0-B5B8-4526-8681-9903C567D2ED}',
                                    0, 1, 0).constants
KompasConst3D = gencache.EnsureModule('{2CAF168C-7961-4B90-9DA2-701419BEEFE3}',
                                      0, 1, 0).constants
KompasObject = Dispatch('Kompas.Application.5', None,
                        KompasAPI5.KompasObject.CLSID)
iApplication = Dispatch(
    'Kompas.Application.7')  # или KompasObject.ksGetApplication7()

iDocument3D = KompasObject.ActiveDocument3D()

iPart = iDocument3D.GetPart(-1)
iPart.name = '1111'
iPart.marking = '001.001'
iPart.Update()
iPart.svoista_8='бч'
SetPropertyValue()

x = IPropertyKeeper.UniqueMetaObjectKey


print(' ')
# print('iDocument ', iDocument3D)
