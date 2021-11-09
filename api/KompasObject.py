import pythoncom
from win32com.client import Dispatch, gencache
import MiscellaneousHelpers as MH


# Kompas API 5
class KompasObject:

    def __init__(self):
        kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
        self.KompasObject = kompas6_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.
                                                         QueryInterface(kompas6_api5_module.KompasObject.CLSID,
                                                                        pythoncom.IID_IDispatch))
        self.iDoc3D = self.KompasObject.ActiveDocument3D()


    def getSizes(self, part) -> list[float]:
        # Получить габариты только тел, без элементов оформления
        parameters = part.GetGabarit(False, True, 0, 0, 0, 0, 0, 0)
        x1 = parameters[1]
        y1 = parameters[2]
        z1 = parameters[3]

        x2 = parameters[4]
        y2 = parameters[5]
        z2 = parameters[6]

        x = float(x2 - x1)
        y = float(y2 - y1)
        z = float(z2 - z1)

        out = list()
        out.append(x)
        out.append(y)
        out.append(z)
        return out

    def getTopPart(self):
        return self.iDoc3D.GetPart(self.iDoc3D.GetObjectType(self.KompasObject))

    # Возвращает словарь плотности детали
    def getDensity(self, part5) -> dict[str, str]:
        out = {"Плотность": "0.00"}
        density = part5.density
        out["Плотность"] = format(density, '.2f')
        return out
