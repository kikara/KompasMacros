import pythoncom
from win32com.client import Dispatch, gencache


# Kompas API 7
class KompasApplication:

    def __init__(self):
        kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
        self.IApplication = kompas_api7_module.IApplication(Dispatch("Kompas.Application.7")._oleobj_
                                               .QueryInterface(kompas_api7_module.IApplication.CLSID,
                                                                     pythoncom.IID_IDispatch))
        self.IKompasDocument = self.IApplication.ActiveDocument
        self.IKompasDocument3D = kompas_api7_module.IKompasDocument3D(self.IKompasDocument)

    # Get iPart main document
    def getTopPart(self, iKompasDocument3D):
        return iKompasDocument3D.TopPart

    # Get Name File from filePath
    def getNameFromFilePath(self, filepath: str):
        name = filepath.replace('\\', '/')
        out = name.split('/').pop()
        return out

    def getModelObject(self, part):
        kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
        part_new = part._oleobj_.QueryInterface(kompas_api7_module.IModelObject.CLSID, pythoncom.IID_IDispatch)
        return kompas_api7_module.IModelObject(part_new)

    def addFile(self, topPart, fileName):
        newPart = topPart.AddFromFile(fileName, True, True)
        return newPart


