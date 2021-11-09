from api.KompasApplication import KompasApplication
from api.KompasObject import KompasObject

# iKompasDoc = KompasApplication().IKompasDocument
kompasObject = KompasObject()
iDoc3D = kompasObject.iDoc3D
topPart = kompasObject.getTopPart()
bodyCollection = topPart.BodyCollection()
l = 0
for i in range(0, bodyCollection.GetCount()):
    body = bodyCollection.GetByIndex(i)
    faceCollection = body.FaceCollection()
    for ir in range(0, faceCollection.GetCount()):
        face = faceCollection.GetByIndex(ir)
        if face.IsCylinder():
            param = face.GetCylinderParam()
            l += round(float(param[1]))
print(l)
