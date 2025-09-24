VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Texture_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Name As String
Public TextureType As TextureTypes
Public Texture As IPictureDisp
Public CordinateType As TextureCordinatesTypes
Public Size As Double

Enum TextureTypes
    Image = 0
    Generated = 1
End Enum

Enum TextureCordinatesTypes
    Generated = 0
    UV = 1
    GObject = 2
End Enum

Sub LoadTextureImage(FileName As String)
    Set Texture = LoadPicture(FileName)
End Sub
