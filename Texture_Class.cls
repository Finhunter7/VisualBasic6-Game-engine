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
Private TextureType As TextureTypes
Public Texture As Object
Public CordinateType As TextureCordinatesTypes
Public Size As Double

Enum TextureTypes
    Image = 0
    ArrayTexture = 1
    Generated = 2
End Enum

Enum TextureCordinatesTypes
    Generated = 0
    UV = 1
    GameObjectSpace = 2
End Enum

Public Property Get GetType() As TextureTypes
    GetType = TextureType
End Property

Public Property Let SetType(ByVal RHS As TextureTypes)
    Set Texture = Nothing
    TextureType = RHS
End Property

Sub LoadTextureImage(FileName As String)
    If TextureType = Image Then
        Set Texture = LoadPicture(FileName)
    ElseIf TextureType = ArrayTexture Then
        Dim ArrayImageClass As New VBCEArrayImage_Class
        Set Texture = ArrayImageClass.CreateFromImage(LoadPicture(FileName))
    End If
End Sub

