VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBCEArrayImage_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Name As String
Private ImageArray() As Long
Private SizeX As Long
Private SizeY As Long
Public ScaleX As Long
Public ScaleY As Long

Sub ResizeImage(NewX As Long, NewY As Long)
    SizeX = NewX
    SizeY = NewY
    ReDim ImageArray(NewX, NewY)
End Sub

Sub SetPixelColor(PosX As Long, PosY As Long, Color As Long)
    ImageArray(PosX, PosY) = Color
End Sub
Function GetPixelColor(PosX As Long, PosY As Long) As Long
    GetPixelColor = ImageArray(PosX, PosY)
End Function

Sub DrawLine(StartPos As Vector_Class, EndPos As Vector_Class, Color As Long)
    
End Sub

Sub CreateFromImage(Picture As IPictureDisp, Optional OSizeX As Long, Optional OSizeY As Long)
    Dim Render As New ImageRenderWindow
    If OSizeX > 0 And OSizeY > 0 Then
        Me.ResizeImage Picture.Width / OSizeX, Picture.Height / OSizeY
    Else
        Me.ResizeImage Picture.Width, Picture.Height
    End If
    
    Render.Width = SizeX
    Render.Height = SizeY
    
    Render.Picture = Picture
    For PLine = 1 To SizeY
        For PCol = 1 To SizeX
            SetPixelColor CLng(PCol), CLng(PLine), CLng(Render.Point(PCol, PLine))
            DoEvents
        Next
    Next
    Render.Show
    RenderTo Render
End Sub

Sub RenderTo(TargetObject As Object)
    TargetObject.Cls
    For PLine = 1 To SizeY
        For PCol = 1 To SizeX
            TargetObject.PSet (PCol, PLine), ImageArray(PCol, PLine)
            DoEvents
        Next
    Next
End Sub

Function GetSizeX() As Long
    GetSizeX = SizeX
End Function
Function GetSizeY() As Long
    GetSizeY = SizeY
End Function
