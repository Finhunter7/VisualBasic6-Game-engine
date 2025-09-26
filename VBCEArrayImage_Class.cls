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
    If SizeX > PosX And SizeY > PosY Then
        ImageArray(PosX, PosY) = Color
    Else
        ImageArray(SizeX, SizeY) = Color
    End If
End Sub
Function GetPixelColor(PosX As Long, PosY As Long) As Long
    GetPixelColor = ImageArray(PosX, PosY)
End Function

Sub DrawLine(StartPos As Vector_Class, EndPos As Vector_Class, Color As Long)
    
End Sub

Sub CreateFromImage(Picture As IPictureDisp, Optional OScaleX As Long, Optional OScaleY As Long)
    Dim Render As New ImageRenderWindow
    If OScaleX > 0 And OScaleY > 0 Then
        Me.ResizeImage Picture.Width / OScaleX, Picture.Height / OScaleY
    Else
        Me.ResizeImage Picture.Width, Picture.Height
    End If
    
    Render.Width = Picture.Width
    Render.Height = Picture.Height
    'Render.Show
    'Render.Visible = False
    Render.Picture = Picture
    For PLine = 1 To SizeY
        For PCol = 1 To SizeX
            SetPixelColor CLng(PCol), CLng(PLine), CLng(Render.Point(PCol * OScaleX, PLine * OScaleY))
            DoEvents
        Next
    Next
    'MsgBox ImageArray(2, 2)
    'RenderTo Render, 20, 20
    Render.Hide
End Sub

Sub RenderTo(TargetObject As Object, newScaleX As Double, newScaleY As Double)
    For PLine = 1 To SizeY
        For PCol = 1 To SizeX
            MainWindow.Engine.GameWindow.PSet (PCol * newScaleX, PLine * newScaleY), ImageArray(PCol, PLine)
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
