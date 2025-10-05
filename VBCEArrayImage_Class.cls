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
Private RealX As Long
Private RealY As Long

Sub ResizeImage(NewX As Long, NewY As Long)
    SizeX = NewX
    SizeY = NewY
    ReDim ImageArray(NewX, NewY)
End Sub

Sub SetPixelColor(PosX As Long, PosY As Long, Color As Long)
    If SizeX > PosX And SizeY > PosY Then
        ImageArray(PosX, PosY) = Color
    Else
        'ImageArray(SizeX, SizeY) = 1
    End If
End Sub

Function GetPixelColor(PosX As Long, PosY As Long) As Long
    GetPixelColor = ImageArray(PosX, PosY)
End Function

Sub DrawLine(StartPos As Vector_Class, EndPos As Vector_Class, Color As Long)
    
End Sub

Function CreateFromImage(Picture As IPictureDisp) 'OScaleX As Long, OScaleY As Long, ScaleFactor As Integer)
    Dim Render As New ImageRenderWindow
    Dim StatusDia As New StatusDialog
    Dim PixelDenX As Double
    Dim PixelDenY As Double
    'Me.ResizeImage Picture.Width / OScaleX, Picture.Height / OScaleY
    
    Render.Caption = "Please Wait... Rendering Image"
    Render.Width = Picture.Width - 690
    Render.Height = Picture.Height
    RealX = Picture.Width - 690
    RealY = Picture.Height
    Render.Picture = Picture
    
    Me.ResizeImage Render.ScaleWidth, Render.ScaleHeight
    PixelDenX = Picture.Width / Render.ScaleWidth '/ OScaleX
    PixelDenY = Picture.Height / Render.ScaleHeight '/ OScaleY
    
    Render.ScaleMode = 1
    'Render.WindowState = 1
    'Render.ScaleX OScaleX
    'Render.ScaleY OScaleY
    'Render.AutoRedraw = True
    Render.Show
    Render.Caption = "Please Wait... Rendering Texture"
    'Render.Visible = False
    StatusDia.PRGBar_Show "Please Wait... Rendering"
    
    For PLine = 1 To SizeY
        For PCol = 1 To SizeX
            SetPixelColor CLng(PCol), CLng(PLine), CLng(Render.Point(PCol * PixelDenX, PLine * PixelDenY))
            DoEvents
        Next
        StatusDia.PRGBar_SetValue CLng(PLine), CLng(SizeY), "Rendering "
    Next
    Render.Hide
    StatusDia.Hide
    Set CreateFromImage = Me
End Function

Sub RenderTo(TargetObject As Object, Pos As Vector_Class, newScaleX As Double, newScaleY As Double)
    For PLine = 1 To SizeY
        For PCol = 1 To SizeX
            'MainWindow.Engine.GameWindow.PSet (PCol * newScaleX, PLine * newScaleY), ImageArray(PCol, PLine)
            TargetObject.Line (Pos.X + (PCol - 1) * newScaleX, Pos.Y + (PLine - 1) * newScaleY)-(Pos.X + ((PCol - 1) * newScaleX) + newScaleX, Pos.Y + ((PLine - 1) * newScaleY) + newScaleY), ImageArray(PCol, PLine), BF
            'DoEvents
        Next
    Next
End Sub

Function ToPicture(ScaleX As Double, ScaleY As Double) As StdPicture
    Dim ConvertRender As New ImageRenderWindow
    Dim Pos As New Vector_Class
    'Render.ScaleMode = 3
    'ConvertRender.WindowState = 1
    ConvertRender.ScaleMode = 3
    ConvertRender.Width = RealX * ScaleX
    ConvertRender.Height = RealY * ScaleY
    ConvertRender.Show
    ConvertRender.Caption = "Please Wait... Converting Texture To Picture"
    ConvertRender.Cls
    RenderTo ConvertRender, Pos, ScaleX, ScaleY
    Set ToPicture = ConvertRender.Image
    ConvertRender.Hide
End Function

Function GetSizeX() As Long
    GetSizeX = SizeX
End Function
Function GetSizeY() As Long
    GetSizeY = SizeY
End Function

Private Sub Class_Initialize()
    'MyType = ArrayTexture
End Sub
