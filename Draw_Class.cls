VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Draw_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public GameEngine As EngineClass

Sub DrawLineV(SVector As Vector_Class, EVector As Vector_Class, Optional Rotation As Matrix_Class, Optional Color = &HFFFFFF)
    Me.GameEngine.GameWindow.Line (SVector.X, SVector.Y)-(EVector.X, EVector.Y), Color
End Sub

Sub DrawLine(SPosX As Double, SPosY As Double, EPosX As Double, EPosY As Double, Optional Rotation As Matrix_Class, Optional Color = &HFFFFFF)
    Me.GameEngine.GameWindow.Line (SPosX, SPosY)-(EPosX, EPosY), Color
End Sub

Sub DrawBox(SPosX As Double, SPosY As Double, EPosX As Double, EPosY As Double, Optional Color = &HFFFFFF)
    Me.GameEngine.GameWindow.Line (SPosX, SPosY)-(EPosX, EPosY), Color, B
End Sub

Sub DrawBoxV(SVector As Vector_Class, EVector As Vector_Class, Optional Color = &HFFFFFF)
    Me.GameEngine.GameWindow.Line (SVector.X, SVector.Y)-(EVector.X, EVector.Y), Color, B
End Sub

Sub DrawSolidBoxV(SVector As Vector_Class, EVector As Vector_Class, Optional Color = &HFFFFFF)
    Me.GameEngine.GameWindow.Line (SVector.X, SVector.Y)-(EVector.X, EVector.Y), Color, BF
End Sub

Sub DrawPoint(XPos As Double, YPos As Double, Optional Color = &HFFFFFF)
    Me.GameEngine.GameWindow.PSet (XPos, YPos), Color
End Sub

Sub DrawPointV(Position As Vector_Class, Optional Color = &HFFFFFF)
    Me.GameEngine.GameWindow.PSet (Position.X, Position.Y), Color
End Sub

Function GetPixelColor(XPos As Single, YPos As Single) As Long
    GetPixelColor = Me.GameEngine.GameWindow.Point(XPos, YPos)
End Function

Sub DrawCircleV(Position As Vector_Class, Radius As Double, Optional Color = &HFFFFFF)
    Me.GameEngine.GameWindow.Circle (Position.X, Position.Y), Radius, Color
End Sub

Sub PrintTextV(Position As Vector_Class, text, Optional TextSize = 8, Optional TextColor = &HFFFFFF)
    Me.GameEngine.GameWindow.FontSize = TextSize
    Me.GameEngine.GameWindow.ForeColor = TextColor
    'Me.GameEngine.GameWindow.PSet (Position.X, Position.Y)
    Me.GameEngine.GameWindow.CurrentX = Position.X
    Me.GameEngine.GameWindow.CurrentY = Position.Y
    Me.GameEngine.GameWindow.Print text
End Sub

Sub PrintText(XPos As Double, YPos As Double, text, Optional TextSize = 8, Optional TextColor = &HFFFFFF)
    Me.GameEngine.GameWindow.FontSize = TextSize
    Me.GameEngine.GameWindow.ForeColor = TextColor
    Me.GameEngine.GameWindow.CurrentX = XPos
    Me.GameEngine.GameWindow.CurrentY = YPos
    Me.GameEngine.GameWindow.Print text
End Sub

Sub DrawCubeCV(SVector As Vector_Class, EVector As Vector_Class, Optional Color = &HFFFFFF, Optional Rotation As Matrix_Class)
    'Me.DrawClass.DrawBoxCV SVector, Me.Position.VClone().AddVector(1000, 1000, 0)
    'Me.DrawClass.DrawLineV SVector, Me.Position.VClone().AddVector(250, 250, 0)
    'Me.DrawClass.DrawLineV Me.Position.VClone().AddVector(1000, 0, 0), Me.Position.VClone().AddVector(1500, 250, 0)
    'Me.DrawClass.DrawBoxCV Me.Position.VClone().AddVector(250, 250, 0), Me.Position.VClone().AddVector(1500, 1500, 0)
    'Me.DrawLineV Me.Position.VClone().AddVector(0, 1000, 0), Me.Position.VClone().AddVector(250, 1500, 0)
End Sub

Sub DrawBoxCV(SVector As Vector_Class, EVector As Vector_Class, Optional Color = &HFFFFFF, Optional Rotation As Matrix_Class)
    Me.GameEngine.GameWindow.Line (SVector.X, SVector.Y)-(EVector.X, SVector.Y), Color
    Me.GameEngine.GameWindow.Line (SVector.X, SVector.Y)-(SVector.X, EVector.Y), Color
    Me.GameEngine.GameWindow.Line (EVector.X, SVector.Y)-(EVector.X, EVector.Y), Color
    Me.GameEngine.GameWindow.Line (SVector.X, EVector.Y)-(EVector.X, EVector.Y), Color
End Sub

Sub DrawVBCETexture(VBCEArrayImage As VBCEArrayImage_Class, Position As Vector_Class, SizeX As Double, SizeY As Double, Optional Rotation As Matrix_Class)
    VBCEArrayImage.RenderTo Me.GameEngine.GameWindow, Position, SizeX, SizeY
End Sub

Sub DrawTextureV(Texture As Texture_Class, Position As Vector_Class, Optional Size As Vector_Class)
    Me.GameEngine.GameWindow.PaintPicture Texture.Texture, Position.X, Position.Y, Size.X, Size.Y
End Sub

Sub DrawPictureV(Pic As StdPicture, Pos As Vector_Class, Optional SizeX As Double, Optional SizeY As Double)
    Me.GameEngine.GameWindow.PaintPicture Pic, Pos.X, Pos.Y, SizeX, SizeY
End Sub
