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
Private GameEngine As EngineClass

Sub SetEngine(Engine As EngineClass)
    Set GameEngine = Engine
End Sub

Sub DrawLineV(SVector, EVector, Optional Rotation As Matrix_Class, Optional Color As Long = 16777215)
Attribute DrawLineV.VB_Description = "Draws Line To GameWindow Using 2-Axis Vector Object"
    GameEngine.GameWindow.Line (SVector.X, SVector.Y)-(EVector.X, EVector.Y), Color
End Sub

Sub DrawLine(SPosX As Double, SPosY As Double, EPosX As Double, EPosY As Double, Optional Rotation As Matrix_Class, Optional Color As Long = 16777215)
Attribute DrawLine.VB_Description = "Draws Line In To GameWindow"
    GameEngine.GameWindow.Line (SPosX, SPosY)-(EPosX, EPosY), Color
End Sub

Sub DrawBox(SPosX As Double, SPosY As Double, EPosX As Double, EPosY As Double, Optional Color As Long = 16777215)
    GameEngine.GameWindow.Line (SPosX, SPosY)-(EPosX, EPosY), Color, B
End Sub

Sub DrawBoxV(SVector, EVector, Optional Color As Long = 16777215)
    GameEngine.GameWindow.Line (SVector.X, SVector.Y)-(EVector.X, EVector.Y), Color, B
End Sub

Sub DrawSolidBoxV(SVector2 As Variant, EVector2 As Variant, Optional Color As Long = 16777215)
    GameEngine.GameWindow.Line (SVector2.X, SVector2.Y)-(EVector2.X, EVector2.Y), Color, BF
End Sub

Sub DrawPoint(XPos As Double, YPos As Double, Optional Color As Long = 16777215)
    GameEngine.GameWindow.PSet (XPos, YPos), Color
End Sub

Sub DrawPointV(Position As Variant, Optional Color As Long = 16777215)
    GameEngine.GameWindow.PSet (Position.X, Position.Y), Color
End Sub

Function GetPixelColor(XPos As Single, YPos As Single) As Long
    GetPixelColor = GameEngine.GameWindow.Point(XPos, YPos)
End Function

Sub DrawCircleV(Position As Variant, Radius As Double, Optional Color As Long = vbWhite)
    GameEngine.GameWindow.Circle (Position.X, Position.Y), Radius, Color
End Sub

Sub PrintTextV(Position As Variant, Text As String, Optional TextSize As Integer = 8, Optional TextColor As Long = 16777215)
    GameEngine.GameWindow.FontSize = TextSize
    GameEngine.GameWindow.ForeColor = TextColor
    'Me.GameEngine.GameWindow.PSet (Position.X, Position.Y)
    GameEngine.GameWindow.CurrentX = Position.X
    GameEngine.GameWindow.CurrentY = Position.Y
    GameEngine.GameWindow.Print Text
End Sub

Sub PrintText(XPos As Double, YPos As Double, Text As String, Optional TextSize As Integer = 8, Optional TextColor As Long = 16777215)
    GameEngine.GameWindow.FontSize = TextSize
    GameEngine.GameWindow.ForeColor = TextColor
    GameEngine.GameWindow.CurrentX = XPos
    GameEngine.GameWindow.CurrentY = YPos
    GameEngine.GameWindow.Print Text
End Sub

Sub DrawText(XPos As Double, YPos As Double, Text As String, Optional TextSize As Integer = 8, Optional TextColor As Long = 16777215)
    PrintText XPos, YPos, Text, TextSize, TextColor
End Sub

Sub DrawTextV(Position As Variant, Text As String, Optional TextSize As Integer = 8, Optional TextColor As Long = 16777215)
    PrintTextV Position, Text, TextSize, TextColor
End Sub

Sub DrawCubeCV(SVector As Vector_Class, EVector As Vector_Class, Optional Color As Long = vbWhite, Optional Rotation As Matrix_Class)
    'Me.DrawClass.DrawBoxCV SVector, Me.Position.VClone().AddVector(1000, 1000, 0)
    'Me.DrawClass.DrawLineV SVector, Me.Position.VClone().AddVector(250, 250, 0)
    'Me.DrawClass.DrawLineV Me.Position.VClone().AddVector(1000, 0, 0), Me.Position.VClone().AddVector(1500, 250, 0)
    'Me.DrawClass.DrawBoxCV Me.Position.VClone().AddVector(250, 250, 0), Me.Position.VClone().AddVector(1500, 1500, 0)
    'Me.DrawLineV Me.Position.VClone().AddVector(0, 1000, 0), Me.Position.VClone().AddVector(250, 1500, 0)
End Sub

Sub DrawBoxCV(SVector, EVector, Optional Color As Long = vbWhite, Optional Rotation As Matrix_Class)
    GameEngine.GameWindow.Line (SVector.X, SVector.Y)-(EVector.X, SVector.Y), Color
    GameEngine.GameWindow.Line (SVector.X, SVector.Y)-(SVector.X, EVector.Y), Color
    GameEngine.GameWindow.Line (EVector.X, SVector.Y)-(EVector.X, EVector.Y), Color
    GameEngine.GameWindow.Line (SVector.X, EVector.Y)-(EVector.X, EVector.Y), Color
End Sub

Sub DrawVBCETexture(VBCEArrayImage As Object, Position As Variant, SizeX As Double, SizeY As Double, Optional Rotation As Matrix_Class)
    'VBCEArrayImage.RenderTo GameEngine.GameWindow, Position, SizeX, SizeY
End Sub

Sub DrawTextureV(Texture As Object, Position As Variant, Optional Size As Variant)
    GameEngine.GameWindow.PaintPicture Texture.Texture, Position.X, Position.Y, Size.X, Size.Y
End Sub

Sub DrawPictureV(StdPic, PosVector2 As Variant, Optional SizeX As Double, Optional SizeY As Double)
    GameEngine.GameWindow.PaintPicture StdPic, PosVector2.X, PosVector2.Y, SizeX, SizeY
End Sub
