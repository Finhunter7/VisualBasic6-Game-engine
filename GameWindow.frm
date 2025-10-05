VERSION 5.00
Begin VB.Form GameWindow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Form3"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5580
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form3"
   ScaleHeight     =   4065
   ScaleWidth      =   5580
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "GameWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Engine As EngineClass
'Public MouseX As Integer
'Public MouseY As Integer
'Public MouseDown As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'On Error Resume Next
    'Me.Engine.InputSystem.KeysDown.Add KeyCode, "Key" & KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'On Error Resume Next
    'Me.Engine.InputSystem.KeysDown.Remove ("Key" & KeyCode)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Me.MouseX = X
    'Me.MouseY = Y
    'Me.MouseDown = Button
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Engine.EndGame
End Sub
