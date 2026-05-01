VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PictureboxForVBS_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private IName As String
Private WithEvents IPictureBox As PictureBox
Attribute IPictureBox.VB_VarHelpID = -1
Private IEvents As New EventManager_Class

Public Property Get Name() As String
    Name = IName
End Property
Friend Property Let Name(RHS As String)
    IEvents.ChangeClassName RHS
    IName = RHS
End Property

Public Property Get PictureBox() As PictureBox
    Set PictureBox = IPictureBox
End Property
Friend Property Set PictureBox(RHS As PictureBox)
    Set IPictureBox = RHS
End Property

Public Property Get Events() As EventManager_Class
    Set Events = IEvents
End Property
Private Property Set Events(RHS As EventManager_Class)
    Set IEvents = RHS
End Property

Private Sub Class_Initialize()
    Events.CreateEvent "Change"
    Events.CreateEvent "Click"
    Events.CreateEvent "DblClick"
    Events.CreateEvent "KeyDown"
    Events.CreateEvent "KeyPress"
    Events.CreateEvent "KeyUp"
    Events.CreateEvent "MouseDown"
    Events.CreateEvent "MouseMove"
    Events.CreateEvent "MouseUp"
    Events.CreateEvent "Paint"
    Events.CreateEvent "Resize"
    Events.CreateEvent "Validate"
End Sub

Private Sub Class_Terminate()
    Events.Clear
End Sub

Private Sub IPictureBox_Change()
    Events.Item("Change").Invoke
End Sub

Private Sub IPictureBox_Click()
    Events.Item("Click").Invoke
End Sub

Private Sub IPictureBox_DblClick()
    Events.Item("DblClick").Invoke
End Sub

Private Sub IPictureBox_KeyDown(KeyCode As Integer, Shift As Integer)
    Events.Item("KeyDown").Invoke KeyCode, Shift
End Sub

Private Sub IPictureBox_KeyPress(KeyAscii As Integer)
    Events.Item("KeyPress").Invoke KeyAscii
End Sub

Private Sub IPictureBox_KeyUp(KeyCode As Integer, Shift As Integer)
    Events.Item("KeyUp").Invoke KeyCode, Shift
End Sub

Private Sub IPictureBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Events.Item("MouseDown").Invoke Button, Shift, X, Y
End Sub

Private Sub IPictureBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Events.Item("MouseMove").Invoke Button, Shift, X, Y
End Sub

Private Sub IPictureBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Events.Item("MouseUp").Invoke Button, Shift, X, Y
End Sub

Private Sub IPictureBox_Paint()
    Events.Item("Paint").Invoke
End Sub

Private Sub IPictureBox_Resize()
    Events.Item("Resize").Invoke
End Sub

Private Sub IPictureBox_Validate(Cancel As Boolean)
    Events.Item("Validate").Invoke
End Sub
