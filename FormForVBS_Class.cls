VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormForVBS_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private IName As String
Private WithEvents IForm As Form
Attribute IForm.VB_VarHelpID = -1
'Events For VBS
Private IEvents As New EventManager_Class

Public Property Get Name() As String
    Name = IName
End Property
Friend Property Let Name(RHS As String)
    Events.ChangeClassName RHS
    IName = RHS
End Property

Private Property Set Events(RHS As EventManager_Class)
    Set IEvents = RHS
End Property
Public Property Get Events() As EventManager_Class
    Set Events = IEvents
End Property

Friend Property Set Form(RHS As Form)
    Set IForm = RHS
End Property
Public Property Get Form() As Form
    Set Form = IForm
End Property

Private Sub IForm_KeyDown(KeyCode As Integer, Shift As Integer)
    Events("KeyDown").Invoke KeyCode, Shift
End Sub

Private Sub IForm_KeyPress(KeyAscii As Integer)
    Events("KeyPress").Invoke KeyAscii
End Sub

Private Sub IForm_KeyUp(KeyCode As Integer, Shift As Integer)
    Events("KeyUp").Invoke KeyCode, Shift
End Sub

Private Sub IForm_Load()
   Events("Load").Invoke
End Sub

Private Sub IForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Events("MouseDown").Invoke Button, Shift, X, Y
End Sub

Private Sub IForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Events("MouseMove").Invoke Button, Shift, X, Y
End Sub

Private Sub IForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Events("MouseUp").Invoke Button, Shift, X, Y
End Sub

Private Sub IForm_Resize()
    Events("Resize").Invoke
End Sub

Private Sub IForm_Unload(Cancel As Integer)
    Events("Unload").Invoke Cancel
End Sub

Private Sub Class_Initialize()
    Events.CreateEvent "Unload"
    Events.CreateEvent "Load"
    Events.CreateEvent "Resize"
    Events.CreateEvent "MouseDown"
    Events.CreateEvent "MouseMove"
    Events.CreateEvent "MouseUp"
    Events.CreateEvent "KeyDown"
    Events.CreateEvent "KeyPress"
    Events.CreateEvent "KeyUp"
End Sub
