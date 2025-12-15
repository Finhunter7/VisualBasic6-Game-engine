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
Public WithEvents Form As Form
Attribute Form.VB_VarHelpID = -1

'Events For VBS
Public Events As New EventManager_Class
'Public OnLoadEvent As New EngineEvent_Class
'Public OnUnloadEvent As New EngineEvent_Class
'Public OnResizeEvent As New EngineEvent_Class

Sub Build(Name As String)
    Me.Name = Name
End Sub

Public Property Get Name() As String
    Name = IName
End Property

Public Property Let Name(RHS As String)
    Events.ChangeClassName RHS
    IName = RHS
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Events("OnKeyDown").Invoke KeyCode, Shift
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Events("OnKeyPress").Invoke KeyAscii
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Events("OnKeyUp").Invoke KeyCode, Shift
End Sub

Private Sub Form_Load()
   Events("OnLoad").Invoke
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Events("OnMouseDown").Invoke Button, Shift, X, Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Events("OnMouseMove").Invoke Button, Shift, X, Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Events("OnMouseUp").Invoke Button, Shift, X, Y
End Sub

Private Sub Form_Resize()
    Events("OnResize").Invoke
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Events("OnUnload").Invoke Cancel
End Sub

Private Sub Class_Initialize()
    Events.CreateEvent "OnUnload"
    Events.CreateEvent "OnLoad"
    Events.CreateEvent "OnResize"
    Events.CreateEvent "OnMouseDown"
    Events.CreateEvent "OnKeyDown"
    Events.CreateEvent "OnKeyPress"
    Events.CreateEvent "OnKeyUp"
End Sub
