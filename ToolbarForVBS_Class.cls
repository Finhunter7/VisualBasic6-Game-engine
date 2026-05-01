VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToolbarForVBS_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private IName As String
Private WithEvents IToolbar As Toolbar
Attribute IToolbar.VB_VarHelpID = -1
Private IEvents As New EventManager_Class

Public Property Get Name() As String
    Name = IName
End Property
Friend Property Let Name(RHS As String)
    IEvents.ChangeClassName RHS
    IName = RHS
End Property

Public Property Get Toolbar() As Toolbar
    Set Toolbar = IToolbar
End Property
Friend Property Set Toolbar(RHS As Toolbar)
    Set IToolbar = RHS
End Property

Public Property Get Events() As EventManager_Class
    Set Events = IEvents
End Property
Private Property Set Events(RHS As EventManager_Class)
    Set IEvents = RHS
End Property

Private Sub Class_Initialize()
    Events.CreateEvent "ButtonClick"
    Events.CreateEvent "Change"
End Sub

Private Sub Class_Terminate()
    Events.Clear
End Sub

Private Sub IToolbar_ButtonClick(ByVal Button As ComctlLib.Button)
    Events.Item("ButtonClick").Invoke Button
End Sub

Private Sub IToolbar_Change()
    Events.Item("Change").Invoke
End Sub
