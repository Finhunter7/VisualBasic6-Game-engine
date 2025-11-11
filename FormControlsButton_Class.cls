VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormControlsButton_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private IName As String
Public Events As New EventManager_Class
Public WithEvents Button As CommandButton
Attribute Button.VB_VarHelpID = -1

Sub Build(Name As String)
    Me.Name = Name
    'OnClickEvent.ExecuteProc = "OnClick"
End Sub

Public Property Get Name() As String
    Name = IName
End Property

Public Property Let Name(RHS As String)
    Events.ChangeClassName RHS
    IName = RHS
End Property

Private Sub Button_Click()
    Events.Item("OnClick").Invoke
End Sub

Private Sub Class_Initialize()
    Events.MyClass = Me.Name
    Events.CreateEvent "OnClick"
End Sub
