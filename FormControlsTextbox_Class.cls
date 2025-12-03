VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormControlsTextbox_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private IName As String
'Public OnChangeEvent As New EngineEvent_Class
Public Events As New EventManager_Class ' Newer Version Of The Event System
Public WithEvents TextBox As TextBox
Attribute TextBox.VB_VarHelpID = -1

Sub Build(Name As String)
    Me.Name = Name
    'OnClickEvent.ExecuteProc = "OnClick"
End Sub

Public Property Get Name() As String
    Name = IName
End Property

Public Property Let Name(RHS As String)
    'OnClickEvent.EventClassName = RHS
    Events.ChangeClassName RHS
    IName = RHS
End Property

Private Sub Class_Initialize()
    Events.ChangeClassName IName
    Events.CreateEvent "OnChange"
End Sub

Private Sub Text_Change()
    Events.Item("OnChange").Invoke
End Sub
