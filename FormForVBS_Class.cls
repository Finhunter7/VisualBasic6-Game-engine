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
Public WithEvents CForm As Form
Attribute CForm.VB_VarHelpID = -1

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

Private Sub CForm_Load()
   Events("OnLoad").Invoke
End Sub

Private Sub CForm_Resize()
    Events("OnResize").Invoke
End Sub

Private Sub CForm_Unload(Cancel As Integer)
    Events("OnUnload").Invoke Cancel
End Sub

Private Sub Class_Initialize()
    Events.CreateEvent "OnUnload"
    Events.CreateEvent "OnLoad"
    Events.CreateEvent "OnResize"
End Sub

