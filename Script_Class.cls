VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Script_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Implements IScript_Class
Private IName As String ' Scriptin Nimi
Public Tag As String
Private IData As String ' Scriptin koodi
Public Index As Long ' Ei Käytössä
Public ScriptPriority As Long
Private IMyModule As Module
Public MyObject As Object
Public Active As Boolean

Public Property Get Name() As String
    Name = IName
End Property

Friend Property Let Name(RHS As String)
    IName = RHS
End Property

Public Property Get Data() As String
    Data = IData
End Property

Public Property Let Data(RHS As String)
    'CountKeysStat
    IData = RHS
End Property

Public Property Get MyModule() As Module
    Set MyModule = IMyModule
End Property
Friend Property Set MyModule(RHS As Module)
    Set IMyModule = RHS
End Property

Sub Load()
    On Error Resume Next
    MyModule.Run "Load"
End Sub

Sub Update()
    On Error Resume Next
    MyModule.Run "Update"
End Sub

Private Sub Class_Initialize()
    Me.Active = True
End Sub
