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
Public Name As String ' Scriptin Nimi
Public Tag As String
Private IData As String ' Scriptin koodi
Public Index As Long ' Ei Käytössä
Public ScriptPriority As Long
Public MyModule As Module
Public MyObject As Object
Public Active As Boolean

Public Property Get Data() As String
    Data = IData
End Property

Public Property Let Data(RHS As String)
    'CountKeysStat
    IData = RHS
End Property

Private Sub CountKeysStat()
    Dim Vars As Variant
    Dim pos As Long
    Vars = Array("Dim", "Private", "Public")
    For Each Var In Vars
        Do Until pos > Len(IData)
            pos = InStr(pos + Len(Var), IData, Var, vbTextCompare)
        Loop
        
    Next
End Sub

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

Sub CreateModule(CodeEngine As ScriptControl)
    Set MyModule = CodeEngine.Modules.Add("Script_" & Me.Name & "_" & Rnd(1) * 10000, MyObject)
    MyModule.AddCode Data
End Sub
