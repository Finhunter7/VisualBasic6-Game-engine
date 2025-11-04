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
Public Data As String ' Scriptin koodi
Public Index As Long ' Ei Käytössä
Public ScriptPriority As Long
Public MyModule As Module
Public MyObject As Object
Public GameObject As Object
'Public MyObjects As New Collection ' Scriptin Objektit Jota Käytetään Me.MyObjects.Item(Index/Key) Komennolla Tai Suoraan Tällä Muuttujalla
'Public MyPropertyBag As New PropertyBag ' Scriptin Arvoja Jota Muut Scriptit Voivat Muuttaa/Käyttää
Public Active As Boolean

Public Function GetVariable(VName As String) As Variant
Attribute GetVariable.VB_Description = "Gets Variable In Script's Module"
Attribute GetVariable.VB_UserMemId = 0
    On Error GoTo end1
    If IsObject(Me.MyModule.Eval(VName)) Then
        Set GetVariable = Me.MyModule.Eval(VName)
    Else
        GetVariable = Me.MyModule.Eval(VName)
    End If
    'Set GetVariable = Me.MyModule.Eval(VName)
    Exit Function
end1:
    
End Function

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
