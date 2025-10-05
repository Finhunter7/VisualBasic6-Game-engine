VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EngineUtils_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private GEngine As EngineClass

Sub Build(VBCEEngine As EngineClass)
    Set GEngine = VBCEEngine
End Sub

Public Function CList() As Collection
    Dim newCol As New Collection
    Set CList = newCol
End Function
Public Function List(ListOfWhat As VbVarType) As CList_Class
    Dim newList As New CList_Class
    newList.SetListType ListOfWhat
    Set List = newList
End Function

Public Function CreateEvent(ProcName As String)
    Dim NEvent As New EngineEvent_Class
    'NEvent.S_Modules.Add Script.MyModule
    NEvent.ExecuteProc = ProcName
    Set CreateEvent = NEvent
End Function

Public Sub WatchEvent(EventScript As Script_Class, EventVarName As String, UserScript As Script_Class) 'Param 1. "Script What Contains The Event" Param 2. "Event Var Name" Param 3. "Script What Uses The Event"
    On Error Resume Next
    EventScript.MyModule.Eval(EventVarName).S_Modules.Add UserScript.MyModule
End Sub

Public Sub Using(Class As String)

End Sub

Public Sub UsingSystem(Class As String)

End Sub

Public Sub Import(ModuleFile As String)

End Sub

