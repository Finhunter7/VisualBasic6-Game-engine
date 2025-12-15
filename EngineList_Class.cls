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

Public Function CList(ParamArray items() As Variant) As Collection
    Dim newCol As New Collection
    For Each Item In items
        newCol.Add Item
    Next
    Set CList = newCol
End Function
Public Function List(ListOfWhat As VbVarType, ParamArray items() As Variant) As CList_Class
    Dim newList As New CList_Class
    newList.SetListType ListOfWhat
    For Each Item In items
        newList.Add Item
    Next
    Set List = newList
End Function

Public Function PropertyBagObject() As PropertyBag
    Dim newList As New VBRUN.PropertyBag
    Set PropertyBagObject = newList
End Function

'Public Function CreateEvent(FromCName As String, ProcName As String) As Object
    'Dim NEvent As New EngineEvent_Class
    'NEvent.S_Modules.Add Script.MyModule
    'NEvent.EventClassName = FromCName
    'NEvent.ExecuteProc = ProcName
    'Set CreateEvent = NEvent
'End Function

Public Function CreateEventManager(ClassName As String) As EventManager_Class
    Dim NEventManager As New EventManager_Class
    NEventManager.ChangeClassName ClassName
    Set CreateEventManager = NEventManager
End Function

'Public Sub EventAddUser(TEvent As EngineEvent_Class, TargetObj As Object) 'Param 1. "Script What Contains The Event" Param 2. "Event Var Name" Param 3. "Script What Uses The Event"
    'On Error Resume Next
    'TEvent.AddUser TargetObj
'End Sub

Public Sub Using(Class As String)

End Sub

Public Sub UsingSystem(Class As String)

End Sub

Public Sub Import(ModuleFile As String)

End Sub

