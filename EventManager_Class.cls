VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EventManager_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private MyClass As String
Private ColEvents As New Collection

'Sub AddUserTo(User, EventID As String)
    'ColEvents.Item(EventID).AddUser User
'End Sub

Sub CreateEvent(Name As String)
    Dim NEvent As New EngineEvent_Class
    NEvent.Name = Name
    NEvent.ExecuteProc = Name
    NEvent.EventClassName = MyClass
    ColEvents.Add NEvent, Name
End Sub

'Sub AddEvent(NEvent As EngineEvent_Class)
    'ColEvents.Add NEvent, NEvent.ExecuteProc
'End Sub

Sub ChangeClassName(NewName As String)
    MyClass = NewName
    For Each EItem In ColEvents
        EItem.EventClassName = NewName
    Next
End Sub

Function Count() As Long
    Count = ColEvents.Count
End Function

Sub Reload()
    For Each EItem In ColEvents
        EItem.Clear
    Next
End Sub

Sub Clear()
    For i = 1 To ColEvents.Count
        ColEvents.Remove i
    Next
End Sub

Sub RemoveEvent(ID)
    ColEvents.Remove ID
End Sub

Function Item(ID) As EngineEvent_Class
Attribute Item.VB_UserMemId = 0
    Set Item = ColEvents.Item(ID)
End Function

