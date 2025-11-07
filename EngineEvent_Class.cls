VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EngineEvent_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Name As String
Public EventClassName As String
Public ExecuteProc As String
Private S_Modules As New Collection
'Option Explicit

Function Invoke(ParamArray IArgs() As Variant) As Variant  'Args() As Variant
    Dim retVal As Variant
    On Error Resume Next
    For i = 1 To S_Modules.Count
        If Not S_Modules.Item(i) Is Nothing Then
            Invoke = CallByName(S_Modules.Item(i), EventClassName & "_" & ExecuteProc, VbMethod, IArgs)
        Else
            RemoveUser i
        End If
    Next
End Function

Sub Clear()
    For i = 1 To S_Modules.Count
        S_Modules.Remove i
    Next
End Sub

Sub RemoveUser(Index)
    On Error Resume Next
    S_Modules.Remove Index
End Sub

Function UsersCount() As Long
    UsersCount = S_Modules.Count
End Function

Function AddUser(UObject, Optional Key As String) As Long ' Returns Pos Index
    On Error Resume Next
    S_Modules.Add UObject, Key
    'AddUser = Me.S_Modules.Count
End Function
