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
Public S_Modules As New Collection
'Option Explicit

Function Invoke(ParamArray IArgs() As Variant) As Variant  'Args() As Variant
    Dim retVal As Variant
    On Error Resume Next
    For Each Item In S_Modules
        If Not Item Is Nothing Then
            Invoke = CallByName(Item, EventClassName & "_" & ExecuteProc, VbMethod, IArgs)
        Else
        
        End If
    Next
End Function

Sub Clear()
    For i = 1 To Me.S_Modules.Count
        Me.S_Modules.Remove i
    Next
End Sub

Sub RemoveUser(Index)
    On Error Resume Next
    Me.S_Modules.Remove Index
End Sub

Function AddUser(UObject As Object, Optional key As String) As Long ' Returns Pos Index
    On Error Resume Next
    S_Modules.Add UObject, key
    'AddUser = Me.S_Modules.Count
End Function
