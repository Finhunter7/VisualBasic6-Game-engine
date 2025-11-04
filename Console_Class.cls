VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Console_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private textBuffer As String

Public Console As New Console

Function WriteLine(text)
    On Error Resume Next
    Me.Console.Print text
    textBuffer = textBuffer & text & vbCrLf
End Function

Sub DisplayMessage()
    'Me.Clear
    'WriteLine "VBCE Engine Version " & App.Major & "." & App.Minor & "." & App.Revision
    'WriteLine "VBCE Made By Sami Nissinen 2025-2025"
    'WriteLine "ScriptControll Ready For Execution Input"
End Sub

Function ReadLine() As String
    If Not Len(InputBuffer) = 0 Then
        ReadLine = InputBuffer
    End If
End Function

Function Clear()
    'Text1.text = ""
    'Me.Cls
    'Me.Point 0, 0
    'textBuffer = ""
    'LinesCount = 0
    'If Not StatementHook = "" Then
        'Me.Caption = "Console - " & StatementHook
    'End If
End Function
