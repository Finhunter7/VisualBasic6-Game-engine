VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InputSystem_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public WithEvents GameEngine As EngineClass
Attribute GameEngine.VB_VarHelpID = -1
Public KeysDown As New Collection
Public KeysUp As New Collection
Public KeysPressed As New Collection
Public MouseX As Single
Public MouseY As Single
Public MouseButtonDown As Integer

' Events For VBS And JS
Public OnMouseDownEvent As New EngineEvent_Class
Public OnKeyDownEvent As New EngineEvent_Class

Public Enum InputSystemKeyConstants
    IKeyUp = 38
    IKeyDown = 40
    IKeyLeft = 37
    IKeyRight = 39
    IKeyW = 87
    IKeyS = 83
    IKeyA = 65
    IKeyD = 68
    IKeySpace = 32
    
End Enum

Function IsKeyDown(Key As Integer) As Boolean
    For i = 1 To Me.KeysDown.Count
        If Me.KeysDown.Item(i) = Key Then
            IsKeyDown = True
            Exit Function
        End If
    Next
    IsKeyDown = False
End Function

Function GetMousePosition() As Vector_Class
    Dim m_Pos As New Vector_Class
    m_Pos.X = Me.MouseX
    m_Pos.Y = Me.MouseY
    Set GetMousePosition = m_Pos
End Function

Function IsMouseButtonDown(Button As Integer) As Boolean
    If Button = Me.MouseButtonDown Then
        IsMouseButtonDown = True
    Else
        IsMouseButtonDown = False
    End If
End Function

Private Sub Class_Initialize()
    ' Create Events
    OnMouseDownEvent.EventClassName = "InputSystem"
    OnMouseDownEvent.ExecuteProc = "OnMouseDown"
    OnKeyDownEvent.EventClassName = "InputSystem"
    OnKeyDownEvent.ExecuteProc = "OnKeyDown"
End Sub

Private Sub GameEngine_KeyDown(KeyCode As Integer)
    On Error Resume Next
    Me.KeysDown.Add KeyCode, "Key_" & KeyCode
    OnKeyDownEvent.Invoke KeyCode
End Sub

Private Sub GameEngine_KeyUp(KeyCode As Integer)
    On Error Resume Next
    Me.KeysDown.Remove "Key_" & KeyCode
End Sub

Private Sub GameEngine_MouseDown(Button As Integer, X As Single, Y As Single)
    MouseButtonDown = Button
    OnMouseDownEvent.Invoke Button, X, Y
End Sub

Private Sub GameEngine_MouseMove(Button As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
End Sub

Private Sub GameEngine_MouseUp(Button As Integer, X As Single, Y As Single)
    MouseButtonDown = 0
End Sub

Private Sub GameEngine_OnGameEnd()
    Me.OnMouseDownEvent.Clear
End Sub

Private Sub GameEngine_OnGameStop()
    For i = 1 To Me.KeysDown.Count
        Me.KeysDown.Remove i
    Next
End Sub
