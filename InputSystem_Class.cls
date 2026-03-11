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
Private WithEvents GameEngine As EngineClass
Attribute GameEngine.VB_VarHelpID = -1
Private IKeysDown As New Collection
Private IKeysUp As New Collection
Private IKeysPressed As New Collection
Private IMouseX As Single
Private IMouseY As Single
Private IMouseButtonDown As Integer

' Events For VBS And JS
Private IEvents As New EventManager_Class
'Public OnMouseDownEvent As New EngineEvent_Class
'Public OnKeyDownEvent As New EngineEvent_Class

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

Public Property Get MouseX() As Single
    MouseX = IMouseX
End Property

Public Property Get MouseY() As Single
    MouseY = IMouseY
End Property

Public Property Get MouseButtonDown() As Integer
    MouseButtonDown = IMouseButtonDown
End Property

Public Property Get KeysDown() As Collection
    Set KeysDown = IKeysDown
End Property

Public Property Get KeysUp() As Collection
    Set KeysUp = IKeysUp
End Property

Public Property Get KeysPressed() As Collection
    Set KeysPressed = IKeysPressed
End Property

Public Property Get Events() As EventManager_Class
    Set Events = IEvents
End Property

Private Property Set Events(RHS As EventManager_Class)
    Set IEvents = RHS
End Property

Sub SetEngine(Engine As EngineClass)
    Set GameEngine = Engine
End Sub

Function IsKeyDown(Key As Integer) As Boolean
    For i = 1 To IKeysDown.Count
        If IKeysDown.Item(i) = Key Then
            IsKeyDown = True
            Exit Function
        End If
    Next
    IsKeyDown = False
End Function

Function GetMousePosition() As Vector_Class
    Dim m_Pos As New Vector_Class
    m_Pos.X = IMouseX
    m_Pos.Y = IMouseY
    Set GetMousePosition = m_Pos
End Function

Function IsMouseButtonDown(Button As Integer) As Boolean
    If Button = IMouseButtonDown Then
        IsMouseButtonDown = True
    Else
        IsMouseButtonDown = False
    End If
End Function

Private Sub Class_Initialize()
    ' Create Events
    IEvents.ChangeClassName "InputSystem"
    IEvents.CreateEvent "OnMouseDown"
    IEvents.CreateEvent "OnKeyDown"
    'OnMouseDownEvent.EventClassName = "InputSystem"
    'OnMouseDownEvent.ExecuteProc = "OnMouseDown"
    'OnKeyDownEvent.EventClassName = "InputSystem"
    'OnKeyDownEvent.ExecuteProc = "OnKeyDown"
End Sub

Private Sub GameEngine_KeyDown(KeyCode As Integer)
    On Error Resume Next
    IKeysDown.Add KeyCode, "Key_" & KeyCode
    'OnKeyDownEvent.Invoke KeyCode
    IEvents.Item("OnKeyDown").Invoke KeyCode
End Sub

Private Sub GameEngine_KeyUp(KeyCode As Integer)
    On Error Resume Next
    IKeysDown.Remove "Key_" & KeyCode
End Sub

Private Sub GameEngine_MouseDown(Button As Integer, X As Single, Y As Single)
    IMouseButtonDown = Button
    IEvents.Item("OnMouseDown").Invoke Button, X, Y
End Sub

Private Sub GameEngine_MouseMove(Button As Integer, X As Single, Y As Single)
    IMouseX = X
    IMouseY = Y
End Sub

Private Sub GameEngine_MouseUp(Button As Integer, X As Single, Y As Single)
    IMouseButtonDown = 0
End Sub

Private Sub GameEngine_OnGameEnd()
    IEvents.Reload
End Sub

Sub ClearKeysDown()
    For i = 1 To IKeysDown.Count
        IKeysDown.Remove i
    Next
End Sub

Private Sub GameEngine_OnGameStop()
    ClearKeysDown
End Sub
