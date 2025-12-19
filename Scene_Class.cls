VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Scene_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' Events
Public Event OnRemove()
Public Event OnObjectRemove(Name As String)

Public GameEngine As EngineClass
' Variables
Public Name As String
Attribute Name.VB_VarUserMemId = 0
Attribute Name.VB_VarProcData = ";Text"
Public Index As Long
Private Layers(1 To 20) As Boolean

Public Objects As New Collection
Attribute Objects.VB_VarDescription = "Collection/List Of GameObjects To Render And Manipulate In Game. Use AddObject Function To Add Objects"
'Private CopyObjects As New Collection
'Public SavedObjects As New Collection
'Public Scripts As New Collection
Public IsPaused As Boolean
Public Active As Boolean

Public WorldCenterPos As New Vector_Class
Public WorldBackGround As Variant
Public SunAngle As New Vector_Class
Public EditorCameraPosition As New Vector_Class
Public Camera As Object
Public Gravity As New Vector_Class


Sub CreateFromScene(SceneObj As Scene_Class)
    With SceneObj
        Me.Name = .Name
        Me.Index = .Index
        'Me.Layer = .Layer
        
        For Each Item In .Objects
            Me.Objects.Add Item, Item.Name
        Next
        Me.WorldCenterPos.SetVectorV .WorldCenterPos
        Me.SunAngle.SetVectorV .SunAngle
        'Me.CameraPos.SetVectorV .CameraPos
        Me.Gravity.SetVectorV .Gravity
        
    End With
End Sub

'Private Intialize()
'Set Intialize() = EndScene

Function EndScene() As Boolean
Attribute EndScene.VB_Description = "Removes The Scene From GameEngine's Scenes Collection And Makes The Scene Inactive"
    RaiseEvent OnRemove
    EndScene = Me.GameEngine.EndScene(Me.Name)
End Function

Sub Load()
    On Error Resume Next
    For Each Item In Me.Objects
        Item.Load
    Next
End Sub

Sub Update()
    On Error Resume Next
    For Each Item In Me.Objects
        If Item.Active Then
            Item.Update
        End If
    Next
End Sub

Sub RedrawObjects()
    On Error Resume Next
    For Each Item In Me.Objects
        CallByName Item, "EditorCallDraw", VbMethod
    Next
End Sub

Function SetCamera(CameraObject As Object) As Object
    Set Camera = CameraObject
    Set SetCamera = CameraObject
End Function

Sub ChangeEditorCameraPosition(AddX As Double, AddY As Double, AddZ As Double)
    Me.EditorCameraPosition.AddVector AddX, AddY, AddZ
End Sub

Function AddObject(Obj As Variant, Optional Name As String = "") As Object
Attribute AddObject.VB_Description = "Adds Object To Scene And Changes The Object Name"
    On Error Resume Next
    Set Obj.MyScene = Me
    If Name = "" Then
        Obj.Name = "Object_" & Me.Objects.Count & "_" & Rnd(1) * 10000 ' Don't Do This. This is a tempoary fix
        Me.Objects.Add Obj, Obj.Name ' Tämä Tulee Muuttumaan
    ElseIf Name = "#ObjectName" Then
        Obj.Name = Obj.Name & "_" & Me.Objects.Count & "_" & Rnd(1) * 10000 ' Don't Do This. This is a tempoary fix
        Me.Objects.Add Obj, Obj.Name
    Else
        Obj.Name = Name
        Me.Objects.Add Obj, Obj.Name
    End If
    If Me.GameEngine.IsEngineRunning Then
        Obj.Load
    End If
    Set AddObject = Obj
End Function

Function GetObject(ID) As Object
    On Error Resume Next
    If Objects.Count > 0 Then
        Set GetObject = Me.Objects.Item(ID)
    End If
End Function

Sub EndObject(ID)
    On Error Resume Next
    Me.Objects.Remove ID
    'RaiseEvent OnObjectRemove(ID)
End Sub

Sub DeleteObject(ID)
Attribute DeleteObject.VB_Description = "Deletes GameObject From Objects Collection."
    On Error Resume Next
    Me.Objects.Remove ID
    'RaiseEvent OnObjectRemove(ID)
End Sub

Private Sub Class_Initialize()
    OnSceneLoad
End Sub
Private Sub OnSceneLoad()
    Me.Gravity.SetVector 0, 9.81, 0
    Me.EditorCameraPosition.SetVector 0, 0, 10
    Me.SunAngle.SetVector 20, 0, 0
    Set Camera = Nothing
    Layers(1) = True
End Sub
