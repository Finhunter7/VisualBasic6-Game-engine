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

Private IGameEngine As EngineClass
' Variables
' Scene
Private IName As String
Attribute IName.VB_VarUserMemId = 0
Attribute IName.VB_VarProcData = ";Text"
Public Index As Long
Private IPriority As Integer
Private Layers(1 To 20) As Boolean


' Objects
Private IObjects As New Collection
Attribute IObjects.VB_VarDescription = "Collection/List Of GameObjects To Render And Manipulate In Game. Use AddObject Function To Add Objects"
Private IObjectsAdded As Long
'Private CopyObjects As New Collection
'Public SavedObjects As New Collection
'Public Scripts As New Collection
Private IIsPaused As Boolean
Private IIsActive As Boolean

Public WorldCenterPos As New Vector_Class
Public WorldBackGround As Variant
Public SunAngle As New Vector_Class
Public EditorCameraPosition As New Vector_Class
Private ICamera As Object
Public Gravity As New Vector_Class

Public Property Get GameEngine() As EngineClass
    Set GameEngine = IGameEngine
End Property
Friend Property Set GameEngine(RHS As EngineClass)
    Set IGameEngine = RHS
End Property

Public Property Get Name() As String
Attribute Name.VB_Description = "Get's The Scene Name. Same Name Is Also Used By The Scenes Collection In Engine"
Attribute Name.VB_UserMemId = 0
    Name = IName
End Property
Friend Property Let Name(RHS As String)
    IName = RHS
End Property

Public Property Get DrawPriority() As Integer
    DrawPriority = IPriority
End Property
Friend Property Let DrawPriority(RHS As Integer)
    IPriority = RHS
End Property

Public Property Get Objects() As Collection
    Set Objects = IObjects
End Property
Private Property Set Objects(RHS As Object)
    Set IObjects = RHS
End Property

Public Property Get AddedObjectsCount() As Long
    AddedObjectsCount = IObjectsAdded
End Property

Public Property Get Camera() As Object
    Set Camera = ICamera
End Property

Public Property Get IsActive() As Boolean
    IsActive = IIsActive
End Property
Friend Property Let IsActive(RHS As Boolean)
    IIsActive = RHS
End Property

'Public Property Get IsPaused() As Boolean
    'IsPaused = IIsPaused
'End Property

'Sub CreateFromScene(SceneObj As Scene_Class)
    'With SceneObj
        'Me.Name = .Name
        'Me.Index = .Index
        'Me.Layer = .Layer
        
        'For Each Item In .Objects
            'Me.Objects.Add Item, Item.Name
        'Next
        'Me.WorldCenterPos.SetVectorV .WorldCenterPos
        'Me.SunAngle.SetVectorV .SunAngle
        'Me.CameraPos.SetVectorV .CameraPos
        'Me.Gravity.SetVectorV .Gravity
        
    'End With
'End Sub

'Private Intialize()
'Set Intialize() = EndScene

Function EndScene() As Boolean
Attribute EndScene.VB_Description = "Removes The Scene From GameEngine's Scenes Collection And Makes The Scene Inactive"
    RaiseEvent OnRemove
    EndScene = GameEngine.EndScene(Me.Name)
End Function

'Sub PauseScene()
    'IIsPaused = True
'End Sub

'Sub ResumeScene()
    'IIsPaused = False
'End Sub

Sub Reset()
    
End Sub

Sub Load()
    On Error Resume Next
    For Each Item In IObjects
        Item.Load
    Next
End Sub

Sub Update()
    On Error Resume Next
    For Each Item In IObjects
        If Item.Active Then
            Item.Update
        End If
    Next
End Sub

Sub DrawObjects() ' Not Curently Used
    On Error Resume Next
    For Each Item In IObjects
        Item.Draw
    Next
End Sub

Sub PauseUpdate() ' Update Is Routed Here If Scene Is Paused
    For Each Item In IObjects
        If Item.Active Then
            Item.PausedUpdate
        End If
    Next
End Sub

Sub RedrawObjects()
    On Error Resume Next
    For Each Item In IObjects
        Item.EditorCallDraw
    Next
End Sub

Function SetCamera(CameraObject As Variant) As Object
    If IsObject(CameraObject) Then
        Set ICamera = CameraObject
        Set SetCamera = CameraObject
    Else
        Set SetCamera = Nothing
    End If
End Function

Sub ChangeEditorCameraPosition(AddX As Single, AddY As Single, AddZ As Single)
    Me.EditorCameraPosition.AddVector AddX, AddY, AddZ
End Sub

Function AddObject(Obj As Variant, Optional Name As String = "") As Object
Attribute AddObject.VB_Description = "Adds Object To Scene And Changes The Object Name"
    On Error Resume Next
    If Not IsObject(Obj) Then
        Err.Raise 1, "AddObject", "Invalid Object"
        Exit Function
    End If
    Set Obj.MyScene = Me
    If Name = "" Then
        Obj.Name = "Object_" & IObjects.Count & "_" & IObjectsAdded 'CInt(Rnd(1) * 10000) ' Don't Do This. This is a tempoary fix. This line will be removed in future
        IObjects.Add Obj, Obj.Name ' Tämä Tulee Muuttumaan
    ElseIf Name = "#ObjectName" Then
        Obj.Name = Obj.Name & "_" & IObjects.Count & "_" & IObjectsAdded 'CInt(Rnd(1) * 10000)  ' Don't Do This. This is a tempoary fix. This line will be removed in future
        IObjects.Add Obj, Obj.Name
    Else
        Dim NObj As Long
        Dim OName As String
        OName = Name
        Do While ObjectExists(OName)
            NObj = NObj + 1
            OName = Name & "_" & NObj
        Loop
        Obj.Name = OName
        IObjects.Add Obj, Obj.Name
    End If
    If GameEngine.IsEngineRunning Then
        Obj.Load
    End If
    IObjectsAdded = IObjectsAdded + 1
    Set AddObject = Obj
End Function

Function ObjectExists(Name As String) As Boolean
    For Each Item In IObjects
        If Item.Name = Name Then
            ObjectExists = True
            Exit Function
        End If
    Next
    ObjectExists = False
End Function

Function GetObject(ID) As Object
    On Error Resume Next
    If IObjects.Count > 0 Then
        Set GetObject = IObjects.Item(ID)
    End If
End Function

Function GetObjectWithVarVal(VarName As String, VarVal As Variant) As Object
    On Error Resume Next
    If IObjects.Count > 0 Then
        For Each Item In IObjects
            If CallByName(Item, VarName, VbGet) = VarVal Then
                Set GetObjectWithVarVal = Item
                Exit Function
            End If
        Next
    End If
    
End Function

Function GetObjectsWithVarVal(VarName As String, VarVal As Variant) As Collection
    On Error Resume Next
    Dim olist As New Collection
    If IObjects.Count > 0 Then
        For Each Item In IObjects
            If CallByName(Item, VarName, VbGet) = VarVal Then
                olist.Add Item, Item.Name
            End If
        Next
        Set GetObjectsWithVarVal = olist
    End If
    
End Function

Sub EndObject(ID)
    On Error Resume Next
    IObjects.Remove ID
    'RaiseEvent OnObjectRemove(ID)
End Sub

Sub DeleteObject(ID)
Attribute DeleteObject.VB_Description = "Deletes GameObject From Objects Collection."
    On Error Resume Next
    IObjects.Remove ID
    'RaiseEvent OnObjectRemove(ID)
End Sub

Sub RemoveAllObjects() ' Removes The Objects Collection From Memory And Redeclares it
Attribute RemoveAllObjects.VB_Description = "Removes The Objects Collection From Memory And Declares New One"
    Set IObjects = Nothing
    Set IObjects = New Collection
End Sub

Private Sub Class_Initialize()
    OnSceneLoad
End Sub
Private Sub OnSceneLoad()
    Me.Gravity.SetVector 0, 9.81, 0
    Me.EditorCameraPosition.SetVector 0, 0, 10
    Me.SunAngle.SetVector 20, 0, 0
    Set ICamera = Nothing
    Layers(1) = True
End Sub
