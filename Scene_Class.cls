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

' Variables
Public Name As String
Attribute Name.VB_VarUserMemId = 0
Attribute Name.VB_VarProcData = ";Text"
Public Index As Long
Public Layer As Long

Public GameEngine As EngineClass
Public Objects As New Collection
Private CopyObjects As New Collection
'Public Scripts As New Collection

Public WorldCenterPos As New Vector_Class
Public WorldBackGround As Variant
Public SunAngle As New Vector_Class
Public CameraPos As New Vector_Class
Public Gravity As New Vector_Class

Public Enum ObjectPresets
    SEmpty = 0
    SPlane = 1
    Cube = 2
    SCircle = 3
End Enum


'Private Intialize()
'Set Intialize() = EndScene

Function EndScene() As Boolean
    RaiseEvent OnRemove
    EndScene = Me.GameEngine.EndScene(Me.Name)
End Function

Sub OnLoad(GameCodeEngine As ScriptControl)
    LoadMeshCode GameCodeEngine
    'For i = 1 To CopyObjects.Count
       ' CopyObjects.Remove i
    'Next
    'For i = 1 To Me.Objects.Count
        'CopyObjects.Add Me.Objects.Item(i), Me.Objects.Item(i).Name
    'Next
End Sub

Sub OnGameStop()
    For i = 1 To Me.Objects.Count
        Me.Objects.Remove i
    Next
    For i = 1 To CopyObjects.Count
        Me.Objects.Add CopyObjects.Item(i), CopyObjects.Item(i).Name
    Next
End Sub

Private Sub LoadMeshCode(BEngine As ScriptControl)
    'Dim GObject As GameObject_Class
    'For i = 1 To Me.Objects.Count()
        'Set GObject = Me.Objects.Item(i)
        'GObject.Load BEngine
    'Next
End Sub

Sub UpdateThis(BEngine As ScriptControl)
    'Dim GObject As GameObject_Class
    'For i = 1 To Me.Objects.Count()
        'Set GObject = Me.Objects.Item(i)
        'GObject.Update BEngine
    'Next
End Sub

Sub ReDrawObjects(BEngine As ScriptControl)
    Dim GObject As GameObject_Class
    For i = 1 To Me.Objects.Count()
        Set GObject = Me.Objects.Item(i)
        GObject.EditorCallDraw BEngine
    Next
End Sub

Sub SetCamera(CObjectPos As Vector_Class)
    Me.CameraPos.SetVectorV CObjectPos
    'ChangeCameraPosition
End Sub

Sub ChangeCameraPosition(AddX As Double, AddY As Double, AddZ As Double)
    For Each Obj In Me.Objects
        Obj.Position.AddVector -AddX, -AddY, -AddZ
    Next
    Me.CameraPos.AddVector AddX, AddY, AddZ
    ReDrawObjects Me.GameEngine.CodeEngine
End Sub

Sub AddObject(TGameEngine As EngineClass, TObject As Object, tKey As String)
    On Error Resume Next
End Sub

Function CreateNewObject(objname As String, Optional SpawnPos As Vector_Class, Optional Parent As GameObject_Class, Optional Preset As ObjectPresets) As GameObject_Class ' Käytä Tätä Vain Uuden Luomisessa
    Dim CnewObject As New GameObject_Class
    Dim CnewObject2 As New GameObject_Class
    Dim NewMesh As String
    
    CnewObject.Name = objname
    CnewObject2.Name = "Hello 12345"
    CnewObject.SetParent CnewObject
    CnewObject.SceneHiearchyPath = Me.Name & "\" & objname
    Set CnewObject.MyScene = Me
    
    If SpawnPos Is Nothing Then
        CnewObject.Position.SetVectorV Me.WorldCenterPos
    Else
        Set CnewObject.Position = SpawnPos.VClone
    End If
    
    On Error Resume Next
    
    Me.Objects.Add CnewObject, CnewObject.Name
    Me.GameEngine.GameObjects.Add CnewObject, Me.Name & "_" & CnewObject.Name
    Dim MeshScript As New Script_Class
    MeshScript.Tag = "RendererScript"
    
    If Preset = SEmpty Then 'Empty
        'CnewObject.ChangeMeshCode CnewObject.MeshCode, Me.GameEngine.CodeEngine
    
    ElseIf Preset = SPlane Then
        MeshScript.Name = Me.Name & "_" & CnewObject.Name & "_Mesh"
        NewMesh = vbNewLine & "Sub Draw() 'Same as Update but executed after Update" & vbNewLine & "Engine.DrawClass.DrawBoxV Me.Position,Me.Position.VClone().AddVector(500,500,0)" & vbNewLine & "End Sub" & vbNewLine & vbNewLine & "Sub EditorDraw()" & vbNewLine & "Draw()" & vbNewLine & "End Sub"
        MeshScript.Data = NewMesh
        Me.GameEngine.Mesh_Scripts.Add MeshScript, MeshScript.Name
        Set MeshScript.MyObject = CnewObject
        CnewObject.Components.Add MeshScript, "Mesh_Code"
    ElseIf Preset = SCircle Then 'Circle
        MeshScript.Name = Me.Name & "_" & CnewObject.Name & "_Mesh"
        NewMesh = vbNewLine & "Sub Draw() 'Same as Update but executed after Update" & vbNewLine & "Engine.GameWindow.Circle 2,Me.Position.X ,Me.Position.Y,100,VbRed,Me.Position.X,Me.Position.Y,1" & vbNewLine & "End Sub" & vbNewLine & vbNewLine & "Sub EditorDraw()" & vbNewLine & "Draw()" & vbNewLine & "End Sub"
        MeshScript.Data = NewMesh
        Me.GameEngine.Mesh_Scripts.Add MeshScript, MeshScript.Name
        Set MeshScript.MyObject = CnewObject
        CnewObject.Components.Add MeshScript, "Mesh_Code"
        
        'CnewObject.ChangeMeshCode NewMesh, Me.GameEngine.CodeEngine
    ElseIf Preset = Cube Then
        MeshScript.Name = Me.Name & "_" & CnewObject.Name & "_Mesh"
        NewMesh = vbNewLine & _
        "Sub Draw() 'Same as Update but executed after Update" & vbNewLine & _
        "Engine.DrawClass.DrawBoxV Me.Position,Me.Position.VClone().AddVector(1000,1000,0)" & vbNewLine & _
        "Engine.DrawClass.DrawLineV Me.Position,Me.Position.VClone().AddVector(250,250,0)" & vbNewLine & _
        "Engine.DrawClass.DrawLineV Me.Position.VClone().AddVector(1000,0,0),Me.Position.VClone().AddVector(1500,250,0)" & vbNewLine & _
        "Engine.DrawClass.DrawBoxV Me.Position.VClone().AddVector(250,250,0),Me.Position.VClone().AddVector(1500,1500,0)" & vbNewLine & _
        "Engine.DrawClass.DrawLineV Me.Position.VClone().AddVector(0,1000,0),Me.Position.VClone().AddVector(250,1500,0)" & vbNewLine & _
        "End Sub" & vbNewLine & vbNewLine & _
        "Sub EditorDraw()" & vbNewLine & _
        "Draw()" & vbNewLine & _
        "End Sub"
        MeshScript.Data = NewMesh
        Me.GameEngine.Mesh_Scripts.Add MeshScript, MeshScript.Name
        Set MeshScript.MyObject = CnewObject
        CnewObject.Components.Add MeshScript, "Mesh_Code"
        
    Else
    
    End If
    
    Me.GameEngine.RefreshChanges
    Set CreateNewObject = CnewObject
End Function

Function CreateObjectInstance(SceneAndObjectName As String, Optional Position As Vector_Class) As GameObject_Class ' Käytä Tätä Pelissä
    Dim ThisObject As New GameObject_Class
    Dim ObjectsCount As Integer
    ObjectsCount = Me.Objects.Count
    
    Set ThisObject = Me.GameEngine.GameObjects.Item(objectName)
    ThisObject.Name = ThisObject.Name & ObjectsCount
    Set ThisObject.MyScene = Me
    
    Me.Objects.Add ThisObject, ThisObject.Name
    Set CreateObjectInstance = ThisObject
End Function


Function CreateObject(ExSceneAndName As String, Optional CallerObject As GameObject_Class, Optional SpawnPos As Vector_Class) As GameObject_Class
    Dim EXObject As GameObject_Class
    
    Set EXObject = Me.GameEngine.GameObjects.Item(ExSceneAndName)
    
    If EXObject Is Nothing Then
        Set CreateObject = Nothing
        Exit Function
    Else
        If CallerObject Is Nothing Then
            Me.Objects.Add EXObject, EXObject.Name
        Else
            CallerObject.MyScene.Objects.Add EXObject, EXObject.Name
        End If
        If SpawnPos Is Nothing Then
            Set EXObject.Position = Me.WorldCenterPos
        Else
            Set EXObject.Position = SpawnPos
        End If
    End If
    
    
End Function

Private Sub Class_Initialize()
    OnSceneLoad
End Sub
Private Sub OnSceneLoad()
    Me.Gravity.SetVector 0, 9.81, 0
    Me.CameraPos.SetVector 0, 0, 10
    Me.SunAngle.SetVector 20, 0, 0
    Me.Layer = 0
End Sub

