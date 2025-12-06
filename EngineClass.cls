VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EngineClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

' This Is The Main Engine Class

'Events
Public Event OnProjectOpen()
Public Event OnDataChanged()
Public Event OnSceneCreated()
Public Event OnSceneActivated()
Public Event OnGameStart()
Public Event OnRenderStart()
Public Event OnRenderEnd()
Public Event OnGameStop()
Public Event OnGameEnd()
Public Event OnGamePause()
Public Event OnGameResume()
Public Event OnError(Error As Error)
Public Event OnTimeout()
Public Event MouseDown(Button As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer)
Public Event KeyUp(KeyCode As Integer)
'Editor Events
Public Event EditorMouseDown(Button As Integer, X As Single, Y As Single)
'Public Event EditorMouseUp(Button As Integer, X As Single, Y As Single)
Public Event EditorMouseMove(Button As Integer, X As Single, Y As Single)
Public Event EditorKeyDown(KeyCode As Integer)
'Public Event EditorKeyUp(KeyCode As Integer)


' Project
Private m_ProjectFile As String
Public ReloadProjectOnStartup As Boolean
Public ProjectName As String
Public ProjectPath As String
Private IsProjectOpened As Boolean
Public ProjectCodeLanguage As String
Public ProjectType As ProjectTypes

Public GameWindowSizeX As Long
Public GameWindowSizeY As Long
Public GameWindowScaleX As Long
Public GameWindowScaleY As Long

' Editor Utils
Public IsInEditor As Boolean

Public BaseScene As Integer
Public EngineRefrences As New Collection
Public EngineComponents As New Collection

Public WithEvents CodeEngine As MSScriptControl.ScriptControl ' Referenssi ScriptControl luokkaan jolla pystyt‰‰n suorittamaan koodia.
Attribute CodeEngine.VB_VarHelpID = -1
Public WithEvents EngineTimer As EngineTimer_Class
Attribute EngineTimer.VB_VarHelpID = -1

Public GlobalScript As New Script_Class
Public GlobalObjects As New Collection

Public WithEvents GameWindow As Form
Attribute GameWindow.VB_VarHelpID = -1
Private newGameWindow As Form
Public WithEvents EditorWindow As Form
Attribute EditorWindow.VB_VarHelpID = -1
Private EditorSelectedObject As Object
'Public EngineTimer As Timer

'Public ProjectBrowser As Form
'Public SceneBrowser As Form
Public SettingsWindow As Form
Public EConsole As Form

Public InputSystem As New InputSystem_Class
Public DrawClass As New Draw_Class
Public FormControls As New FormControls

Private EVectorClass As New Vector_Class
Private MathClass As New Math_Class
Private AudioSystem As New AudioSystem_Class
Private fileSystemHandler As Object 'FileSystemObject
Public Direct11Lib As Object

Private EngineUtil_Class As New EngineUtils_Class

'Public SaveLoadGameClass As New GameLoader_Class

Public Scenes As New Collection ' Aktiiviset Scenet
Public ScenesInactive As New Collection
'Private SceneRestorePoint As New Collection

Public Forms As New Collection
Public GameObjects As New Collection
Public Scripts As New Collection
Public ClassComponents As New Collection
Public Mesh_Scripts As New Collection
Public Textures As New Collection
Public VBCETextures As New Collection
Public Components As New Collection

Public IsEngineRunning As Boolean
Public RunGameInEditor As Boolean

Private TargetFramerate As Integer
Private FrameTime As Double
Private SimulationTime As Double 'In Milliseconds
Private RenderStart As Double
Private CurrentFrame As Long

' Constants And Enums

Enum ProjectTypes
    VBCEGameProject = 0
    VBCEScriptProject = 1
    VBCEVideoProject = 2
    VBCEDX8GameProject = 3
    VBCEDX11GameProject = 4
End Enum

Enum SceneConstants
    Default = 0
    Overlay = 1
    Underlay = 2
End Enum

Enum ScriptTypes
    DefaultScript = 0
    ComponentScript = 1
    MeshScript = 2
End Enum

Enum CodeEditorEnums
    ObjectCode = 1
    MeshData = 2
    ScriptCode = 3
End Enum

Enum SavingEnums
    Script = 0
    Scene = 1
    GameObject = 2
    Project = 3
    Animation = 4
End Enum

Enum ComponentTypes
    ScriptComponent = 0
    SpriteRendererComponent = 1
    PhysicsComponent = 2
    CustomComponent = 3
End Enum

Enum TextureClases
    ArrayTexture = 0
    PictureTexture = 1
    PictureOrImage = 2
End Enum

Function GetCurrentScene() As Scene_Class ' Palauttaa ensimm‰isen Scenen
    If Me.Scenes.Count < 1 Then
        Set GetCurrentScene = Nothing
    Else
        Set GetCurrentScene = Me.Scenes.Item(BaseScene)
    End If
End Function

Function GetReference(NameOrID) As Object
    If Me.EngineRefrences.Count < 1 Then
        Set GetReference = Nothing
    Else
        Set GetReference = Me.EngineRefrences.Item(NameOrID).Refrence
    End If
End Function

Function EditorGetSelectedObject(MPosX As Single, MPosY As Single) As Object
    On Error Resume Next
    Dim curScene As Scene_Class
    Dim SceneObject As Object
    If Me.IsEngineRunning Then
        Set EditorSelectedObject = Nothing
        Set EditorGetSelectedObject = Nothing
        Exit Function
    End If
    For i = 1 To Me.Scenes.Count
        Set curScene = Me.Scenes(i)
        For O = 1 To curScene.Objects.Count
            Set SceneObject = curScene.Objects(O)
            If MPosX > (SceneObject.Position.X - 125) And MPosX < (SceneObject.Position.X + 125) Then
                If MPosY > (SceneObject.Position.Y - 125) And MPosY < (SceneObject.Position.Y + 125) Then
                    Me.EditorRefreshScenes
                    DrawSelectedLines SceneObject
                    DrawSelectedInfo SceneObject
                    Set EditorSelectedObject = SceneObject
                    Set EditorGetSelectedObject = SceneObject
                    Exit Function
                End If
            End If
        Next
    Next
    If Me.Scenes.Count > 0 And Not EditorSelectedObject Is Nothing Then
        Me.EditorRefreshScenes
    End If
    Set EditorSelectedObject = Nothing
    Set EditorGetSelectedObject = Nothing
End Function

Function GetSelectedObject() As Object
    Set GetSelectedObject = EditorSelectedObject
End Function

Sub DrawSelectedInfo(SceneObject As Object)
    'Me.EditorRefreshScenes
    Me.EditorWindow.FontSize = 8
    Me.EditorWindow.ForeColor = &HFFFFFF
    Me.GameWindow.CurrentX = 0
    Me.GameWindow.CurrentY = 0
    Me.EditorWindow.Print "Object Name: " & SceneObject.Name
    Me.EditorWindow.Print "Object X: " & SceneObject.Position.X
    Me.EditorWindow.Print "Object Y: " & SceneObject.Position.Y
End Sub

Sub DrawSelectedLines(SceneObject As Object)
    On Error Resume Next
    Me.EditorWindow.FontSize = 8
    Me.EditorWindow.ForeColor = &HFFFFFF
    Me.EditorWindow.Line (SceneObject.Position.X - 125, SceneObject.Position.Y - 125)-(SceneObject.Position.X + 125, SceneObject.Position.Y + 125), vbGreen, B
    Me.EditorWindow.Line (SceneObject.Position.X, SceneObject.Position.Y)-(SceneObject.Position.X + 250, SceneObject.Position.Y), vbRed
    Me.EditorWindow.Print "X"
    Me.EditorWindow.Line (SceneObject.Position.X, SceneObject.Position.Y)-(SceneObject.Position.X, SceneObject.Position.Y + 250), vbGreen
    Me.EditorWindow.Print "Y"
End Sub

Function SelectObject(GameObject As Object)
    If Me.IsEngineRunning Then
        Exit Function
    End If
    Me.EditorRefreshScenes
    DrawSelectedLines GameObject
    DrawSelectedInfo GameObject
    'Me.EditorWindow.Line (GameObject.Position.X - (GameObject.EditorColliderSize.X / 2), GameObject.Position.Y - (GameObject.EditorColliderSize.Y / 2))-(GameObject.Position.X + (GameObject.EditorColliderSize.X / 2), GameObject.Position.Y + (GameObject.EditorColliderSize.Y / 2)), vbGreen, B
    'Me.EditorWindow.Line (SceneObject.Position.X, SceneObject.Position.Y)-(SceneObject.Position.X + (SceneObject.EditorColliderSize.X), SceneObject.Position.Y), vbRed
    'Me.EditorWindow.Line (SceneObject.Position.X, SceneObject.Position.Y)-(SceneObject.Position.X, SceneObject.Position.Y + SceneObject.EditorColliderSize.Y), vbGreen
    Set EditorSelectedObject = GameObject
End Function

Function EditorRefreshScenes()
    Dim curScene As Scene_Class
    Me.EditorWindow.Cls
    For i = 1 To Me.Scenes.Count
        Set curScene = Me.Scenes(i)
        curScene.ReDrawObjects Me.CodeEngine
    Next
End Function
Function RefreshChanges()
    'Me.ProjectBrowser.Update
    'Me.SceneBrowser.Update
    Me.EditorRefreshScenes
    RaiseEvent OnDataChanged
End Function

Function EditorMoveCamera(DirX As Double, DirY As Double)
    If Me.IsEngineRunning Then
        Exit Function
    End If
    'Me.GetCurrentScene.CameraPos.AddVector DirX, DirY
    Me.GetCurrentScene.ChangeCameraPosition DirX, DirY, 0
    'Me.EditorRefreshScenes
End Function

Function CreateNewScene(thisName As String) As Scene_Class ' Luo uuden Scenen ScenesInactive kokoelmaan Scene luokan pohjalta
Attribute CreateNewScene.VB_Description = "Creates New Inactive Scene"
    Dim newScene As New Scene_Class
    
    If thisName = "" Then
        Exit Function
    End If
    
    newScene.Name = thisName
    Set newScene.GameEngine = Me
    
    newScene.WorldCenterPos.X = Me.GameWindow.Width / 2
    newScene.WorldCenterPos.Y = Me.GameWindow.Height / 2
    
    ScenesInactive.Add newScene, thisName
    Set CreateNewScene = newScene
    RaiseEvent OnDataChanged
    RaiseEvent OnSceneCreated
End Function

Function CreateCustomScene(SceneObject As Object, Optional Name As String) As Object
    On Error Resume Next
    Set SceneObject.MyEngine = Me
    If Name = "" Then
        Me.ScenesInactive.Add SceneObject, SceneObject.Name
    Else
        SceneObject.Name = Name
        Me.ScenesInactive.Add SceneObject, SceneObject.Name
    End If
    Set CreateCustomScene = SceneObject
End Function

Sub AddGlobalObject(Name As String, TObject As Object, Optional AddMembers As Boolean = False)
    Dim newObject As New CodeEngineObject_Class
    newObject.Build Name, TObject
    Me.GlobalObjects.Add newObject, Name
    'Me.CodeEngine.AddObject Name, TObject, AddMembers
End Sub

Function RenameScene(SceneName As String, NewName As String)
    Dim thisScene As Scene_Class
    Set thisScene = ScenesInactive.Item(SceneName)
    If Me.Scenes.Count < 1 Then
    
    Else
        Me.EndScene SceneName
    End If
    If thisScene Is Nothing Then
        Me.EConsole.WriteLine "Failed To Rename Scene " & SceneName & " To " & NewName & "Scene Does Not Exist"
        Exit Function
    Else
        thisScene.Name = NewName
        ScenesInactive.Add thisScene, NewName
        ScenesInactive.Remove SceneName
    End If
    RaiseEvent OnDataChanged
End Function

Function AddScene(Name, Optional OverOrUnder As SceneConstants = Default) As Object ' Lis‰‰ uuden Scenen ScenesInactive kokoelmasta Scenes kokoelmaan
Attribute AddScene.VB_Description = "Makes The Inactive Scene Active"
    On Error Resume Next
    Dim newScene As Scene_Class
    If Me.ScenesInactive.Count < 1 Then
        'Me.EConsole.WriteLine "AddScene " & Str(Name) & " Failed"
        Set AddScene = Nothing
        Exit Function
    End If
    Set newScene = ScenesInactive.Item(Name)
    If newScene Is Nothing Then
        Set AddScene = Nothing
        Exit Function
    End If
    
    If OverOrUnder = Underlay And Me.ScenesInactive.Count > 0 Then
        Scenes.Add newScene, newScene.Name, 1
        'newScene.OnLoad Me.CodeEngine
        'Set AddScene = newScene
    ElseIf OverOrUnder = Overlay Then
        Scenes.Add newScene, newScene.Name, , Scenes.Count
        'newScene.OnLoad Me.CodeEngine
        'Set AddScene = newScene
    Else
        Scenes.Add newScene, newScene.Name
        'newScene.OnLoad Me.CodeEngine
        'Set AddScene = newScene
    End If
    Set AddScene = newScene
    
    RaiseEvent OnDataChanged
    RaiseEvent OnSceneActivated
End Function

Function EndScene(Name) As Boolean
    Dim Scene As Scene_Class
    If Me.Scenes.Count < 1 Then
        EndScene = False
        Exit Function
    End If
    Set Scene = Me.Scenes.Item(Name)
    If Scene Is Nothing Then
        EndScene = False
    Else
        Me.Scenes.Remove Name
        EndScene = True
    End If
    RaiseEvent OnDataChanged
End Function

Function CopyScene(Name As String) As Scene_Class
    Dim newSceneCopy As New Scene_Class
    newSceneCopy.CreateFromScene Me.ScenesInactive.Item(Name)
    Set CopyScene = newSceneCopy
End Function

Sub UnloadScene(SceneName As String)
    EndScene SceneName
    Me.ScenesInactive.Remove SceneName
    RaiseEvent OnDataChanged
End Sub

Function RemoveScript(TScriptName As String)
    On Error Resume Next
    Dim ThisItem As Script_Class
    If Me.Scripts.Count > 0 Then
        Set ThisItem = Me.Scripts.Item(TScriptName)
        If Not ThisItem Is Nothing Then
            Me.Scripts.Remove TScriptName
            Exit Function
        End If
    End If
    If Me.Mesh_Scripts.Count > 0 Then
        Set ThisItem = Me.Mesh_Scripts.Item(TScriptName)
        If Not ThisItem Is Nothing Then
            Me.Mesh_Scripts.Remove TScriptName
            Exit Function
        End If
    End If
    RaiseEvent OnDataChanged
End Function

Function UnloadGameObject(GameObjectName As String)
    Dim thisGameObject As GameObject_Class
    Dim ThisGameObjModule As Module
    Set thisGameObject = Me.GameObjects.Item(GameObjectName)
    'Set ThisGameObjModule = Me.CodeEngine.Modules(GameObjectName)
    'Unload ThisGameObjModule
    'thisGameObject.EndObject
    Me.GameObjects.Remove GameObjectName
    'Unload ThisGameObject
    RaiseEvent OnDataChanged
End Function

Function CreateNewScript(Name As String, Optional ExposedObject As Object = Nothing, Optional Priority As Long = 0) As Script_Class
    Dim newClass As New Script_Class
    
    newClass.Name = Name
    Set newClass.MyObject = ExposedObject
    newClass.Tag = "Script"
    If Priority = 0 Or Scripts.Count = 0 Then
        Scripts.Add newClass, Name
    Else
        Scripts.Add newClass, Name, Priority
    End If
    
    Set CreateNewScript = newClass
    RaiseEvent OnDataChanged
End Function

Function CreateNewClassComponent(Obj As Object, Name As String) As Object
    On Error Resume Next
    If Not IsEmpty(Obj.Name) Then
        Obj.Name = Name
    End If
    ClassComponents.Add Obj, Name
    Set CreateNewClassComponent = Obj
End Function

Sub EditScript(Name As String)
    Dim editableScripts As New Collection
    Dim CodeEditorInstance As New CodeEditor
    Dim thisObj As Script_Class
    
    editableScripts.Add Me.GlobalScript, Me.GlobalScript.Name
    For Each Item In Me.Scripts
        editableScripts.Add Item, Item.Name
    Next
    For Each Item In Me.Mesh_Scripts
        editableScripts.Add Item, Item.Name
    Next
    
    Set thisObj = editableScripts.Item(Name)
    
    Set CodeEditorInstance.GameEngine = Me
    CodeEditorInstance.EditData thisObj, Name, ScriptCode, Me
    
End Sub

Function CreateGameObject(Name As String, SceneID) As GameObject_Class
    Dim TargetScene As Scene_Class
    Dim NGObject As New GameObject_Class
    
    Set TargetScene = Me.ScenesInactive.Item(SceneID)
    
    NGObject.Name = Name
    Set NGObject.MyScene = TargetScene
    TargetScene.Objects.Add NGObject, NGObject.Name
    Set CreateGameObject = NGObject
End Function

Function CreateCustomGameObject(GObject As Variant, Name As String, SceneID) As Object
    Dim TargetScene As Scene_Class
    Set TargetScene = Me.ScenesInactive.Item(SceneID)
    TargetScene.AddObject GObject, Name
    Set CreateCustomGameObject = GObject
End Function

Function SetGlobalScriptData(Data As String) As Boolean
    On Error Resume Next
    Me.GlobalScript.Data = Data
    Me.CodeEngine.AddCode Me.GlobalScript.Data
End Function

Function AddGlobalScriptData(Data As String) As Boolean
    On Error Resume Next
    Me.GlobalScript.Data = Me.GlobalScript.Data & Data & vbCrLf
    Me.CodeEngine.AddCode Me.GlobalScript.Data
End Function

Sub EditObjectCode(Edit As CodeEditorEnums, Object As GameObject_Class, Optional Scene As Scene_Class)
    
End Sub

Private Sub Class_Initialize()
    ScriptEngine_Intialize
End Sub
Private Sub ScriptEngine_Intialize()
    On Error Resume Next
    'Intialize ScriptEngine
    If Me.CodeEngine Is Nothing Then
        Set CodeEngine = New MSScriptControl.ScriptControl
    Else
        Me.CodeEngine.Reset
    End If
    If Me.CodeEngine Is Nothing Then
        MsgBox "This Program Requires MSScriptControl.ScriptControl ActiveX/COM Object. Unable To Load VBCE", vbExclamation
        MsgBox "MSScriptControl Failed To Load. Unable To Load VBCE", vbExclamation
        'Unload Me
        End
    End If
    Set fileSystemHandler = New Scripting.FileSystemObject
    If fileSystemHandler Is Nothing Then
        MsgBox "Scripting.FileSystemObject ActiveX Component Failed To Load", vbExclamation
    End If
    Set Direct11Lib = New DirectX11_Class
    Direct11Lib.Initialize Me
    If Direct11Lib Is Nothing Then
        MsgBox "DirectX11 Component Failed To Load. DirectX11 Features Will Not Be Usable.", vbExclamation
    End If
    
    AudioSystem.Build Me
    Me.SetCodeLaunguage "VBScript"
    Me.CodeEngine.AddObject "Engine", Me, True
    Me.CodeEngine.AddObject "Vector", EVectorClass, True
    Me.CodeEngine.AddObject "Math", MathClass, True
    Me.CodeEngine.AddObject "AudioSystem", AudioSystem, True
    
    Me.GlobalScript.Name = "Global"
    Me.ProjectName = "Project1"
    Me.ProjectType = VBCEGameProject
    Me.GameWindowSizeX = 640
    Me.GameWindowSizeY = 480
    TargetFramerate = 60
    FrameTime = 16
    
    Dim GlobalObject As CodeEngineObject_Class
    
    Set GlobalObject = New CodeEngineObject_Class
    GlobalObject.Build "Engine", Me
    Me.GlobalObjects.Add GlobalObject, GlobalObject.Name
    
    Set GlobalObject = New CodeEngineObject_Class
    GlobalObject.Build "Vector", EVectorClass
    Me.GlobalObjects.Add GlobalObject, GlobalObject.Name
    
    Set GlobalObject = New CodeEngineObject_Class
    GlobalObject.Build "AudioSystem", AudioSystem
    Me.GlobalObjects.Add GlobalObject, GlobalObject.Name
    
    Set GlobalObject = New CodeEngineObject_Class
    GlobalObject.Build "EngineUtils", EngineUtil_Class
    EngineUtil_Class.Build Me
    Me.GlobalObjects.Add GlobalObject, GlobalObject.Name
    
    Set Me.EConsole = VBCE.Console.CreateNewConsoleInstance(Me)
    Me.CodeEngine.AddObject "Console", Me.EConsole, True
    
    Dim GOConsole As New CodeEngineObject_Class
    GOConsole.Build "Console", Me.EConsole
    Me.GlobalObjects.Add GOConsole, GOConsole.Name
    
    'Set GlobalObject = New CodeEngineObject_Class
    'GlobalObject.Build "VBA", VBA
    'Me.GlobalObjects.Add GlobalObject, GlobalObject.Name
    
    'Dim newRef As New EngineRefrence_Class
    'newRef.Name = "FileSystemObject"
    'newRef.Path = ""
    'newRef.RefrenceClass = "Scripting.FileSystemObject"
    'Set newRef.Refrence = fileSystemHandler
    'Me.EngineRefrences.Add newRef, newRef.Name
    'Me.CodeEngine.AddObject "NullClass", NullObject, True
    
    Me.InputSystem.SetEngine Me
    Me.DrawClass.SetEngine Me
    Set EngineTimer = New EngineTimer_Class
    Set newGameWindow = New GameWindow
    Set Me.GameWindow = newGameWindow
    
    OnEngineLoad
    
End Sub

Private Function OnEngineLoad()
    If Len(Command) > 0 Then
        OnProjectLoad Command
    Else
        Me.BaseScene = 1
    End If
    IsEngineRunning = False
    SimulationTime = 0
End Function

Private Function OnProjectLoad(Com As String)
    Dim Commands() As String
    
    Commands = Split(Com, " ")
    
    If UBound(Commands) > 0 Then
        If Commands(0) = "-P" Then
            If UBound(Commands) > 1 Then
                OpenProject Commands(1), True
            End If
        ElseIf Commands(0) = "-C" Then
            If UBound(Commands) > 1 Then
                LoadStartupCode Commands(1)
            End If
        Else
            OpenProject Commands(0), False
        End If
    End If
End Function
Private Sub LoadStartupCode(FilePath As String)
    Me.CodeEngine.AddCode Me.GetTextFromFile(FilePath)
End Sub

Function LoadEngineInEditorMode(TGameWindow As Form, Optional TEditorWindow As Form)
    Me.IsInEditor = True
    Set Me.GameWindow = TGameWindow
    Set Me.EditorWindow = TEditorWindow
    
End Function

Function EndGame()
    If Me.IsInEditor Then
        OnGameStop
    Else
        'Unload Me.GameWindow
        'Unload Me.EditorWindow
        End
    End If
    RaiseEvent OnGameEnd
End Function

Function PauseGame()
    Me.EngineTimer.Kill
    Me.GameWindow.Caption = Me.ProjectName & " [Paused]"
    RaiseEvent OnGamePause
End Function

Public Property Get IsProjectOnDisk() As Boolean
    IsProjectOnDisk = IsProjectOpened
End Property

Public Property Get ProjectFilePath() As String
    ProjectFilePath = m_ProjectFile
End Property

Function ResumeGame()
    RaiseEvent OnGameResume
    RenderStart = Timer
    Me.GameWindow.Caption = Me.ProjectName
    Me.EngineTimer.Start
End Function

Function GetCodeLanguage() As String
    GetCodeLanguage = ProjectCodeLanguage
End Function

Function SetCodeLaunguage(Language As String)
    Me.CodeEngine.Language = Language
    ProjectCodeLanguage = Language
    RaiseEvent OnDataChanged
End Function

Function SetProjectType(PType As ProjectTypes)
    Me.ProjectType = PType
End Function

Function SetTargetFramerate(Framerate As Integer)
    TargetFramerate = Framerate
    FrameTime = (1000 / Framerate)
    Me.EngineTimer.Interval = FrameTime
    RaiseEvent OnDataChanged
End Function

Function GetTargetFramerate() As Integer
    GetTargetFramerate = TargetFramerate
End Function

Function GetTargetFrameTime() As Double
    GetTargetFrameTime = FrameTime
End Function

Function GetRenderedFrames() As Long
    GetRenderedFrames = CurrentFrame
End Function

Function GetFrameTime() As Double
    GetFrameTime = Me.EngineTimer.GetRealIntervalTime()
End Function

Function GetFrameRate() As Double
    GetFrameRate = (((TargetFramerate / (Timer - RenderStart)) * SimulationTime) / 1000)
    'GetFrameRate = (1000 / Me.GetRealFrameTime())
End Function

Function GetSimulationTime() As Double
    GetSimulationTime = SimulationTime
End Function

Function GetRealTime() As Double
    GetRealTime = Timer - RenderStart
End Function

Function GetAppPath() As String
    GetAppPath = App.Path
End Function

Function GetApp() As App
    Set GetApp = App
End Function

Function GetGamePath() As String
    GetGamePath = Me.ProjectPath
End Function

Function GetWindowCenter() As Vector_Class
    Dim newVec As New Vector_Class
    newVec.SetVector Me.GameWindow.Width / 2, Me.GameWindow.Height / 2
    Set GetWindowCenter = newVec
End Function

Function GetWindowSize() As Vector_Class
    Dim newVec As New Vector_Class
    newVec.SetVector Me.GameWindow.Width, Me.GameWindow.Height
    Set GetWindowSize = newVec
End Function

Function PrintEngineInfo()
    Me.EConsole.WriteLine "ScriptControl Information"
    Me.EConsole.WriteLine "Engine Information"
    Me.EConsole.WriteLine "Engine Version " & App.Major & "." & App.Minor & " " & App.Revision
    Me.EConsole.WriteLine "Active Scenes Count " & Me.Scenes.Count
    Me.EConsole.WriteLine "Inactive Scenes Count " & Me.ScenesInactive.Count
    Me.EConsole.WriteLine "Scripts Count " & Me.Scripts.Count
End Function

Private Function OnGameStop()
    If Me.RunGameInEditor = False Then
        Unload Me.GameWindow
        Set Me.GameWindow = Me.EditorWindow
    End If
    If Me.EngineTimer.IsActive() Then
        Me.EngineTimer.Kill
    End If
    IsEngineRunning = False
    Me.GameWindow.Caption = "Engine [Desing]"
    Me.EConsole.WriteLine "Engine Stopped"
    Set Me.GameWindow = Me.EditorWindow
    Me.EditorRefreshScenes
    'Me.CodeEngine.Reset
    'For Each Item In Me.Scenes
        'Item.OnGameStop
    'Next
    If IsProjectOpened Then
        ClearGameEngineData
        Me.OpenProject m_ProjectFile, False
    Else
        ReloadCodeEngine
    End If
    
    RaiseEvent OnGameStop
End Function

Function ClearGameEngineData()
    On Error Resume Next
    Me.ProjectName = "Untiled"
    'ReloadCodeEngine
    For i = 1 To Me.Scenes.Count
        Me.Scenes.Remove i
    Next
    Set Me.Scenes = Nothing
    Set Me.Scenes = New Collection
    For i = 1 To Me.ScenesInactive.Count
        Me.ScenesInactive.Remove i
    Next
    Set Me.ScenesInactive = Nothing
    Set Me.ScenesInactive = New Collection
    'For i = 1 To Me.GameObjects.Count
        'Me.GameObjects.Remove i
    'Next
    For i = 1 To ClassComponents.Count
        ClassComponents.Remove i
    Next
    Set ClassComponents = Nothing
    Set ClassComponents = New Collection
    For i = 1 To Scripts.Count
        Scripts.Remove i
    Next
    Set Scripts = Nothing
    Set Scripts = New Collection
    
    For i = 1 To Components.Count
        Components.Remove i
    Next
    ReloadCodeEngine
    RefreshChanges
End Function

Private Sub ReloadCodeEngine()
    Me.CodeEngine.Reset
    Me.CodeEngine.Language = GetCodeLanguage()
    'Me.CodeEngine.AddObject "Engine", Me, True
    'Me.CodeEngine.AddObject "Vector", EVectorClass, True
    'Me.CodeEngine.AddObject "Math", MathClass, True
    'Me.CodeEngine.AddObject "AudioSystem", AudioSystem, True
    'Me.CodeEngine.AddObject "Console", Me.EConsole, True
    For Each Obj In Me.GlobalObjects
        Me.CodeEngine.AddObject Obj.Name, Obj.MyObject, True
    Next
    For Each Ref In Me.EngineRefrences
        If Not Ref.HoldData Then
            Set Ref.Refrence = Nothing
            Set Ref.Refrence = CreateObject(Ref.RefrenceClass, Ref.Server)
        End If
    Next
End Sub

Function CreateModuleFromFile(FileName As String, Optional ExposedObject As Object = Nothing) As Module
Attribute CreateModuleFromFile.VB_Description = "Creates ScriptControl Module From File"
    Dim FSystem As New FileSystemObject
    Dim SFile As file
    Dim dataStream As TextStream
    Dim CModule As Module
    
    On Error Resume Next
    If FSystem.FileExists(FileName) Then
        Set SFile = FSystem.GetFile(FileName)
        Set dataStream = SFile.OpenAsTextStream(ForReading, TristateUseDefault)
        Set CModule = Me.CodeEngine.Modules.Add(SFile.Name & "_FileResource_" & Me.CodeEngine.Modules.Count, ExposedObject)
        CModule.AddCode dataStream.ReadAll()
        dataStream.Close
        Set CreateModuleFromFile = CModule
    End If
End Function

Function OpenProject(FileName As String, localPath As Boolean) As Boolean
    Dim FileSystem As New FileSystemObject
    Dim ScriptFile As file
    Dim ScriptFileDataStream As TextStream
    Dim ProjectFile As file
    Dim ProjectFileDataStream As TextStream
    Dim Data As String
    Dim SModule As Module
    
    RaiseEvent OnProjectOpen
    
    On Error Resume Next
    If FileSystem.FileExists(FileName) Then
        Set ProjectFile = FileSystem.GetFile(FileName)
        Me.ProjectPath = ProjectFile.ParentFolder.Path
        m_ProjectFile = FileName
        IsProjectOpened = True
        Set ProjectFileDataStream = ProjectFile.OpenAsTextStream(ForReading, TristateUseDefault)
        Do Until ProjectFileDataStream.AtEndOfStream
            Data = ProjectFileDataStream.ReadLine()
            If localPath Then
                If FileSystem.FileExists(Me.GetGamePath() & Data) Then
                    Set ScriptFile = FileSystem.GetFile(Me.GetGamePath() & Data)
                    Set ScriptFileDataStream = ScriptFile.OpenAsTextStream(ForReading, TristateUseDefault)
                    Set SModule = Me.CodeEngine.Modules.Add(ScriptFile.Name & "_" & Me.CodeEngine.Modules.Count)
                    SModule.AddCode ScriptFileDataStream.ReadAll()
                    SModule.Run "Initialize"
                    ScriptFileDataStream.Close
                End If
            Else
                If FileSystem.FileExists(ProjectFile.ParentFolder.Path & Data) Then
                    Set ScriptFile = FileSystem.GetFile(ProjectFile.ParentFolder.Path & Data)
                    Set ScriptFileDataStream = ScriptFile.OpenAsTextStream(ForReading, TristateUseDefault)
                    Set SModule = Me.CodeEngine.Modules.Add(ScriptFile.Name & "_" & Me.CodeEngine.Modules.Count)
                    SModule.AddCode ScriptFileDataStream.ReadAll()
                    SModule.Run "Initialize"
                    ScriptFileDataStream.Close
                End If
            End If
        Loop
        ProjectFileDataStream.Close
        OpenProject = True
    Else
        OpenProject = False
    End If
    Exit Function
End Function

Function GetTextFromFile(FileName As String) As String
    Dim FileObj As New FileSystemObject
    Dim Data As TextStream
    If FileObj.FileExists(FileName) Then
        Set Data = FileObj.OpenTextFile(FileName, ForReading, False)
        GetTextFromFile = Data.ReadAll()
        Data.Close
    Else
        
    End If
End Function

Function GetFolder(Path As String) As Folder
    Dim FileObj As New FileSystemObject
    Dim CFolder As Folder
    'Dim Files As New Collection
    
    If FileObj.FolderExists(Path) Then
        Set CFolder = FileObj.GetFolder(Path)
        
        Set GetFilesInFolder = CFolder
        'For Each Item In CFolder.Files
            'Files.Add Item, Item.Name
        'Next
    Else
        Set GetFilesInFolder = Nothing
    End If
    'Set GetFilesInFolder = Files
End Function

Private Function OnGameLoad()
    IsEngineRunning = True
    OpenGameWindow
    Me.GameWindow.Caption = Me.ProjectName
    If Me.IsInEditor Then
        Me.EConsole.WriteLine "Engine Started"
        Me.EditorWindow.Caption = "Engine Running"
        If Me.GetCurrentScene() Is Nothing Then
            Me.EditorWindow.StatusBar1.Panels(1).text = "Scene: Nothing"
        Else
            Me.EditorWindow.StatusBar1.Panels(1).text = "Scene: " & Me.GetCurrentScene().Name
        End If
        If Me.GetTargetFramerate > 0 Then
            Me.EditorWindow.StatusBar1.Panels(3).text = "Target Framerate: " & Me.GetTargetFramerate
        Else
            Me.EditorWindow.StatusBar1.Panels(3).text = "Target Framerate: Uncapped"
        End If
        
        MainWindow.StatusBar1.Panels(1).text = ""
    End If
    
    CurrentFrame = 0
    SimulationTime = 0
    RenderStart = Timer
    RaiseEvent OnGameStart
End Function

Private Function OpenGameWindow()
    If RunGameInEditor = False And Me.IsInEditor Then
        Set Me.GameWindow = newGameWindow
        Set Me.GameWindow.Engine = Me
        Me.GameWindow.Show
        Me.GameWindow.Refresh
    ElseIf Me.IsInEditor = False Then
        Set Me.GameWindow.Engine = Me
        Me.GameWindow.Show
        Me.GameWindow.Refresh
    End If
End Function

Private Function CreateStatusDialog(Caption As String) As Object
    Dim DialogInstance As New StatusDialog
    DialogInstance.PRGBar_Show Caption
    Set CreateStatusDialog = DialogInstance
End Function

Sub CheckScriptCode(Script As Script_Class, CodeEditor As Object)
    Dim procs As Procedure
    Dim thisModule As Module
    Me.EConsole.Show
    ReloadCodeEngine
    
    On Error GoTo Error1
    Set thisModule = Me.CodeEngine.Modules.Add("Script_" & thisScript.Name, thisScript.MyObject)
    Set Script.MyModule = thisModule
    Script.MyModule.AddCode Script.Data
    
    For i = 1 To Script.MyModule.Procedures.Count
        Set procs = Script.MyModule.Procedures.Item(i)
        If procs.NumArgs < 1 Then
            Script.MyModule.Run procs.Name
        End If
    Next
    Exit Sub
Error1:
    
End Sub

Private Sub CreateScriptModules(StatusDialog As Form)
    Dim thisScript As Script_Class
    Dim thisModule As Module
    Dim curScene As Scene_Class
    Dim baseObject As IObject_Class
    Dim curGObject As GameObject_Class
    'OnGameLoad
    ReloadCodeEngine ' Resetoi Koodimoottorin
    'CreateModules
    
    On Error Resume Next
    For i = 1 To Me.Scripts.Count ' K‰sittelee Kaikki Skriptit, Luo Koodinsuoritus Moottorille Moduulin Ja Tekee Skriptille Referensin T‰h‰n Moduuliin Josta Sit‰ On Helppo Ohjata
        Set thisScript = Me.Scripts.Item(i)
        If thisScript.MyObject Is Nothing Then ' T‰m‰ On Myˆs Hyv‰ Kun T‰t‰ Operoidaan Windows 9x K‰yttˆj‰rjestelmiss‰
            Set thisModule = Me.CodeEngine.Modules.Add("Script_" & thisScript.Name & "_" & CLng(Rnd(1) * 10000), thisScript)
        Else
            Set thisModule = Me.CodeEngine.Modules.Add("Script_" & thisScript.Name & "_" & CLng(Rnd(1) * 10000), thisScript.MyObject)
        End If
        Set thisScript.MyModule = thisModule
        If Me.IsInEditor Then
            StatusDialog.PRGBar_SetValue CLng(i), CLng(Me.Scripts.Count), "CodeEngine Modules Adding Code "
        End If
        thisModule.AddCode thisScript.Data
    Next
    
    For i = 1 To Me.Mesh_Scripts.Count
        Set thisScript = Me.Mesh_Scripts.Item(i)
        If thisScript.MyObject Is Nothing Then
            Set thisModule = Me.CodeEngine.Modules.Add(thisScript.Name & "_" & CLng(Rnd(1) * 10000), thisScript)
        Else
            Set thisModule = Me.CodeEngine.Modules.Add(thisScript.Name & "_" & CLng(Rnd(1) * 10000), thisScript.MyObject)
        End If
        Set thisScript.MyModule = thisModule
        If Me.IsInEditor Then
            StatusDialog.PRGBar_SetValue CLng(i), CLng(Me.Mesh_Scripts.Count), "Objects Loading Mesh Code "
        End If
        thisModule.AddCode thisScript.Data
    Next
    
    'loadStatusDialog.Hide
End Sub

Function StartEngine()
Attribute StartEngine.VB_Description = "Starts The Engine"
    Dim thisScript As Script_Class
    Dim thisModule As Module
    Dim curScene As Scene_Class
    'Dim baseObject As IObject_Class
    Dim curGObject As Object
    
    If Me.IsInEditor Then
        Dim loadStatusDialog As StatusDialog
        Set loadStatusDialog = CreateStatusDialog("Please Wait... Starting Game")
    End If
    
    OnGameLoad
    Me.GameWindow.Cls
    'ReloadCodeEngine ' Resetoi Koodimoottorin
    'CreateModules
    
    On Error GoTo GlobalError
    If IsProjectOpened = False Then
        Me.CodeEngine.AddCode Me.GlobalScript.Data
    End If
    On Error GoTo CodeLoadError
    For i = 1 To Me.Scripts.Count ' K‰sittelee Kaikki Skriptit, Luo Koodinsuoritus Moottorille Moduulin Ja Tekee Skriptille Referensin T‰h‰n Moduuliin Josta Sit‰ On Helppo Ohjata
        Set thisScript = Me.Scripts.Item(i)
        If thisScript.MyObject Is Nothing Then ' T‰m‰ On Myˆs Hyv‰ Kun T‰t‰ Operoidaan Windows 9x K‰yttˆj‰rjestelmiss‰
            Set thisModule = Me.CodeEngine.Modules.Add("Script_" & thisScript.Name & "_" & Me.CodeEngine.Modules.Count, thisScript)
        Else
            Set thisModule = Me.CodeEngine.Modules.Add("Script_" & thisScript.Name & "_" & Me.CodeEngine.Modules.Count, thisScript.MyObject)
        End If
        Set thisScript.MyModule = thisModule
        If Me.IsInEditor Then
            loadStatusDialog.PRGBar_SetValue CLng(i), CLng(Me.Scripts.Count), "CodeEngine Creating Modules "
        End If
        thisModule.AddCode thisScript.Data
    Next
    
    On Error GoTo MError
    For i = 1 To Me.Mesh_Scripts.Count
        Set thisScript = Me.Mesh_Scripts.Item(i)
        If thisScript.MyObject Is Nothing Then
            Set thisModule = Me.CodeEngine.Modules.Add(thisScript.Name & "_" & Me.CodeEngine.Modules.Count, thisScript)
        Else
            Set thisModule = Me.CodeEngine.Modules.Add(thisScript.Name & "_" & Me.CodeEngine.Modules.Count, thisScript.MyObject)
        End If
        Set thisScript.MyModule = thisModule
        If Me.IsInEditor Then
            loadStatusDialog.PRGBar_SetValue CLng(i), CLng(Me.Mesh_Scripts.Count), "Objects Loading Mesh Code "
        End If
        thisModule.AddCode thisScript.Data
    Next
    On Error Resume Next
    For i = 1 To Me.Scenes.Count
        Set curScene = Me.Scenes.Item(i)
        curScene.Load
        'For o = i To curScene.Objects.Count
            'Set curGObject = curScene.Objects(o)
            'curGObject.Load 'Me.CodeEngine
        'Next
    Next
    'On Error GoTo SError
    For i = 1 To Me.Scripts.Count
        Set thisScript = Me.Scripts.Item(i)
        If thisScript.Active Then
            thisScript.MyModule.Run "Load"
        End If
    Next
    
    loadStatusDialog.Hide
    Me.EngineTimer.Start
    Exit Function
    
    
GlobalError:
    MsgBox Me.CodeEngine.Error.Description & " In Global Script On Col: " & Me.CodeEngine.Error.Column & " In Line: " & Me.CodeEngine.Error.Line, vbExclamation
    If Me.IsInEditor Then
        loadStatusDialog.Hide
    End If
    OnGameStop
    Exit Function
MError:
    MsgBox Me.CodeEngine.Error.Description & " In Object Mesh Code On Line: " & Me.CodeEngine.Error.Line, vbExclamation
    'Me.EConsole.WriteLine Me.CodeEngine.Error.Description & " In Object Mesh Code On Line: " & Me.CodeEngine.Error.Line
    If Me.IsInEditor Then
        loadStatusDialog.Hide
    End If
    OnGameStop
    Exit Function
    
SError:
    Exit Function
    
CodeLoadError:
    MsgBox Me.CodeEngine.Error.Description & " On Line " & Me.CodeEngine.Error.Line, vbExclamation
    'Me.EConsole.WriteLine Me.CodeEngine.Error.Description & " In " & thisScript.Name & " On Line " & Me.CodeEngine.Error.Line
    If Me.IsInEditor Then
        loadStatusDialog.Hide
    End If
    OnGameStop
    Exit Function

End Function

Private Function EngineUpdate()
    'Dim baseObject As Object
    'Dim thisScript As Script_Class
    'Dim curScene As Scene_Class
    'Dim curGObject As GameObject_Class
    'Scripts = listObj.Count()
    CurrentFrame = CurrentFrame + 1
    'On Error GoTo SErr32
    On Error Resume Next
    For Each Item In Me.Scripts
        'Me.CodeEngine.Modules.Item("Script_" & thisScript.Name).Run "Update"
        If Item.Active Then
            Item.MyModule.Run "Update"
        End If
    Next
    For Each cScene In Me.Scenes
        cScene.Update
        For Each cObject In cScene.Objects
            If cObject.Active Then
                cObject.Update
            End If
        Next
    Next
    SimulationTime = SimulationTime + Me.GetTargetFrameTime()
    Me.DrawClass.PrintText 1, 1, "FPS: " & Me.GetFrameRate()
    Me.DrawClass.PrintText 1, 200, "FrameTime: " & Me.GetFrameTime()
End Function

Function RenderScene(Frames As Long, Optional DestinationPath As String)
    Dim thisScript As Script_Class
    Dim curScene As Scene_Class
    Dim IGameObject As IObject_Class
    Dim curGObject As GameObject_Class
    
    Dim loadStatusDialog As StatusDialog
    Set loadStatusDialog = CreateStatusDialog("Please Wait... Rendering Scenes")
    
    CurrentFrame = 0
    SimulationTime = 0
    RenderStart = Timer
    
    Me.EditorWindow.Cls
    Me.EditorWindow.Caption = "Rendering Scenes"
    ReloadCodeEngine ' Resetoi Koodimoottorin
    'CreateModules
    RaiseEvent OnRenderStart
    
    On Error Resume Next
    Me.CodeEngine.AddCode Me.GlobalScript.Data
    For i = 1 To Me.Scripts.Count
        Set thisScript = Me.Scripts.Item(i)
        Set thisModule = Me.CodeEngine.Modules.Add("Script_" & thisScript.Name, thisScript.MyObject)
        Set thisScript.MyModule = thisModule
        If Me.IsInEditor Then
            loadStatusDialog.PRGBar_SetValue CLng(i), CLng(Me.Scripts.Count), "CodeEngine Modules Adding Code "
        End If
        thisModule.AddCode thisScript.Data
    Next
    
    For i = 1 To Me.Mesh_Scripts.Count
        Set thisScript = Me.Mesh_Scripts.Item(i)
        Set thisModule = Me.CodeEngine.Modules.Add(thisScript.Name, thisScript.MyObject)
        Set thisScript.MyModule = thisModule
        If Me.IsInEditor Then
            loadStatusDialog.PRGBar_SetValue CLng(i), CLng(Me.Mesh_Scripts.Count), "Objects Loading Mesh Code "
        End If
        thisModule.AddCode thisScript.Data
    Next
    
    For i = 1 To Me.Scenes.Count
        Set curScene = Me.Scenes.Item(i)
        curScene.OnLoad Me.CodeEngine
        For O = i To curScene.Objects.Count
            Set curGObject = curScene.Objects(O)
            Set IGameObject = curGObject
            loadStatusDialog.PRGBar_SetValue CLng(O), CLng(curScene.Objects.Count), "Objects Loading Mesh Code "
            IGameObject.Load Me.CodeEngine
        Next
    Next
    
    For i = 1 To Me.Scripts.Count
        Set thisScript = Me.Scripts.Item(i)
        'Me.CodeEngine.Modules.Item("Script_" & thisScript.Name).Run "Load"
        thisScript.MyModule.Run "Load"
    Next
    
    For f = 1 To Frames
        CurrentFrame = CurrentFrame + 1
        loadStatusDialog.PRGBar_SetValue CLng(f), Frames, "Rendering Frames "
        For i = 1 To Me.Scripts.Count
            Set thisScript = Me.Scripts.Item(i)
            thisScript.MyModule.Run "Update"
        Next
        For i = 1 To Me.Scenes.Count
            Set curScene = Me.Scenes.Item(i)
            For O = 1 To curScene.Objects.Count
                Set curGObject = curScene.Objects(O)
                Set IGameObject = curGObject
                'curGObject.Update Me.CodeEngine
                IGameObject.Update Me.CodeEngine
            Next
        Next
        SimulationTime = SimulationTime + Me.GetTargetFrameTime()
        
        If Not DestinationPath = "" Then
            On Error GoTo SaveError
            SavePicture Me.GameWindow.Image, DestinationPath & "\" & f & ".bmp"
        End If
    Next
    MsgBox "Render Complete", vbInformation
    loadStatusDialog.Hide
    Me.EditorWindow.Caption = "Render Result"
    RaiseEvent OnRenderEnd
    Exit Function
    
SaveError:
    Me.EConsole.WriteLine "Failed To Save Frame " & f & DestinationPath & f & ".bmp" '& " Reason " & ErrObject.Description
    Resume Next
End Function

Private Sub CodeEngine_Error()
    Me.EConsole.WriteLine "Warning. " & Me.CodeEngine.Error.Source & " " & Me.CodeEngine.Error.Description & " In Line " & CodeEngine.Error.Line
    RaiseEvent OnError(Me.CodeEngine.Error)
End Sub

Private Sub CodeEngine_Timeout()
    If Me.IsEngineRunning Then
        Me.EndGame
    End If
    RaiseEvent OnTimeout
End Sub

Private Sub EditorWindow_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.IsEngineRunning Then
        Exit Sub
    End If
    If KeyCode = vbKeyUp Then
        Me.EditorMoveCamera 0, -100
    ElseIf KeyCode = vbKeyDown Then
        Me.EditorMoveCamera 0, 100
    End If
    If KeyCode = vbKeyLeft Then
        Me.EditorMoveCamera -100, 0
    ElseIf KeyCode = vbKeyRight Then
        Me.EditorMoveCamera 100, 0
    End If
    Me.EditorRefreshScenes
    RaiseEvent EditorKeyDown(KeyCode)
End Sub

Private Sub EditorWindow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.IsEngineRunning = False Then
        If Button = 1 Then
            Me.EditorGetSelectedObject X, Y
            If Me.GetSelectedObject() Is Nothing Then
                'Me.EditorWindow.StatusBar1.Panels(5).text = "Selected Object: Nothing"
                Exit Sub
            Else
                'Me.EditorWindow.StatusBar1.Panels(5).text = "Selected Object: " & Me.GetSelectedObject().Name
            End If
        End If
        RaiseEvent EditorMouseDown(Button, X, Y)
    End If
End Sub

Private Sub EditorWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.IsEngineRunning = False Then
        If Button = 1 Then
            If Not Me.GetSelectedObject() Is Nothing Then
                Me.EditorWindow.Cls
                Me.GetSelectedObject().Position.X = X
                Me.GetSelectedObject().Position.Y = Y
                Me.EditorRefreshScenes
            End If
        Else
    
        End If
        RaiseEvent EditorMouseMove(Button, X, Y)
    End If
End Sub

Private Sub EngineTimer_Update()
    EngineUpdate
End Sub

Private Sub GameWindow_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.IsEngineRunning Then
        If KeyCode = 27 Then
            Me.EndGame
        End If
        If KeyCode = 112 Then
            Me.EConsole.Show
        End If
        If KeyCode = vbKeyPrint Then
            Me.EngineTimer.Kill
            Me.GameWindow.PrintForm
            Me.EngineTimer.Start
        End If
    Else
        
    End If
    RaiseEvent KeyDown(KeyCode)
End Sub

Private Sub GameWindow_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode)
End Sub

Private Sub GameWindow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, X, Y)
End Sub

Private Sub GameWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, X, Y)
End Sub

Private Sub GameWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, X, Y)
End Sub
