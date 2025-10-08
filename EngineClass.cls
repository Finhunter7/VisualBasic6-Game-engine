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
Public Event OnDataChanged()
Public Event OnSceneCreated()
Public Event OnSceneActivated()
Public Event OnGameStart()
Public Event OnRenderStart()
Public Event OnRenderEnd()
Public Event OnGameStop()
Public Event OnGameEnd()
Public Event OnError()
Public Event MouseDown(Button As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer)
Public Event KeyUp(KeyCode As Integer)

' Project
Public ProjectName As String
Public ProjectPath As String
Public ProjectCodeLanguage As String
Public ProjectType As ProjectTypes
Public GameWindowSizeX As Long
Public GameWindowSizeY As Long

' Editor Utils
Public IsInEditor As Boolean
Public WorkspaceUtilClass As Object

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
Private EditorSelectedObject As GameObject_Class
'Public EngineTimer As Timer

Public ProjectBrowser As Form
Public SceneBrowser As Form
Public SettingsWindow As Form
Public EConsole As Form

Public InputSystem As New InputSystem_Class
Public DrawClass As New Draw_Class

Private EVectorClass As New Vector_Class
Private MathClass As New Math_Class
Private AudioSystem As New AudioSystem_Class
Private fileSystemHandler As FileSystemObject
Public Direct8Lib As Object
Private Direct1LIB As Object

Private EngineUtil_Class As New EngineUtils_Class

Public SaveLoadGameClass As New GameLoader_Class

Public Scenes As New Collection ' Aktiiviset Scenet
Public ScenesInactive As New Collection
Private SceneRestorePoint As New Collection

Public GameObjects As New Collection
Public Scripts As New Collection
Public Mesh_Scripts As New Collection
Public Textures As New Collection
Public VBCETextures As New Collection

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

Function EditorGetSelectedObject(MPosX As Single, MPosY As Single)
    Dim curScene As Scene_Class
    Dim SceneObject As GameObject_Class
    If Me.IsEngineRunning Then
        Set EditorSelectedObject = Nothing
        Exit Function
    End If
    For i = 1 To Me.Scenes.Count
        Set curScene = Me.Scenes(i)
        For o = 1 To curScene.Objects.Count
            Set SceneObject = curScene.Objects(o)
            If MPosX > (SceneObject.Position.X - (SceneObject.EditorColliderSize.X / 2)) And MPosX < (SceneObject.Position.X + (SceneObject.EditorColliderSize.X / 2)) Then
                If MPosY > (SceneObject.Position.Y - (SceneObject.EditorColliderSize.Y / 2)) And MPosY < (SceneObject.Position.Y + (SceneObject.EditorColliderSize.Y / 2)) Then
                    Me.EditorRefreshScenes
                    Me.EditorWindow.FontSize = 8
                    Me.EditorWindow.ForeColor = &HFFFFFF
                    Me.GameWindow.CurrentX = 0
                    Me.GameWindow.CurrentY = 0
                    Me.EditorWindow.Print "Object Name: " & SceneObject.Name
                    Me.EditorWindow.Print "Object Tag: " & SceneObject.Tag
                    Me.EditorWindow.Print "Object X: " & SceneObject.Position.X
                    Me.EditorWindow.Print "Object Y: " & SceneObject.Position.Y
                    Me.EditorWindow.Print "Object Z: " & SceneObject.Position.Z
                    Me.EditorWindow.Print "Object ScaleX: " & SceneObject.SScale.X
                    Me.EditorWindow.Print "Object ScaleY: " & SceneObject.SScale.Y
                    Me.EditorWindow.Print "Object ScaleZ: " & SceneObject.SScale.Z
                    Me.EditorWindow.Print "Object Size: " & SceneObject.Size
                    Me.EditorWindow.Line (SceneObject.Position.X - (SceneObject.EditorColliderSize.X / 2), SceneObject.Position.Y - (SceneObject.EditorColliderSize.Y / 2))-(SceneObject.Position.X + (SceneObject.EditorColliderSize.X / 2), SceneObject.Position.Y + (SceneObject.EditorColliderSize.Y / 2)), vbGreen, B
                    Me.EditorWindow.Line (SceneObject.Position.X, SceneObject.Position.Y)-(SceneObject.Position.X + (SceneObject.EditorColliderSize.X), SceneObject.Position.Y), vbRed
                    Me.EditorWindow.Print "X"
                    Me.EditorWindow.Line (SceneObject.Position.X, SceneObject.Position.Y)-(SceneObject.Position.X, SceneObject.Position.Y + SceneObject.EditorColliderSize.Y), vbGreen
                    Me.EditorWindow.Print "Y"
                    Set EditorSelectedObject = SceneObject
                    Exit Function
                End If
            End If
        Next
    Next
    If Me.Scenes.Count > 0 And Not EditorSelectedObject Is Nothing Then
        Me.EditorRefreshScenes
    End If
    Set EditorSelectedObject = Nothing
End Function

Function GetSelectedObject() As GameObject_Class
    Set GetSelectedObject = EditorSelectedObject
End Function

Function SelectObject(GameObject As GameObject_Class)
    If Me.IsEngineRunning Then
        Exit Function
    End If
    Me.EditorRefreshScenes
    Me.EditorWindow.Line (GameObject.Position.X - (GameObject.EditorColliderSize.X / 2), GameObject.Position.Y - (GameObject.EditorColliderSize.Y / 2))-(GameObject.Position.X + (GameObject.EditorColliderSize.X / 2), GameObject.Position.Y + (GameObject.EditorColliderSize.Y / 2)), vbGreen, B
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
    Me.ProjectBrowser.Update
    Me.SceneBrowser.Update
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

Sub AddGlobalObject(Name As String, TObject As Object)
    Dim newObject As New CodeEngineObject_Class
    newObject.Build Name, TObject
    Me.GlobalObjects.Add newObject, Name
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

Function AddScene(Name, Optional OverOrUnder As SceneConstants) As Scene_Class ' Lis‰‰ uuden Scenen ScenesInactive kokoelmasta Scenes kokoelmaan
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
    
    If OverOrUnder = Default Or Me.Scenes.Count < 1 Then
        Scenes.Add newScene, newScene.Name
        newScene.OnLoad Me.CodeEngine
        Set AddScene = newScene
    ElseIf OverOrUnder = Overlay Then
        Scenes.Add newScene, newScene.Name, , Scenes.Count
        newScene.OnLoad Me.CodeEngine
        Set AddScene = newScene
    Else
        Scenes.Add newScene, newScene.Name, 1
        newScene.OnLoad Me.CodeEngine
        Set AddScene = newScene
    End If
    
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

Function CreateNewScript(Name As String, SType As ScriptTypes, Optional PresetData As String) As Script_Class
    Dim newClass As New Script_Class
    Dim SClass As Script_Class
    Dim EXObject As GameObject_Class
    Dim modul As Module

    newClass.Name = Name
    
    If SType = DefaultScript Then
        newClass.Data = vbNewLine + "Sub Load() 'Called Before First Frame" + vbNewLine + vbNewLine + "End Sub" + vbNewLine + vbNewLine + "Sub Update() 'Called On Every Frame Update" + vbNewLine + vbNewLine + "End Sub"
        newClass.Tag = "Script"
        Me.Scripts.Add newClass, Name
    ElseIf SType = MeshScript Then
        newClass.Data = vbNewLine & "Sub Draw() 'Same as Update but executed after Update" & vbNewLine & "" & vbNewLine & "End Sub" & vbNewLine & vbNewLine & "Sub EditorDraw()" & vbNewLine & "" & vbNewLine & "End Sub"
        newClass.Tag = "MeshScript"
        Me.Mesh_Scripts.Add newClass, Name
    ElseIf SType = ComponentScript Then
        newClass.Data = vbNewLine + "Sub Load() 'Called Before First Frame" + vbNewLine + vbNewLine + "End Sub" + vbNewLine + vbNewLine + "Sub Update() 'Called On Every Frame Update" + vbNewLine + vbNewLine + "End Sub"
        newClass.Tag = "ComponentScript"
        Me.Scripts.Add newClass, Name
    End If
    If Not PresetData = "" Then
        newClass.Data = PresetData
    End If
    
    'Me.ProjectBrowser.UpdateProjectBrowser
    Set CreateNewScript = newClass
    RaiseEvent OnDataChanged
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

Sub SceneAddObject(PresetNum As ObjectPresets, Optional Scene As Scene_Class)
    Dim thisScene As Scene_Class
    Dim objectName As String
    
    Set thisScene = Scene
    If thisScene Is Nothing Then
        Set thisScene = Me.GetCurrentScene()
    End If
    If thisScene Is Nothing Then
        MsgBox "No Active Scenes Found", vbExclamation
        Exit Sub
    End If
    
    objectName = InputBox("Give Object Name")
    If Len(objectName) = 0 Then
        Exit Sub
    End If
    
    Set newObject = thisScene.CreateNewObject(objectName, , , PresetNum)
    Set newObject.MyScene = thisScene
    RaiseEvent OnDataChanged
End Sub

Sub EditObjectCode(Edit As CodeEditorEnums, Object As GameObject_Class, Optional Scene As Scene_Class)
    
End Sub

Private Sub Class_Initialize()
    ScriptEngine_Intialize
End Sub
Private Sub ScriptEngine_Intialize()
    'Intialize ScriptEngine
    If Me.CodeEngine Is Nothing Then
        Set CodeEngine = New MSScriptControl.ScriptControl
    Else
        Me.CodeEngine.Reset
    End If
    If Me.CodeEngine Is Nothing Then
        MsgBox "MSScriptControl Failed To Load. Unable To Load VBCE", vbExclamation
        'Unload Me
        End
    End If
    Set fileSystemHandler = New Scripting.FileSystemObject
    AudioSystem.Build Me
    Me.SetCodeLaunguage "VBScript"
    Me.CodeEngine.AddObject "Engine", Me, True
    Me.CodeEngine.AddObject "Vector", EVectorClass, True
    Me.CodeEngine.AddObject "Math", MathClass, True
    Me.CodeEngine.AddObject "AudioSystem", AudioSystem, True
    
    Me.GlobalScript.Name = "Global"
    Me.ProjectName = "Project1"
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
    
    If Commands(0) = "/P" Then
        
    ElseIf Commands(0) = "/C" Then
        LoadStartupCode Commands(1)
    End If
End Function
Private Sub LoadStartupCode(FilePath As String)
    Dim Data As String
    Dim curLine As String
    Dim curFile As Object
    'Open (FilePath) For Input As #1
        'Do Until EOF(1)
            'Input #1, curLine
            'Data = Data & curLine & vbCrLf
        'Loop
    'Close #1
    If fileSystemHandler Is Nothing Then
        Exit Sub
    End If
    
    If fileSystemHandler.FileExists(FilePath) Then
        Set curFile = fileSystemHandler.OpenTextFile(FilePath, ForReading, False)
        Do Until curFile.AtEndOfStream
            curLine = curFile.ReadLine()
            Data = Data & curLine & vbCrLf
        Loop
        curFile.Close
        Me.CodeEngine.AddCode Data
    Else
        MsgBox "File " & FilePath & " Does Not Exists", vbExclamation
    End If
End Sub

Function LoadEngine(TIsInEditor As Boolean, TGameWindow As Form, TConsole As Form, Optional TEditorWindow As Form, Optional TProjectBrowser As Form, Optional TSceneBrowser As Form, Optional TSettingsWindow As Form)
    Me.IsInEditor = TIsInEditor
    Set Me.GameWindow = TGameWindow
    
    If Me.IsInEditor Then
        ProjectType = VBCEGameProject
        Set Me.EditorWindow = TEditorWindow
        'Set Me.CodeEngine = ScriptEngine
        'Set Me.EngineTimer = TEngineTimer
        'Set Me.EngineTimer = New EngineTimer_Class
        Set newGameWindow = New GameWindow
        Set Me.SettingsWindow = TSettingsWindow
        Set Me.ProjectBrowser = TProjectBrowser
        Set Me.SceneBrowser = TSceneBrowser
        Set Me.ProjectBrowser.GameEngine = Me
        Set Me.SceneBrowser.GameEngine = Me
        Set Me.WorkspaceUtilClass = New VBCEWorkspace_Class
        Me.WorkspaceUtilClass.LoadDeveloperTools Me
    End If
    
    Set Me.InputSystem.GameEngine = Me
    Set Me.EConsole = VBCE.Console
    Set Me.EConsole.GameEngine = Me
    
    Set Me.DrawClass.GameEngine = Me
    Set Me.EngineTimer = New EngineTimer_Class
    
    Me.CodeEngine.AddObject "Console", Me.EConsole, True
    Dim GOConsole As New CodeEngineObject_Class
    GOConsole.Build "Console", Me.EConsole
    Me.GlobalObjects.Add GOConsole, GOConsole.Name
    
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
    MsgBox "Game Paused"
    Me.EngineTimer.Start
End Function

Function ResumeGame()
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

Function SetTargetFramerate(Framerate As Integer)
    TargetFramerate = Framerate
    FrameTime = (1000 / Framerate)
    Me.EngineTimer.Interval = FrameTime
    RaiseEvent OnDataChanged
End Function

Function GetTargetFramerate() As Integer
    GetTargetFramerate = TargetFramerate
End Function

Function GetFrameTime() As Double
    GetFrameTime = FrameTime
End Function

Function GetRenderedFrames() As Long
    GetRenderedFrames = CurrentFrame
End Function

Function GetRealFrameTime() As Double
    GetRealFrameTime = Me.EngineTimer.GetRealIntervalTime()
End Function

Function GetFrameRate() As Double
    GetFrameRate = (((TargetFramerate / (Timer - RenderStart)) * SimulationTime) / 1000)
End Function

Function GetSimulationTime() As Double
    GetSimulationTime = SimulationTime
End Function

Function GetRealTime() As Double
    GetRealTime = Timer - RenderStart
End Function

Function GetGamePath() As String
    GetGamePath = App.Path
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
    Me.GameWindow.Caption = "Engine Stopped"
    Me.EConsole.WriteLine "Engine Stopped"
    Set Me.GameWindow = Me.EditorWindow
    Me.EditorRefreshScenes
    'Me.CodeEngine.Reset
    RaiseEvent OnGameStop
End Function

Sub SaveObjectToDisk(SaveData As SavingEnums, FileName As String, EObject As Object)
    SaveLoadGameClass.SaveObjectToDiskVBCFormat SaveData, FileName, EObject
End Sub

Function SaveProject()
    'Dim fileSystemHandler As New Scripting.FileSystemObject
    Dim projectFile As TextStream
    Dim dataFolder As Folder
    Dim setProjectPath As String
    
    If Me.ProjectPath = "" Then
        setProjectPath = FileBrowser1.SaveToDir(MainWindow)
        If Len(setProjectPath) = 0 Then
            Exit Function
        End If
        Me.ProjectPath = setProjectPath
        
        If Not fileSystemHandler.FolderExists(Me.ProjectPath) Then
            fileSystemHandler.CreateFolder Me.ProjectPath
        End If
    ElseIf fileSystemHandler.FolderExists(Me.ProjectPath) Then
        
    Else
        fileSystemHandler.CreateFolder Me.ProjectPath
    End If
    
    Set projectFile = fileSystemHandler.CreateTextFile(Me.ProjectPath & "\" & Me.ProjectName & ".VbceProject", True)
    SaveProjectGeneralInfo projectFile
    
    If fileSystemHandler.FolderExists(Me.ProjectPath & "\Scripts") Then
        fileSystemHandler.DeleteFolder Me.ProjectPath & "\Scripts", True
        Set dataFolder = fileSystemHandler.CreateFolder(Me.ProjectPath & "\Scripts")
        SaveProjectScriptData dataFolder
    Else
        Set dataFolder = fileSystemHandler.CreateFolder(Me.ProjectPath & "\Scripts")
        SaveProjectScriptData dataFolder
    End If
    
    If fileSystemHandler.FolderExists(Me.ProjectPath & "\Scenes") Then
        fileSystemHandler.DeleteFolder Me.ProjectPath & "\Scenes"
        Set dataFolder = fileSystemHandler.CreateFolder(Me.ProjectPath & "\Scenes")
        SaveProjectSceneData dataFolder
    Else
        Set dataFolder = fileSystemHandler.CreateFolder(Me.ProjectPath & "\Scenes")
        SaveProjectSceneData dataFolder
    End If
    
    projectFile.Close
End Function

Function ClearGameEngineData()
    For i = 1 To Me.Scenes.Count
        Me.Scenes.Remove i
    Next
    For i = 1 To Me.ScenesInactive.Count
        Me.ScenesInactive.Remove i
    Next
    For i = 1 To Me.GameObjects.Count
        Me.GameObjects.Remove i
    Next
    For i = 1 To Me.Scripts.Count
        Me.Scripts.Remove i
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

Function OpenProject(FileName As String)
    Const projectVarsCount = 2
    Dim ProjectVars(projectVarsCount) As String
    Dim projectFile As file
    Dim projectFileTextStream As TextStream
    Dim sceneFileTextStream As TextStream
    Dim curLine As String
    Dim curDataSelector As Integer
    Dim curScene As Scene_Class
    ProjectVars(1) = "Project_Name"
    ProjectVars(2) = "Project_Path"
    
    If fileSystemHandler.FileExists(FileName) Then
        Set projectFile = fileSystemHandler.GetFile(FileName)
        Set projectFileTextStream = projectFile.OpenAsTextStream(ForReading, TristateUseDefault)
        
        curLine = projectFileTextStream.ReadLine()
        
        If curLine = "VBCE_ProjectFile                      " Then
            ClearGameEngineData
        Else
            MsgBox "File Is Not VBCE ProjectFile. Open Canceled", vbExclamation
            projectFileTextStream.Close
            Exit Function
        End If
        
        Do Until projectFileTextStream.AtEndOfStream
            curLine = projectFileTextStream.ReadLine()
            If RTrim(curLine) = "[Project]" Then
                curDataSelector = 1
                curLine = projectFileTextStream.ReadLine()
            End If
            If RTrim(curLine) = "[InactiveScenes]" Then
                curDataSelector = 3
                curLine = projectFileTextStream.ReadLine()
            End If
            If RTrim(curLine) = "[ActiveScenes]" Then
                curDataSelector = 2
                curLine = projectFileTextStream.ReadLine()
            End If
            If RTrim(curLine) = "}" Or RTrim(curLine) = "[Objects]" Then
                curDataSelector = 0
                curLine = projectFileTextStream.ReadLine()
            End If
            
            If curDataSelector = 1 Then
                For i = 1 To projectVarsCount
                    Select Case Left(curLine, Len(ProjectVars(i)))
                    
                        Case "Project_Name"
                            Me.ProjectName = RTrim(Replace(curLine, "Project_Name", ""))
                        Case "Project_Path"
                            Me.ProjectPath = RTrim(Replace(curLine, "Project_Path", ""))
                        Case ""
                            
                    End Select
                Next
            ElseIf curDataSelector = 3 Then
                Set curScene = Me.CreateNewScene(curLine)
            ElseIf curDataSelector = 2 Then
                Me.AddScene curLine
            End If
            
        Loop
        projectFileTextStream.Close
    Else
        MsgBox "Projectfile Is Invalid", vbExclamation
        Exit Function
    End If
End Function

Private Function SaveProjectGeneralInfo(TStream As TextStream)
    Dim curScene As Scene_Class
    TStream.WriteLine "VBCE_ProjectFile                      "
    TStream.WriteLine "[Project]"
    TStream.WriteLine "Project_Version " & App.Major & "." & App.Minor & "." & App.Revision
    TStream.WriteLine "Project_Name " & Me.ProjectName
    TStream.WriteLine "Project_Path " & Me.ProjectPath
    TStream.WriteLine "Project_BaseSceneNum " & Me.BaseScene
    TStream.WriteLine "Project_ScenesInactive_Count " & Me.ScenesInactive.Count
    TStream.WriteLine "Project_Scenes_Count " & Me.Scenes.Count
    TStream.WriteLine "Project_TargetFramerate " & Me.GetTargetFramerate()
    TStream.WriteLine "Project_Code_Language " & Me.GetCodeLanguage()
    TStream.WriteLine "[InactiveScenes]"
    For i = 1 To Me.ScenesInactive.Count
        TStream.WriteLine Me.ScenesInactive.Item(i).Name
    Next
    TStream.WriteLine "[ActiveScenes]"
    For i = 1 To Me.Scenes.Count
        TStream.WriteLine Me.Scenes.Item(i).Name
    Next
    TStream.WriteLine "[Objects]"
    For i = 1 To Me.ScenesInactive.Count
        Set curScene = Me.ScenesInactive.Item(i)
        For o = 1 To curScene.Objects.Count
            TStream.WriteLine curScene.Name & "_" & curScene.Objects.Item(o).Name
        Next
    Next
    
End Function
Private Function SaveProjectScriptData(sFolder As Folder)
    Dim curScript As Script_Class
    Dim sData As TextStream
    If Me.Scripts.Count > 0 Then
        For i = 1 To Me.Scripts.Count
            Set curScript = Me.Scripts.Item(i)
            Set sData = sFolder.CreateTextFile(curScript.Name & ".VbceScript", True)
            sData.WriteLine "[General]"
            sData.WriteLine "Script_Name " & curScript.Name
            sData.WriteLine "[Data]"
            sData.Write curScript.Data
        Next
        sData.Close
    End If
End Function

Private Function SaveProjectSceneData(scFolder As Folder)
    Dim curScene As Scene_Class
    Dim sData As TextStream
    If Me.ScenesInactive.Count > 0 Then
        For i = 1 To Me.ScenesInactive.Count
            Set curScene = Me.ScenesInactive.Item(i)
            Set sData = scFolder.CreateTextFile(curScene.Name & ".VbceScene", True)
            sData.WriteLine "[General]"
            sData.WriteLine "Scene_Name " & curScene.Name
            sData.WriteLine "[Gravity]"
            sData.WriteLine "X " & curScene.Gravity.X
            sData.WriteLine "Y " & curScene.Gravity.Y
            sData.WriteLine "Z " & curScene.Gravity.Z
            sData.WriteLine "[Objects]"
            For o = 1 To curScene.Objects.Count
                sData.WriteLine curScene.Objects.Item(o).Name
            Next
        Next
        sData.Close
    End If
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
    Dim baseObject As IObject_Class
    Dim curGObject As GameObject_Class
    
    If Me.IsInEditor Then
        Dim loadStatusDialog As StatusDialog
        Set loadStatusDialog = CreateStatusDialog("Please Wait... Starting Game")
    End If
    
    OnGameLoad
    Me.GameWindow.Cls
    ReloadCodeEngine ' Resetoi Koodimoottorin
    'CreateModules
    
    On Error GoTo GlobalError
    Me.CodeEngine.AddCode Me.GlobalScript.Data
    On Error GoTo CodeLoadError
    For i = 1 To Me.Scripts.Count ' K‰sittelee Kaikki Skriptit, Luo Koodinsuoritus Moottorille Moduulin Ja Tekee Skriptille Referensin T‰h‰n Moduuliin Josta Sit‰ On Helppo Ohjata
        Set thisScript = Me.Scripts.Item(i)
        If thisScript.MyObject Is Nothing Then ' T‰m‰ On Myˆs Hyv‰ Kun T‰t‰ Operoidaan Windows 9x K‰yttˆj‰rjestelmiss‰
            Set thisModule = Me.CodeEngine.Modules.Add("Script_" & thisScript.Name & "_" & CLng(Rnd(1) * 10000), thisScript)
        Else
            Set thisModule = Me.CodeEngine.Modules.Add("Script_" & thisScript.Name & "_" & CLng(Rnd(1) * 10000), thisScript.MyObject)
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
            Set thisModule = Me.CodeEngine.Modules.Add(thisScript.Name & "_" & CLng(Rnd(1) * 10000), thisScript)
        Else
            Set thisModule = Me.CodeEngine.Modules.Add(thisScript.Name & "_" & CLng(Rnd(1) * 10000), thisScript.MyObject)
        End If
        Set thisScript.MyModule = thisModule
        If Me.IsInEditor Then
            loadStatusDialog.PRGBar_SetValue CLng(i), CLng(Me.Mesh_Scripts.Count), "Objects Loading Mesh Code "
        End If
        thisModule.AddCode thisScript.Data
    Next
    
    For i = 1 To Me.Scenes.Count
        Set curScene = Me.Scenes.Item(i)
        curScene.OnLoad Me.CodeEngine
        For o = i To curScene.Objects.Count
            Set curGObject = curScene.Objects(o)
            
            Set baseObject = curGObject
            baseObject.Load Me.CodeEngine
        Next
    Next
    'On Error GoTo SError
    On Error Resume Next
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
    Dim baseObject As IObject_Class
    'Dim thisScript As Script_Class
    'Dim curScene As Scene_Class
    'Dim curGObject As GameObject_Class
    'Scripts = listObj.Count()
    CurrentFrame = CurrentFrame + 1
    'On Error GoTo SErr32
    On Error Resume Next
    For Each thisScript In Me.Scripts
        'Me.CodeEngine.Modules.Item("Script_" & thisScript.Name).Run "Update"
        If thisScript.Active Then
            thisScript.MyModule.Run "Update"
        End If
    Next
    
    'On Error GoTo DErr32
    On Error Resume Next
    For Each curScene In Me.Scenes
        For Each curGObject In curScene.Objects
            Set baseObject = curGObject
            If baseObject.Active Then
                baseObject.Update Me.CodeEngine
            End If
        Next
    Next
    SimulationTime = SimulationTime + Me.GetFrameTime()
    Me.DrawClass.PrintText 1, 1, "FPS: " & Me.GetFrameRate()
    Me.DrawClass.PrintText 1, 200, "FrameTime: " & Me.GetRealFrameTime()
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
        For o = i To curScene.Objects.Count
            Set curGObject = curScene.Objects(o)
            Set IGameObject = curGObject
            loadStatusDialog.PRGBar_SetValue CLng(o), CLng(curScene.Objects.Count), "Objects Loading Mesh Code "
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
            For o = 1 To curScene.Objects.Count
                Set curGObject = curScene.Objects(o)
                Set IGameObject = curGObject
                'curGObject.Update Me.CodeEngine
                IGameObject.Update Me.CodeEngine
            Next
        Next
        SimulationTime = SimulationTime + Me.GetFrameTime()
        
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
    If Me.IsInEditor Then
        MainWindow.StatusBar1.Panels(1).text = Me.CodeEngine.Error.Source & " " & CodeEngine.Error.Description & " In Line " & CodeEngine.Error.Line
    End If
    Me.EConsole.WriteLine "Warning. " & Me.CodeEngine.Error.Source & " " & Me.CodeEngine.Error.Description & " In Line " & CodeEngine.Error.Line
    RaiseEvent OnError
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
