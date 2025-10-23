VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBCEWorkspace_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private GameEngine As EngineClass

Enum BrowserTypes
    SolutionBrowser = 0
    ProjectBrowser = 1
    SceneBrowser = 2
End Enum

Enum SBTActions
    CreateNewitem = 0
    OpenNew = 1
    SaveItem = 2
    RemoveItem = 3
    SBTBringFront = 4
    SBTBringBack = 5
End Enum


Public Sub LoadDeveloperTools(TargetGameEngine As EngineClass)
    Set GameEngine = TargetGameEngine
End Sub

Private Sub CreateProjectTree(TreeViewTarget As Object)
    Dim curNode As Object
    
    Set curNode = TreeViewTarget.Nodes.Add(, 1, "c_Project_1", GameEngine.ProjectName, 20, 20)
    curNode.Tag = "Project"
    Set curNode = TreeViewTarget.Nodes.Add("c_Project_1", 4, "Engine", "Engine", 21, 21)
    curNode.Tag = "Object"
    Set curNode = TreeViewTarget.Nodes.Add("Engine", 4, "F_References", "References", 4, 4)
    curNode.Tag = "Object"
    
    With GameEngine
        For i = 1 To .EngineRefrences.Count
            Set curNode = TreeViewTarget.Nodes.Add("F_References", 4, "Reference_" & .EngineRefrences.Item(i).Name, .EngineRefrences.Item(i).Name, 5, 5)
            curNode.Tag = "Object"
            Set curNode = TreeViewTarget.Nodes.Add("Reference_" & .EngineRefrences.Item(i).Name, 4, "Reference_Class_" & .EngineRefrences.Item(i).Name, "Class: " & .EngineRefrences.Item(i).RefrenceClass, 25, 25)
            curNode.Tag = "Propety"
            Set curNode = TreeViewTarget.Nodes.Add("Reference_" & .EngineRefrences.Item(i).Name, 4, "Reference_FilePath_" & .EngineRefrences.Item(i).Name, "FilePath: " & .EngineRefrences.Item(i).Path, 25, 25)
            curNode.Tag = "Propety"
        Next
    End With
    
    Set curNode = TreeViewTarget.Nodes.Add("c_Project_1", 4, "Program", "Program", 21, 21)
    curNode.Tag = "Object"
    Set curNode = TreeViewTarget.Nodes.Add("Program", 4, "AppPath", "Path: " & GameEngine.GetGamePath, 25, 25)
    curNode.Tag = "Propety"
    Set curNode = TreeViewTarget.Nodes.Add("Program", 4, "ProjectPath", "ProjectPath: " & GameEngine.ProjectPath, 25, 25)
    curNode.Tag = "Propety"
    Set curNode = TreeViewTarget.Nodes.Add("Program", 4, "TaskVisible", "TaskVisible: " & App.TaskVisible, 25, 25)
    curNode.Tag = "Propety"
    Set curNode = TreeViewTarget.Nodes.Add("Program", 4, "ThreadID", "ThreadID: " & App.ThreadID, 25, 25)
    curNode.Tag = "Propety"
    
    Set curNode = TreeViewTarget.Nodes.Add("c_Project_1", 4, "ScriptEngine", "MSScriptControl", 7, 7)
    curNode.Tag = "CodeEngine"
    Set curNode = TreeViewTarget.Nodes.Add("ScriptEngine", 4, "ScriptEngine_Propetites", "Propertites", 15, 15)
    curNode.Tag = "Propetites"
    
    
    ' CodeEngine Propertites
    With GameEngine.CodeEngine
        Set curNode = TreeViewTarget.Nodes.Add("ScriptEngine_Propetites", 4, "AllowUI", "AllowUI: " & .AllowUI, 25, 25)
        curNode.Tag = "Propety"
        Set curNode = TreeViewTarget.Nodes.Add("ScriptEngine_Propetites", 4, "Language", "Language: " & .Language, 25, 25)
        curNode.Tag = "Propety"
        Set curNode = TreeViewTarget.Nodes.Add("ScriptEngine_Propetites", 4, "Global_Procedures_Count", "Global Procedures Count: " & .Procedures.Count, 25, 25)
        curNode.Tag = "Propety"
        Set curNode = TreeViewTarget.Nodes.Add("ScriptEngine_Propetites", 4, "SitehWnd", "SitehWnd: " & .SitehWnd, 25, 25)
        curNode.Tag = "Propety"
        Set curNode = TreeViewTarget.Nodes.Add("ScriptEngine_Propetites", 4, "Timeout", "Timeout: " & .Timeout, 25, 25)
        curNode.Tag = "Propety"
        Set curNode = TreeViewTarget.Nodes.Add("ScriptEngine_Propetites", 4, "Modules_Count", "Modules Count: " & .Modules.Count, 25, 25)
        curNode.Tag = "Propety"
        Set curNode = TreeViewTarget.Nodes.Add("ScriptEngine_Propetites", 4, "UseSafeSubset", "UseSafeSubset: " & .UseSafeSubset, 25, 25)
        curNode.Tag = "Propety"
    End With
    
    ' CodeEngine Global Objects
    Set curNode = TreeViewTarget.Nodes.Add("ScriptEngine", 4, "F_Objects", "Objects", 1, 1)
    curNode.Tag = "Folder"
    Set curNode = TreeViewTarget.Nodes.Add("ScriptEngine", 4, "F_Library", "Library", 1, 1)
    curNode.Tag = "Folder"
    Set curNode = TreeViewTarget.Nodes.Add("F_Library", 4, "VB_Lib", "VB", 12, 12)
    curNode.Tag = "Library"
    Set curNode = TreeViewTarget.Nodes.Add("F_Library", 4, "VBA_Lib", "VBA", 12, 12)
    curNode.Tag = "Library"
    Set curNode = TreeViewTarget.Nodes.Add("F_Library", 4, "VBRUN_Lib", "VBRUN", 12, 12)
    curNode.Tag = "Library"
    
    For Each Obj In GameEngine.GlobalObjects
        Set curNode = TreeViewTarget.Nodes.Add("F_Objects", 4, "Global_" & Obj.Name, Obj.Name, 5, 5)
        curNode.Tag = "Object"
    Next
    
    Set curNode = TreeViewTarget.Nodes.Add("ScriptEngine", 4, "F_Modules", "Modules", 1, 1)
    curNode.Tag = "Folder"
    Set curNode = TreeViewTarget.Nodes.Add("ScriptEngine", 4, "F_GlobalProcedures", "GlobalProcedures", 1, 1)
    curNode.Tag = "Folder"
    Set curNode = TreeViewTarget.Nodes.Add("ScriptEngine", 4, "F_Global", "Global", 5, 5)
    curNode.Tag = "Folder"
    Set curNode = TreeViewTarget.Nodes.Add("F_Global", 4, "Script_Global", "Global", 24, 24)
    curNode.Tag = "Script"
    
    Set curNode = TreeViewTarget.Nodes.Add("c_Project_1", 4, "F_Scenes", "Scenes", 1, 1)
    curNode.Tag = "Folder"
    Set curNode = TreeViewTarget.Nodes.Add("c_Project_1", 4, "F_Scripts", "Scripts", 1, 1)
    curNode.Tag = "Folder"
    Set curNode = TreeViewTarget.Nodes.Add("c_Project_1", 4, "F_Comp_Scripts", "Component_Scripts", 1, 1)
    curNode.Tag = "Folder"
    Set curNode = TreeViewTarget.Nodes.Add("c_Project_1", 4, "F_Mesh_Scripts", "Mesh_Scripts", 1, 1)
    curNode.Tag = "Folder"
    Set curNode = TreeViewTarget.Nodes.Add("c_Project_1", 4, "F_GameObjects", "GameObjects", 1, 1)
    curNode.Tag = "Folder"
    Set curNode = TreeViewTarget.Nodes.Add("c_Project_1", 4, "F_Textures", "Textures", 1, 1)
    curNode.Tag = "Folder"
    
    With GameEngine
        For Each Item In .Mesh_Scripts
            Set curNode = TreeViewTarget.Nodes.Add("F_Mesh_Scripts", 4, "Mesh_" & Item.Name, Item.Name, 14, 14)
            curNode.Tag = "Script"
        Next
        For Each Item In .Textures
            Set curNode = TreeViewTarget.Nodes.Add("F_Textures", 4, "Texture_" & Item.Name, Item.Name, 14, 14)
            curNode.Tag = "Texture"
        Next
        For Each Item In .VBCETextures
            Set curNode = TreeViewTarget.Nodes.Add("F_Textures", 4, "Texture_" & Item.Name, Item.Name, 14, 14)
            curNode.Tag = "Texture"
        Next
    End With
    
End Sub

Sub BrowseEngineData(TargetTreeView As Object)
    On Error Resume Next
    TargetTreeView.Nodes.Clear
    CreateProjectTree TargetTreeView
    Dim thisNode As Object
    Dim thisModule As Module
    Dim thisprocedure As Procedure
    
    For Each Item In GameEngine.ScenesInactive
        Set thisNode = TargetTreeView.Nodes.Add("F_Scenes", 4, "Sc_" & Item.Name, Item.Name, 17, 17)
        thisNode.Tag = "Scene"
        Set thisNode = TargetTreeView.Nodes.Add("Sc_" & Item.Name, 4, "Sc_" & Item.Name & "_ObjectFolder", "Objects", 1, 1)
        thisNode.Tag = "Folder"
        Set thisNode = TargetTreeView.Nodes.Add("F_GameObjects", 4, Item.Name & "_Objects", Item.Name, 17, 17)
        thisNode.Tag = "Folder"
        For Each Obj In Item.Objects
            Set thisNode = TargetTreeView.Nodes.Add("Sc_" & Item.Name & "_ObjectFolder", 4, "Sc_" & Item.Name & "_Object_" & Obj.Name, Obj.Name, 7, 7)
            thisNode.Tag = "GameObject"
        Next
    Next
    For i = 1 To GameEngine.CodeEngine.Modules.Count
        Set thisModule = GameEngine.CodeEngine.Modules.Item(i)
        Set thisNode = TargetTreeView.Nodes.Add("F_Modules", 4, "M_" & thisModule.Name, thisModule.Name, 5, 5)
        thisNode.Tag = "Module"
        For p = 1 To thisModule.Procedures.Count
            Set thisprocedure = thisModule.Procedures.Item(p)
            Set thisNode = TargetTreeView.Nodes.Add("M_" & thisModule.Name, 4, thisModule.Name & "_Propety_Name_" & thisprocedure.Name, "Procedure: " & thisprocedure.Name, 25, 25)
            thisNode.Tag = "Procedure"
            Set thisNode = TargetTreeView.Nodes.Add(thisModule.Name & "_Propety_Name_" & thisprocedure.Name, 4, thisModule.Name & "_Propety_NumArgs_" & thisprocedure.Name, "Procedure NumArgs: " & thisprocedure.NumArgs, 25, 25)
            thisNode.Tag = "Procedure"
            Set thisNode = TargetTreeView.Nodes.Add(thisModule.Name & "_Propety_Name_" & thisprocedure.Name, 4, thisModule.Name & "_Propety_HasReturnVal_" & thisprocedure.Name, "Procedure HasReturnValue: " & thisprocedure.HasReturnValue, 25, 25)
            thisNode.Tag = "Procedure"
        Next
    Next
    For i = 1 To GameEngine.Scripts.Count
        Set thisNode = TargetTreeView.Nodes.Add("F_Scripts", 4, "S_" & GameEngine.Scripts.Item(i).Name, GameEngine.Scripts.Item(i).Name, 14, 14)
        thisNode.Tag = "Script"
    Next
    For Each GObj In GameEngine.GameObjects
        Set thisNode = TargetTreeView.Nodes.Add(GObj.MyScene.Name & "_Objects", 4, "G_" & GObj.MyScene.Name & "_" & GObj.Name, GObj.Name, 7, 7)
        thisNode.Tag = "GameObject"
        For Each CObj In GObj.GetMyChildren()
            Set thisNode = TargetTreeView.Nodes.Add(CObj.MyScene.Name & "_Objects", 4, "G_" & CObj.MyScene.Name & "_" & CObj.Name, CObj.Name, 7, 7)
            thisNode.Tag = "GameObject"
        Next
        Set thisNode = TargetTreeView.Nodes.Add("G_" & GObj.MyScene.Name & "_" & GObj.Name, 4, "G_" & GObj.MyScene.Name & "_" & GObj.Name & "_" & "AComponents", "ActiveComponents", 1, 1)
        thisNode.Tag = "Folder"
        Set thisNode = TargetTreeView.Nodes.Add("G_" & GObj.MyScene.Name & "_" & GObj.Name, 4, "G_" & GObj.MyScene.Name & "_" & GObj.Name & "_" & "Components", "Components", 1, 1)
        thisNode.Tag = "Folder"
        For Each Comp In GObj.Components
            Set thisNode = TargetTreeView.Nodes.Add("G_" & GObj.MyScene.Name & "_" & GObj.Name & "_" & "Components", 4, GObj.MyScene.Name & "_" & GObj.Name & "_" & Comp.Name, Comp.Name, 5, 5)
            thisNode.Tag = "Component"
        Next
        For Each Comp In GObj.MyComponents
            Set thisNode = TargetTreeView.Nodes.Add("G_" & GObj.MyScene.Name & "_" & GObj.Name & "_" & "AComponents", 4, GObj.MyScene.Name & "_" & GObj.Name & "_A_" & Comp.Name, Comp.Name, 5, 5)
            thisNode.Tag = "Component"
        Next
        
    Next
    For i = 1 To GameEngine.CodeEngine.Procedures.Count
        Set thisprocedure = GameEngine.CodeEngine.Procedures.Item(i)
        Set thisNode = TargetTreeView.Nodes.Add("F_GlobalProcedures", 4, "Procedure_" & thisprocedure.Name, thisprocedure.Name)
        thisNode.Tag = "Procedure"
        Set thisNode = TargetTreeView.Nodes.Add("Procedure_" & thisprocedure.Name, 4, "Propety_NumArgs_" & thisprocedure.Name, "NumArgs: " & thisprocedure.NumArgs, 25, 25)
        thisNode.Tag = "Propety"
        Set thisNode = TargetTreeView.Nodes.Add("Procedure_" & thisprocedure.Name, 4, "Propety_HasReturnValue_" & thisprocedure.Name, "HasReturnValue: " & thisprocedure.HasReturnValue, 25, 25)
        thisNode.Tag = "Propety"
    Next
End Sub

Sub UpdateSceneBrowser(TreeViewObject As Object)
    On Error Resume Next
    TreeViewObject.Nodes.Clear
    Dim curGObject As GameObject_Class
    Dim curGobjChild As GameObject_Class
    Dim curScene As Scene_Class
    Dim curCol As Collection
    
    If GameEngine.Scenes.Count < 1 Then
        TreeViewObject.Nodes.Add , 0, "None", "No Active Scenes"
        Exit Sub
    End If
    For curItem = 1 To GameEngine.Scenes.Count
        Set thisNode = TreeViewObject.Nodes.Add(, 0, GameEngine.Scenes.Item(curItem).Name, GameEngine.Scenes.Item(curItem).Name, 17, 17)
        thisNode.Tag = "Scene"
        
        For curObject = 1 To GameEngine.Scenes.Item(curItem).Objects.Count
            Set curGObject = GameEngine.Scenes.Item(curItem).Objects.Item(curObject)
            Set curScene = GameEngine.Scenes.Item(curItem)
            Set thisNode = TreeViewObject.Nodes.Add(curScene.Name, 4, curScene.Name & "_" & curGObject.Name, curGObject.Name, 7, 7)
            
            thisNode.Tag = "Object"
            For Each CItem In curGObject.GetMyChildren()
                Set thisNode = TreeViewObject.Nodes.Add(curGObject.MyScene.Name & "_" & curGObject.Name, 4, curGObject.Name & "_" & CItem.Name, CItem.Name, 7, 7)
                'Set curGobjChild = curGObject.Child.Item(curChild)
                'Set thisNode = TreeViewObject.Nodes.Add(curGObject.MyScene.Name & "_" & curGObject.Name, 4, curGObject.Name & "_" & CItem.Name, CItem.Name, 7, 7)
                thisNode.Tag = "Object"
            Next
        Next
    Next
End Sub

Sub TreeViewBrowsersOnItemClick(TreeViewObject As Object, BrowserType As BrowserTypes)
    Dim tGameObject As GameObject_Class
    If TreeViewObject.selectedItem Is Nothing Then
        Exit Sub
    End If
    
    If BrowserType = ProjectBrowser Then
        With TreeViewObject.selectedItem
            If .Tag = "Script" Then
                GameEngine.EditScript .text
            ElseIf .Tag = "Scene" Then
    
            ElseIf .Tag = "Module" Then
            'Me.EditModule Me.GameEngine.CodeEngine.Modules(.Key)
            ElseIf .Tag = "MeshCode" Then
            
            End If
        End With
    ElseIf BrowserType = SceneBrowser Then
        With TreeViewObject.selectedItem
            If .Tag = "Object" Then
                If .Parent.Tag = "Object" Then
                    Set tGameObject = GameEngine.GetCurrentScene().Objects(.Parent.text).Child.Item(.text)
                    GameEngine.SelectObject tGameObject
                Else
                    Set tGameObject = GameEngine.GetCurrentScene().Objects(.text)
                    GameEngine.SelectObject tGameObject
                End If
            ElseIf .Tag = "Scene" Then
                
            ElseIf .Tag = "Module" Then
            'Me.EditModule Me.GameEngine.CodeEngine.Modules(.Key)
            ElseIf .Tag = "MeshCode" Then
            
            End If
        End With
    Else
        MsgBox "Not Implemented"
    End If
    
End Sub

Sub TreeViewBrowserHandleActions(TreeViewObject As Object, BrowserType As BrowserTypes, Action As SBTActions)
    Dim tGameObject As GameObject_Class
    
    If TreeViewObject.selectedItem Is Nothing Then
        Exit Sub
    End If
    If BrowserType = ProjectBrowser Then
        With TreeViewObject.selectedItem
            If Action = RemoveItem Then
                If .Tag = "Script" Then
                    GameEngine.RemoveScript .text
                ElseIf .Tag = "Scene" Then
                    GameEngine.UnloadScene .text
                ElseIf .Tag = "GameObject" Then
                    GameEngine.UnloadGameObject .Parent & "_" & .text
                Else
                    MsgBox "Item cannot be removed", vbExclamation
                End If
            End If
        End With
    ElseIf BrowserType = SceneBrowser Then
        With TreeViewObject.selectedItem
            If Action = RemoveItem Then
                If .Tag = "Object" Then
                    GameEngine.Scenes(.root).Objects(.text).EndObject
                ElseIf .Tag = "Scene" Then
                    GameEngine.Scenes(.text).EndScene
                Else
                    MsgBox "Item cannot be removed", vbExclamation
                End If
            ElseIf Action = SBTBringBack Then
            
            End If
        End With
    Else
        MsgBox "Not Implemented"
    End If
End Sub

Sub TreeViewRightClick(TMainWindow As MDIForm, TreeViewObject As Object, MenuObject As Object)
        If Not TreeViewObject.selectedItem Is Nothing Then
            TMainWindow.BmnuRemove1.Caption = "Remove " & TreeViewObject.selectedItem.text & " " & TreeViewObject.selectedItem.Tag
        Else
            TMainWindow.BmnuRemove1.Caption = "Remove"
        End If
        TMainWindow.PopupMenu MenuObject
End Sub

Sub BrowseObject(TreeViewObject As Object, TObject As Object)
    TreeViewObject.Nodes.Clear
    If TObject Is Nothing Then
        TreeViewObject.Nodes.Add , 0, "None", "No Items Selected Or No Items Avaible"
        Exit Sub
    End If
End Sub

