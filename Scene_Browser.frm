VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Scene_Browser 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Scene-Browser"
   ClientHeight    =   3060
   ClientLeft      =   165
   ClientTop       =   210
   ClientWidth     =   3600
   Icon            =   "Scene_Browser.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   Tag             =   "SceneBrowser"
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2805
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   0
            Text            =   "Item Type:"
            TextSave        =   "Item Type:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   0
            Text            =   "Item Class:"
            TextSave        =   "Item Class:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   4471
      _Version        =   327682
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   24
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Scene_Browser.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Scene_Browser.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Scene_Browser.frx":093E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Scene_Browser.frx":0C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Scene_Browser.frx":0F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Scene_Browser.frx":128C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Scene_Browser.frx":15A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Scene_Browser.frx":18C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Scene_Browser.frx":1BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Scene_Browser.frx":1EF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Scene_Browser.frx":220E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Scene_Browser.frx":2528
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Scene_Browser.frx":2842
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Scene_Browser.frx":2B5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Scene_Browser.frx":2E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Scene_Browser.frx":3190
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Scene_Browser.frx":34AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Scene_Browser.frx":37C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Scene_Browser.frx":3ADE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Scene_Browser.frx":3DF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Scene_Browser.frx":4112
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Scene_Browser.frx":442C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Scene_Browser.frx":4746
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Scene_Browser.frx":4A60
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Scene_Browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private GameEngine As EngineClass
Private EditorClass As VBCEWorkspace_Class

Sub InitializeThis(GEngine As Object, WorkSpaceClass As Object)
    Set GameEngine = GEngine
    Set EditorClass = WorkSpaceClass
End Sub

Private Sub Form_Load()
    'Set GameEngine = Form1.engClass
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    TreeView1.Top = 0
    TreeView1.Left = 0
    TreeView1.Width = Me.Width - 115
    TreeView1.Height = Me.Height - 500 - Me.StatusBar1.Height
End Sub

Private Sub mnuCircle1_Click()
    Dim thisScene As Scene_Class
    Dim newObject As GameObject_Class
    Dim curObject As GameObject_Class
    Dim NewObjectName As String
    Dim selObject As Node
    
    'Dim NewMesh As String
    'NewMesh = vbNewLine & "Sub Draw()" & vbNewLine & "Engine.GameWindow.Circle 2,Me.Position.X ,Me.Position.Y,100,VbRed,Me.Position.X,Me.Position.Y,1" & vbNewLine & "End Sub"
    
    Set selObject = TreeView1.selectedItem
    
    If Not selObject Is Nothing Then
        With selObject
            If .Tag = "Object" Then
                Set thisScene = GameEngine.Scenes(.Parent)
                NewObjectName = InputBox("Child Object Name")
                
                Dim newChild As New GameObject_Class
                'Set newChild.Position = thisScene.Objects(.text).Position
                Set curObject = thisScene.Objects(.text)
                Set newChild.ParentObject = curObject
                newChild.Name = NewObjectName
                curObject.Child.Add newChild, NewObjectName
                
            ElseIf .Tag = "Scene" Then
                Set thisScene = GameEngine.Scenes(.text)
                NewObjectName = InputBox("Object Name")
                
                Set newObject = thisScene.CreateNewObject(NewObjectName, , , 1)
                Set newObject.MyScene = thisScene
                'newObject.ChangeMeshCode NewMesh, Me.GameEngine.CodeEngine
            End If
        End With
        Me.Update
    Else
        MsgBox "Please select valid scene", vbExclamation
    End If
    
End Sub

Private Sub mnuClose1_Click()
    Me.Hide
End Sub

Private Sub mnuEmpty1_Click()
    Dim thisScene As Scene_Class
    Dim newObject As GameObject_Class
    Dim curObject As GameObject_Class
    Dim IGameObject As IObject_Class
    Dim NewObjectName As String
    Dim selObject As Node
    
    Set selObject = TreeView1.selectedItem
    
    If Not selObject Is Nothing Then
        With selObject
            If .Tag = "Object" Then
                Set thisScene = GameEngine.Scenes(.Parent)
                NewObjectName = InputBox("Child Object Name")
                
                Dim newChild As New GameObject_Class
                'Set newChild.Position = thisScene.Objects(TreeView1.selectedItem.text).Position
                Set curObject = thisScene.Objects(TreeView1.selectedItem.text)
                Set newChild.ParentObject = curObject
                newChild.Name = NewObjectName
                curObject.Child.Add newChild, NewObjectName
            ElseIf .Tag = "Scene" Then
                
                Set thisScene = GameEngine.Scenes(TreeView1.selectedItem.text)
                NewObjectName = InputBox("Object Name")
                
                Set newObject = thisScene.CreateNewObject(NewObjectName)
                Set newObject.MyScene = thisScene
                'Set IGameObject = newObject
                'IGameObject.Load Me.GameEngine.CodeEngine
            End If
        End With
        Me.Update
    Else
        MsgBox "Please select valid scene", vbExclamation
    End If
End Sub

Function AddObject(PresetNum As ObjectPresets)
    Dim thisScene As Scene_Class
    Dim newObject As GameObject_Class
    Dim curObject As GameObject_Class
    Dim NewObjectName As String
    Dim selObject As Node
    
    'Dim NewMesh As String
    'NewMesh = vbNewLine & "Sub Draw()" & vbNewLine & "Engine.GameWindow.Circle 2,Me.Position.X ,Me.Position.Y,100,VbRed,Me.Position.X,Me.Position.Y,1" & vbNewLine & "End Sub"
    
    Set selObject = TreeView1.selectedItem
    
    If Not selObject Is Nothing Then
        With selObject
            If .Tag = "Object" Then
                Set thisScene = GameEngine.Scenes(.Parent)
                NewObjectName = InputBox("Child Object Name")
                
                Dim newChild As New GameObject_Class
                'Set newChild.Position = thisScene.Objects(.text).Position
                Set curObject = thisScene.Objects(.text)
                Set newChild.ParentObject = curObject
                newChild.Name = NewObjectName
                curObject.Child.Add newChild, NewObjectName
                
            ElseIf .Tag = "Scene" Then
                Set thisScene = GameEngine.Scenes(.text)
                NewObjectName = InputBox("Object Name")
                
                Set newObject = thisScene.CreateNewObject(NewObjectName, , , PresetNum)
                Set newObject.MyScene = thisScene
                'newObject.ChangeMeshCode NewMesh, Me.GameEngine.CodeEngine
            End If
        End With
        Me.Update
    Else
        MsgBox "Please select valid scene", vbExclamation
    End If
End Function

Private Sub mnuOC1_Click()
    Dim CodeEditorInstance As New CodeEditor
    Dim thisScene As Scene_Class
    
    Dim thisGameObject As GameObject_Class
    Dim selObject As Node
    
    Set selObject = TreeView1.selectedItem
    
    If Not selObject Is Nothing Then
        If selObject.Tag = "Object" Then
            Set thisScene = GameEngine.Scenes(selObject.Parent)
            Set thisGameObject = thisScene.Objects(selObject.text)
            
            'CodeEditorInstance.EditData thisGameObject.MyScript.Data, thisGameObject.Name, "GameObjectCode", Me.GameEngine, thisGameObject
        End If
    Else
    
    End If
End Sub

Sub EditObjectCode(EditData As CodeEditorEnums)
    Dim CodeEditorInstance As New CodeEditor
    Dim thisScene As Scene_Class
    
    Dim thisGameObject As GameObject_Class
    Dim selObject As Node
    
    Set selObject = TreeView1.selectedItem
    
    If Not selObject Is Nothing Then
        If selObject.Tag = "Object" Then
            Set thisScene = GameEngine.Scenes(selObject.Parent)
            Set thisGameObject = thisScene.Objects(selObject.text)
            Me.GameEngine.EditObjectCode EditData, thisGameObject, thisScene
            'CodeEditorInstance.EditData ThisGameObject.MyScript.Data, ThisGameObject.Name, "GameObjectCode", Me.GameEngine, ThisGameObject
        End If
    Else
    
    End If
End Sub

Private Sub mnuOMD1_Click()
    Dim CodeEditorInstance As New CodeEditor
    Dim thisScene As Scene_Class
    
    Dim thisGameObject As GameObject_Class
    Dim selObject As Node
    
    Set selObject = TreeView1.selectedItem
    
    If Not selObject Is Nothing Then
        If selObject.Tag = "Object" Then
            Set thisScene = GameEngine.Scenes(selObject.Parent)
            Set thisGameObject = thisScene.Objects(selObject.text)
            
            'Set CodeEditorInstance.GameObject = ThisGameObject
            
            'CodeEditorInstance.Text1.Text = ThisGameObject.MeshCode
            'CodeEditorInstance.curClassName = ThisGameObject.Name
            'CodeEditorInstance.Caption = ThisGameObject.Name
            'CodeEditorInstance.Session = "MeshData"
            'CodeEditorInstance.Show
        End If
    Else
    
    End If
End Sub

Sub SaveAs()
    Dim selObject As Node
    Dim thisGameObject As GameObject_Class
    Dim thisScene As Scene_Class
    Dim FileName As String
    
    Set selObject = TreeView1.selectedItem
    If selObject Is Nothing Then
        MsgBox "No Items Selected", vbExclamation
        Exit Sub
    End If
    
    FileName = FileBrowser1.SaveFile(MainWindow, selObject.text)
    
    If Len(FileName) = 0 Then
        Exit Sub
    End If
    
    If selObject.Tag = "Object" Then
        Set thisScene = GameEngine.Scenes(selObject.Parent)
        Set thisGameObject = thisScene.Objects(selObject.text)
        Me.GameEngine.SaveObjectToDisk GameObject, FileName, thisGameObject
    ElseIf selObject.Tag = "Scene" Then
        Set thisScene = GameEngine.Scenes(selObject.text)
        Me.GameEngine.SaveObjectToDisk Scene, FileName, thisScene
    Else
    
    End If
End Sub

Sub EditMeshdata()
    Dim CodeEditorInstance As New CodeEditor
    Dim thisScene As Scene_Class
    
    Dim thisGameObject As GameObject_Class
    Dim selObject As Node
    
    Set selObject = TreeView1.selectedItem
    
    If Not selObject Is Nothing Then
        If selObject.Tag = "Object" Then
            Set thisScene = GameEngine.Scenes(selObject.Parent)
            Set thisGameObject = thisScene.Objects(selObject.text)
            
            'Set CodeEditorInstance.GameObject = ThisGameObject
            
            'CodeEditorInstance.Text1.Text = ThisGameObject.MeshCode
            'CodeEditorInstance.curClassName = ThisGameObject.Name
            'CodeEditorInstance.Caption = ThisGameObject.Name
            'CodeEditorInstance.Session = "MeshData"
            'CodeEditorInstance.EditData ThisGameObject.MeshCode, ThisGameObject.Name, "MeshData", Me.GameEngine, ThisGameObject
            'CodeEditorInstance.Show
        End If
    Else
    
    End If
End Sub

Private Sub mnuRefresh1_Click()
    Me.Update
End Sub

Private Sub mnuRemove1_Click()
    Dim thisScene As Scene_Class
    Dim thisGameObject As GameObject_Class
    Dim selObject As Node
    
    Set selObject = TreeView1.selectedItem
    
    If Not selObject Is Nothing Then
        With selObject
            If .Tag = "Object" Then
                Set thisScene = GameEngine.Scenes(.Parent)
                Set thisGameObject = thisScene.Objects(.text)
                thisGameObject.EndObject
                Me.Update
            ElseIf .Tag = "Scene" Then
                'If Form1.engClass.Scenes.Count > 1 Then
                    GameEngine.Scenes.Remove .text
                    Me.Update
                'Else
                    'MsgBox "Base Scenes cannot be removed", vbExclamation, "Scene_Browser"
                'End If
            End If
        End With
    End If
End Sub

Sub RemoveItem()
    Dim thisScene As Scene_Class
    Dim thisGameObject As GameObject_Class
    Dim selObject As Node
    
    Set selObject = TreeView1.selectedItem
    
    If Not selObject Is Nothing Then
        With selObject
            If .Tag = "Object" Then
                Set thisScene = GameEngine.Scenes(.Parent)
                Set thisGameObject = thisScene.Objects(.text)
                thisGameObject.EndObject
                Me.Update
            ElseIf .Tag = "Scene" Then
                'If Form1.engClass.Scenes.Count > 1 Then
                    GameEngine.Scenes.Remove .text
                    Me.Update
                'Else
                    'MsgBox "Base Scenes cannot be removed", vbExclamation, "Scene_Browser"
                'End If
            End If
        End With
    End If
End Sub

Private Sub mnuScene1_Click()
    Dim addName As String
    Dim AddScene As Scene_Class
    
    addName = InputBox("Scene Name")
    If Len(addName) = 0 Then
        Exit Sub
    End If
    GameEngine.AddScene addName
    Me.Update
End Sub

Private Sub TreeView1_AfterLabelEdit(Cancel As Integer, NewString As String)
    If Cancel = 0 Then
        If TreeView1.selectedItem.Tag = "Scene" Then
        
        ElseIf TreeView1.selectedItem.Tag = "Object" Then
        
        End If
    End If
End Sub

Private Sub TreeView1_Click()
    On Error Resume Next
    If TreeView1.selectedItem.Tag = "Scene" Then
        StatusBar1.Panels(1).text = "Item Type: Scene"
        StatusBar1.Panels(2).text = "Item Class: Scene_Class"
        MainWindow.mnuAddObject1.Caption = "Add Object To Scene " & TreeView1.selectedItem
        
        MainWindow.mnuSaveW1.Caption = "Save " & TreeView1.selectedItem.text
        MainWindow.mnuSaveAs1.Caption = "Save " & TreeView1.selectedItem.text & " As..."
        'Me.mnuOMD1.Enabled = False
        'Me.mnuOC1.Enabled = False
    ElseIf TreeView1.selectedItem.Tag = "Object" Then
        StatusBar1.Panels(1).text = "Item Type: GameObject"
        StatusBar1.Panels(2).text = "Item Class: GameObject_Class"
        'Me.mnuOMD1.Enabled = True
        MainWindow.mnuSaveW1.Caption = "Save " & TreeView1.selectedItem.text
        MainWindow.mnuSaveAs1.Caption = "Save " & TreeView1.selectedItem.text & " As..."
        'Me.mnuOC1.Enabled = True
    Else
        StatusBar1.Panels(1).text = "Item Type: None"
        StatusBar1.Panels(2).text = "Item Class: None"
        'Me.mnuOMD1.Enabled = False
        'Me.mnuOC1.Enabled = False
    End If
    MainWindow.mnuRemove1.Caption = "Remove " & TreeView1.selectedItem.Tag & " " & TreeView1.selectedItem.text
End Sub

Function Update()
    UpdateSceneBrowser
End Function

Function UpdateSceneBrowser()
    EditorClass.UpdateSceneBrowser Me.TreeView1
End Function

Private Sub TreeView1_DblClick()
    Dim selObject As GameObject_Class
    Dim thisScene As Scene_Class
    On Error Resume Next
    If TreeView1.selectedItem.Tag = "Scene" Then
    ElseIf TreeView1.selectedItem.Tag = "Object" Then
        Set thisScene = GameEngine.Scenes.Item(TreeView1.selectedItem.Parent)
        Set selObject = thisScene.Objects.Item(TreeView1.selectedItem.text)
        GameEngine.SelectObject selObject
    End If
End Sub
