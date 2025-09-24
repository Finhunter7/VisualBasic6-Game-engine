VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form2 
   Caption         =   "Project Browser"
   ClientHeight    =   4830
   ClientLeft      =   165
   ClientTop       =   210
   ClientWidth     =   5790
   Icon            =   "TreeViewTest1.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   5790
   Tag             =   "ProjectBrowser"
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4575
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   0
            Text            =   "Item: None"
            TextSave        =   "Item: None"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   0
            Text            =   "Type: Test_Class "
            TextSave        =   "Type: Test_Class "
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   0
            Text            =   "In Scene:"
            TextSave        =   "In Scene:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5530
      _Version        =   327682
      Indentation     =   529
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   4440
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":0442
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5160
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   37
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":0984
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":0C9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":0FB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":12D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":15EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":1906
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":1C20
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":1F3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":2254
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":256E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":2888
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":2BA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":2EBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":31D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":34F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":380A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":3B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":3E3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":4158
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":4472
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":478C
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":4AA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":4DC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":50DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":53F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":5506
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":5618
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":572A
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":583C
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":594E
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":5C68
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":5F82
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":629C
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":65B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":68D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":6BEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeViewTest1.frx":6F04
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public GameEngine As EngineClass

Private Sub Form_Load()
    'Me.CreateProjectTree
    Me.UpdateProjectBrowser
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    TreeView1.Top = 0
    TreeView1.Left = 0
    TreeView1.Width = Me.Width - 100
    TreeView1.Height = Me.Height - 500 - Me.StatusBar1.Height
End Sub

Private Sub mnuClose1_Click()
    Me.Hide
End Sub

Private Sub mnuFolder1_Click()
    Dim Name As String
    Dim thisNode As Node
    Name = InputBox("Folder Name")

    Set thisNode = TreeView1.Nodes.Add(TreeView1.selectedItem.Key, 4, "F_" & Name, Name, 1, 1)
    thisNode.ExpandedImage = 2

End Sub

Private Sub mnuInfo1_Click()
    Dim This_Object As Object
    If TreeView1.selectedItem Is Nothing Then
        Exit Sub
    End If
    
    With TreeView1.selectedItem
        Select Case .Tag
            Case "Script"
               Set This_Object = Me.GameEngine.Scripts.Item(.text)
               
            Case "Scene"
                Set This_Object = Me.GameEngine.ScenesInactive.Item(.text)
                MsgBox This_Object.Name
            Case "CodeEngine"
                
        End Select
    End With
End Sub

Private Sub mnuInputSys1_Click()
    Dim Name As String
    Dim thisNode As Node

    Name = InputBox("InputSystem Name")

    Set thisNode = TreeView1.Nodes.Add(TreeView1.selectedItem.Key, 4, Name, Name, 10, 10)
    thisNode.Tag = "InputSystem"
End Sub

Private Sub mnuRefresh1_Click()
    Me.UpdateProjectBrowser
End Sub

Private Sub mnuRemove1_Click()
    
End Sub

Sub RemoveItem()
    If TreeView1.selectedItem Is Nothing Then
        Exit Sub
    End If
    
    With TreeView1.selectedItem
        Select Case .Tag
            Case "Script"
                Me.GameEngine.RemoveScript .text
            Case "Scene"
                Me.GameEngine.UnloadScene .text
            Case "CodeEngine"
                MsgBox "CodeEngine Cannot Be Removed", vbExclamation
            Case "GameObject"
                Me.GameEngine.UnloadGameObject .text
        End Select
    End With
    Me.Update
End Sub

Sub SaveAs()
    Dim selObject As Node
    Dim thisGameObject As GameObject_Class
    Dim thisScene As Scene_Class
    Dim thisScript As Script_Class
    Dim Filename As String
    
    Set selObject = TreeView1.selectedItem
    If selObject Is Nothing Then
        MsgBox "No Items Selected", vbExclamation
        Exit Sub
    End If

    If selObject.Tag = "Scene" Then
        Filename = FileBrowser1.SaveFile(MainWindow, selObject.text, "*.VBScene")
        If Len(Filename) = 0 Then
            Exit Sub
        End If
        Set thisScene = GameEngine.ScenesInactive(selObject.text)
        Me.GameEngine.SaveObjectToDisk Scene, Filename, thisScene
    ElseIf selObject.Tag = "Script" Then
        Filename = FileBrowser1.SaveFile(MainWindow, selObject.text, "*.VBCEScript")
        If Len(Filename) = 0 Then
            Exit Sub
        End If
        Set thisScript = GameEngine.Scripts(selObject.text)
        Me.GameEngine.SaveObjectToDisk Script, Filename, thisScript
    ElseIf selObject.Tag = "GameObject" Then
        Filename = FileBrowser1.SaveFile(MainWindow, selObject.text, "*.VBGObject")
        If Len(Filename) = 0 Then
            Exit Sub
        End If
        Set thisGameObject = GameEngine.GameObjects(selObject.text)
        Me.GameEngine.SaveObjectToDisk GameObject, Filename, thisGameObject
    Else
        MsgBox selObject.text & " Item Cannot Be Save On Disk", vbExclamation
    End If
End Sub

Private Sub mnuScene1_Click()
    Dim Name As String
    'Dim thisNode As Node
    Name = InputBox("Scene Name")
    If Len(Name) < 1 Then
        Exit Sub
    End If
    'Set thisNode = TreeView1.Nodes.Add(TreeView1.SelectedItem.Key, 4, Name, Name, 17, 17)
    'thisNode.Tag = "Scene"
    GameEngine.CreateNewScene Name
    Me.UpdateProjectBrowser
End Sub

Private Sub mnuScript1_Click()

End Sub

Private Sub TreeView1_AfterLabelEdit(Cancel As Integer, NewString As String)
    Dim thisNode As Node
    Dim oldName As String
    If Cancel = 0 Then
        Set thisNode = TreeView1.selectedItem
        With thisNode
            If thisNode Is Nothing Then
                Exit Sub
            Else
                oldName = .text
                If .Tag = "Script" Then
                    
                ElseIf .Tag = "Scene" Then
                    Me.GameEngine.RenameScene .text, NewString
                ElseIf .Tag = "Project" Then
                    Me.GameEngine.ProjectName = NewString
                End If
            End If
        
        End With
    Else

    End If
    'Me.Update
End Sub

Private Sub TreeView1_Click()
    Dim SelGObject As GameObject_Class
    If TreeView1.selectedItem Is Nothing Then
        Exit Sub
    End If
    
    With TreeView1.selectedItem
    Me.StatusBar1.Panels(1).text = "Item Type: " & .Tag
    Me.StatusBar1.Panels(3).text = ""
        'Select Case .Tag
        If .Tag = "Script" Then
            Me.StatusBar1.Panels(2).text = "Class: Script_Class"
            MainWindow.mnuRemove1.Caption = "Remove " & .Tag & " " & .text
            MainWindow.mnuSaveW1.Caption = "Save " & TreeView1.selectedItem.text
            MainWindow.mnuSaveAs1.Caption = "Save " & TreeView1.selectedItem.text & " As..."
        ElseIf .Tag = "Scene" Then
            Me.StatusBar1.Panels(2).text = "Class: Scene_Class"
            MainWindow.mnuRemove1.Caption = "Remove " & .Tag & " " & .text
            
            MainWindow.mnuSaveW1.Caption = "Save " & TreeView1.selectedItem.text
            MainWindow.mnuSaveAs1.Caption = "Save " & TreeView1.selectedItem.text & " As..."
            
        ElseIf .Tag = "CodeEngine" Then
            Me.StatusBar1.Panels(2).text = "Class: MSScriptControl"
        ElseIf .Tag = "GameObject" Then
            Me.StatusBar1.Panels(2).text = "Class: GameObject_Class"
            Set SelGObject = Me.GameEngine.GameObjects.Item(.text)
            If Not SelGObject.MyScene Is Nothing Then
                Me.StatusBar1.Panels(3).text = "In Scene: " & SelGObject.MyScene.Name
            End If
        Else
        End If
        'End Select
    End With
End Sub

Private Sub TreeView1_DblClick()
    
    If TreeView1.selectedItem Is Nothing Then
        Exit Sub
    End If
    
    With TreeView1.selectedItem
        If .Tag = "Script" Then
            Me.GameEngine.EditScript .text
        ElseIf .Tag = "Scene" Then
    
        ElseIf .Tag = "Module" Then
            'Me.EditModule Me.GameEngine.CodeEngine.Modules(.Key)
        ElseIf .Tag = "MeshCode" Then
            
        End If
    End With
    
End Sub

Function Update()
    UpdateProjectBrowser
End Function

Function UpdateProjectBrowser()
    Me.GameEngine.BrowseEngineData Me.TreeView1
End Function

Private Sub TreeView1_OLEStartDrag(Data As ComctlLib.DataObject, AllowedEffects As Long)
    If Not Me.TreeView1.selectedItem Is Nothing Then
        If Me.TreeView1.selectedItem.Tag = "GameObject" Then
            'Data.SetData Me.TreeView1.selectedItem.text, 5
        Else
        
        End If
    End If
End Sub
