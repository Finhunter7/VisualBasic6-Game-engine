VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MainWindow 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   9315
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   14205
   Icon            =   "MainWindow.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   14205
      _ExtentX        =   25056
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   3
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "TNew1"
            Object.ToolTipText     =   "New"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "TOpen1"
            Object.ToolTipText     =   "Open Active Window"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "TSave1"
            Object.ToolTipText     =   "Save Active Window"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   9015
      Width           =   14205
      _ExtentX        =   25056
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   0
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox SolutionBrowserBar1 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   8595
      Left            =   10575
      Negotiate       =   -1  'True
      ScaleHeight     =   8595
      ScaleWidth      =   3630
      TabIndex        =   0
      Top             =   420
      Width           =   3630
      Begin ComctlLib.Toolbar ToolbarRightCaptionBar1 
         Height          =   390
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         ImageList       =   "ImageList1"
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   9
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "SBTNew1"
               Object.Tag             =   ""
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "SBTOpen1"
               Object.Tag             =   ""
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "SBTSave1"
               Object.Tag             =   ""
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "SBTRemove1"
               Object.Tag             =   ""
               ImageIndex      =   14
            EndProperty
            BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "SBTFolderUp1"
               Object.Tag             =   ""
               ImageIndex      =   15
            EndProperty
            BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               Object.Width           =   1e-4
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "SBTFront1"
               Object.Tag             =   ""
               ImageIndex      =   16
            EndProperty
            BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "SBTBack1"
               Object.Tag             =   ""
               ImageIndex      =   17
            EndProperty
         EndProperty
      End
      Begin ComctlLib.TreeView TreeViewBrowser2 
         Height          =   1095
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1931
         _Version        =   327682
         Indentation     =   529
         LineStyle       =   1
         Style           =   7
         ImageList       =   "TreeImageList1"
         Appearance      =   1
         OLEDropMode     =   1
      End
      Begin ComctlLib.TabStrip SolutionBrowser1 
         Height          =   2295
         Left            =   0
         TabIndex        =   1
         Top             =   360
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   4048
         MultiRow        =   -1  'True
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   6
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Solution"
               Key             =   "TSolution"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Project-Browser"
               Key             =   "TProjectB"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Scene-Browser"
               Key             =   "TSceneBrowser"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Object-Browser"
               Key             =   "TObjectB"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Folder-Browser"
               Key             =   "TFolderB1"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Component-Browser"
               Key             =   "TComB1"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6000
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   18
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":0984
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":0EC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":1408
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":151A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":162C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":173E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":1850
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":1D92
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":1EA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":1FB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":20C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":21D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":2714
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":281E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":2930
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":2A42
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":2B54
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList TreeImageList1 
      Left            =   6000
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   38
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":2C66
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":2F80
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":329A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":35B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":38CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":3BE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":3F02
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":421C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":4536
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":4850
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":4B6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":4E84
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":519E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":54B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":57D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":5AEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":5E06
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":6120
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":643A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":6754
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":6A6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":6D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":70A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":73BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":76D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":77E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":78FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":7A0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":7B1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":7C30
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":7F4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":8264
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":857E
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":8898
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":8BB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":8ECC
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":91E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainWindow.frx":92F0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile1 
      Caption         =   "File"
      Begin VB.Menu mnuNewProject1 
         Caption         =   "New Project"
      End
      Begin VB.Menu mnuOpenProject1 
         Caption         =   "Open Project..."
      End
      Begin VB.Menu mnuSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddProject1 
         Caption         =   "Add Project..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCloseProject1 
         Caption         =   "Close Project"
      End
      Begin VB.Menu mnuSpace11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveProject1 
         Caption         =   "Save Project..."
      End
      Begin VB.Menu mnuSaveProjectAs1 
         Caption         =   "Save Project As..."
      End
      Begin VB.Menu mnuSpace3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveW1 
         Caption         =   "Save"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSaveAs1 
         Caption         =   "Save As..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuspace9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBuildProject1 
         Caption         =   "Build Project..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuRenderTo1 
         Caption         =   "Render Project..."
      End
      Begin VB.Menu mnuSpace5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit1 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit1 
      Caption         =   "Edit"
      Begin VB.Menu mnuObjectCode1 
         Caption         =   "Object Code"
      End
      Begin VB.Menu mnuObjectComponents1 
         Caption         =   "Object Components"
      End
   End
   Begin VB.Menu mnuView1 
      Caption         =   "View"
      Begin VB.Menu mnuRefresh1 
         Caption         =   "Refresh ActiveWindow"
      End
      Begin VB.Menu mnuSpace8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolbars1 
         Caption         =   "Toolbar"
         Begin VB.Menu mnuEngineBrowser1 
            Caption         =   "Engine Browser"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuInspector1 
            Caption         =   "Inspector"
         End
      End
      Begin VB.Menu mnuspace10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewPort1 
         Caption         =   "View-Port Window..."
      End
      Begin VB.Menu mnuProjectBrowser1 
         Caption         =   "Project-Browser..."
      End
      Begin VB.Menu mnuSceneBrowser1 
         Caption         =   "Scene-Browser..."
      End
      Begin VB.Menu mnuCodeEditor1 
         Caption         =   "Code-Editor..."
      End
      Begin VB.Menu mnuConsole1 
         Caption         =   "Console..."
      End
      Begin VB.Menu mnuErrors1 
         Caption         =   "Errors..."
      End
   End
   Begin VB.Menu mnuProject 
      Caption         =   "Project"
      Begin VB.Menu mnuProgram1 
         Caption         =   "Program"
         Begin VB.Menu mnuAddForm1 
            Caption         =   "Add Form"
         End
      End
      Begin VB.Menu mnuScripts1 
         Caption         =   "Scripts"
         Begin VB.Menu mnuPresets1 
            Caption         =   "Presets"
            Begin VB.Menu mnuExPlayerMScript1 
               Caption         =   "Player Movement Script"
            End
            Begin VB.Menu mnuExCCompScript1 
               Caption         =   "Custom Component Script"
            End
         End
         Begin VB.Menu mnuSpace13 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAddScript1 
            Caption         =   "Add Script"
         End
         Begin VB.Menu mnuAddMeshScript1 
            Caption         =   "Add Mesh Script"
         End
         Begin VB.Menu mnuAddComScript1 
            Caption         =   "Add Component Script"
         End
      End
      Begin VB.Menu mnuScenes1 
         Caption         =   "Scenes"
         Begin VB.Menu mnuScene1 
            Caption         =   "Add Scene"
         End
         Begin VB.Menu mnuAcScene1 
            Caption         =   "Activate Scene..."
         End
         Begin VB.Menu mnuAddObject1 
            Caption         =   "Add Object (Current Scene)"
            Begin VB.Menu mnuEmpty1 
               Caption         =   "Empty"
            End
            Begin VB.Menu mnuSpace7 
               Caption         =   "-"
            End
            Begin VB.Menu mnuPlane1 
               Caption         =   "Plane"
            End
            Begin VB.Menu mnuCube1 
               Caption         =   "Cube"
            End
            Begin VB.Menu mnuCircle1 
               Caption         =   "Circle"
            End
            Begin VB.Menu mnuPicture1 
               Caption         =   "Picture"
            End
         End
         Begin VB.Menu mnuAddExistingObj1 
            Caption         =   "Add Existing Object (Current Scene)"
         End
      End
      Begin VB.Menu mnuAnimsAndEffects1 
         Caption         =   "Animation And Effects"
         Begin VB.Menu mnuAddEffect1 
            Caption         =   "Add Effect..."
         End
         Begin VB.Menu mnuAddAnimation1 
            Caption         =   "Add Animation..."
         End
      End
      Begin VB.Menu mnuAddFile1 
         Caption         =   "Add File..."
      End
      Begin VB.Menu mnuSpace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemove1 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuSpace6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefrences1 
         Caption         =   "Engine Refrences..."
      End
      Begin VB.Menu mnuComponents1 
         Caption         =   "Components..."
      End
      Begin VB.Menu mnuSpace4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProjectProp1 
         Caption         =   "Project Propetites..."
      End
   End
   Begin VB.Menu mnuFormat1 
      Caption         =   "Format"
   End
   Begin VB.Menu mnuDebug1 
      Caption         =   "Debug"
   End
   Begin VB.Menu mnuRun1 
      Caption         =   "Run"
      Begin VB.Menu mnuStart1 
         Caption         =   "Start Engine"
      End
      Begin VB.Menu mnuStartInNewWindow1 
         Caption         =   "Start In Window"
      End
      Begin VB.Menu mnuResume1 
         Caption         =   "Resume Engine"
      End
      Begin VB.Menu mnuPause1 
         Caption         =   "Pause Engine"
      End
      Begin VB.Menu mnuStop1 
         Caption         =   "Stop Engine"
      End
      Begin VB.Menu mnuRender1 
         Caption         =   "Render Frames"
      End
   End
   Begin VB.Menu mnuWindow1 
      Caption         =   "Window"
      Begin VB.Menu mnuCascade1 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuTitleHoriz1 
         Caption         =   "Tile Horizontal"
      End
      Begin VB.Menu mnuTileVertic1 
         Caption         =   "Tile Vertical"
      End
      Begin VB.Menu mnuArrangeIcons1 
         Caption         =   "Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp1 
      Caption         =   "Help"
   End
   Begin VB.Menu mnuTest1 
      Caption         =   "Test"
      Visible         =   0   'False
   End
   Begin VB.Menu Browsermnu1 
      Caption         =   "Browser"
      Visible         =   0   'False
      Begin VB.Menu BmnuCreateNew1 
         Caption         =   "Create New"
         Begin VB.Menu BmnuScript1 
            Caption         =   "Script"
         End
         Begin VB.Menu BmnuCRef1 
            Caption         =   "Reference"
         End
         Begin VB.Menu BmnuCComp1 
            Caption         =   "Component"
         End
         Begin VB.Menu BmnuSpace2 
            Caption         =   "-"
         End
         Begin VB.Menu BmnuCScene1 
            Caption         =   "Scene"
         End
      End
      Begin VB.Menu BmnuAddNew1 
         Caption         =   "Add New From File"
         Enabled         =   0   'False
         Begin VB.Menu BmnuAddScript1 
            Caption         =   "Script"
         End
         Begin VB.Menu BmnuAddReference1 
            Caption         =   "Reference"
         End
         Begin VB.Menu BmnuComponent1 
            Caption         =   "Component"
         End
         Begin VB.Menu BmnuSpace1 
            Caption         =   "-"
         End
         Begin VB.Menu BmnuAddScene1 
            Caption         =   "Scene"
         End
      End
      Begin VB.Menu BmnuRemove1 
         Caption         =   "Remove"
      End
      Begin VB.Menu BmnuSpace3 
         Caption         =   "-"
      End
      Begin VB.Menu BmnuPropertites1 
         Caption         =   "Propertites"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents Engine As EngineClass
Attribute Engine.VB_VarHelpID = -1
Private EditorClass As New VBCEWorkspace_Class

Private Sub List1_Click()

End Sub

Private Sub BmnuCScene1_Click()
    Engine.CreateNewScene InputBox("Name", "Scene Name")
End Sub

Private Sub BmnuRemove1_Click()
    With Me.SolutionBrowser1.selectedItem
        If .Key = "TProjectB" Then
            EditorClass.TreeViewBrowserHandleActions Me.TreeViewBrowser2, ProjectBrowser, RemoveItem
        ElseIf .Key = "TSceneBrowser" Then
            EditorClass.TreeViewBrowserHandleActions Me.TreeViewBrowser2, SceneBrowser, RemoveItem
        End If
    End With
End Sub

Private Sub BmnuScript1_Click()
    Engine.CreateNewScript InputBox("Name", "Create Script"), DefaultScript
End Sub

Private Sub Engine_OnDataChanged()
    'ProjectChanged = True
End Sub

Private Sub Engine_OnGamePause()
    Me.Caption = Me.Engine.ProjectName & " - " & "Visual Basic Code Engine [Paused]"
End Sub

Private Sub Engine_OnGameResume()
    Me.Caption = Me.Engine.ProjectName & " - " & "Visual Basic Code Engine [run]"
End Sub

Private Sub Engine_OnGameStart()
    Me.Caption = Me.Engine.ProjectName & " - " & "Visual Basic Code Engine [run]"
End Sub

Private Sub Engine_OnGameStop()
    Me.Caption = Me.Engine.ProjectName & " - " & "Visual Basic Code Engine [design]"
End Sub

Private Sub MDIForm_Load()
    Set Engine = New EngineClass
    OnLoad
End Sub

Private Sub OnLoad()
    EditorClass.LoadDeveloperTools Engine
    Form2.InitializeThis Engine, EditorClass
    Scene_Browser.InitializeThis Engine, EditorClass
    
    Form1.Show
    ProjectTypeSelectorDialog.Show vbModal, Me
    Engine.LoadEngine True, Form1, Console, VBCEGameProject, Form1, Nothing
    Set ActiveWindow = Me.ActiveForm
    Me.Caption = Me.Engine.ProjectName & " - " & "Visual Basic Code Engine [design]"
End Sub

Private Sub mnuAcScene1_Click()
    If Engine.ScenesInactive.Count > 0 Then
        SceneAcWindow.ShowScenes Me, Me.Engine
    Else
        MsgBox "No Scenes To Activate", vbExclamation
    End If
End Sub

Private Sub mnuAddComScript1_Click()
    Engine.CreateNewScript InputBox("Script Name"), ComponentScript
End Sub

Private Sub mnuAddFile1_Click()
    Dim curFile As String
    Dim TScene As Scene_Class
    Dim curObject As GameObject_Class
    curFile = FileBrowser1.OpenFile(Me, "Open", "*.*|*.VBGObject|*.VBScene|*.VBCEProject")
    If FileBrowser1.GetFormat() = "VBScene" Then
        Set TScene = Me.Engine.SaveLoadGameClass.LoadObjectFromDiskVBCFormat(Scene, curFile)
        Me.Engine.ScenesInactive.Add TScene, TScene.Name
    ElseIf FileBrowser1.GetFormat() = "VBGObject" Then
        Set TScene = Me.Engine.GetCurrentScene()
        Set curObject = Me.Engine.SaveLoadGameClass.LoadObjectFromDiskVBCFormat(GameObject, curFile)
        'TScene.AddObject Me.Engine, curObject, curObject.Name
    End If
End Sub

Private Sub mnuAddForm1_Click()
    Engine.Forms.Add New NewForm
End Sub

Private Sub mnuAddMeshScript1_Click()
    Engine.CreateNewScript InputBox("Script Name"), MeshScript
End Sub

Private Sub mnuAddScript1_Click()
    Engine.CreateNewScript InputBox("Script Name"), DefaultScript
End Sub

Private Sub mnuBox1_Click()
    On Error Resume Next
    If Me.ActiveForm.Tag = "SceneBrowser" Then
        Me.ActiveForm.AddObject 2
    Else
        Me.Engine.SceneAddObject 2
    End If
End Sub

Private Sub mnuArrangeIcons1_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuCascade1_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuCircle1_Click()
    On Error Resume Next
    If Me.ActiveForm.Tag = "SceneBrowser" Then
        Me.ActiveForm.AddObject SCircle
    Else
        Me.Engine.SceneAddObject SCircle
    End If
End Sub

Private Sub mnuCloseProject1_Click()
    If Not Me.Engine Is Nothing Then
        If MsgBox("Are you sure", vbYesNo, "Close Project") = vbYes Then
            Set Engine = Nothing
        End If
    End If
End Sub

Private Sub mnuCodeEditor1_Click()
    'Dim CodeEditorInstance As New CodeEditor
    
    'CodeEditorInstance.Show
End Sub

Private Sub mnuConsole1_Click()
    Engine.EConsole.Show
    Engine.EConsole.DisplayMessage
    If Engine.EConsole.WindowState = 1 Then
        Engine.EConsole.WindowState = 0
    End If
End Sub

Private Sub mnuCube1_Click()
    On Error Resume Next
    If Me.ActiveForm.Tag = "SceneBrowser" Then
        Me.ActiveForm.AddObject Cube
    Else
        Me.Engine.SceneAddObject Cube
    End If
End Sub

Private Sub mnuEmpty1_Click()
    On Error Resume Next
    If Me.ActiveForm.Tag = "SceneBrowser" Then
        Me.ActiveForm.AddObject SEmpty
    Else
        Me.Engine.SceneAddObject SEmpty
    End If
End Sub

Private Sub mnuEngineBrowser1_Click()
    If SolutionBrowserBar1.Visible = True Then
        mnuEngineBrowser1.Checked = False
        SolutionBrowserBar1.Visible = False
    Else
        mnuEngineBrowser1.Checked = True
        SolutionBrowserBar1.Visible = True
    End If
End Sub

Private Sub mnuExCCompScript1_Click()
    Engine.CreateNewScript InputBox("Script Name"), ComponentScript, vbNewLine & _
    "Class CustomComponent" & vbNewLine & vbNewLine & _
    "Public Name" & vbNewLine & _
    "Public self" & vbNewLine & vbNewLine & _
    "Private Sub Class_Initialize()" & vbNewLine & _
    "'Name = " & vbNewLine & _
    "End Sub" & vbNewLine & vbNewLine & _
    "Private Sub Class_Terminate()" & vbNewLine & _
    vbNewLine & _
    "End Sub" & vbNewLine & vbNewLine & _
    "End Class" & vbNewLine & _
    "Set Component = New CustomComponent"
End Sub

Private Sub mnuExit1_Click()
    Unload Me
    End
End Sub

Private Sub SetAcForm()
    
End Sub

Private Sub mnuNewProject1_Click()
    If MsgBox("Make new project", vbYesNo) = vbYes Then
        Set Engine = Nothing
        Set Engine = New EngineClass
        OnLoad
    Else
    
    End If
    'Engine.ClearGameEngineData
    
    'Do Until Me.ActiveForm Is Nothing
    'Unload Me.ActiveForm
    'Loop
End Sub

Private Sub mnuObjectCode1_Click()
    Dim curObject As GameObject_Class
    On Error Resume Next
    Set curObject = Me.Engine.GetSelectedObject()
    If Me.ActiveForm.Tag = "ViewPort" Then
        If Not curObject Is Nothing Then
            Me.Engine.EditObjectCode ObjectCode, curObject, curObject.MyScene
        Else
            MsgBox "No Objects Selected Or Object Does Not Exist", vbExclamation
        End If
    ElseIf Me.ActiveForm.Tag = "SceneBrowser" Then
        Me.ActiveForm.EditObjectCode ObjectCode
    Else
        MsgBox "Current Window Does Not Support This Method", vbExclamation
    End If
End Sub

Private Sub mnuObjectMesh1_Click()
    Dim curObject As GameObject_Class
    On Error Resume Next
    
    Set curObject = Me.Engine.GetSelectedObject()
    
    If Me.ActiveForm.Tag = "ViewPort" Then
        If Not curObject Is Nothing Then
            For Each Item In curObject.Components
                MsgBox "Component " & Item.Name
            Next
        Else
            MsgBox "No Objects Selected Or Object Does Not Exist", vbExclamation
        End If
    ElseIf Me.ActiveForm.Tag = "SceneBrowser" Then
        Me.ActiveForm.EditObjectCode MeshData
    Else
        MsgBox "Current Window Does Not Support This Method", vbExclamation
    End If
End Sub

Private Sub mnuOpenProject1_Click()
    Engine.OpenProject FileBrowser1.OpenFile(Me), False
End Sub

Private Sub mnuPause1_Click()
    If Me.Engine.IsEngineRunning = True Then
        Me.Engine.PauseGame
    End If
End Sub

Private Sub mnuPlane1_Click()
    On Error Resume Next
    If Me.ActiveForm.Tag = "SceneBrowser" Then
        Me.ActiveForm.AddObject SPlane
    Else
        Me.Engine.SceneAddObject SPlane
    End If
End Sub

Private Sub mnuProjectBrowser1_Click()
    Form2.Show
    Form2.Update
    If Form2.WindowState = 1 Then
        Form2.WindowState = 0
    End If
    SetAcForm
End Sub

Private Sub mnuProjectProp1_Click()
    'EngineSettings.Show vbModal, Me
End Sub

Private Sub mnuRefrences1_Click()
    EngineRefrences.RefrencesShow Me, Me.Engine
End Sub

Private Sub mnuRefresh1_Click()
    On Error Resume Next
    Me.ActiveForm.Update
End Sub

Private Sub mnuRemove1_Click()
    On Error Resume Next
    If Me.ActiveForm.Tag = "ViewPort" Or Me.ActiveForm.Tag = "CodeEditor" Then
        
    Else
        Me.ActiveForm.RemoveItem
    End If
End Sub

Private Sub mnuRender1_Click()
    Engine.RenderScene InputBox("Render Frames")
    Form1.Show
End Sub

Private Sub mnuRenderTo1_Click()
    Engine.RenderScene InputBox("Render Frames"), FileBrowser1.SaveToDir(Me)
End Sub

Private Sub mnuResume1_Click()
    If Me.Engine.IsEngineRunning = True Then
        Me.Engine.ResumeGame
    End If
End Sub

Private Sub mnuSaveAs1_Click()
    On Error Resume Next
    If Me.ActiveForm.Tag = "SceneBrowser" Or Me.ActiveForm.Tag = "ProjectBrowser" Or Me.ActiveForm.Tag = "CodeEditor" Then
        Me.ActiveForm.SaveAs
    Else
        MsgBox "Current Window Does Not Support This Method", vbExclamation
    End If
End Sub

Private Sub mnuSaveProject1_Click()
    'Engine.SaveProject
    EditorClass.SaveProject FileBrowser1.SaveToDir(Me)
End Sub

Private Sub mnuSaveW1_Click()
    On Error Resume Next
    Me.ActiveForm.Save
End Sub

Private Sub mnuScene1_Click()
    Engine.CreateNewScene InputBox("Scene Name")
    Engine.ProjectBrowser.Update
End Sub

Private Sub mnuSceneBrowser1_Click()
    Scene_Browser.Show
    Scene_Browser.Update
    If Scene_Browser.WindowState = 1 Then
        Scene_Browser.WindowState = 0
    End If
    SetAcForm
End Sub

Private Sub mnuStart1_Click()
    If Not Engine.IsEngineRunning Then
        Engine.RunGameInEditor = True
        Engine.StartEngine
    End If
End Sub

Private Sub mnuStartInNewWindow1_Click()
    If Not Engine.IsEngineRunning Then
        Engine.RunGameInEditor = False
        Engine.StartEngine
    End If
End Sub

Private Sub mnuStop1_Click()
    If Me.Engine.IsEngineRunning Then
        Me.Engine.EndGame
    End If
End Sub

Private Sub mnuTest1_Click()
    'Me.Engine.SaveLoadGameClass.SaveObjectToDiskXML Scene, "C:\tmp\Test.xml", Me.Engine.GetCurrentScene()
    ModelEditor1.Show
End Sub

Private Sub mnuTileVertic1_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuTitleHoriz1_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuViewPort1_Click()
    Engine.EditorWindow.Show
    SetAcForm
End Sub

Private Sub ToolBarRight1_Resize()
    
End Sub

Private Sub SolutionBrowser1_Click()
    Me.TreeViewBrowser2.Nodes.Clear
    
    Me.ToolbarRightCaptionBar1.Buttons(1).ToolTipText = "Create New Item"
    Me.ToolbarRightCaptionBar1.Buttons(2).ToolTipText = "Open File From Disk"
    Me.ToolbarRightCaptionBar1.Buttons(3).ToolTipText = "Save Item To Disk"
    Me.ToolbarRightCaptionBar1.Buttons(4).ToolTipText = "Remove Item From Project"
    Me.ToolbarRightCaptionBar1.Buttons(8).Visible = False
    Me.ToolbarRightCaptionBar1.Buttons(9).Visible = False
    
    If Me.SolutionBrowser1.selectedItem.Key = "TSceneBrowser" Then
        EditorClass.UpdateSceneBrowser Me.TreeViewBrowser2
        Me.ToolbarRightCaptionBar1.Buttons(1).ToolTipText = "Create New Object"
        Me.ToolbarRightCaptionBar1.Buttons(2).ToolTipText = "Open Object/Scene From Disk"
        Me.ToolbarRightCaptionBar1.Buttons(3).ToolTipText = "Save Object/Scene To Disk"
        Me.ToolbarRightCaptionBar1.Buttons(4).ToolTipText = "Make Scene Inactive Or Remove Object From Scene"
        Me.ToolbarRightCaptionBar1.Buttons(8).ToolTipText = "Change Drawing Priority Of Scene/Object To Front"
        Me.ToolbarRightCaptionBar1.Buttons(9).ToolTipText = "Change Drawing Priority Of Scene/Object To Back"
        Me.ToolbarRightCaptionBar1.Buttons(8).Visible = True
        Me.ToolbarRightCaptionBar1.Buttons(9).Visible = True
            
    ElseIf Me.SolutionBrowser1.selectedItem.Key = "TSolution" Then
    
            
    ElseIf Me.SolutionBrowser1.selectedItem.Key = "TProjectB" Then
        EditorClass.BrowseEngineData Me.TreeViewBrowser2
    ElseIf Me.SolutionBrowser1.selectedItem.Key = "TFolderB1" Then
        EditorClass.UpdateFolderBrowser Me.TreeViewBrowser2
    End If
End Sub

Private Sub SolutionBrowserBar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        SolutionBrowserBar1.Left = X
    End If
End Sub

Private Sub SolutionBrowserBar1_Resize()
    On Error Resume Next
    SolutionBrowser1.Top = Me.ToolbarRightCaptionBar1.Top + Me.ToolbarRightCaptionBar1.Height
    SolutionBrowser1.Left = 0 '120
    ToolbarRightCaptionBar1.Left = 0 '120
    SolutionBrowser1.Height = Me.SolutionBrowserBar1.Height - (Me.ToolbarRightCaptionBar1.Top + Me.ToolbarRightCaptionBar1.Height)
    TreeViewBrowser2.Height = Me.SolutionBrowser1.Height - 550 - (Me.ToolbarRightCaptionBar1.Top + Me.ToolbarRightCaptionBar1.Height)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        
        Case "TNew1"
        
        Case "TOpen1"
            Me.ActiveForm.OpenNew
        Case "TSave1"
            Me.ActiveForm.Save
    End Select
    
End Sub

Private Sub ToolbarRightCaptionBar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        
        Case "SBTNew1"
            SBTHandler CreateNewitem
        Case "SBTOpen1"
            SBTHandler OpenNew
        Case "SBTSave1"
            SBTHandler SaveItem
        Case "SBTRemove1"
            SBTHandler RemoveItem
    End Select
End Sub

Private Sub SBTHandler(Action As SBTActions)
    With Me.SolutionBrowser1.selectedItem
        If .Key = "TProjectB" Then
            EditorClass.TreeViewBrowserHandleActions Me.TreeViewBrowser2, ProjectBrowser, Action
        ElseIf .Key = "TSceneBrowser" Then
            EditorClass.TreeViewBrowserHandleActions Me.TreeViewBrowser2, SceneBrowser, Action
        End If
    End With
    
End Sub

Private Sub TreeViewBrowser2_AfterLabelEdit(Cancel As Integer, NewString As String)
    Dim thisNode As Node
    Dim oldName As String
    If Cancel = 0 Then
        Set thisNode = TreeViewBrowser2.selectedItem
        With thisNode
            If thisNode Is Nothing Then
                Exit Sub
            Else
                oldName = .text
                If .Tag = "Script" Then
                    MsgBox "Not Implemented"
                ElseIf .Tag = "Scene" Then
                    Me.Engine.RenameScene .text, NewString
                ElseIf .Tag = "Project" Then
                    Me.Engine.ProjectName = NewString
                ElseIf .Tag = "GameObject" Then
                    MsgBox "Not Implemented"
                End If
            End If
        
        End With
    Else

    End If
End Sub

Private Sub TreeViewBrowser2_DblClick()
    If Me.SolutionBrowser1.selectedItem.Key = "TProjectB" Then
        EditorClass.TreeViewBrowsersOnItemClick TreeViewBrowser2, ProjectBrowser
    ElseIf Me.SolutionBrowser1.selectedItem.Key = "TSceneBrowser" Then
        EditorClass.TreeViewBrowsersOnItemClick TreeViewBrowser2, SceneBrowser
    End If
End Sub

Private Sub TreeViewBrowser2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        EditorClass.TreeViewRightClick Me, Me.TreeViewBrowser2, Browsermnu1
    End If
End Sub
