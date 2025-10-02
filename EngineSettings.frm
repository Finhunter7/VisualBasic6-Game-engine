VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form EngineSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6240
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   6240
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Engine"
      Height          =   4575
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   5775
      Begin VB.ComboBox PName_Combo 
         Height          =   315
         Left            =   1560
         TabIndex        =   15
         Top             =   360
         Width           =   3975
      End
      Begin VB.ComboBox StartupProcs_Combo1 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "EngineSettings.frx":0000
         Left            =   2040
         List            =   "EngineSettings.frx":0007
         TabIndex        =   13
         Text            =   "Main"
         Top             =   1080
         Width           =   3495
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "EngineSettings.frx":0011
         Left            =   1560
         List            =   "EngineSettings.frx":001E
         TabIndex        =   11
         Text            =   "VBGame Project"
         Top             =   720
         Width           =   3975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Uncapped Framerate"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   5295
      End
      Begin VB.ComboBox FramerateCombo1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1035
            SubFormatType   =   1
         EndProperty
         Height          =   315
         ItemData        =   "EngineSettings.frx":0057
         Left            =   1560
         List            =   "EngineSettings.frx":0070
         TabIndex        =   6
         Text            =   "60"
         Top             =   1440
         Width           =   3975
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "EngineSettings.frx":0092
         Left            =   1560
         List            =   "EngineSettings.frx":009C
         TabIndex        =   5
         Text            =   "VBScript"
         Top             =   2160
         Width           =   3975
      End
      Begin VB.Label Label3 
         Caption         =   "Project Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label L_Startup_Procs 
         Caption         =   "Startup Procedure Sub:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label L_Project_Type 
         Caption         =   "Project Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Framerate:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Code Laugage:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   1215
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   5055
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   8916
      MultiRow        =   -1  'True
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Project"
            Key             =   "TProject"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "MSScriptControl"
            Key             =   "TMSScript1"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apply"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   5280
      Width           =   1095
   End
End
Attribute VB_Name = "EngineSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    If Me.Check1.Value = 1 Then
        Me.FramerateCombo1.Enabled = False
    Else
        FramerateCombo1.Enabled = True
    End If
End Sub

Private Sub Command1_Click()
    ApplySettings
    Me.Hide
End Sub

Private Sub Command2_Click()
    ApplySettings
End Sub

Sub ApplySettings()
    MainWindow.Engine.SetTargetFramerate Me.FramerateCombo1.text
    MainWindow.Engine.SetCodeLaunguage Me.Combo2.text
    If Check1.Value = 1 Then
        MainWindow.Engine.SetTargetFramerate -1
    End If
End Sub

Private Sub Command3_Click()
    Me.Hide
End Sub

Private Sub TabStrip1_Click()
    Me.Frame1.Visible = False
    Select Case Me.TabStrip1.selectedItem.Key
    
        Case "TProject"
            Me.Frame1.Visible = True
        Case "TMSScript1"
            
    
    End Select
End Sub

Sub ProjectSettingsShow(ownerForm As Object)
    Me.Show vbModal, ownerForm
End Sub
