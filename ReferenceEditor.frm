VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ReferenceEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add/Create Reference"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6000
   Icon            =   "ReferenceEditor.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6000
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Propertites"
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   5775
      Begin VB.CheckBox Check1 
         Caption         =   "Reload Object On Engine Start"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Width           =   5535
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   12
         Top             =   1440
         Width           =   4455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   255
         Left            =   5280
         TabIndex        =   11
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Text            =   "C:\"
         Top             =   720
         Width           =   4095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   1080
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Text            =   "Reference1"
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label L_ServerName 
         Caption         =   "ServerName:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label L_Path 
         Caption         =   "File:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Format Ex: Libname.ClassName"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Width           =   5415
      End
      Begin VB.Label L_R_Class 
         Caption         =   "Class:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label L_Name 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   3840
      Width           =   1095
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6588
      MultiRow        =   -1  'True
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ActiveX Object"
            Key             =   "TObject1"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ReferenceEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ReturnObject As Object
Private Handle As RefAddTypes
Enum RefAddTypes
    CreateNew = 0
    GetFromFile = 1
    View = 2
End Enum
Private Sub Command1_Click()
    Set ReturnObject = Nothing
    Me.Hide
End Sub

Private Sub Command2_Click()
    If Handle = View Then
        Me.Hide
    ElseIf Handle = CreateNew Then
        CreateReference
    Else
        CreateReference
    End If
End Sub

Private Sub Command3_Click()
    Dim fileName As String
    fileName = FileBrowser1.OpenFile(MainWindow, "Open File", "*.*")
    Text2.text = fileName
End Sub

Function CreateNewReference(TVBCEEngine As EngineClass, ownerForm As Form) As EngineRefrence_Class
    ListReferenceClasses TVBCEEngine
    Command2.Caption = "Create"
    Combo2.Enabled = True
    Combo1.Enabled = True
    Combo2.text = ""
    Combo1.text = ""
    Text2.Enabled = False
    Text2.text = ""
    Text1.Enabled = True
    Command3.Enabled = False
    Handle = CreateNew
    Me.Show vbModal, ownerForm
    Set CreateNewReference = ReturnObject
End Function

Function GetReferenceFromFile(TVBCEEngine As EngineClass, ownerForm As Form) As EngineRefrence_Class
    ListReferenceClasses TVBCEEngine
    Command2.Caption = "Get"
    Combo2.Enabled = False
    Combo1.Enabled = True
    Combo2.text = ""
    Combo1.text = ""
    Text2.Enabled = True
    Text2.text = ""
    Text1.Enabled = True
    Command3.Enabled = True
    Handle = GetFromFile
    Me.Show vbModal, ownerForm
    Set GetReferenceFromFile = ReturnObject
End Function

Sub ViewReference(VBCEEngineReference As EngineRefrence_Class, ownerForm As Form)
    Handle = View
    Command2.Caption = "Ok"
    Combo2.Enabled = False
    Text2.Enabled = False
    Text1.Enabled = False
    Command3.Enabled = False
    Combo1.Enabled = False
    Combo2.text = VBCEEngineReference.Server
    Combo1.text = VBCEEngineReference.RefrenceClass
    Text2.text = VBCEEngineReference.Path
    Text1.text = VBCEEngineReference.Name
    Me.Show vbModal, ownerForm
End Sub

Sub ListReferenceClasses(TVBCEEngine As EngineClass)
    Combo1.Clear
    For Each Item In TVBCEEngine.EngineRefrences
        Combo1.AddItem Item.RefrenceClass
    Next
End Sub

Private Sub CreateReference()
    Dim newRefrence As New EngineRefrence_Class
    newRefrence.Name = Text1.text
    If Len(newRefrence.Name) = 0 Then
        MsgBox "Please Insert Valid Name For Reference", vbExclamation
        Exit Sub
    End If
    newRefrence.RefrenceClass = Combo1.text
    newRefrence.Path = Text2.text
    newRefrence.Server = Combo2.text
    newRefrence.HoldData = CBool(Check1.Value)
    Set ReturnObject = newRefrence
    Me.Hide
End Sub

