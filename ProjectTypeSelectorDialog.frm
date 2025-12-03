VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ProjectTypeSelectorDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Please Select Project Type"
   ClientHeight    =   3750
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "ProjectTypeSelectorDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "ProjectTypeSelectorDialog.frx":030A
      Left            =   4680
      List            =   "ProjectTypeSelectorDialog.frx":0314
      TabIndex        =   3
      Text            =   "VBScript"
      Top             =   1560
      Width           =   1215
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6165
      View            =   2
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Code Launguage:"
      Height          =   495
      Left            =   4680
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   1575
      Left            =   4680
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5040
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ProjectTypeSelectorDialog.frx":032E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ProjectTypeSelectorDialog.frx":0648
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ProjectTypeSelectorDialog.frx":0962
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ProjectTypeSelectorDialog.frx":0C7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ProjectTypeSelectorDialog.frx":0F96
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ProjectTypeSelectorDialog.frx":12B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ProjectTypeSelectorDialog.frx":15CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ProjectTypeSelectorDialog.frx":18E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ProjectTypeSelectorDialog.frx":25F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ProjectTypeSelectorDialog.frx":2910
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ProjectTypeSelectorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Engine As EngineClass
Option Explicit

Sub OpenThis(MMainWindow As MDIForm, GameEngine As EngineClass)
    Set Engine = GameEngine
    Me.Show vbModal, MMainWindow
End Sub

Private Sub CancelButton_Click()
    Me.Hide
End Sub

Private Sub Combo1_Click()
    MsgBox "Not Yet Supported/Implemented", vbInformation
End Sub

Private Sub Form_Load()
    Me.ListView1.ListItems.Add(, "Stantard32", "Stantard 32Bit EXE", 1, 1).Tag = "EXE32"
    Me.ListView1.ListItems.Add(, "Stantard16", "Stantard 16Bit EXE", 8, 8).Tag = "EXE16"
    Me.ListView1.ListItems.Add(, "VBGame", "VBCE Game Project (2D)", 7, 7).Tag = "VBGame"
    Me.ListView1.ListItems.Add(, "VBDXGame", "VBCE Game Project With Direct8 Lib (2D/3D)", 7, 7).Tag = "VBCEDGame"
    Me.ListView1.ListItems.Add(, "VBDX11Game", "VBCE Game Project With Direct11 Lib (2D/3D)", 7, 7).Tag = "VBCED11Game"
    Me.ListView1.ListItems.Add(, "VidPro", "VBCE Animation Project (2D)", 9, 9).Tag = "Video"
    
    
End Sub

Private Sub ListView1_Click()
    With Me.ListView1.selectedItem
        If .Tag = "VBCEDGame" Then
            Me.Label1.Caption = "Allows Usage Of Direct8 Library To Generate Graphics If Supported."
        ElseIf .Tag = "VBCED11Game" Then
            Me.Label1.Caption = "Allows Usage Of Direct11 Library To Generate Graphics If Supported."
        ElseIf .Tag = "VBGame" Then
            Me.Label1.Caption = "This Project Type Uses The Form's Build In Draw Functions To Create Graphics."
        ElseIf .Tag = "Video" Then
            Me.Label1.Caption = "This Project Type Uses The Form's Build In Draw Functions To Create Graphics For Animations."
        Else
            Me.Label1.Caption = ""
        End If
    End With
End Sub

Private Sub ListView1_DblClick()
    If Me.ListView1.selectedItem Is Nothing Then
        Exit Sub
    End If
    With Me.ListView1.selectedItem
        If .Tag = "VBGame" Then
            Engine.ProjectType = VBCEGameProject
            Me.Hide
        ElseIf .Tag = "VBCED11Game" Then
            If Engine.Direct11Lib Is Nothing Then
                MsgBox "This Project Type Is Not Supported In Your Computer", vbInformation
            Else
                Engine.ProjectType = VBCEDX11GameProject
                Me.Hide
            End If
        Else
            MsgBox "Not Yet Supported/Implemented", vbInformation
        End If
    End With
End Sub

Private Sub OKButton_Click()
    Me.Hide
End Sub
