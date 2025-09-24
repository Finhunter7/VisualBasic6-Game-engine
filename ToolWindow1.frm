VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ToolWindow1 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   2280
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   2280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ListView ListView1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   4260
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1080
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ToolWindow1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ToolWindow1.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ToolWindow1.frx":0634
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ToolWindow1.frx":094E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ToolWindow1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MCodeEditor As Form
Private Commands(2, 2) As String

Function CreateTable()
    Commands(0, 0) = "Engine"
    Commands(0, 1) = "Vector"
    Commands(0, 2) = "Console"
End Function


Function ToolWindowCreateList(Optional ClassName = "")
    Me.ListView1.ListItems.Clear
    If ClassName = "" Then
        
        For c = 1 To 3
            For i = 1 To 3
                Me.ListView1.ListItems.Add , Commands(c - 1, i - 1), Commands(c - 1, i - 1), 1, 1
            Next
        Next
        
    ElseIf ClassName = "Vector" Then
        
    End If
End Function

Function ToolWindowOpen(TCodeEditor As Form)
    Set MCodeEditor = TCodeEditor
    Me.CreateTable
    Me.ToolWindowCreateList "Vector"
    Me.Show
End Function

Private Sub Form_Resize()
    Me.ListView1.Top = 0
    Me.ListView1.Left = 0
    Me.ListView1.Height = Me.Height - 1100
End Sub

