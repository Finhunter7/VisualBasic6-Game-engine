VERSION 5.00
Begin VB.Form StatusDialog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please Wait... Building Game"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6615
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.PictureBox PRGBar1 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   6315
      TabIndex        =   0
      Top             =   240
      Width           =   6375
      Begin VB.PictureBox PRGStat1 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   3135
         TabIndex        =   1
         Top             =   0
         Width           =   3135
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Creating Files"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   6375
   End
End
Attribute VB_Name = "StatusDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CancelOp As Boolean

Enum StatusdialogTypes
    StatusdlgDefault = 0
    StatusdlgCancel = 1
End Enum

Function PRGBar_SetValue(curVal As Long, max As Long, text As String)
    PRGStat1.Width = PRGBar1.Width * (curVal / max)
    Label2.Caption = text & (100 * (curVal / max)) & "% Complete "
    Me.Refresh
End Function

Sub PRGBar_Show(CaptionText As String, Optional dlgType As StatusdialogTypes)
    If dlgType = StatusdlgCancel Then
        Command1.Visible = True
        Me.Height = 2190
    Else
        Me.Height = 1650
        Command1.Visible = False
    End If
    Me.Caption = CaptionText
    Me.Show
End Sub

Public Property Get IsOperationCanceled() As Boolean
    IsOperationCanceled = CancelOp
End Property

Private Sub Command1_Click()
    CancelOp = True
End Sub

Private Sub Form_Load()
    CancelOp = False
End Sub
