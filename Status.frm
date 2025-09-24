VERSION 5.00
Begin VB.Form StatusDialog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please Wait... Building Game"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6630
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
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
Function PRGBar_SetValue(curVal As Long, max As Long, text As String)
    PRGStat1.Width = PRGBar1.Width * (curVal / max)
    Label2.Caption = text & (100 * (curVal / max)) & "% Complete "
    Me.Refresh
    DoEvents
End Function

Sub PRGBar_Show(CaptionText As String)
    Me.Caption = CaptionText
    Me.Show
End Sub

