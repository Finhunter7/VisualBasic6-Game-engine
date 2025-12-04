VERSION 5.00
Begin VB.UserControl Script_LogicBrick 
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3870
   ScaleHeight     =   1335
   ScaleWidth      =   3870
   Begin VB.Frame Frame1 
      Caption         =   "Script"
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Execute Statement:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Script Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Script_LogicBrick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Name As String

Public Sub SetName(TName As String)
    Name = TName
    Frame1.Caption = Name & " Script Controller"
End Sub

