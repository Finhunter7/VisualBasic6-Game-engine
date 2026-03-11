VERSION 5.00
Begin VB.Form BuildWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8055
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   8055
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   5175
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2415
      Begin VB.Frame Frame2 
         Caption         =   "File"
         Height          =   1455
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2175
         Begin VB.OptionButton Option1 
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1815
         End
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Build"
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6720
      TabIndex        =   0
      Top             =   5880
      Width           =   1215
   End
End
Attribute VB_Name = "BuildWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
