VERSION 5.00
Begin VB.Form SceneAcWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activate Scene"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4335
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "SceneAcWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private GameEngine As EngineClass

Private Sub Command1_Click()
    Me.Hide
End Sub

Private Sub List1_Click()
    Me.Caption = "Activate Scene " & List1.text
End Sub

Private Sub List1_DblClick()
    GameEngine.AddScene List1.text, Default
    Me.Hide
End Sub

Public Sub ShowScenes(ModalForm As Form, TGameEngine As EngineClass)
    Set GameEngine = TGameEngine
    ListScenes
    
    Me.Show vbModal, ModalForm
    
End Sub

Private Sub ListScenes()
    Me.List1.Clear
    For i = 1 To GameEngine.ScenesInactive.Count
        Me.List1.AddItem GameEngine.ScenesInactive.Item(i).Name
    Next
End Sub
