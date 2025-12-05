VERSION 5.00
Begin VB.UserControl LogicEditor 
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6420
   ControlContainer=   -1  'True
   ScaleHeight     =   4455
   ScaleWidth      =   6420
   Begin VB.Menu mnuFile1 
      Caption         =   "File"
   End
   Begin VB.Menu mnuLogic1 
      Caption         =   "Logic"
      Begin VB.Menu mnuAddLogic1 
         Caption         =   "Add Logic"
         Begin VB.Menu mnuUpdateLogic1 
            Caption         =   "Update Logic"
         End
         Begin VB.Menu mnuScriptLogic1 
            Caption         =   "Script Logic"
         End
      End
   End
End
Attribute VB_Name = "LogicEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private MyBricks As New Collection

Private Sub mnuScriptLogic1_Click()
    Dim TName As String
    TName = InputBox("Insert Name For Control")
    Set thisControl = Controls.Add("VBCE.Script_LogicBrick", TName)
    thisControl.SetName TName
    thisControl.Visible = True
    MyBricks.Add thisControl, TName
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = 2 Then
            PopupMenu mnuLogic1
        End If
End Sub

Sub ConnectControllers(ActivatorCont As Object, ControllerCont As Object)
    Set ActivatorCont.ConnectedCont = ControllerCont
    Set ControllerCont.Activator = ActivatorCont
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If Not ActiveControl Is Nothing Then
            ActiveControl.Move X, Y
        Else
        
        End If
    Else
    
    End If
End Sub
