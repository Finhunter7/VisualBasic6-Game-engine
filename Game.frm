VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Viewport1"
   ClientHeight    =   4125
   ClientLeft      =   165
   ClientTop       =   210
   ClientWidth     =   5220
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Game.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   4125
   ScaleWidth      =   5220
   Tag             =   "ViewPort"
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3870
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   0
            Text            =   "Scene:"
            TextSave        =   "Scene:"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   0
            Text            =   "Layer: 1"
            TextSave        =   "Layer: 1"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   0
            Text            =   "Target Framerate:"
            TextSave        =   "Target Framerate:"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim list As New Collection
'Dim Inputs As New InputSystem_Class
'Public CodeEngine As New ScriptControl ' Referenssi ScriptControl luokkaan jolla pystyt‰‰n suorittamaan koodia.
'Public engClass As New EngineClass
'Dim VecClass As New Vector_Class
'Public MouseX As Double
'Public MouseY As Double
'Public MouseDown As Integer
'Dim curScene As Long
'Public Engine As EngineClass

Private Sub Form_Load()
    On Error Resume Next
    'MsgBox Command
    
    'Me.LoadEngine
    
    'Form2.Show
    'Me.OnFormLoad
End Sub
Function LoadEngine()

    'Dim eClass As EngineClass
    'Dim startScene As Scene_Class
    
    'Set eClass = engClass
    'eClass.ProjectName = "Project1"
    'eClass.IsInEditor = True
    
    'eClass.LoadEngine Me, Me, Form2, Scene_Browser, EngineSettings, Console
    
    'Set startScene = eClass.CreateNewScene("SampleScene")
    'eClass.AddScene startScene.Name
    'startScene.Name = "SampleScene"
    'startScene.CreateNewObject "Circle", , , 1
    'eClass.Scenes.Add startScene, "SampleScene"
    'eClass.ScenesInactive.Add startScene, "SampleScene"
End Function

Function OnFormLoad()

End Function

Private Sub mnuExit1_Click()
    'Unload Me
    'End
End Sub

Function Update()
    Me.Cls
End Function

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Data.GetFormat(1) Then
        MsgBox "GameObject " & Data.GetData(5)
    End If
End Sub

Private Sub Form_Resize()
    'Me.ScaleWidth = 5000
    'Me.ScaleHeight = 4000
End Sub
