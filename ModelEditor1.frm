VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ModelEditor1 
   BackColor       =   &H8000000C&
   Caption         =   "Model-Texture-UV Editor"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11670
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   11670
   Tag             =   "ModelEditor"
   Begin ComctlLib.Toolbar DrawToolsBar1 
      Align           =   3  'Align Left
      Height          =   7965
      Left            =   0
      TabIndex        =   2
      Top             =   420
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   14049
      Appearance      =   1
      ImageList       =   "DrawModeImages1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   4
            Style           =   1
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   6
            Style           =   1
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   741
      Appearance      =   1
      _Version        =   327682
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   11670
      TabIndex        =   3
      Top             =   8385
      Width           =   11670
      Begin VB.HScrollBar AnimationScroll1 
         Height          =   495
         Left            =   0
         Max             =   20
         TabIndex        =   4
         Top             =   0
         Width           =   11655
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   8880
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Current Frame:"
            TextSave        =   "Current Frame:"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Text            =   "Mouse X:"
            TextSave        =   "Mouse X:"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Text            =   "Mouse Y:"
            TextSave        =   "Mouse Y:"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   5520
      Picture         =   "ModelEditor1.frx":0000
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   0
      Top             =   4440
      Width           =   1080
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3360
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ModelEditor1.frx":05FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ModelEditor1.frx":0B40
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ModelEditor1.frx":1082
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList DrawModeImages1 
      Left            =   2760
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ModelEditor1.frx":15C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ModelEditor1.frx":16D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ModelEditor1.frx":17E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ModelEditor1.frx":18FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ModelEditor1.frx":1A0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ModelEditor1.frx":1B16
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ModelEditor1.frx":1C20
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ModelEditor1.frx":1D2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ModelEditor1.frx":1E34
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ModelEditor1.frx":1F36
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ModelEditor1.frx":2048
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ModelEditor1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private curObject As Object
Private NewMeshData As String
Private CodeLoader As New MSScriptControl.ScriptControl
Private ScaleZoom As Double

Private Sub AnimationScroll1_Change()
    SetAnimationFrame AnimationScroll1.Value
End Sub

Private Sub Form_Load()
    ScaleZoom = 1
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Picture1.Top = (Me.Height / 2) - (Picture1.Height / 2)
    Picture1.Left = (Me.Width / 2) - (Picture1.Width / 2)
End Sub


Sub DrawObject()
    CodeLoader.AddCode NewMeshData
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 109 Then
        ChangeZoom -0.05
    ElseIf KeyCode = 107 Then
        ChangeZoom 0.05
    End If
End Sub

Sub ChangeZoom(Val As Double)
    ScaleZoom = ScaleZoom + Val
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim NewRender As New VBCEArrayImage_Class
    NewRender.CreateFromImage Me.Picture1.Picture, 79, 32
    
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.StatusBar1.Panels(2).text = "Mouse X: " & (X * ScaleZoom)
    Me.StatusBar1.Panels(3).text = "Mouse Y: " & (Y * ScaleZoom)
End Sub

Private Sub Picture2_Resize()
    On Error Resume Next
    AnimationScroll1.Top = 0
    AnimationScroll1.Left = 0
    AnimationScroll1.Height = Picture2.Height
    AnimationScroll1.Width = Picture2.Width
End Sub

Sub SetAnimationFrame(Frame As Long)
    Me.StatusBar1.Panels(1).text = "Current Frame: " & Frame
End Sub

