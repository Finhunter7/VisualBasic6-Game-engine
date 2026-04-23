VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form TimelineEditor_Form 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Untiled Action - Keyframe Editor"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7365
   ClipControls    =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   1410
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   0
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   7335
   End
   Begin VB.PictureBox Timeline_Frame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   487
      TabIndex        =   0
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "TimelineEditor_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents Engine As EngineClass
Attribute Engine.VB_VarHelpID = -1
Private drawToFrame As Long
Private drawFrames As Long
Private offset As Long
Private curFrame As Long

Sub Build(tEngine As EngineClass)
    Set Engine = tEngine
End Sub

Private Sub Engine_OnUpdate()
    curFrame = Engine.GetRenderedFrames()
    Draw
End Sub

Private Sub Form_Load()
    drawToFrame = 10
    drawFrames = 100
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.Timeline_Frame.Top = 0
    Me.Timeline_Frame.Left = 0
    Me.Timeline_Frame.Width = Me.ScaleWidth
    Me.Timeline_Frame.Height = Me.ScaleHeight - 375 - Me.StatusBar1.Height
    Me.HScroll1.Top = Me.Timeline_Frame.Height
    Me.HScroll1.Left = 0
    Me.HScroll1.Width = Me.ScaleWidth
    Me.Timeline_Frame.ScaleHeight = 100
    Me.Timeline_Frame.ScaleWidth = 500
    Draw
End Sub

Private Sub HScroll1_Change()
    offset = HScroll1.Value
    Draw
End Sub

Private Sub HScroll1_Scroll()
    offset = HScroll1.Value
    Draw
End Sub

Private Sub Timeline_Frame_Click()
    CreateTimeline
End Sub

Sub Draw()
    If curFrame > drawFrames + offset And Engine.IsEngineRunning Then
        offset = offset + drawFrames
    End If
    CreateTimeline
    Me.StatusBar1.Panels(1).text = "Curent Frame: " & curFrame
End Sub

Sub CreateTimeline()
    Me.Timeline_Frame.Cls
    Dim dis As Long
    dis = (Timeline_Frame.ScaleWidth / (drawFrames / drawToFrame))
    For i = 1 To drawFrames
        Timeline_Frame.PSet (dis * i, Timeline_Frame.Top)
        Timeline_Frame.Print CStr((drawToFrame * i) + offset)
        Timeline_Frame.Line (dis * i, Timeline_Frame.Top)-(dis * i, Timeline_Frame.ScaleHeight), vbBlack
    Next
    Me.Timeline_Frame.Line ((Timeline_Frame.ScaleWidth / drawFrames) * (curFrame - offset), Me.Timeline_Frame.Top)-((Timeline_Frame.ScaleWidth / drawFrames) * (curFrame - offset), Me.Timeline_Frame.ScaleHeight), vbRed
End Sub

Private Sub Timeline_Frame_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyLeft
            curFrame = curFrame - 1
        Case vbKeyRight
            curFrame = curFrame + 1
    End Select
    Draw
End Sub

Private Sub Timeline_Frame_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        curFrame = CLng(((X / Timeline_Frame.ScaleWidth) * drawFrames) + offset)
        Draw
    End If
End Sub

