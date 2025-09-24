VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Code"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   8910
   Icon            =   "CodeEditor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   1
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   8055
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            TextSave        =   "NUM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
            Enabled         =   0   'False
            TextSave        =   "INS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6495
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   11456
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"CodeEditor.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8280
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   8280
      Top             =   960
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
            Picture         =   "CodeEditor.frx":04C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CodeEditor.frx":07DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CodeEditor.frx":0AF8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile1 
      Caption         =   "File"
      Begin VB.Menu mnuNew1 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpen1 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave1 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveAs1 
         Caption         =   "Save as"
      End
      Begin VB.Menu mnuSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose1 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuEdit1 
      Caption         =   "Edit"
      Begin VB.Menu mnuCreGaFunc1 
         Caption         =   "Create Game Subs"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim curFile As String
Dim textChanged As Boolean
Dim dlgConst As ErrorConstants


Private Sub Form_Resize()
On Error Resume Next
RichTextBox1.Top = 480
RichTextBox1.Left = 0
RichTextBox1.Height = Me.Height - 1500
RichTextBox1.Width = Me.Width - 100
End Sub

Private Sub mnuClose1_Click()
Me.Hide
End Sub

Private Sub mnuCreGaFunc1_Click()
RichTextBox1.Text = RichTextBox1.Text + "Sub Load()" + vbNewLine + "" + vbNewLine + "End Sub" + vbNewLine + vbNewLine
RichTextBox1.Text = RichTextBox1.Text + "Sub Update()" + vbNewLine + "" + vbNewLine + "End Sub" + vbNewLine + vbNewLine
textChanged = False
End Sub

Private Sub mnuNew1_Click()
If textChanged = True Then
    choise = MsgBox("Changes has been made. Do you want to save", vbExclamation + vbYesNoCancel)
    If choise = vbYes Then
        If curFile = "" Then
            CommonDialog1.DialogTitle = "Save"
            CommonDialog1.ShowSave
            RichTextBox1.SaveFile CommonDialog1.FileName
            RichTextBox1.Text = ""
            curFile = ""
            textChanged = False
        Else
            RichTextBox1.SaveFile curFile
            textChanged = False
        End If
    ElseIf choise = vbNo Then
        RichTextBox1.Text = ""
        curFile = ""
        textChanged = False
    Else
    
    End If
Else
    RichTextBox1.Text = ""
    curFile = ""
    textChanged = False
End If

End Sub

Private Sub mnuOpen1_Click()
CommonDialog1.CancelError = True
CommonDialog1.DialogTitle = "Open"
CommonDialog1.ShowOpen
RichTextBox1.LoadFile CommonDialog1.FileName
curFile = CommonDialog1.FileName
textChanged = False
End Sub

Private Sub mnuSave1_Click()
If curFile = "" Then
    CommonDialog1.DialogTitle = "Save"
    CommonDialog1.ShowSave
    RichTextBox1.SaveFile CommonDialog1.FileName
    curFile = CommonDialog1.FileName
    textChanged = False
Else
    RichTextBox1.SaveFile curFile
    textChanged = False
End If

End Sub

Private Sub mnuSaveAs1_Click()
CommonDialog1.DialogTitle = "Save as"
CommonDialog1.ShowSave
RichTextBox1.SaveFile CommonDialog1.FileName
curFile = CommonDialog1.FileName
textChanged = False
End Sub

Private Sub RichTextBox1_Change()
textChanged = True
End Sub

