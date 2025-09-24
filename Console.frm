VERSION 5.00
Begin VB.Form Console 
   BackColor       =   &H00000000&
   Caption         =   "Console"
   ClientHeight    =   5805
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Console.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   5805
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   5775
      Left            =   7800
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Menu mnuExecute1 
      Caption         =   "Execute"
      Begin VB.Menu mnuStatement1 
         Caption         =   "Statement"
      End
      Begin VB.Menu mnuHook1 
         Caption         =   "Hook Statement"
      End
   End
   Begin VB.Menu mnuview1 
      Caption         =   "View"
      Begin VB.Menu mnuClear1 
         Caption         =   "Clear"
      End
   End
End
Attribute VB_Name = "Console"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private textBuffer As String
Private InputBuffer As String
Private StatementHook As String
Private curText As Long
'Private line() As Long
Private LinesCount As Integer
Public GameEngine As EngineClass

Private Sub Form_Load()
    'Clear
    curText = 1
    VScroll1.Min = 1
End Sub

Private Sub Form_Paint()
    Repaint
End Sub

Sub Repaint()
    Me.Cls
    Print Mid(textBuffer, curText)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    VScroll1.Top = 0
    VScroll1.Left = Me.Width - VScroll1.Width - 115
    VScroll1.Height = Me.Height - 750
    'Text1.Top = 0
    'Text1.Left = 0
    'Text1.Height = Me.Height - 800
    'Text1.Width = Me.Width - 200
End Sub

Function WriteLine(text)
    On Error Resume Next
    'Text1.text = textBuffer
    'Text1.text = Text1.text & text & vbNewLine
    Print text
    LinesCount = LinesCount + 1
    textBuffer = textBuffer & text & vbCrLf
    VScroll1.max = LinesCount
    'textBuffer = Text1.text
End Function

Sub DisplayMessage()
    'Me.Clear
    'WriteLine "VBCE Engine Version " & App.Major & "." & App.Minor & "." & App.Revision
    'WriteLine "VBCE Made By Sami Nissinen 2025-2025"
    'WriteLine "ScriptControll Ready For Execution Input"
End Sub

Function ReadLine() As String
    If Not Len(InputBuffer) = 0 Then
        ReadLine = InputBuffer
    End If
End Function

Function Clear()
    'Text1.text = ""
    Me.Cls
    Me.Point 0, 0
    textBuffer = ""
    LinesCount = 0
    If Not StatementHook = "" Then
        Me.Caption = "Console - " & StatementHook
    End If
End Function
Function CreateNewConsoleInstance() As Console
    Dim newConsole As New Console
    Set CreateNewConsoleInstance = newConsole
    newConsole.Show
End Function

Private Sub mnuClear1_Click()
    Me.Clear
End Sub

Private Sub mnuHook1_Click()
    Dim Stat As String
    Stat = InputBox("Hook Statement")
    If Len(Stat) = 0 Then
        StatementHook = ""
        Me.Caption = "Console"
        Exit Sub
    End If
    
    Me.WriteLine "Hooked To " & Stat
    StatementHook = Stat
    Me.Caption = "Console - " & StatementHook
End Sub

Private Sub mnuStatement1_Click()
    Dim Stat As String
    Stat = InputBox("Execute Statement")
    If Len(Stat) = 0 Then
        Exit Sub
    End If
    Me.Caption = "Console - " & StatementHook & Stat
    Me.WriteLine Stat
    On Error GoTo err32
    Me.GameEngine.CodeEngine.ExecuteStatement StatementHook & Stat
    If StatementHook = "" Then
        Me.Caption = "Console"
    Else
        Me.Caption = "Console - " & StatementHook
    End If
    Exit Sub
    
err32:
    Console.WriteLine Me.GameEngine.CodeEngine.Error.Description
End Sub


Private Sub VScroll1_Change()
    curText = VScroll1.Value
    Repaint
End Sub
