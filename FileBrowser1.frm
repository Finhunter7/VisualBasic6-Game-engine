VERSION 5.00
Begin VB.Form FileBrowser1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6975
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4080
      TabIndex        =   6
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4200
      TabIndex        =   5
      Top             =   3600
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FileBrowser1.frx":0000
      Left            =   4200
      List            =   "FileBrowser1.frx":0007
      TabIndex        =   4
      Text            =   "*.*"
      Top             =   3960
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   4440
      Width           =   1335
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
   Begin VB.DirListBox Dir1 
      Height          =   4365
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.FileListBox File1 
      Height          =   3405
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "FileName:"
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "FileType:"
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   3960
      Width           =   735
   End
End
Attribute VB_Name = "FileBrowser1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selectedItem As String
Dim Handle As FileBrowserHandler
Private curFileName As String

Enum FileBrowserHandler
    FOpen = 1
    FSave = 2
    DSave = 3
    DOpen = 4
End Enum

Private Sub Combo1_Change()
    File1.Pattern = Combo1.text
End Sub

Private Sub Combo1_Click()
    File1.Pattern = Combo1.text
End Sub

Public Property Get GetFileName() As String
    SelFileName = curFileName
End Property

Private Sub Command1_Click()
    If Handle = FOpen Then
        selectedItem = File1.Path & "\" & File1.FileName
        curFileName = File1.FileName
    ElseIf Handle = DSave Then
        selectedItem = Dir1.Path
    ElseIf Handle = FSave Then
        selectedItem = Dir1.Path & "\" & Text1.text
        curFileName = Text1.text
    Else
        selectedItem = Dir1.Path
    End If
    
    If Text1.text = "" And Handle = FSave Then
        MsgBox "Please Insert Filename", vbExclamation
    ElseIf Handle = FOpen And File1.FileName = "" Then
        MsgBox "Please Select File", vbExclamation
    Else
        Me.Hide
    End If
End Sub

Private Sub Command2_Click()
    selectedItem = ""
    curFileName = ""
    Me.Hide
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    Me.Caption = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    Me.Caption = File1.Path & "\" & File1.FileName
    Text1.text = File1.FileName
End Sub

Sub AddFilters(formatsText As String)
    Combo1.Clear
    If formatsText = "" Then
        Combo1.text = "*.*"
        Combo1.AddItem "*.*"
        Exit Sub
    End If
    For Each Item In Split(RTrim(formatsText), "|")
        Combo1.AddItem Item
    Next
    Combo1.text = Combo1.List(0)
End Sub

Function GetFormat() As String
    If Handle = FOpen Then
        cArray = Split(File1.FileName, ".")
        GetFormat = cArray(1)
    ElseIf Handle = FSave Then
        cArray = Split(Combo1.text, ".")
        GetFormat = cArray(1)
    End If
End Function

Function OpenFile(MainForm As Form, Optional WindowCaption As String, Optional formats As String) As String
    Handle = FOpen
    Me.Caption = WindowCaption
    Me.Command1.Caption = "Open"
    AddFilters formats
    Me.Show vbModal, MainForm
    OpenFile = selectedItem
End Function

Function OpenFolder(MainForm As Form, Optional WindowCaption As String) As String
    Handle = DOpen
    Me.Caption = WindowCaption
    Me.Command1.Caption = "Open"
    AddFilters "None"
    Me.Show vbModal, MainForm
    OpenFolder = selectedItem
End Function

Function SaveFile(MainForm As Form, Optional Suggest As String, Optional formats As String) As String
    Handle = FSave
    Me.Command1.Caption = "Save"
    Text1.text = Suggest
    AddFilters formats
    Me.Show vbModal, MainForm
    SaveFile = selectedItem
End Function

Function SaveToDir(MainForm As Form, Optional formats As String) As String
    Handle = DSave
    Me.Command1.Caption = "Save"
    AddFilters formats
    Me.Show vbModal, MainForm
    SaveToDir = selectedItem
End Function

