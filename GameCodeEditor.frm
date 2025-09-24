VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form CodeEditor 
   Caption         =   "Code Editor"
   ClientHeight    =   5415
   ClientLeft      =   165
   ClientTop       =   210
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "GameCodeEditor.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   5415
   ScaleWidth      =   4470
   Tag             =   "CodeEditor"
   Begin ComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   13
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "TCut1"
            Description     =   ""
            Object.ToolTipText     =   "Cut"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "TCopy1"
            Description     =   ""
            Object.ToolTipText     =   "Copy"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "TPaste1"
            Description     =   ""
            Object.ToolTipText     =   "Paste"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "TFind1"
            Description     =   ""
            Object.ToolTipText     =   "Find"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "TUndo1"
            Description     =   ""
            Object.ToolTipText     =   "Undo"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "TRedo1"
            Description     =   ""
            Object.ToolTipText     =   "Redo"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "TRun1"
            Description     =   ""
            Object.ToolTipText     =   "Start Game In Window"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "TCheck1"
            Description     =   ""
            Object.ToolTipText     =   "Check Code For Errors"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Check For Errors"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "TProp1"
            Description     =   ""
            Object.ToolTipText     =   "View Propertites"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
      EndProperty
   End
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   2
      Top             =   420
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   688
      _Version        =   327682
      Begin VB.ComboBox ScriptCombo1 
         Height          =   360
         Left            =   60
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "Script_Name"
         Top             =   15
         Width           =   2775
      End
      Begin VB.ComboBox MethodCombo1 
         Height          =   360
         Left            =   3000
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "Script_Methods"
         Top             =   15
         Width           =   8055
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5160
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
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
            AutoSize        =   2
            Object.Width           =   2778
            Text            =   "Exposed Object:       "
            TextSave        =   "Exposed Object:       "
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Text            =   "Line:"
            TextSave        =   "Line:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4471
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"GameCodeEditor.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   10560
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   13
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GameCodeEditor.frx":04CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GameCodeEditor.frx":0A0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GameCodeEditor.frx":0F50
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GameCodeEditor.frx":1492
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GameCodeEditor.frx":15A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GameCodeEditor.frx":16B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GameCodeEditor.frx":17C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GameCodeEditor.frx":18DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GameCodeEditor.frx":1E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GameCodeEditor.frx":1F2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GameCodeEditor.frx":2040
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GameCodeEditor.frx":214A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GameCodeEditor.frx":225C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "CodeEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const KeyWordsCount = 51
Public curClassName As String
Public GameObject As GameObject_Class
Public GameEngine As EngineClass
Public Session As CodeEditorEnums
Private inEditObjectData As Object
Public cChanged As Boolean

Private colorAdded As Boolean
Private Items(KeyWordsCount) As String

Private Sub Form_Activate()
    MainWindow.mnuSaveW1.Caption = "Save " & Me.curClassName & " Code"
    MainWindow.mnuSaveAs1.Caption = "Save " & Me.curClassName & " Code As"
End Sub

Private Sub Form_Load()
    LoadItems
    If GameObject Is Nothing Then
        Me.StatusBar1.Panels(3).text = "Exposed Object: None"
    Else
        Me.StatusBar1.Panels(3).text = "Exposed Object: " & GameObject.Name
    End If
    
    Me.Caption = Me.curClassName
    'ToolWindow1.ToolWindowOpen Me
End Sub

Function LoadItems()
    Items(1) = "Sub"
    Items(2) = "Function"
    Items(3) = "Exit"
    Items(4) = "End"
    Items(5) = "Private"
    Items(6) = "Public"
    Items(7) = "And"
    Items(8) = "Or"
    Items(9) = "Not"
    Items(10) = "Is"
    Items(11) = "To"
    Items(12) = "Do"
    Items(13) = "Until"
    Items(14) = "Loop"
    Items(15) = "For"
    Items(16) = "Next"
    Items(17) = "True"
    Items(18) = "False"
    Items(19) = "If"
    Items(20) = "Then"
    Items(21) = "Set"
    Items(22) = "As"
    Items(23) = "On"
    Items(24) = "Return"
    Items(25) = "Resume"
    Items(26) = "Optional"
    Items(27) = "Double"
    Items(28) = "Integer"
    Items(29) = "Long"
    Items(30) = "Boolean"
    Items(31) = "String"
    Items(32) = "Nothing"
    Items(33) = "Else"
    Items(34) = "ElseIf"
    Items(35) = "New"
    Items(36) = "Goto"
    Items(37) = "Error"
    Items(38) = "Const"
    Items(39) = "Enum"
    Items(40) = "Friend"
    Items(41) = "Let"
    Items(42) = "Get"
    Items(43) = "Property"
    Items(44) = "Dim"
    Items(45) = "Declare"
    Items(46) = "Class"
    Items(47) = "Call"
    Items(48) = "Lib"
    Items(49) = "Alias"
    Items(50) = "ByVal"
    Items(51) = "Each"
    ' ConvertFuncs
    'Items(52) = "VbString"
    'Items(53) = "VbSingle"
    'Items(54) = "CLng"
    'Items(55) = "CStr"
    'Items(56) = "CVar"
    'Items(57) = "CDbl"
    'Items(58) = "CSng"
End Function
Private Sub Form_Paint()
    'AddColor
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Text1.Top = Me.Toolbar1.Height + Me.Toolbar2.Height
    Text1.Left = 0
    Text1.Height = Me.Height - 780 - (Me.Toolbar1.Height + Me.Toolbar2.Height)
    Text1.Width = Me.Width - 125
End Sub

Private Sub mnuAddC1_Click()
    AddColor
End Sub

Private Sub mnuAddColorCheck1_Click()
    If mnuAddColorCheck1.Checked Then
        mnuAddColorCheck1.Checked = False
    Else
        mnuAddColorCheck1.Checked = True
    End If
End Sub

Private Sub mnuClose1_Click()
    Me.Hide
End Sub

Sub Update()
    AddColor
End Sub

Function EditData(Script As Script_Class, TName As String, TSession As CodeEditorEnums, MGameEngine As EngineClass, Optional GObject As GameObject_Class)
    Set Me.GameObject = GObject
    Set Me.GameEngine = MGameEngine
    Set inEditObjectData = Script
    Me.Text1.text = Script.Data
    Me.Session = TSession
    Me.curClassName = TName
    Me.Caption = Me.curClassName
    Me.Show
    AddColor
End Function

Private Sub mnuOpen1_Click()
    'Dim file As String
    'On Error Resume Next
    'With CommonDialog1
        '.DialogTitle = "Open"
        '.CancelError = True
        '.Filter = "All Files (*.*)|*.*|*.vbge|*.vbge"
        '.ShowOpen
        'If Len(.FileName) = 0 Then
        '    Exit Sub
        'End If
        'file = .FileName
    'End With
    'file = FileBrowser1.OpenFile(MainWindow)
    'Text1.LoadFile file
End Sub

Private Sub mnuSave1_Click()
    If Me.GameEngine.IsEngineRunning = False Then
        SaveData
    Else
        Dim Choise As VbMsgBoxResult
        Choise = MsgBox("Saving Will Stop GameEngine. Do You Want To Proceed", vbYesNo + vbExclamation)
        If Choise = vbYes Then
            Me.GameEngine.EndGame
            SaveData
        Else
            Exit Sub
        End If
    End If
    
    cChanged = False
End Sub

Sub Save()
    If Me.GameEngine.IsEngineRunning = False Then
        SaveData
    Else
        Dim Choise As VbMsgBoxResult
        Choise = MsgBox("Saving Will Stop GameEngine. Do You Want To Proceed", vbYesNo + vbExclamation)
        If Choise = vbYes Then
            Me.GameEngine.EndGame
            SaveData
        Else
            Exit Sub
        End If
    End If
    
    cChanged = False
End Sub

Private Sub SaveData()
    If Me.Session = ScriptCode Then
        inEditObjectData.Data = Me.Text1.text
    ElseIf Me.Session = MeshData Then
        'Me.GameObject.MeshCode = Text1.text
        inEditObjectData.Data = Text1.text
    ElseIf Me.Session = ObjectCode Then
        'Me.GameObject.MyScript.Data = Text1.text
    Else
        MsgBox "Error Saving Code. Session Info Not Set", vbExclamation
    End If
    Me.Caption = curClassName
End Sub

Sub SaveAs()
    'Dim file As String
    
    'file = FileBrowser1.SaveFile(MainWindow)
    'If Len(file) = 0 Then
        'Exit Sub
    'End If
    'Text1.SaveFile file
    Save
End Sub

Private Sub mnuSaveTo1_Click()
    Dim file As String
    On Error Resume Next
    'With CommonDialog1
        '.DialogTitle = "Save"
        '.CancelError = True
        '.Filter = "All Files (*.*)|*.*|*.vbge|*.vbge"
        '.ShowSave
        'If Len(.FileName) = 0 Then
            'Exit Sub
        'End If
        'file = .FileName
    'End With
    file = FileBrowser1.SaveFile(MainWindow)
    Text1.SaveFile file
End Sub

Private Sub Text1_Change()
    cChanged = True
    colorAdded = False
    Me.Caption = curClassName & " - Changes Made"
    'FindItems
End Sub

Private Function AddColor()
    On Error Resume Next
    Dim OldPos As Integer
    Dim Pos As Integer
    Dim TChanged As Boolean
    
    TChanged = cChanged
    OldPos = Text1.SelStart
    
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.text)
    Text1.SelColor = vbBlack
    
    For Item = 1 To KeyWordsCount
        Pos = 0
        For ILine = 1 To Len(Text1.text)
            Pos = Text1.Find(Items(Item), ILine, , rtfWholeWord)
            If Pos > 0 Then
                Text1.SelStart = Pos
                Text1.SelLength = 1
                Text1.SelText = UCase(Text1.SelText)
                Text1.SelStart = Pos
                Text1.SelLength = Len(Items(Item))
                Text1.SelColor = vbBlue
                ILine = Pos '+ Len(Items(Item)) - 1
            Else
                Exit For
            End If
        Next
    Next
    Text1.SelStart = OldPos
    Text1.SelColor = vbBlack
    colorAdded = True
    cChanged = TChanged
End Function

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    'me.StatusBar1.Panels(4).Text = me.Text1.sel
    Text1.SelColor = vbBlack
    If KeyCode = vbKeyReturn Then
        AddColor
    End If
    
End Sub

Private Sub SelectToolTip(SelectedText As String)
    If SelectedText = "Vector" Or SelectedText = "Position" Or SelectedText = "WorldCenterPos" Or SelectedText = "SScale" Then
        Me.Text1.ToolTipText = SelectedText & " Is Part Of Vector_Class. Avaible Methods: VClone(), NewVector(X,Y,Z),AddVectorV(Vector()),AddVector(X,Y,Z),SetVector(X,Y,Z),SetVectorV(Vector())"
    ElseIf SelectedText = "Console" Then
        Me.Text1.ToolTipText = SelectedText & " Avaible Methods: WriteLine , Clear"
    ElseIf SelectedText = "WriteLine" Then
        Me.Text1.ToolTipText = SelectedText & "(Text)"
    Else
        Me.Text1.ToolTipText = "No ToolTip Avaible For This Method"
    End If
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If Me.mnuAddColorCheck1.Checked And Button = 1 Then
        'AddColor
    'End If
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelectToolTip Me.Text1.SelText
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
    
        Case "TCut1"
        
        Case "TCopy1"
        
        Case "TPaste1"
        
        Case "TRun1"
            Me.GameEngine.RunGameInEditor = False
            Me.GameEngine.StartGame
    End Select
End Sub
