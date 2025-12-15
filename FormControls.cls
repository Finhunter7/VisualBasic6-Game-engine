VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function CreateForm(Name As String) As FormForVBS_Class
    Dim CForm As New FormForVBS_Class
    CForm.Name = Name
    Set CForm.Form = New NewForm
    Set CreateForm = CForm
End Function

Function AddLabel(Form As Form, Name As String) As Label
    Dim CLabel As Label
    Set CLabel = Form.Controls.Add("VB.Label")
    Set AddLabel = CLabel
End Function

Function AddFrame(Form As Form, Name As String) As Frame
    Dim CFrame As Frame
    Set CFrame = Form.Controls.Add("VB.Frame")
    Set AddFrame = CFrame
End Function

Function AddMenu(Form As Form, Name As String) As Menu
    Dim NMenu As Menu
    
    Set NMenu = Form.Controls.Add("VB.Menu")
    Set AddMenu = NMenu
End Function

Function AddButton(Form As Form, Name As String) As FormControlsButton_Class
    Dim CButton As New FormControlsButton_Class
    CButton.Build Name
    Set CButton.Button = Form.Controls.Add("VB.CommandButton", Name)
    Set AddButton = CButton
End Function

Function AddTextBox(Form As Form, Name As String) As FormControlsTextbox_Class
    Dim CTextBox As New FormControlsTextbox_Class
    CTextBox.Build Name
    Set CTextBox.TextBox = Form.Controls.Add("VB.TextBox", Name)
    Set AddTextBox = CTextBox
End Function

Function AddUserControl(Form As Form, Name As String) As Object
    Dim CUserControl As Object
    Set CUserControl = Form.Controls.Add("VB.UserControl", Name)
    Set AddUserControl = CUserControl
End Function
