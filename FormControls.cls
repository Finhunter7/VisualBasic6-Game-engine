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

Function CreateForm(Name As String)
    
End Function

Function AddButton(Form As Form, Name As String)
    Dim CButton As New FormControlsButton_Class
    CButton.Build Name
    Set CButton.Button = Form.Controls.Add("VB.CommandButton", Name)
    Set AddButton = CButton
End Function
