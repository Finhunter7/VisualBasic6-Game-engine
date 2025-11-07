VERSION 5.00
Begin VB.Form NewForm 
   Caption         =   "Form"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "NewForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public OnLoadEvent As New EngineEvent_Class

Private Sub Form_Initialize()
    OnLoadEvent.EventClassName = "Form"
    OnLoadEvent.ExecuteProc = "Load"
End Sub

Private Sub Form_Load()
    OnLoadEvent.Invoke
End Sub

Public Sub SetName(NewName As String)
    OnLoadEvent.EventClassName = NewName
End Sub

