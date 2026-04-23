VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuForVBS_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private IName As String
Private WithEvents tMenu As Menu
Attribute tMenu.VB_VarHelpID = -1
Private mEvents As New EventManager_Class

Public Property Get Name() As String
    Name = IName
End Property
Friend Property Let Name(RHS As String)
    mEvents.ChangeClassName RHS
    IName = RHS
End Property

Public Property Get Menu() As Menu
    Set Menu = tMenu
End Property
Friend Property Set Menu(RHS As Menu)
    Set tMenu = RHS
End Property

Public Property Get Events() As EventManager_Class
    Set Events = mEvents
End Property

Private Property Set Events(RHS As EventManager_Class)
    Set mEvents = RHS
End Property

Private Sub Class_Initialize()
    mEvents.CreateEvent "Click"
End Sub

Private Sub mMenu_Click()
    mEvents("Click").Invoke
End Sub
