VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InternetControls_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents WSock As Winsock
Attribute WSock.VB_VarHelpID = -1
Public WithEvents INetControl As InetCtlsObjects.Inet
Attribute INetControl.VB_VarHelpID = -1
Public Events As New EventManager_Class

Private Sub Class_Initialize()
    Events.ChangeClassName "NetControls"
    Events.CreateEvent "INetControl_StateChanged"
    Events.CreateEvent "WSock_Error"
    Events.CreateEvent "WSock_Close"
    Events.CreateEvent "WSock_Connect"
    Events.CreateEvent "WSock_ConnectionRequest"
    Events.CreateEvent "WSock_DataArrival"
    Events.CreateEvent "WSock_SendComplete"
    Events.CreateEvent "WSock_SendProgress"
End Sub

Private Sub INetControl_StateChanged(ByVal State As Integer)
    Events.Item("INetControl_StateChanged").Invoke State
End Sub

Private Sub WSock_Close()
    Events.Item("WSock_Close").Invoke
End Sub

Private Sub WSock_Connect()
    Events.Item("WSock_Connect").Invoke
End Sub

Private Sub WSock_ConnectionRequest(ByVal requestID As Long)
    Events.Item("WSock_ConnectionRequest").Invoke requestID
End Sub

Private Sub WSock_DataArrival(ByVal bytesTotal As Long)
    Events.Item("").Invoke
End Sub

Private Sub WSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Events.Item("WSock_Error").Invoke Number, Description, Scode, Source, HelpFile, HelpContext, CancelDisplay
End Sub

Private Sub WSock_SendComplete()
    Events.Item("WSock_SendComplete").Invoke
End Sub

Private Sub WSock_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    Events.Item("WSock_SendProgress").Invoke bytesSent, bytesRemaining
End Sub
