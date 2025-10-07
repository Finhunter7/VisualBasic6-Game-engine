VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AudioSystem_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private GameEngine As EngineClass
Private Declare Function CPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal UFlags As Long) As Long
Private Declare Function DirectSound Lib "Dsound" Alias "DirectSoundCreate8" (guid As String) As Object

Sub Build(Engine As EngineClass)
    Set GameEngine = Engine
End Sub

Public Function EmitSound(FileName As String, UFlags As Long)
    CPlaySound FileName, UFlags
End Function

Public Function DirectCreateSoundDevice(guid As String) As Object
    Set DirectCreateSoundDevice = DirectSound(guid) 'GameEngine.Direct8Lib.DirectSoundCreate(guid)
End Function

Public Function CreateSoundBuffer(Device As Object, FileName As String, bufferDesc) As Object
    Set CreateSoundBufferFromFile = Device.CreateSoundBufferFromFile(FileName, bufferDesc)
End Function

Public Sub PlayBuffer(Buffer As Object, Flags)
    Buffer.Play Flags
End Sub

Public Sub StopBuffer(Buffer As Object)
    Buffer.Stop
End Sub
