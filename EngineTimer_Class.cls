VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EngineTimer_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event Update()
Public Interval As Double
Private IntervalTime As Double
Private IntervalStart As Double
Private Cycles As Long
Private Active As Boolean
Private startTime As Double

Public Sub Start()
    Active = True
    startTime = (Timer * 1000)
    IntervalStart = (Timer * 1000)
    If Me.Interval < 0 Then
        TimingEngineUncapped
    Else
        TimingEngine
    End If
End Sub

Private Sub TimingEngine()
    Do Until Active = False
        If (Timer * 1000) >= startTime Then
            startTime = ((Timer * 1000) + Me.Interval)
            IntervalTime = (Timer * 1000) - IntervalStart
            IntervalStart = (Timer * 1000)
            RaiseEvent Update
            Cycles = 0
        End If
        Cycles = Cycles + 1
        DoEvents
    Loop
End Sub

Private Sub TimingEngineUncapped()
    Do Until Active = False
        startTime = ((Timer * 1000) + Me.Interval)
        IntervalTime = (Timer * 1000) - IntervalStart
        IntervalStart = (Timer * 1000)
        RaiseEvent Update
        DoEvents
    Loop
End Sub

Public Sub Kill()
    Active = False
End Sub

Private Sub Class_Initialize()
    Me.Interval = 16
    Active = False
End Sub

Public Function GetRealIntervalTime() As Double
    If Active Then
        GetRealIntervalTime = IntervalTime
    Else
        GetRealIntervalTime = Interval
    End If
End Function

Public Function IsActive() As Boolean
    IsActive = Active
End Function

Public Function GetCycles()
    If Me.Interval < 0 Then
        GetCycles = 0
    Else
        GetCycles = Cycles
    End If
End Function
