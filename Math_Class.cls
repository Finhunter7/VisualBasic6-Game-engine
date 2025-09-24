VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Math_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Function Approach(CurentNumber As Double, TargetNumber As Double, Step As Double) As Double
    If TargetNumber - Step < CurentNumber Then
        Approach = CurentNumber - Step
    ElseIf TargetNumber + Step > CurentNumber Then
        Approach = CurentNumber + Step
    End If
End Function

