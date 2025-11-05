VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Matrix_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private IMatrix(2, 2, 2) As Double
'Public Matrix(4, 4) As Double

Private Enum AxisEnum
    AX = 0
    AY = 1
    AZ = 2
End Enum

Public Sub Rotate(Almount As Double, Axis)

End Sub

Public Function To_Euler()
    
End Function

Public Function To_Vector()

End Function
