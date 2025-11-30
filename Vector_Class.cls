VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Vector_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public X As Double
Public Y As Double
Public Z As Double

Function To_Array()
Attribute To_Array.VB_Description = "Converts Vector To Array"
    Dim TArray(2) As Double
    TArray(0) = X
    TArray(1) = Y
    TArray(2) = Z
    To_Array = TArray
End Function

Function To_Collection()
Attribute To_Collection.VB_Description = "Converts Vector To Collection"
    Dim NCol As New Collection
    NCol.Add X
    NCol.Add Y
    NCol.Add Z
    Set To_Collection = NCol
End Function

Function DistanceToVector(TargetVector As Vector_Class) As Double
    Distance = Sqr(((TargetVector.X - Me.X) * (TargetVector.X - Me.X)) + ((TargetVector.Y - Me.Y) * (TargetVector.Y - Me.Y)) + ((TargetVector.Z - Me.Z) * (TargetVector.Z - Me.Z)))
End Function

Function VClone() As Vector_Class ' Luo Kopion Vektorista
    Dim CVector As New Vector_Class
    CVector.X = Me.X
    CVector.Y = Me.Y
    CVector.Z = Me.Z
    Set VClone = CVector
End Function

Function NewVector(Optional NX As Double, Optional NY As Double, Optional NZ As Double) As Vector_Class 'Luo Uuden Vektorin
Attribute NewVector.VB_Description = "Returns A New Vector Instance"
Attribute NewVector.VB_UserMemId = 0
    Dim NewVectorInstance As New Vector_Class
    NewVectorInstance.X = NX
    NewVectorInstance.Y = NY
    NewVectorInstance.Z = NZ
    Set NewVector = NewVectorInstance
End Function

Function AddVectorV(Vector As Object) As Vector_Class ' Yhdist‰‰ Vektorin Arvot Ja Palauttaa Muutuneen Vektorin
    Me.X = Me.X + Vector.X
    Me.Y = Me.Y + Vector.Y
    Me.Z = Me.Z + Vector.Z
    Set AddVectorV = Me
End Function
Function AddVector(Optional AddX As Double, Optional AddY As Double, Optional AddZ As Double) As Vector_Class
    Me.X = Me.X + AddX
    Me.Y = Me.Y + AddY
    Me.Z = Me.Z + AddZ
    Set AddVector = Me
End Function

Function SetVector(Optional NewX As Double, Optional NewY As Double, Optional NewZ As Double) As Vector_Class
    Me.X = NewX
    Me.Y = NewY
    Me.Z = NewZ
    Set SetVector = Me
End Function

Function SetV(NVector As Object) As Vector_Class
    Me.X = NVector.X
    Me.Y = NVector.Y
    Me.Z = NVector.Z
    Set SetV = Me
End Function

Sub SetVectorV(NVector As Object) ' Legacy feature this will be removed in future updates
    SetV NVector
End Sub

