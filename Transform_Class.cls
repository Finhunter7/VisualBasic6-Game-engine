VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Transform_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Position As Vector_Class
'public Rotation as Rot_Class
Public SScale As Vector_Class

Public Function Translate(AddVector As Vector_Class)
    Me.Position.X = Me.Position.X + AddVector.X
    Me.Position.Y = Me.Position.Y + AddVector.Y
    Me.Position.Z = Me.Position.Z + AddVector.Z
End Function

Public Function Move(Optional X As Double, Optional Y As Double, Optional Z As Double)
    Me.Position.X = Me.Position.X + X
    Me.Position.Y = Me.Position.Y + Y
    Me.Position.Z = Me.Position.Z + Z
End Function

