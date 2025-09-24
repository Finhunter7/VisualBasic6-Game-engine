VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EngineRefrence_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Name As String
Public Path As String
Public Refrence As Object
Public RefrenceClass As String
Public Server As String
Public Active As Boolean
Public HoldData As Boolean

Sub Build(TName As String, TPath As String, TRefrenceClass As String, TActive As Boolean)
    Name = TName
    Path = TPath
    RefrenceClass = TRefrenceClass
    Active = TActive
End Sub
