VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Component_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Name As String
Private Owner As Object
Private MyObject As Object
'Private ComponentScriptData As String
Sub SetObject(TObject As Object)
    Set MyObject = TObject
End Sub

Sub Load(GameObject)
    Set Owner = GameObject
    MyObject.Load
End Sub

Sub Update()
    MyObject.Update
End Sub

