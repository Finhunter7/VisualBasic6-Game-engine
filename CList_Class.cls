VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CList_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private LCol As New Collection
Private LVarType As VbVarType

Sub SetListType(LType As VbVarType)
    If LCol.Count < 1 Then
        LVarType = LType
    End If
End Sub

Sub Remove(NameOrID)
    LCol.Remove NameOrID
End Sub

Sub RemoveAll()
    For i = 1 To LCol.Count
        LCol.Remove i
    Next
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = 0
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = LCol.[_NewEnum]
End Property

Sub Add(NItem, Optional Key)
    If VarType(NItem) = LVarType Then
        LCol.Add NItem, Key
    Else
        Err.Raise 13
    End If
End Sub
Sub Insert(NItem, Index, Optional Key)
    If VarType(NItem) = LVarType Then
        LCol.Add NItem, Key, Index
    Else
        'Err.Raise 13
    End If
End Sub

Function Item(NameOrID)
    If LCol.Count > 0 Then
        If LVarType = vbObject Then
            Set Item = LCol.Item(NameOrID)
        Else
            Item = LCol.Item(NameOrID)
        End If
    End If
End Function

Public Property Get Count() As Long
    Count = LCol.Count
End Property

Function CountInstances(Data, Optional InObjectVar) As Long
    Dim items As Long
    If LVarType = vbObject Then
        For i = 1 To LCol.Count
            If CallByName(LCol.Item(i), InObjectVar, VbGet) = Data Then
                items = i
            End If
        Next
    Else
        For i = 1 To LCol.Count
            If LCol.Item(i) = Data Then
                items = i
            End If
        Next
    End If
    CountInstances = items
End Function

Function GetInstance(InstanceData, Optional InObjectVar) As Variant
    If LVarType = vbObject Then
        For i = 1 To LCol.Count
            If CallByName(LCol.Item(i), InObjectVar, VbGet) = Data Then
                Set GetInstance = LCol.Item(i)
                Exit For
            End If
        Next
    Else
        For i = 1 To LCol.Count
            If LCol.Item(i) = Data Then
                GetInstance = LCol.Item(i)
                Exit For
            End If
        Next
    End If
End Function

Function Find(text) As Long
    Dim curText As String
    If LCol.Count > 0 Then
        If LVarType = vbString Then
            For i = 1 To LCol.Count
                curText = LCol.Item(i)
                If curText = text Then
                    Find = i
                    Exit Function
                End If
            Next
            Find = -1
        Else
            Find = -1
        End If
    Else
        Find = -1
    End If
End Function
