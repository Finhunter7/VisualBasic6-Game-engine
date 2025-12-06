VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GameObject_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Implements IObject_Class

Public Name As String 'Read Only
Attribute Name.VB_VarUserMemId = 0
Attribute Name.VB_VarProcData = ";Text"
Attribute Name.VB_VarDescription = "Name Of The GameObject"
Public SceneHiearchyPath As String
Public Tag As String
Public MyScene As Scene_Class ' Asetetaan kun objekti luodaan
Public Mass As Double
Public ParentObject As GameObject_Class
Public Child As New Collection ' of GameObject_Class
'Public Transform As New Transform_Class ' Yhdistelm‰ Sijainista , Pyˆrimisest‰ ja Koosta
Private IPosition As New Vector_Class
Public ScenePosition As New Vector_Class ' Scene's Camera position affects this
Public Rotation As New Matrix_Class
Public SScale As New Vector_Class
Public Size As Double 'Kokonais Mittakoko
Public EditorColliderSize As New Vector_Class ' Editorissa n‰kyv‰ valittsemis alue
Public Velocity As New Vector_Class '
Public MyComponents As New Collection ' Kaikki Komponentit Jotka Objekti luo
Public Components As New Collection
Public MyPropertites As New PropertyBag ' K‰ytet‰‰n kun objektista tehd‰‰n peliss‰ kopioita
Public Visible As Boolean
Public Active As Boolean
Public text As Variant ' K‰ytet‰‰n Teksti Objekteissa

Public Property Get Position() As Vector_Class
    Set Position = IPosition
End Property

Private Sub Initialize()
    Visible = True
    Active = True
    Me.Mass = 1
    Me.Size = 1
    Me.SScale.SetVector 1, 1, 1
    Me.EditorColliderSize.SetVector 500 * Me.SScale.X, 500 * Me.SScale.Y, 500 * Me.SScale.Z
    'Set Me.Transform.Position = Me.Position
    'Set Me.Transform.SScale = Me.SScale
    'Me.MeshCode = vbNewLine & "Sub Draw() 'Same As Update But Executed After Update" & vbNewLine & vbNewLine & "End Sub" & vbNewLine & vbNewLine & "Sub EditorDraw()" & vbNewLine & "Draw()" & vbNewLine & "End Sub"
    'Me.MyScript.Data = vbNewLine & "Set This = Me.MyObjects.Item(1)" & vbNewLine & "Sub Load()" & vbNewLine & vbNewLine & "End Sub" & vbNewLine & vbNewLine & "Sub Update()" & vbNewLine & vbNewLine & "End Sub"
End Sub

Public Sub BuildFromObject(GObject As GameObject_Class)
    With GObject
        Me.Name = .Name
        Me.SceneHiearchyPath = .SceneHiearchyPath
        Me.Tag = .Tag
        Set Me.MyScene = .MyScene
        Me.Mass = .Mass
        Set Me.ParentObject = .ParentObject
        For Each Item In .Child
            Me.Child.Add Item, Item.Name
        Next
        Me.Position.SetVectorV .Position
        Me.SScale.SetVectorV .SScale
        Me.Size = .Size
        Me.EditorColliderSize.SetVectorV .EditorColliderSize
        For Each Item In .Components
            Me.Components.Add Item, Item.Name
        Next
        Me.text = .text
    End With
End Sub

Public Function HasProperty(Name) As Boolean
    On Error GoTo end1
    Dim Information As Variant
    
    Information = MyPropertites.ReadProperty(Name)
    
    HasProperty = True
    Exit Function
end1:
    HasProperty = False
    Exit Function
End Function

Function GetAllChildren(IsActivator As Boolean) As Collection
    Dim newCol As New Collection
    Dim curCol As Collection
    
    For Each Item In Me.Child
        If Item.GetAllChildren() Is Nothing Then
            newCol.Add Item, Item.Name
        Else
            newCol.Add Item.GetAllChildren(False), Item.Name
        End If
    Next
    
    If newCol.Count < 1 Then
        Set GetAllChildren = Nothing
    Else
        Set GetAllChildren = newCol
    End If
    
End Function

Function GetMyChildren() As Collection
    Dim newCol As New Collection
    
    For Each Item In Me.Child
        newCol.Add Item, Item.Name
    Next
    If newCol.Count < 1 Then
        Set GetMyChildren = Nothing
    Else
        Set GetMyChildren = newCol
    End If
End Function

Sub SetParent(GObject As GameObject_Class)
    Set Me.ParentObject = GObject
    GObject.Child.Add Me, Me.Name
    Me.MyScene.Objects.Remove Me.Name
End Sub

Sub RemoveParent()
    Me.ParentObject.Child.Remove Me.Name
    Set Me.ParentObject = Nothing
End Sub

Public Function GetComponent(NameOrIndex) As Object
    On Error Resume Next
    If MyComponents.Count > 0 Then
        Set GetComponent = MyComponents.Item(NameOrIndex)
    Else
        Set GetComponent = Nothing
    End If
End Function

Public Function GetComponents() As Collection
    Set GetComponents = MyComponents
End Function

Public Function GetRuntimeComponent(NameOrID) As Object
    On Error Resume Next
    
    If Me.Components.Count > 0 Then
        Set GetRuntimeComponent = Me.Components.Item(NameOrID)
    Else
        Set GetRuntimeComponent = Nothing
    End If
End Function

Public Function RemoveComponent(NameOrID) As Boolean
    On Error GoTo DError
    Dim curComp As Object
    
    If MyComponents.Count > 0 Then
        Set curComp = MyComponents.Item(NameOrID)
        Set curComp.MyObject = curComp
        'curComp.IsComponent = False
        MyComponents.Remove NameOrID
        RemoveComponent = True
    Else
        RemoveComponent = False
    End If
    Exit Function
DError:
    RemoveComponent = False
End Function

Public Sub AddComponent(NComponent As Script_Class)
    On Error Resume Next
    Dim newComponent As Object
    'Set NComponent.MyObject = Me
    'Set newComponent = NComponent
    'Set newComponent.MyObject = Me
    Set newComponent = NComponent.MyModule.Eval("Component")
    If newComponent Is Nothing And Me.MyScene.GameEngine.IsEngineRunning = True Then
        Err.Raise 8, Me.Name & ".AddComponent", "Script " & NComponent.Name & " Is Not Component Supported"
    Else
        MyComponents.Add NComponent, NComponent.Name
        Debug.Print "Added Component " & NComponent.Name
    End If
End Sub

Private Sub Load()
    Me.EditorColliderSize.SetVector 500 * Size, 500 * Size, 500 * Size
    Me.ResetPhysics
    'LoadMeshCode GameCodeEngine
    'LoadCode GameCodeEngine
    'InitializeComponents
    LoadComponents 'GameCodeEngine
    Draw 'GameCodeEngine
End Sub

Sub ResetPhysics()
    Me.Velocity.SetVector 0, 0, 0
End Sub

Sub SetPosition(NewX As Double, Optional NewY As Double, Optional NewZ As Double)
    Me.Position.X = NewX
    Me.Position.Y = NewY
    Me.Position.Z = NewZ
End Sub

Private Sub LoadMeshCode()
    On Error Resume Next
    'GameCodeEngine.Modules(MyScene.Name & "_" & Me.Name & "_MeshCode").AddCode MeshCode
End Sub

Private Sub LoadCode()
    On Error Resume Next
    'GameCodeEngine.Modules(MyScene.Name & "_" & Me.Name & "_Code").AddCode MyScript.Data
    'GameCodeEngine.Modules(MyScene.Name & "_" & Me.Name & "_Code").Run "Load"
End Sub

Private Sub Update()
    Draw 'GameCodeEngine
    'CallUpdate 'GameCodeEngine
    UpdateComponents 'GameCodeEngine
    UpdatePhysics
End Sub

Sub EditorCallDraw() ' K‰ytet‰‰n Kun Editorin Tarvitsee Piirt‰‰ Objekti Uudelleen
    On Error Resume Next
    For Each Comp In Me.Components
        If Comp.Tag = "RendererScript" Then
            Comp.MyModule.Run "EditorDraw"
            Exit For
        End If
        'Comp.EditorDraw
    Next
End Sub

Sub Draw()
    On Error Resume Next
    If Visible Then
        For Each Comp In Me.Components
            If Comp.Tag = "RendererScript" Then
                Comp.MyModule.Run "Draw"
                Exit For
            End If
            'Comp.Draw
        Next
    End If
End Sub

Private Sub InitializeComponents()
    On Error Resume Next
    Dim curComponent As Object
    For Each Item In MyComponents
        Set curComponent = Item.MyModule.Eval("Component")
        Me.Components.Add curComponent, curComponent.Name
        Set curComponent.self = Me
    Next
End Sub

Private Function LoadComponents()
    On Error Resume Next
    For Each Item In Me.Components
        Item.Load
    Next
End Function

Private Function UpdateComponents()
    On Error Resume Next
    For Each Item In Me.Components
        Item.Update
    Next
End Function
Private Function UpdatePhysics()
    If Me.ParentObject Is Nothing Then
        Me.Position.AddVectorV Me.Velocity
    Else
        Me.Velocity.SetVectorV Me.ParentObject.Velocity
        'Me.Position.SetVectorV Me.ParentObject.Position
        'Me.Position.SetVector Me.ParentObject.Position.X
    End If
End Function

Function EndObject() As Boolean
    For i = 1 To Me.Components.Count
        Components.Remove i
    Next
    If Me.ParentObject Is Nothing Then
        If Me.MyScene Is Nothing Then
            EndObject = False
            Exit Function
        Else
            Me.MyScene.Objects.Remove Me.Name
            Set Me.MyScene = Nothing
            EndObject = True
        End If
    Else
        Me.ParentObject.Child.Remove Me.Name
        Set Me.MyScene = Nothing
        Set Me.ParentObject = Nothing
        EndObject = True
    End If
End Function

Private Sub Class_Initialize()
    Initialize
End Sub
