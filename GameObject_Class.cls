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
Implements IObject_Class

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
Public Transform As New Transform_Class ' Yhdistelm‰ Sijainista , Pyˆrimisest‰ ja Koosta
Public Position As New Vector_Class
Public Rotation As New Matrix_Class
Public SScale As New Vector_Class
Public Size As Double 'Kokonais Mittakoko
Public EditorColliderSize As New Vector_Class ' Editorissa n‰kyv‰ valittsemis alue
Public Velocity As New Vector_Class '
Public Components As New Collection
Public MyPropertites As New PropertyBag ' K‰ytet‰‰n kun objektista tehd‰‰n peliss‰ kopioita
'Public MeshCode As Script_Class ' Referenssi Scriptiin
Private IVisible As Boolean
Private IActive As Boolean
Public text As Variant ' K‰ytet‰‰n Teksti Objekteissa

Private Sub Initialize()
    IVisible = True
    IActive = True
    Me.Mass = 1
    Me.Size = 1
    Me.SScale.SetVector 1, 1, 1
    Me.EditorColliderSize.SetVector 500 * Me.SScale.X, 500 * Me.SScale.Y, 500 * Me.SScale.Z
    Set Me.Transform.Position = Me.Position
    Set Me.Transform.SScale = Me.SScale
    'Me.MeshCode = vbNewLine & "Sub Draw() 'Same As Update But Executed After Update" & vbNewLine & vbNewLine & "End Sub" & vbNewLine & vbNewLine & "Sub EditorDraw()" & vbNewLine & "Draw()" & vbNewLine & "End Sub"
    'Me.MyScript.Data = vbNewLine & "Set This = Me.MyObjects.Item(1)" & vbNewLine & "Sub Load()" & vbNewLine & vbNewLine & "End Sub" & vbNewLine & vbNewLine & "Sub Update()" & vbNewLine & vbNewLine & "End Sub"
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

Function GetAllChildren() As Variant
    Dim newCol As New Collection
    
    For Each Item In Me.Child
        If Not Item.GetAllChildren() Is Nothing Then
            newCol.Add Item.GetAllChildren(), Item.Name
        Else
            MsgBox Item.Name
            newCol.Add Item, Item.Name
            Exit For
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
        Set GetChildren = Nothing
    Else
        Set GetChildren = newCol
    End If
End Function

Sub SetParent(GObject As GameObject_Class)
    Set Me.ParentObject = GObject
    GObject.Child.Add Me, Me.Name
End Sub

Sub RemoveParent()
    Me.ParentObject.Child.Remove Me.Name
    Set Me.ParentObject = Nothing
End Sub

Public Function GetComponent(NameOrIndex) As Object
    On Error Resume Next
    If Me.Components.Count > 0 Then
        Set GetComponent = Me.Components.Item(NameOrIndex)
    Else
        Set GetComponent = Nothing
    End If
End Function

Public Function RemoveComponent(NameOrID) As Boolean
    On Error GoTo DError
    Dim curComp As Object
    
    If Me.Components.Count > 0 Then
        Set curComp = Me.Components.Item(NameOrID)
        Set curComp.MyObject = Nothing
        'curComp.IsComponent = False
        Me.Components.Remove NameOrID
        RemoveComponent = True
    Else
        RemoveComponent = False
    End If
    Exit Function
DError:
    RemoveComponent = False
End Function

Public Sub AddComponent(NComponent As Object)
    Dim newComponent As Object
    'Set NComponent.MyObject = Me
    'Set newComponent = NComponent
    'Set newComponent.MyObject = Me
    Me.Components.Add newComponent, newComponent.Name
End Sub

Private Sub Load(GameCodeEngine As ScriptControl)
    Me.EditorColliderSize.SetVector 500 * Size, 500 * Size, 500 * Size
    Me.ResetPhysics
    'LoadMeshCode GameCodeEngine
    'LoadCode GameCodeEngine
    LoadComponents GameCodeEngine
    Draw GameCodeEngine
End Sub

Sub ResetPhysics()
    Me.Velocity.SetVector 0, 0, 0
End Sub

Sub SetPosition(NewX As Double, Optional NewY As Double, Optional NewZ As Double)
    Me.Position.X = NewX
    Me.Position.Y = NewY
    Me.Position.Z = NewZ
End Sub

Private Sub LoadMeshCode(GameCodeEngine As ScriptControl)
    On Error Resume Next
    'GameCodeEngine.Modules(MyScene.Name & "_" & Me.Name & "_MeshCode").AddCode MeshCode
End Sub

Private Sub LoadCode(GameCodeEngine As ScriptControl)
    On Error Resume Next
    'GameCodeEngine.Modules(MyScene.Name & "_" & Me.Name & "_Code").AddCode MyScript.Data
    'GameCodeEngine.Modules(MyScene.Name & "_" & Me.Name & "_Code").Run "Load"
End Sub

Private Sub Update(GameCodeEngine As ScriptControl)
    Draw GameCodeEngine
    CallUpdate GameCodeEngine
    UpdateComponents GameCodeEngine
    UpdatePhysics
End Sub

Function ChangeMeshCode(NewCode As String, GameCodeEngine As ScriptControl)
    On Error Resume Next
    'Me.MeshCode = NewCode
    'GameCodeEngine.Modules(MyScene.Name & "_" & Me.Name & "_MeshCode").AddCode MeshCode
End Function

Function EditorCallDraw(GameCodeEngine As ScriptControl) ' K‰ytet‰‰n Kun Editorin Tarvitsee Piirt‰‰ Objekti Uudelleen
    On Error Resume Next
    For Each Comp In Me.Components
        If Comp.Tag = "RendererScript" Then
            Comp.MyModule.Run "EditorDraw"
            Exit For
        End If
    Next
End Function

Private Function Draw(GameCodeEngine As ScriptControl)
    On Error Resume Next
    If IVisible Then
        For Each Comp In Me.Components
        If Comp.Tag = "RendererScript" Then
            Comp.MyModule.Run "Draw"
            Exit For
        End If
    Next
    End If
End Function

Private Function CallUpdate(GameCodeEngine As ScriptControl)
    'GameCodeEngine.Modules(MyScene.Name & "_" & Me.Name & "_Code").Run "Update"
End Function

Private Function LoadComponents(CodeEngine As ScriptControl)
    On Error Resume Next
    Dim curComponent As Object
    For i = 1 To Me.Components.Count()
        Set curComponent = Me.Components.Item(i)
        curComponent.Load
    Next
End Function

Private Function UpdateComponents(CodeEngine As ScriptControl)
    On Error Resume Next
    Dim curComponent As Object
    For i = 1 To Me.Components.Count()
        Set curComponent = Me.Components.Item(i)
        curComponent.Update
    Next
End Function
Private Function UpdatePhysics()
    Me.Position.AddVectorV Me.Velocity
End Function

Function EndObject()
    For i = 1 To Me.Components.Count
        Me.RemoveComponent i
    Next
    Me.MyScene.Objects.Remove Me.Name
End Function

Private Sub Class_Initialize()
    Initialize
End Sub

Private Property Let IObject_Class_Active(ByVal RHS As Boolean)
    IActive = RHS
End Property

Private Property Get IObject_Class_Active() As Boolean
    IObject_Class_Active = IActive
End Property

Private Sub IObject_Class_Load(CEngine As MSScriptControl.IScriptControl)
    Load CEngine
End Sub

Private Sub IObject_Class_Update(CEngine As MSScriptControl.IScriptControl)
    Update CEngine
End Sub

Private Property Let IObject_Class_Visible(ByVal RHS As Boolean)
    IVisible = RHS
End Property

Private Property Get IObject_Class_Visible() As Boolean
    IObject_Class_Visible = IVisible
End Property
