VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextObject_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements GameObject_Class

Private Property Get GameObject_Class_Active() As Boolean

End Property

Private Property Let GameObject_Class_Active(ByVal RHS As Boolean)

End Property

Private Sub GameObject_Class_AddComponent(NComponent As Object)

End Sub

Private Function GameObject_Class_ChangeMeshCode(NewCode As String, GameCodeEngine As MSScriptControl.IScriptControl) As Variant

End Function

Private Property Set GameObject_Class_Child(ByVal RHS As Collection)

End Property

Private Property Get GameObject_Class_Child() As Collection

End Property

Private Property Get GameObject_Class_Components() As Collection

End Property

Private Property Set GameObject_Class_Components(ByVal RHS As Collection)

End Property

Private Function GameObject_Class_EditorCallDraw(GameCodeEngine As MSScriptControl.IScriptControl) As Variant

End Function

Private Property Get GameObject_Class_EditorColliderSize() As Vector_Class

End Property

Private Property Set GameObject_Class_EditorColliderSize(ByVal RHS As Vector_Class)

End Property

Private Function GameObject_Class_EndObject() As Variant

End Function

Private Function GameObject_Class_GetAllChildren() As Object

End Function

Private Function GameObject_Class_GetComponent(NameOrIndex As Variant) As Object

End Function

Private Function GameObject_Class_HasProperty(Name As Variant) As Boolean

End Function

Private Sub GameObject_Class_Load(GameCodeEngine As MSScriptControl.IScriptControl)

End Sub

Private Property Get GameObject_Class_Mass() As Double

End Property

Private Property Let GameObject_Class_Mass(ByVal RHS As Double)

End Property

Private Property Get GameObject_Class_MyPropertites() As PropertyBag

End Property

Private Property Set GameObject_Class_MyPropertites(ByVal RHS As PropertyBag)

End Property

Private Property Get GameObject_Class_MyScene() As Scene_Class

End Property

Private Property Set GameObject_Class_MyScene(ByVal RHS As Scene_Class)

End Property

Private Property Let GameObject_Class_Name(ByVal RHS As String)

End Property

Private Property Get GameObject_Class_Name() As String

End Property

Private Property Set GameObject_Class_ParentObject(ByVal RHS As GameObject_Class)

End Property

Private Property Get GameObject_Class_ParentObject() As GameObject_Class

End Property

Private Property Set GameObject_Class_Position(ByVal RHS As Vector_Class)

End Property

Private Property Get GameObject_Class_Position() As Vector_Class

End Property

Private Function GameObject_Class_RemoveComponent(NameOrID As Variant) As Boolean

End Function

Private Sub GameObject_Class_ResetPhysics()

End Sub

Private Property Let GameObject_Class_SceneHiearchyPath(ByVal RHS As String)

End Property

Private Property Get GameObject_Class_SceneHiearchyPath() As String

End Property

Private Sub GameObject_Class_SetPosition(NewX As Double, Optional NewY As Double, Optional NewZ As Double)

End Sub

Private Property Let GameObject_Class_Size(ByVal RHS As Double)

End Property

Private Property Get GameObject_Class_Size() As Double

End Property

Private Property Set GameObject_Class_SScale(ByVal RHS As Vector_Class)

End Property

Private Property Get GameObject_Class_SScale() As Vector_Class

End Property

Private Property Let GameObject_Class_Tag(ByVal RHS As String)

End Property

Private Property Get GameObject_Class_Tag() As String

End Property

Private Property Set GameObject_Class_text(RHS As Variant)

End Property

Private Property Let GameObject_Class_text(RHS As Variant)

End Property

Private Property Get GameObject_Class_text() As Variant

End Property

Private Property Set GameObject_Class_Transform(ByVal RHS As Transform_Class)

End Property

Private Property Get GameObject_Class_Transform() As Transform_Class

End Property

Private Sub GameObject_Class_Update(GameCodeEngine As MSScriptControl.IScriptControl)

End Sub

Private Property Set GameObject_Class_Velocity(ByVal RHS As Vector_Class)

End Property

Private Property Get GameObject_Class_Velocity() As Vector_Class

End Property

Private Property Let GameObject_Class_Visible(ByVal RHS As Boolean)

End Property

Private Property Get GameObject_Class_Visible() As Boolean

End Property
