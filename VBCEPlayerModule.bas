Attribute VB_Name = "VBCEPlayerModule"
Private Engine As New EngineClass
Sub Main()
    If Engine.OpenProject(App.Path & "\Project.VBCEProject", True) Then
       Engine.StartEngine
    End If
End Sub
