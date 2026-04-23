Attribute VB_Name = "VBCEPlayerModule"
Private Engine As New EngineClass
Sub Main()
    If Engine.OpenProject(App.Path & "\Project.VBCEProject", True) Then
        If Engine.ProjectType = VBCEGameProject Then
            Engine.StartEngine
        End If
    Else
        End
    End If
End Sub
