Attribute VB_Name = "Editor_Module"
' Module Test
' Tätä Moduulia Tullaan myöhemmin käyttämään
' Tämä on käyttöjärjestelmän sisään rakennettu ominaisuus
' luo 1 uuden peli-ikkunan käyttäen New avainsanaa
Private CEngine As New MSScriptControl.ScriptControl
Private fControls As New FormControls
Private MainForm As New MainWindow

Sub Main() ' Myöhemmin Tästä Käynistyy Moottori
    ' Näyttää Pelimoottorilogon 1 sekunnin ja sen jälkeen itse ohjelman
    OnApplicationStart
    frmSplash.ShowSplash 1, MainForm
End Sub

Private Sub OnApplicationStart()
    Dim globalCode As String
    Dim tStream As TextStream
    Dim globalCodeFolder As Folder
    Dim moduleFolder As Folder
    Dim CModule As Module
    Dim editorDataPath As String
    Dim FileSystem As New FileSystemObject
    
    CEngine.Language = "VBScript"
    CEngine.AddObject "MainWindow", MainForm, True
    CEngine.AddObject "FormUtils", fControls, True
    editorDataPath = App.Path & "\VBSE"
    On Error GoTo cerror
    If FileSystem.FolderExists(editorDataPath) Then
        If FileSystem.FolderExists(editorDataPath & "\Global") Then
            Set globalCodeFolder = FileSystem.GetFolder(editorDataPath & "\Global")
            For Each IFile In globalCodeFolder.Files
                Set tStream = IFile.OpenAsTextStream(ForReading)
                globalCode = globalCode & tStream.ReadAll() & vbCrLf
                tStream.Close
            Next
            CEngine.AddCode globalCode
        End If
        If FileSystem.FolderExists(editorDataPath & "\Modules") Then
            Set moduleFolder = FileSystem.GetFolder(editorDataPath & "\Modules")
            For Each IFile In moduleFolder.Files
                Set CModule = CEngine.Modules.Add(IFile.Name)
                Set tStream = IFile.OpenAsTextStream(ForReading)
                CModule.AddCode tStream.ReadAll()
                tStream.Close
                CModule.Run "Initialize"
            Next
        End If
    End If
    Exit Sub
    
cerror:
    MsgBox CEngine.Error.Description, vbExclamation, CEngine.Error.Source
    Resume Next
End Sub
