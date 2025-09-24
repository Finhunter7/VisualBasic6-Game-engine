VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GameLoader_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const GameObjectProcsCount = 12
Const SceneObjectProcsCount = 9

Private GameObjectProcs(GameObjectProcsCount) As String
Private SceneObjectProcs(SceneObjectProcsCount) As String

Private GameEngine As EngineClass

Private fileSystemHandler As New FileSystemObject
Enum DataField
    GeneralVars = 0
    Data = 1
    Collection = 2
    DArray = 3
    None = 4
    Object = 5
    Pointers = 6
    method = 7
End Enum

Sub SaveObjectToDiskVBCFormat(SaveData As SavingEnums, fileName As String, EObject As Object)
    Dim curGObject As GameObject_Class
    Dim curScript As Script_Class
    Dim curScene As Scene_Class
    Dim textFile As TextStream
    
    If SaveData = Scene Then
        Set curScene = EObject
        If Not curScene Is Nothing Then
            Set textFile = fileSystemHandler.CreateTextFile(fileName & ".VBScene", True)
            With curScene
                textFile.WriteLine "[General]"
                textFile.WriteLine "Scene_Name " & .Name
                textFile.WriteLine "Scene_Index " & .Index
                textFile.WriteLine "Scene_Layer " & .Layer
                textFile.WriteLine "[Collection]"
                For i = 1 To .Objects.Count
                    textFile.WriteLine "AddObject " & .Objects.Item(i).Name
                Next
                textFile.WriteLine "}"
            End With
            
        Else
            Exit Sub
        End If
    ElseIf SaveData = GameObject Then
        Set curGObject = EObject
        If Not curGObject Is Nothing Then
            Set textFile = fileSystemHandler.CreateTextFile(fileName & ".VBGObject", True)
            With curGObject
                textFile.WriteLine "[General]"
                textFile.WriteLine "Object_Name " & .Name
                textFile.WriteLine "Object_Tag " & .Tag
                textFile.WriteLine "Object_Mass " & .Mass
                textFile.WriteLine "Object_Size " & .Size
                'textFile.WriteLine "Object_Active " & .Active
                'textFile.WriteLine "Object_Visible " & .Visible
                textFile.WriteLine "Object_MeshCode {"
                textFile.WriteLine .GetComponent("Mesh_Code").Data
                textFile.WriteLine "}"
                textFile.WriteLine "[Method]"
                textFile.WriteLine "Object_SetPosition {"
                textFile.WriteLine .Position.X
                textFile.WriteLine .Position.Y
                textFile.WriteLine .Position.Z
                textFile.WriteLine "}"
                textFile.WriteLine "[Objects]"
                textFile.WriteLine "Object_MyScene " & .MyScene.Name
                textFile.WriteLine "[Position]"
                textFile.WriteLine "Position.X " & .Position.X
                textFile.WriteLine "Position.Y " & .Position.Y
                textFile.WriteLine "Position.Z " & .Position.Z
                textFile.WriteLine "[SScale]"
                textFile.WriteLine "X " & .SScale.X
                textFile.WriteLine "Y " & .SScale.Y
                textFile.WriteLine "Z " & .SScale.Z
                textFile.WriteLine "[Components]"
                For i = 1 To .Components.Count
                    textFile.WriteLine "Components.Add " & .Components.Item(i).Name
                Next
                textFile.WriteLine "}"
            End With
        Else
            Exit Sub
        End If
    ElseIf SaveData = Script Then
        Set curScript = EObject
        If Not curScript Is Nothing Then
            Set textFile = fileSystemHandler.CreateTextFile(fileName & ".VBCEScript", True)
            With curScript
                textFile.WriteLine "[General]"
                textFile.WriteLine "Script_Name " & .Name
                textFile.WriteLine "Script_MyModule " & .MyModule
                textFile.WriteLine "[Collection]"
                textFile.WriteLine "}"
                textFile.WriteLine "[Data]"
                textFile.Write .Data
            End With
        Else
            Exit Sub
        End If
    End If
    MsgBox "File Save Complete", vbInformation
    textFile.Close
End Sub
Sub SaveObjectToDiskXML(SaveData As SavingEnums, fileName As String, EObject As Object)
    Dim XMLO As New DOMDocument30
    Dim XMLElement As IXMLDOMElement
    Dim XMLNode As IXMLDOMNode
    Dim XMLTextNode As IXMLDOMText
    Dim curGObject As GameObject_Class
    Dim curScript As Script_Class
    Dim curScene As Scene_Class
    
    
    If SaveData = Scene Then
        Set curScene = EObject
        If Not curScene Is Nothing Then
            With curScene
                
            End With
        Else
            XMLO.abort
            Exit Sub
        End If
    ElseIf SaveData = Script Then
        Set curScript = EObject
        
    ElseIf SaveData = GameObject Then
        Set curGObject = EObject
    Else
    
    End If
    XMLO.Save fileName
    XMLO.abort
End Sub

Function LoadObjectFromDiskVBCFormat(ByRef LoadData As SavingEnums, fileName As String) As Object
    Dim objectFolderPath As String
    Dim tempObject As Object
    Dim dataStream As TextStream
    Dim curGObject As GameObject_Class
    Dim curScript As Script_Class
    Dim curScene As Scene_Class
    Dim curObject As Object
    Dim curArea As DataField
    Dim curLine As String
    Dim curCollection As String
    Dim curArray() As String
    Dim arrayCount As Long
    Dim loopData As String
    Dim methodArgs As Variant
    Dim msgboxChoise As VbMsgBoxResult
    msgboxChoise = vbIgnore
    
    If Not fileSystemHandler.FileExists(fileName) Then
        Exit Function
    End If
    
    If LoadData = Scene Then
        Set curScene = New Scene_Class
        Set dataStream = fileSystemHandler.OpenTextFile(fileName, ForReading)
        
        Do Until dataStream.AtEndOfStream
            curLine = dataStream.ReadLine
            
            If Left(curLine, 1) = "[" Then
                curArea = None
            End If
            
            If curLine = "[General]" Then
                curArea = GeneralVars
                curLine = dataStream.ReadLine
            ElseIf curLine = "[Collection]" Then
                curArea = Collection
                curLine = dataStream.ReadLine
            End If
            If curLine = "}" Then
                curArea = None
            End If
            
            If curArea = GeneralVars Then
                If Not curLine = "" Then
                    curArray = Split(Replace(curLine, "Scene_", ""), " ")
                    arrayCount = 0
                    For Each Item In curArray
                        arrayCount = arrayCount + 1
                        If arrayCount = 2 Then
                            CallByName curScene, curArray(0), VbLet, curArray(1)
                            Exit For
                        End If
                    Next
                End If
            ElseIf curArea = Collection Then
                If Not curLine = "" Then
                    If objectFolderPath = "" And msgboxChoise = vbIgnore Then
                        msgboxChoise = MsgBox("Warning. This Scene Has Refrences To Objects. Do You Want To Load The Objects", vbExclamation + vbYesNoCancel)
                        If msgboxChoise = vbCancel Then
                            Exit Function
                        ElseIf msgboxChoise = vbNo Then
                            curArea = None
                            'Return
                        Else
                            objectFolderPath = FileBrowser1.OpenFolder(MainWindow, "Open Folder")
                            If objectFolderPath = "" Then
                                curArea = None
                                'Return
                            End If
                        End If
                    End If
                    If msgboxChoise = vbYes Then
                        curArray = Split(curLine, " ")
                        arrayCount = 0
                        For Each Item In curArray
                            arrayCount = arrayCount + 1
                            If arrayCount = 2 Then
                                Set tempObject = Me.LoadObjectFromDiskVBCFormat(GameObject, objectFolderPath & "\" & curArray(1) & ".VBGObject")
                                CallByName curScene, curArray(0), VbMethod, GameEngine, tempObject, tempObject.Name
                                Exit For
                            End If
                        Next
                    End If
                End If
            Else
            
            End If
            
        Loop
        
        Set curObject = curScene
    ElseIf LoadData = Script Then
    
    ElseIf LoadData = GameObject Then 'Lataa Peli Objektin
        Set curGObject = New GameObject_Class
        Set dataStream = fileSystemHandler.OpenTextFile(fileName, ForReading)
        
        Do Until dataStream.AtEndOfStream
            curLine = dataStream.ReadLine
            methodArgsCount = 0
            If Left(curLine, 1) = "[" Then
                curArea = None
            End If
            
            If curLine = "[General]" Then
                curArea = GeneralVars
                curLine = dataStream.ReadLine
            ElseIf curLine = "[Collection]" Then
                curArea = Collection
                curLine = dataStream.ReadLine
            ElseIf curLine = "[Objects]" Then
                curArea = Pointers
            ElseIf curLine = "[Method]" Then
                curArea = method
                curLine = dataStream.ReadLine
            End If
            
            If curLine = "}" Then
                curArea = None
            End If
            
             If curArea = GeneralVars Then
                If Not curLine = "" Then
                    curArray = Split(Replace(curLine, "Object_", ""), " ")
                    arrayCount = 0
                    For Each Item In curArray
                        arrayCount = arrayCount + 1
                        If arrayCount = 2 Then
                            If curArray(1) = "Tosi" Or curArray(1) = "True" Then
                                CallByName curGObject, curArray(0), VbLet, True
                            ElseIf curArray(1) = "Epätosi" Or curArray(1) = "False" Then
                                CallByName curGObject, curArray(0), VbLet, False
                            ElseIf curArray(1) = "{" Then
                                Do Until RTrim(curLine) = "}" Or dataStream.AtEndOfStream
                                    curLine = dataStream.ReadLine
                                    loopData = loopData & curLine & vbCrLf
                                Loop
                                CallByName curGObject, curArray(0), VbLet, loopData
                            Else
                                CallByName curGObject, curArray(0), VbLet, curArray(1)
                            End If
                            Exit For
                        End If
                    Next
                End If
            ElseIf curArea = method Then
                
            ElseIf curArea = Pointers Then
            
            Else
            
            End If
        Loop
        Set curObject = curGObject
    Else
        Exit Function
    End If
    dataStream.Close
    Set LoadObjectFromDiskVBCFormat = curObject
End Function

Private Sub Class_Initialize()
    SceneObjectProcs(1) = "Name"
    SceneObjectProcs(2) = "Index"
    SceneObjectProcs(3) = "Layer"
    Set GameEngine = MainWindow.Engine
End Sub
