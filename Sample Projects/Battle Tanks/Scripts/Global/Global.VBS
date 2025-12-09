' This Is The Second Most Important Script To Load. If You Have Your Own Classes or Constants Ex: Vector2 Or Const Gravity = 9.81
Private ScriptsData
Private GlobalObjectsFolder
Private DataStream
Sub Initialize()
	ScriptsData = ScriptsData & Engine.GetTextFromFile(Engine.GetGamePath() & "\Scripts\Global\Global.BNK")
	If not Engine.IsInEditor then
		Set GlobalObjectsFolder = Engine.GetFolder(Engine.GetGamePath() & "\Objects\Global")
		for each item In GlobalObjectsFolder.Files
			Set DataStream = Item.OpenAsTextStream(1,-2)
			ScriptsData = ScriptsData & DataStream.ReadAll() & vbCrlf
			DataStream.Close
		next
	end if
	'Engine.SetGlobalScriptData(Engine.GetTextFromFile(Engine.GetGamePath() & "\Scripts\Global\Global.BNK")) ' Sets The Global Script Data
	Engine.SetGlobalScriptData(ScriptsData)
End Sub