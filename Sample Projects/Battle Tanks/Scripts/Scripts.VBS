
Sub Initialize()
	Set ThisScript = Engine.CreateNewScript("Script1") ' Creates Script Named Script1.
	ThisScript.Data = Engine.GetTextFromFile(Engine.GetGamePath() & "\Scripts\Script1.BNK") ' Adds The Script Code From File
	Set ThisScript = Engine.CreateNewScript("GameManager") ' Creates Script Named GameManager
	ThisScript.Data = Engine.GetTextFromFile(Engine.GetGamePath() & "\Scripts\GameManager.BNK") ' Adds The Script Code From File
end sub
