' This Script Is Important To Be Called First Because It Sets The Game Project Information
' Second Most Important Script To Load Is The Global Script Data. See \Scripts\Global\Globals.VBS
Sub Initialize()
	Engine.ProjectName = "Battle Tanks" ' Setting The Project's Name. This Will Be Displayed In GameWindow
	Engine.GameWindow.Width = 11000 
	Engine.GameWindow.Height = 5000
	'Engine.GameWindow.AutoRedraw = True
	Engine.GameWindow.ClipControls = True
	Engine.SetTargetFrameRate 60 '60 ' Setting The Target Framerate. If Set To -1 The Framerate Will Be Uncapped
	Engine.EConsole.Show ' Shows The Console
End sub
