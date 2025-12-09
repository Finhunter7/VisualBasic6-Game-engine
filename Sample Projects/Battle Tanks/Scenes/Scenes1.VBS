
' Add Module Variables Here. These Variables Are Only Visible And Usable For This Module

Sub Initialize() ' Called
	Set ThisScene = Engine.CreateNewScene("MainGame") ' Creates The Scene 
	Set ThisObject = ThisScene.AddObject(New Player,"Player") ' Adding Player GameObject/Class From Global Script
	ThisObject.Position.SetVectorV Vector(1000,Engine.GetWindowCenter().Y,0) ' Setting The Player Position
	Set ThisObject = ThisScene.AddObject(New Enemy1,"Enemy") ' Adding Enemy1 GameObject/Class From Global Script
	ThisObject.Position.SetVectorV Vector(Engine.GetWindowSize().X - 1000,Engine.GetWindowCenter().Y,0) ' Setting The Enemy1 Position
	Engine.AddScene "MainGame" ' Makes The Scene Active
	
	'Set ThisScene = Engine.CreateNewScene("HUD")
	'Set ThisObject = ThisScene.AddObject(New Player,"Player")
	'ThisObject.Position.SetVectorV Vector(3000,200,0)
	'Engine.AddScene "HUD"
End sub