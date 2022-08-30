Dim objShell
Set objShell = Wscript.CreateObject("WScript.Shell")

objShell.Run "rickrollSound.vbs"

Res = GetScreenResolution
ResX = Res(0)
ResY = Res(1)

Function GetScreenResolution()
	Set oIE = CreateObject("InternetExplorer.Application")
	With oIE
		.Navigate("about:blank")
		Do Until .readyState = 4: WScript.Sleep 100: Loop
			width = .document.ParentWindow.screen.width
			height = .document.ParentWindow.screen.height
	End With

	oIE.Quit
	GetScreenResolution = array(width,height)
End Function

For i = 1 to 30
    objShell.Run "rickrollDialog.vbs " & ResX & " " & ResY
    ' WScript.Sleep 100
Next

Set objShell = Nothing