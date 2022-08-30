Dim oPlayer
Set oPlayer = CreateObject("WMPlayer.OCX")

oPlayer.URL = "rickroll.mp3"

Do
  oPlayer.controls.play
  While oPlayer.playState <> 1
    WScript.Sleep 100
  Wend

  oPlayer.close
Loop