Dim maxX, maxY, X, Y

resX = WScript.Arguments.Item(0)
resY = WScript.Arguments.Item(1)

maxX = (resX * 15 - 1500) * 0.8
maxY = (resY * 15 - 1500) * 0.8

Do
    ' box = MsgBox("Never gonna let you down", 0, "Never gonna give you up")
    Randomize
    X = Int((maxX+1)*Rnd)

    Randomize
    Y = Int((maxY+1)*Rnd)

    box = InputBox("Never gonna let you down", "Never gonna give you up", "Never gonna run around an desert you", X, Y)
Loop