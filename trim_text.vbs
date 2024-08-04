a = WScript.Arguments.Count


msgbox a
text = ""

For i = 0 to a - 2
    msgbox WScript.Arguments.item(i)
    temp = Trim(WScript.Arguments.item(i))
    msgbox temp
    ' temp = Replace(temp, chr(10), "")
    text = text & " " & temp
    ' msgbox i
    ' msgbox temp
Next

WScript.StdOut.WriteLine(text)