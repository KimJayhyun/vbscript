''' Error Handling
On Error Resume Next

' input_String = WScript.Arguments.item(0)

temp = ""

For i = 1 To len(input_String)
    temp = temp & Hex(ASC(Mid(input_String, i , 1)))
Next

result = ""

For i = 1 To len(temp)   
    result = result & Mid(temp, i , 1)

    If i Mod 2 = 0 Then
        result = result & "%"
    End If
    
Next

result = Mid(result, 1, len(result) - 1)

msgbox result

' WScript.StdOut.WriteLine(result)

' ''' Error Handling
' Err.Clear