''' Error Handling
On Error Resume Next

startTime = WScript.Arguments.Item(1)
endTime = WScript.Arguments.Item(3)

s_H = Mid(startTime, 1, 2)
s_M = Mid(startTime, 4, 2)
s_S = Mid(startTime, 7, 2)

e_H = Mid(endTime, 1, 2)
e_M = Mid(endTime, 4, 2)
e_S = Mid(endTime, 7, 2)

r_S = e_S - s_S
If (r_S < 0) Then
    r_S = r_S + 60
    e_M = e_M - 1
End If

r_M = e_M - s_M
If(r_M < 0) Then
    r_M = r_M + 60
    e_H = e_H - 1
End If

r_H = e_H - s_H

result = ""

If (r_S <> 0) Then
    result = result & r_S & "초"
End If

If (r_M <> 0) Then
    result = r_M & "분 " & result
End If

If (r_H <> 0) Then
    result = r_H & "시간 " & result
End If

WScript.StdOut.WriteLine(result)

''' Error Handling
Err.Clear