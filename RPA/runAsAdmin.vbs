Set shl = WScript.CreateObject("WScript.Shell")

' Yes : 6, No : 7, Cancel : 2
Res = MsgBox ("Is This Client??", vbYesNoCancel, "Hello World")

If (Res = 6) Then
    currentPath = shl.CurrentDirectory & "\client"
    vbsPath = currentPath & "\patch.vbs"
ElseIf (Res = 7) Then
    currentPath = shl.CurrentDirectory & "\server"
    vbsPath = currentPath & "\patch.vbs"
Else
    vbsPath = "exit"
End If

Set oShell = CreateObject("Shell.Application")
oShell.ShellExecute "cmd.exe", "/c cd " & currentPath & " & " & vbsPath , , "runas", 1
