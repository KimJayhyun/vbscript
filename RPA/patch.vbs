Set fso = CreateObject("Scripting.FileSystemObject")
folderPath = FSO.GetAbsolutePathName(".")

version5 = folderPath & "\rpaclient-upgrade-11.0.0.5\upgradeRPAClient.bat"
response5 = folderPath & "\rpaclient-upgrade-11.0.0.5\responsefile\response.properties"

version9 = folderPath & "\rpaclient-upgrade-11.0.0.9\upgradeRPAClient.bat"
response9 = folderPath & "\rpaclient-upgrade-11.0.0.9\responsefile\response.properties"



' Set oShell = CreateObject("wscript.shell")
' Set oExec = oShell.Exec("cmd.exe")

' Set oShell = CreateObject("Shell.Application")
Set oShell = CreateObject("WScript.Shell")
' Set oShellWindows = oShell.Shell_Windows

' oShell.ShellExecute version9," -response " & response9, , "runas", 1
oShell.Run version5 & " -response " & response5, 1, True

oShell.Run version9 & " -response " & response9, 1, True


