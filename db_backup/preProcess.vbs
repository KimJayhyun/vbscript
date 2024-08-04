''' Error Handling
On Error Resume Next

compID = Wscript.arguments.Item(0)
today = Wscript.arguments.Item(1)

set objFSO = createobject("Scripting.Filesystemobject")

''''''''''' Check CSV File Exists '''''''''''

''' Assign File Path '''

backupCSVFolder = "C:\ProgramData\MySQL\MySQL Server 5.7\Uploads"

dataBackup = backupCSVFolder & "\data-" & today & ".csv"
keywordBackup = backupCSVFolder & "\keyword-" & today & ".csv"

if objFSO.fileExists(dataBackup) Then
	objFSO.deleteFile(dataBackup)
End If

if objFSO.fileExists(keywordBackup) Then
	objFSO.deleteFile(keywordBackup)
End If

''' Error Handling
Err.Clear