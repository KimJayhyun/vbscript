compID = Wscript.arguments.Item(0)
today = Wscript.arguments.Item(1)

''' Error Handling
On Error Resume Next


set objFSO = createobject("Scripting.Filesystemobject")

''''''''''' Check And Move CSV Backup Files '''''''''''
folderName = "D:\RPA\DB_Backup\Backup\public_" & compID

backupCSVFolder = "C:\ProgramData\MySQL\MySQL Server 5.7\Uploads"

''' Assign File Path '''

' rpa_literature_data
dataBackup = backupCSVFolder & "\data-" & today & ".csv"
moveDataDir = folderName & "\" & today & "\data-" & today & ".csv"

' rpa_literature_keyword
keywordBackup = backupCSVFolder & "\keyword-" & today & ".csv"
moveKeywordDir = folderName & "\" & today & "\keyword-" & today & ".csv"


''' Check <rpa_literature_data> Backup csv file exist
if objFSO.fileExists(dataBackup) Then
    dataExist = Round(objFSO.GetFile(dataBackup).Size / 1024) & " KB"
    objFSO.copyfile dataBackup, moveDataDir, True
    objFSO.deleteFile dataBackup
Else 
    dataExist = "백업 실패"
End If

''' Check <rpa_literature_Keyword> Backup csv file exist
if objFSO.fileExists(keywordBackup) Then
    keywordExist = Round(objFSO.GetFile(keywordBackup).Size / 1024) & " KB"
    objFSO.copyfile keywordBackup, moveKeywordDir, True
    objFSO.deleteFile keywordBackup 
Else  
    keywordExist = "백업 실패"
End If 

''''''''''' Check Dump Backup File Exist '''''''''''

dumpFile = folderName & "\" & today & "\dump-" & today & ".sql"

if objFSO.fileExists(dumpFile) Then
    dumpExist = Round(objFSO.GetFile(dumpFile).Size / 1024) & " KB"
Else
    dumpExist = "백업 실패"
End If

''''''''''' Delete Old Files '''''''''''

''' Delete untill 2(<numBackup>) folders left '''
Set objFolder = objFSO.GetFolder(folderName)
Set folders = objFolder.SubFolders

numFolder = folders.Count

numBackup = 2

If numFolder > numBackup Then
''' Assign Folder Array '''
    Dim arr()
    ReDim arr(numFolder - 1)    
    i = 0
    For each ele in folders
        arr(i) = ele
        i = i + 1
    Next

''' Sort Array By Create Day '''    
    for a = UBound(arr) - 1 To 0 Step -1
        for j = 0 to a
            ' fj = objFSO.GetFolder(arr(j)).DateCreated
            ' fjj = objFSO.GetFolder(arr(j+1)).DateCreated
            fj = arr(j)
            fjj = arr(j + 1)
            If fj > fjj Then
                temp = arr(j)
                arr(j+1) = arr(j)
                arr(j) = temp
            End IF
        Next
    Next
    
''' Delete Old Folder '''    
    for t = 0 to numFolder - numBackup - 1
        tmp = arr(t)
        ' objFSO.MoveFolder tmp, "D:\RPA\DB_Backup\Backup\test11111"
	    objFSO.DeleteFolder(tmp)
    Next

End If

''''''''''' Output Result '''''''''''
IF compID = "super" or compID = "template" Then
    message = "<public_" & compID  & ">" & chr(13) & _
    "   - DB" & chr(13) & _
    "     Dump : " & dumpExist
Else
    message = "<public_" & compID  & ">" & chr(13) & _
    "   - DB" & chr(13) & _
    "     Dump : " & dumpExist & chr(13) & _
    "   - TABLE" & chr(13) & _
    "     rpa_literature_data : " & dataExist & chr(13) & _
    "     rpa_literature_keyword : " & keywordExist
End IF

Wscript.StdOut.WriteLine(message)
''' Error Handling
Err.clear
