Sub Move_FolderFolder()
    FolderName = "C:\A2OMSP\Test\AIDA_Data\Test\Marketing\ModelTest\2-2\run_01"
    dest = "C:\A2OMSP\Test\AIDA_Data\Test\Marketing\ModelTest\2-2\run_02"
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(FolderName)
    Set Folders = objFolder.subFolders
    For each ele in Folders
        file_split = Split(ele,"\")
        last_num = Ubound(file_split) 
        folderName = file_split(last_num)
     
        If (folderName Mod 2 = 0) Then 
            'Msgbox folderName
             objFSO.MoveFolder ele , dest  & "\" & folderName
        End If 
 
    Next
    

End Sub 

call Move_FolderFolder