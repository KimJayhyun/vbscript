Sub Move_FileRename(FolderName)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(FolderName)
    Set Files = objFolder.Files
    For each ele in Files 
        file_split = Split(ele,"\")
        last_num = Ubound(file_split) 
        fileName = file_split(last_num)
        new_fileName = Replace(fileName, ".png", "_part2.png")
        objFSO.MoveFile ele, FolderName & "\" & target_num & "\" & new_filename
          
       
    Next
End Sub 

FolderName = "C:\A2OMSP\AIDA_Design\Marketing\2-1\aaaaaa"

call Move_FileRename(FolderName)
