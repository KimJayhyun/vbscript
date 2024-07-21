Sub Move_FileFolder(FolderName)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(FolderName)
    Set Files = objFolder.Files
    For each ele in Files 
        file_split = Split(ele,"\")
        last_num = Ubound(file_split) 
        fileName = file_split(last_num)
        target_split = Split(fileName,".")(0)
        last_split = Split(target_split,"-")
        last_num = Ubound(last_split) 
        target_num = last_split(last_num)
        
        If NOT (objFSO.FolderExists(FolderName & "\" & target_num)) Then 
            objFSO.CreateFolder(FolderName & "\" & target_num) 
            If objFSO.FileExists(ele) Then 

               objFSO.MoveFile ele, FolderName & "\" & target_num & "\" & fileName
           End If 
        End If 
       
    Next
    

End Sub 

FolderName = "C:\Users\ISPark\Desktop\DGB\bmt\make_service\card_ocr_text\makeData\run"

call Move_FileFolder(FolderName)
