Sub Move_FileFolder()
    FolderName = "C:\Users\ISPark\Desktop\DGB\cardData\PDFtoPNG\card_all\1"
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(FolderName)
    Set subFolders = objFolder.SubFolders
    
    
    For each ele in subFolders
        Set objFolder2 = objFSO.GetFolder(ele)
        set files = objFolder2.files
    
        For each file in files
            
            file_split = Split(file,"\")
            last_num = Ubound(file_split) 
            fileName = file_split(last_num)

            objFSO.MoveFile file, objFolder & "\" & fileName
        Next

       objFSO.DeleteFolder(ele) 
       
    Next
    

End Sub 

call Move_FileFolder()
