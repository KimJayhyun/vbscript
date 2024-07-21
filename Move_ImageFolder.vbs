Sub Move_ImageFolder(FolderName)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(FolderName)
    Set subFolders = objFolder.SubFolders
    
    i = 0

    For each ele in subFolders
        Set objFolder2 = objFSO.GetFolder(ele)
        Set subFolders2 = objFolder2.SubFolders
        
        For each folder in subFolders2
            
            Set objFolder3 = objFSO.GetFolder(Folder)
            Set files = objFolder3.files

            For each file in files  
                objFSO.MoveFile file, objFolder & "\" & i & ".png"
                i = i + 1
            Next
    
        Next    

       objFSO.DeleteFolder(ele) 
    Next
End Sub 

FolderName = "C:\Users\ISPark\Desktop\DGB\bmt\make_service\card_ocr_text\makeData\img"

call Move_ImageFolder(FolderName)
