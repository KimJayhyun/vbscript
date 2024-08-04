FolderName = "C:\Users\user1\Desktop\test\data"

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(FolderName)
Set Files = objFolder.Files

Set answerFile = objFSO.CreateTextFile(FolderName + "\label.txt", True)

For each ele in Files
    file_split = Split(ele,"\")
    last_num = Ubound(file_split) 
    fileName = file_split(last_num)

    If (fileName <> "label.txt") Then    
        temp = Split(fileName, ".")(0)
        answer = Split(temp, "_")(1)

        answerFile.WriteLine(fileName + Chr(9) + answer)    
    End If
    

    ' msgbox fileName
    ' msgbox answer

Next

answerFile.Closes