sub fileLoop(folderPath)
    Set oFSO = CreateObject("Scripting.FileSystemObject")

    For Each oFile In oFSO.GetFolder(folderPath).Files
        MsgBox oFile.Name
    Next
End Sub

call fileLoop("C:\Users\ISPark\Desktop\code\Module")