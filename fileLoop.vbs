sub fileLoop(folderPath)
    Set oFSO = CreateObject("Scripting.FileSystemObject")

    For Each oFile In oFSO.GetFolder(folderPath).Files
        MsgBox oFile.Name
    Next
End Sub

folderPath = "C:\Users\ISPark\Desktop\code\Module"

call fileLoop(folderPath)