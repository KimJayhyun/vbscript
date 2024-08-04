set objFSO = createobject("Scripting.Filesystemobject")

folderName = "C:\Users\User\Desktop\test"
Set objFolder = objFSO.GetFolder(folderName)
Set folders = objFolder.SubFolders


Dim arr(1)

i = 0
For each ele in folders
    msgbox(ele)
    arr(i) = ele
    i = i + 1
Next

For each ele in arr
    objFSO.DeleteFolder(ele)
Next


