resultFile = "C:\Users\User\Desktop\zipTest\des\test.zip"
saveFolder = "C:\Users\User\Desktop\zipTest\test"


Set fso = CreateObject("Scripting.FileSystemObject")

With fso.CreateTextFile(resultFile, True)
    .Write Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, chr(0))
End With

With CreateObject("Shell.Application")
    .NameSpace(resultFile).CopyHere .NameSpace(saveFolder).Items

    Do Until .NameSpace(resultFile).Items.Count = _
            .NameSpace(saveFolder).Items.Count
        Wscript.Sleep 1000
    Loop
End With

''''''''''''''''''''''''''''''''''''''''''''''  refer to : https://stackoverflow.com/questions/15139761/zip-a-folder-up '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''