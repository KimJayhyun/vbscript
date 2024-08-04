''''' vbs 파일이 위치한 경로 '''''
lv_Dir = WScript.Arguments.Item(0)
''''' delimeter '''''
lv_Delimeter = WScript.Arguments.Item(1)
''''' Script 파일 이름 '''''
lv_ScriptName = WScript.Arguments.Item(2)

''''' delimeter 삭제한 뒤 경로 '''''
lv_File_dir = Replace(lv_Dir,lv_Delimeter," ")

lv_query = "cmd /c " & lv_ScriptName  & ".bat"

lv_final_query = "'" & lv_query & "'"

'''' bach file 실행함수 '''''
Sub Execute_Batch(lv_File_dir,lv_final_query)
    Set oShell = CreateObject("WScript.Shell")
    oShell.CurrentDirectory = lv_File_dir & "\" 
    oShell.Run lv_final_query , 1, True    
    oShell.Run "cmd /k exit"   
End Sub 

call Execute_Batch(lv_File_dir,lv_query)






