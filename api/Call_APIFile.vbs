''''' vbs ������ ��ġ�� ��� '''''
lv_File_dir = "C:\Users\JayHyun_VM\Desktop\api_test"
''''' delimeter '''''

''''' Script ���� �̸� '''''
lv_ScriptName = "API_Call_test"


lv_query = "cmd /c WScript.exe " & lv_ScriptName  & ".vbs"

lv_final_query = "'" & lv_query & "'"


'''' bach file �����Լ� '''''
Sub Execute_Batch(lv_File_dir,lv_final_query)
    Set oShell = CreateObject("WScript.Shell")
    oShell.CurrentDirectory = lv_File_dir & "\" 
    oShell.Run  lv_final_query , 1, True    
    'oShell.Run "cmd /k exit"   
End Sub 

call Execute_Batch(lv_File_dir,lv_query)






