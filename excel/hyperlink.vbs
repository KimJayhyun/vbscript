excelPath = WScript.Arguments.Item(0)

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True 

Set objWorkbook = objExcel.Workbooks.Open(excelPath)
Set objWorkSheet = objWorkbook.Worksheets("인재목록")

i = 1
Do While True
    url = objWorkSheet.Cells(i,15).Value
    
    If (url = "") Then
        Exit Do
    End If
    
    objWorkSheet.Hyperlinks.Add objWorkSheet.Cells(i,2), url

    i = i + 1
Loop
