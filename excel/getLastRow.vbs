''' Error Handling
On Error Resume Next

excelPath = WScript.Arguments.Item(0)

' excelPath = "D:\ISPARK_RPA\인사쟁이\Temp\Output.xlsx"

Set objExcel = CreateObject("Excel.Application")

xlDown = -4121
xlUp = -4162

Set objWorkbook = objExcel.Workbooks.Open(excelPath)
Set objWorkSheet = objWorkbook.WorkSheets("게시글보고")

' lrow  = sht.Range(sht.Cells(row, col), sht.Cells(row, col)).End(xlDown).row
lastRow_1 = objWorkSheet.Range("B" & objWorkSheet.Rows.Count).End(xlUp).Row

Set objWorkSheet = objWorkbook.WorkSheets("댓글보고")

lastRow_2 = objWorkSheet.Range("B" & objWorkSheet.Rows.Count).End(xlUp).Row


WScript.StdOut.WriteLine(lastRow_1 & ";" & lastRow_2)

objWorkbook.Save 
objWorkbook.Close
objExcel.Quit

''' Error Handling
Err.Clear