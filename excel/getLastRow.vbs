''' Error Handling
On Error Resume Next

excelPath = WScript.Arguments.Item(0)

' excelPath = "D:\ISPARK_RPA\�λ�����\Temp\Output.xlsx"

Set objExcel = CreateObject("Excel.Application")

xlDown = -4121
xlUp = -4162

Set objWorkbook = objExcel.Workbooks.Open(excelPath)
Set objWorkSheet = objWorkbook.WorkSheets("�Խñۺ���")

' lrow  = sht.Range(sht.Cells(row, col), sht.Cells(row, col)).End(xlDown).row
lastRow_1 = objWorkSheet.Range("B" & objWorkSheet.Rows.Count).End(xlUp).Row

Set objWorkSheet = objWorkbook.WorkSheets("��ۺ���")

lastRow_2 = objWorkSheet.Range("B" & objWorkSheet.Rows.Count).End(xlUp).Row


WScript.StdOut.WriteLine(lastRow_1 & ";" & lastRow_2)

objWorkbook.Save 
objWorkbook.Close
objExcel.Quit

''' Error Handling
Err.Clear