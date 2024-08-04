filePath = "C:\Users\User\Desktop\test.xlsx"

Set objExcel = CreateObject("Excel.Application")
objExcel.DisplayAlerts = False

startRow = 1
endRow = 10

Set objWorkbook = objExcel.Workbooks.Open(filePath)
Set objWorkSheet = objWorkbook.Worksheets("sheet1")

Set targetRange = objWorkSheet.Range("B" & startRow, "B" & endRow)

targetRange.Merge

objWorkbook.Save 
objWorkbook.Close

objExcel.Quit