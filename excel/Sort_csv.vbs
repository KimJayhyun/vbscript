sub Sort_csv(filePath)
    Const xlAscending = 1
    Const xlDescending = 2  
    Const xlYes = 1

    COnst xlLocalSessionChanges	= 2
    Const xlCSV = 6

    sort_column1 = "A1"
    sort_column2 = "F1"

    Set objExcel = CreateObject("Excel.Application")

    objExcel.DisplayAlerts = False
    
    
    Set objworkbook = objExcel.Workbooks.Open(filePath) 
    Set objWorksheet = objWorkbook.Worksheets(1)
    Set objRange = objWorksheet.UsedRange

    Set objRange2 = objExcel.Range(sort_column1)
    Set objRange3 = objExcel.Range(sort_column2)
    
    objRange.Sort  objRange2,xlAscending, objRange3,,xlAscending,,,xlYes

    objExcel.ActiveWorkBook.SaveAs filePath, xlCSV
    objExcel.ActiveWorkBook.Close False
    objExcel.quit


    Set oShell = CreateObject("WScript.Shell")
    oShell.Popup "Sort_csv", 2

End Sub



filePath = "C:\Users\ISPark\Desktop\test\sort_test.csv"
call Sort_csv(filePath)