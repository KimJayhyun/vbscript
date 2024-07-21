Sub Add_answer(csvPath, answerArray)
    Set objExcel = CreateObject("Excel.Application")

    Set objworkbook = objExcel.Workbooks.Open(csvPath) 
    Set objWorksheet = objWorkbook.Worksheets(1)

    objExcel.DisplayAlerts = False

    xlCSV = 6

    objWorksheet.Cells(1,13) = "Answer"

    length = UBound(answerArray)
  
    For i = 0 To length
        objWorksheet.Cells(i + 2, 13) = answerArray(i)
    Next

    objExcel.ActiveWorkBook.SaveAs csvPath, xlCSV
    objExcel.ActiveWorkBook.Close False
    objExcel.quit


End Sub

answerArray = Array(6, 5, 4, 3)
csvPath = "C:\Users\ISPark\Desktop\test\result.csv"

call Add_answer(csvPath, answerArray)

