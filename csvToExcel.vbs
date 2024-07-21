sub csvToExcel(filePath)
  Dim xlApp, workBook1, workBook2,aSheets, fileName, aInfo2,aInfo1,oExcel
    Const XlPlatform = "xlWindows"
    Const xlDelimited = 1
    Const xlTextQualifierDoubleQuote = -4142
    Const xlTextFormat = 2
    Const xlGeneralFormat = 1
    Const xlOpenXMLWorkbook  = 51

    Set xl = CreateObject("Excel.Application")

    xl.DisplayAlerts = False

    xl.Workbooks.OpenText filePath, , , xlDelimited _
    , xlTextQualifierDoubleQuote, False, False, False, True, False, False, _
    , Array(Array(1,2), Array(2,2), Array(3,2), Array(4,2), Array(5,2) _
    , Array(6,2), Array(7,2), Array(8,2), Array(9,2), Array(10,2), Array(11,2), Array(12, 2))
  
    Set wb = xl.ActiveWorkbook

    new_filePath = Replace(filePath, "csv", "xlsx")

    wb.SaveAs new_filePath, xlOpenXMLWorkbook, , , , False
    wb.Close

    xl.Quit
End Sub

filePath = "C:\Users\ISPark\Desktop\DGB\bmt\PoC_Checkbox_Test\result\checkbox_test.csv"

Call csvToExcel(filePath)