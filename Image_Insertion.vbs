const xlMoveAndSize =1

Sub Image_Insert(filePath)
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = True 

    Set objWorkbook = objExcel.Workbooks.Open(filePath)
    Set objWorkSheet = objWorkbook.Worksheets("sheet1")
    Set objRngTarget = objWorkSheet.UsedRange

    For Each ele In objRngTarget 
        If ( ele.Row > 1) Then
            ele.RowHeight = ele.RowHeight + 10
            If ( ele.Column = 13 ) Then
                imgPath =  ele
            ' Set pic = objWorkSheet.Pictures.Insert(imgPath)
                Set pic =  objWorkSheet.Shapes.AddPicture(imgPath,0,1,0,0,0,0)
                With pic
                .Left = objWorkSheet.Cells(ele.Row,ele.Column).Left + 40    
                .Height =  150
                .Width = 100
                .Top = objWorkSheet.Cells(ele.Row,ele.Column).Top + 10
                .Placement = xlMoveAndSize
                                
                End With
                
            
            
            End If 
        End If  
    Next 

    objWorkbook.Save 
    objWorkbook.Close
    objExcel.Quit
End Sub 

filePath = "C:\Users\ISPark\Desktop\DGB\bmt\ocrTo106_test\result\yn_result.csv"
Call Image_Insert(filePath)