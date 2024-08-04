''''''''''''''''''''' Input variables From Automation Anywhere ''''''''''''''''''''' 
''' Define Pdf Path : Convert Pdf file to Binary format '''
'Path_Input_PdfFile = "C:\Users\JayHyun_VM\Desktop\api_test\10545_21022_´«¾ØÁ¡¾È¾×0.5£¥_0.5mL.pdf"
''' file_id for the REST API : Convert Binary to Base64 '''
'file_id = "P_11111_110099"
''' file_name for the REST API : Convert Binary to Base64 '''
'file_name = "10545_21023_´«¾ØÁ¡¾È¾×0.5£¥_0.5mL.pdf"
''' product_name for the REST API : Convert Binary to Base64 '''
'product_name = "10545_21023_´«¾ØÁ¡¾È¾×0.5£¥_0.5mL.pdf"

''' url for the REST API : Convert Binary to Base64 '''
'url = "http://hanmi-nlp.42maru.com/api/v1/nlp/coa"


''''''''''''''''''''' User variables  ''''''''''''''''''''' 
Base_Dir ="C:\Users\JayHyun_VM\Desktop\api_test\"
outputFile = "C:\Users\JayHyun_VM\Desktop\api_test\result.txt"

Excel_Name = "C:\Users\JayHyun_VM\Desktop\api_test\result_2021_10_27.xlsx"
sheetName = "Á¦Ç°"

col_file_name = "A"
row_file_name = "2"

col_file_id = "K"
row_file_id = "2" 

''''''''''''''''''''' Object : Excel  '''''''''''''''''''''
Set oExcel = CreateObject("Excel.Application")
Set oWorkbook = oExcel.Workbooks.Open(Excel_Name)
Set oWorksheet = oWorkbook.Worksheets(sheetName)
Set oRng = oWorksheet.UsedRange
Msgbox oRng.Rows.Count
file_name = oWorksheet.Range(col_file_name & row_file_name)
product_name = file_name
Path_Input_PdfFile = Base_Dir & file_name 
data_type = "product"
file_id = oWorksheet.Range(col_file_id & row_file_id)
url = "http://hanmi-nlp.42maru.com/api/v1/nlp/coa"

Msgbox Path_Input_PdfFile
oExcel.Quit

''''''''''''''''''''' Object : Message box  '''''''''''''''''''''
Set oShell = CreateObject("WScript.Shell")
 
''''''''''''''''''''' Define Constant and Object ''''''''''''''''''''' 
Const adTypeBinary = 1
Const adTypeText = 2
Const adSaveCreateOverWrite = 2

''''''''''''''''''''' Object : Convert Pdf file to Binary format '''''''''''''''''''''
Set BinaryStream = CreateObject("ADODB.Stream")
Set oXML = CreateObject("Msxml2.DOMDocument")
Set oNode = oXML.CreateElement("base64")

''''''''''''''''''''' Convert Pdf file to Binary format '''''''''''''''''''''
BinaryStream.Type = adTypeBinary
BinaryStream.Open 

BinaryStream.LoadFromFile Path_Input_PdfFile
''''''''''''''''''''' Binary format '''''''''''''''''''''
ReadFile = BinaryStream.Read
BinaryStream.Close

oNode.dataType="bin.base64"
oNode.nodeTypedValue= ReadFile
''''''''''''''''''''' Base64 format '''''''''''''''''''''
Base64Encode = oNode.text
Set oNode = Nothing

''''''''''''''''''''' Remove "\n" '''''''''''''''''''''
file_data = Replace(Base64Encode, chr(10), "")

Set Base64Encode = Nothing


''''''''''''''''''''' Convert Binary to Base64 format '''''''''''''''''''''
''''''''''''''''''''' Object : Convert Binary to Base64 format '''''''''''''''''''''
 Set oXMLHttp = CreateObject("MSXML2.SERVERXMLHTTP.6.0")

''''''''''''''''''''' Json : Convert Binary to Base64 format ''''''''''''''''''''' 

strJson = "{ ""file_id"" :""" & file_id & """, " &_
"""file_name"" :"""& file_name &"""," &_
"""data_type"" : """& data_type & """," &_
"""product_name"" :"""& product_name &"""," &_
"""file_data"" :"""& file_data &"""}"
 

''''''''''''''''''''' Call API : Convert Binary to Base64 format '''''''''''''''''''''  
oXMLHttp.Open "POST" , url, false
oXMLHttp.setRequestHeader "Content-Type", "application/json;charset=utf-8"

oShell.Popup "Send Json", 3
 
oXMLHttp.send strJson

''' Result From the API Call : OK = "200"  '''
answer_post = oXMLHttp.responseText

oShell.Popup "Got the answer", 3

''' return the Result '''


Set fso = CreateObject("Scripting.Filesystemobject")
Set output = fso.CreateTextFile(outputFile,1)
output.Write answer_post
output.Close