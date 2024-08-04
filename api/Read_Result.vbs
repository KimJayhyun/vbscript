''''''''''''''''''''' Input variables From Automation Anywhere ''''''''''''''''''''' 
''' Define Pdf Path : Convert Pdf file to Binary format '''
Path_Input_PdfFile = "C:\Users\JayHyun_VM\Desktop\api_test\10545_21022_´«¾ØÁ¡¾È¾×0.5£¥_0.5mL.pdf"
''' file_id for the REST API : Convert Binary to Base64 '''
file_id = "P_11111_110099"
''' file_name for the REST API : Convert Binary to Base64 '''
file_name = "10545_21023_´«¾ØÁ¡¾È¾×0.5£¥_0.5mL.pdf"
''' product_name for the REST API : Convert Binary to Base64 '''
product_name = "10545_21023_´«¾ØÁ¡¾È¾×0.5£¥_0.5mL.pdf"
data_type = "product"
''' url for the REST API : Convert Binary to Base64 '''
url = "http://hanmi-nlp.42maru.com/api/v1/nlp/coa"


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
'Msgbox Base64Encode
''''''''''''''''''''' Remove "\n" '''''''''''''''''''''
'Base64Encode_Trimmed = Replace(Base64Encode, chr(10), "")

Base64Encode_Array = Split(Base64Encode,chr(10))

Base64Encode_Trimmed = Join(Base64Encode_Array,"")

Set Base64Encode_Array = Nothing


file_data = Base64Encode_Trimmed

''''''''''''''''''''' Convert Binary to Base64 format '''''''''''''''''''''
''''''''''''''''''''' Object : Convert Binary to Base64 format '''''''''''''''''''''
'Set oXMLHttp = CreateObject("MSXML2.XMLHTTP")
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
 
oXMLHttp.send strJson

''' Result From the API Call : OK = "200"  '''
answer_post = oXMLHttp.responseText

''' return the Result '''

outputFile = "C:\Users\JayHyun_VM\Desktop\api_test\result.txt"
Set fso = CreateObject("Scripting.Filesystemobject")
Set output = fso.CreateTextFile(outputFile,1)
output.Write answer_post
output.Close

