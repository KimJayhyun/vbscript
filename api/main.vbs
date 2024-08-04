Sub PdfToBase64(pdf_path, output_path)
    Const adTypeBinary = 1
    Const adTypeText = 2
    Const adSaveCreateOverWrite = 2


    INPUTFILE = pdf_path
    Set BinaryStream = CreateObject("ADODB.Stream")
    Set oXML=CreateObject("Msxml2.DOMDocument")
    Set oNode=oXML.CreateElement("base64")

    BinaryStream.Type = adTypeBinary
    BinaryStream.Open

    BinaryStream.LoadFromFile INPUTFILE
    ReadFile = BinaryStream.Read
    BinaryStream.Close


    oNode.dataType="bin.base64"
    oNode.nodeTypedValue= ReadFile
    Base64Encode = oNode.text

    a = Replace(Base64Encode, Chr(10), "")
    ' b = Replace(Base64Encode, Chr(13), "")
    ' c = Replace(Base64Encode, Chr(10) & Chr(13), "")
    ' d = Replace(Base64Encode, Chr(10) + Chr(13), "")

    outputFile = output_path
    Set fso = CreateObject("Scripting.Filesystemobject")
    Set output = fso.CreateTextFile(outputFile,1)
    output.Write a
    ' output.Write b
    ' output.Write c
    ' output.Write d

    output.Close
End Sub

Sub PostAPI(file_id, file_name, product_name, text_path)
    Set fso = CreateObject("Scripting.Filesystemobject")

    textfile = fso.OpenTextFile(text_path, 1)
    file_data = textfile.ReadAll
    textfile.Close

    fso = Nothing

    Set xmlhttp = CreateObject("MSXML2.serverXMLHTTP")
    xmlhttp.open "POST", "hanmi-nlp.42maru.com/api/v1/nlp/coa", false

    xmlhttp.setRequestHeader "Content-Type", "application/json;charset=utf-8"

   
    strJson = "{ ""file_id"" :""" & file_id & """, " &_
    """file_name"" :"""& file_name &"""," &_
    """data_type"" : """& data_type & """," &_
    """product_name"" :"""& product_name &"""," &_
    """file_data"" :"""& file_data &"""}"  

    xmlhttp.send(strJson)
    
    respon = xmlhttp.responText

    Msgbox respon
End Sub

output = "C:\Users\User\Desktop\test\base64.txt"

PdfToBase64 "C:\Users\User\Desktop\test\test.pdf", "C:\Users\User\Desktop\test\base64.txt"

MsgBox "done"

' PostAPI "1", "1", "1", output




