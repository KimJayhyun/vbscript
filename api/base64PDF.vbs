Const adTypeBinary = 1
Const adTypeText = 2
Const adSaveCreateOverWrite = 2


INPUTFILE = "C:\Users\ISPark\Desktop\test\test.pdf"
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


outputFile = "C:\Users\ISPark\Desktop\test\result.txt"
Set fso = CreateObject("Scripting.Filesystemobject")
Set output = fso.CreateTextFile(outputFile,1)
output.Write Base64Encode
output.Close

