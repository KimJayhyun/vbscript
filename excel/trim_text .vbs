a = WScript.Arguments.Count

content = ""

ForReading = 1
ForWriting = 2
ForAppending = 8

Set fs = CreateObject("Scripting.FileSystemObject")
Set objStream = CreateObject("ADODB.Stream")
Set f = fs.OpenTextFile("C:\Users\User\Desktop\code\local\temp.txt", ForWriting)

Set test_object = CreateObject("System.Text.Encoding.ASCII")

msgbox "zz"

' objStream.CharSet = "EUC-KR"
' objStream.open


For i = 0 to a - 2
    temp = Trim(WScript.Arguments.item(i))
    temp = Replace(temp, chr(10), "")
    
    ' content = content + " " + temp

    msgbox i & "//" & asc(temp)
    
    ' test = System.Text.Encoding.Unicode.Getbytes temp
    encode_num = System.Text.Encoding.GetEncoding("EUC-KR")
    ' System.Text.Encoding.Convert System.Text.Encoding.Unicode, encode_num, test 

    msgbox encode_num

    f.writeLine(temp)

Next

objStream.WriteText content

objStream.SaveToFile "C:\Users\User\Desktop\code\local\temp.txt"


Set f = fs.OpenTextFile("C:\Users\User\Desktop\code\local\temp.txt", ForReading)

content = f.ReadAll

f.close

msgbox content
WScript.StdOut.WriteLine content

'On Error Resume Next
' WScript.StdOut.WriteLine Mid(content1, 1, 2)

'Err.Clear