Function ParseJson(strJson)
    Set html = CreateObject("htmlfile")
    Set window = html.parentWindow
    window.execScript "var json = " & strJson, "JScript"
    Set ParseJson = window.JSON
End Function

Function Getjsonobject (Strjson)
    Set Sc4json = CreateObject("Msscriptcontrol.scriptcontrol")
    Sc4json.addcode "var jsonobject =" & Strjson, "JScript"
    Set Getjsonobject = Sc4json.CodeObject.jsonObject
End Function



sJson = "{ ""user_list"" : [""jhkim2""]}"



