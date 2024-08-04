Set xmlhttp = CreateObject("MSXML2.serverXMLHTTP")




xmlhttp.open "POST", "https://api.hiworks.com/office/v2/notify", false

xmlhttp.setRequestHeader "Authorization", "Bearer 122cba6b968f753a77c48450081d73fd"
xmlhttp.setRequestHeader "Content-Type", "application/json"

strJson = "{ ""user_list"" : [""jhkim2""], " &_
"""message"" : ""¾È³ç?""                             ," &_
"""link"" : ""https://developers.hiworks.com""," &_
"""mlink"" : ""https://m.hiworks.com""," &_
"""solution_name"" : ""name""," &_
"""solution_image_url"" : ""https://www.hiworks.com/static/images/logo.png""," &_
"""solution_default_url"" : ""https://www.hiworks.com""}" 

xmlhttp.send(strJson)

' WScript.sleep 3000

' Msgbox xmlhttp.Status
