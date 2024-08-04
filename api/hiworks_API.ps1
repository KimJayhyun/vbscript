Function hiworks-alarm {

param (
[Parameter(Mandatory=$true, 
           Position = 0)] 
[string] 
$id ,

 
[Parameter(Mandatory=$true, 
           Position = 1)] 
[string] 
$message ,

[Parameter(Mandatory=$true, 
           Position = 2)] 
[string] 
$title
) 


$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization", "Bearer 122cba6b968f753a77c48450081d73fd ")
$headers.Add("Content-Type", "application/json;charset=utf-8")

################################################################################# $id , $message , $title Ãß°¡ ######################################################################
$json = "{`n   `"user_list`":[`n      `"$id`"`n   ],`n   `"message`":`"$message`",`n   `"link`":`"https://developers.hiworks.com`",`n   `"mlink`":`"https://m.hiworks.com`",`n   `"solution_name`":`"$title`",`n   `"solution_image_url`":`"https://www.hiworks.com/static/images/logo.png`",`n   `"solution_default_url`":`"https://www.hiworks.com`"`n}"

$body = [System.Text.Encoding]::UTF8.GetBytes($json)

$response = Invoke-RestMethod 'api.hiworks.com/office/v2/notify' -Method 'POST' -Headers $headers -Body $body 
$response | ConvertTo-Json

}


