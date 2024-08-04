Function hanmi_42_api{
param (
    [Parameter(Mandatory=$true, 
               Position = 0)] 
    [string] 
    $file_id ,
    
     
    [Parameter(Mandatory=$true, 
               Position = 1)] 
    [string] 
    $file_name ,
    
    [Parameter(Mandatory=$true, 
               Position = 2)] 
    [string] 
    $product_name,
    
    [Parameter(Mandatory=$true, 
               Position = 3)] 
    [string] 
    $file_data
    ) 






$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Content-Type", "application/json;charset=utf-8")



################################################################################# $id , $message , $title �߰� ######################################################################
$json = "{`n   `"file_id`" : `"$file_id`"`n,
`n   `"file_name`":`"$file_name`",
`n   `"data_type`":`"product`",
`n   `"product_name`":`"$product_name`",
`n   `"file_data`":`"$file_data`"`n}"

$body = [System.Text.Encoding]::UTF8.GetBytes($json)

$response = Invoke-RestMethod 'hanmi-nlp.42maru.com/api/v1/nlp/coa' -Method 'POST' -Headers $headers -Body $body 
$response | ConvertTo-Json

}