$pdf = Get-Content C:\Users\User\Desktop\API���� �䱸����\10545_21022_�������Ⱦ�0.5��_0.5mL.pdf -raw
$bytes = [System.Text.Encoding]::ASCII.GetBytes($pdf)
$base64 = [Convert]::ToBase64String($bytes)

# $pdf = Get-Content C:\Users\User\Desktop\API���� �䱸����\10545_21022_�������Ⱦ�0.5��_0.5mL.pdf -Encoding Byte
# $base64 =[Convert]::ToBase64String($pdf)

echo $base64


# $input = ��text to be encoded��
# $By = [System.Text.Encoding]::Unicode.GetBytes($input)
# $output =[Convert]::ToBase64String($By)
# $output
# $a = 'a'
# echo $a