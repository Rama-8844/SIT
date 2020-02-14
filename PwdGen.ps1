(Get-Credential).Password |
ConvertFrom-SecureString|
Out-File "D:\Rama\PSScripts\MyPwd.txt"