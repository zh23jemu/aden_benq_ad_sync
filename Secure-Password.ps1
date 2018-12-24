$PlainPassword = Read-Host "Enter the admin password"
$File = "D:\Work\Scripts\aden_benq_ad_sync\adminpwd"
[Byte[]] $key = (1..16)
$Password = $PlainPassword | ConvertTo-SecureString -AsPlainText -Force
$Password | ConvertFrom-SecureString -key $key | Out-File $File
