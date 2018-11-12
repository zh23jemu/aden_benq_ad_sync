#############################################
# ManageLicense.ps1
# Created: 3-Jul-2018
# Updated: 29-Sep-2018
# Billy Zhou
# This script is used to manage o365 license automatically for end users.
#############################################

get-pssession | remove-pssession

#############################################
## Prepare O365 and Exchange            #####
#############################################

$File = "c:\scripts\adminpwd"
[Byte[]] $key = (1..16) 

$Office365Username = "admin@adengroup.onmicrosoft.com"
$SecureOffice365Password = Get-Content $file | ConvertTo-SecureString -Key $Key 
$Office365Credentials  = New-Object System.Management.Automation.PSCredential $Office365Username, $SecureOffice365Password  

$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Office365Credentials -Authentication Basic -AllowRedirection
Import-PSSession $exchangeSession -AllowClobber | Out-Null

Connect-MsolService -Credential $Office365Credentials

#############################################
## Prepare Log File						#####
#############################################

$today = Get-Date
$batchNo =  Get-Date -Format 'yyyyMMddHH'

#log file
$LicenseLog = "C:\log\ManageLic\ManageLicLog" + $batchNo +".log"
$RunningStatusLog = "C:\log\ManageLic\RunningStatus.log"

# clear running status log
"" > $RunningStatusLog
#license type
$stdLicense = "adengroup:STANDARDPACK"

# Specify users in which OU will be assgined lic
$benqOU = "OU=BENQ,OU=ADEN-Users,DC=CHOADEN,DC=COM"
$cadenaOU = "OU=CADENA,OU=ADEN-Users,DC=CHOADEN,DC=COM"

########## Remove Lic Part ##########
# only get users with standard license
$msolusers = Get-MsolUser -All | Where-Object {$_.Licenses.AccountSkuId -eq $stdLicense}
$count = 1
foreach ($item in $msolusers)
{
    if ($null -ne (get-mailboxstatistics $item.UserPrincipalName).LastLogonTime)
    {
        $lastLogonTime = (Get-MailboxStatistics $item.UserPrincipalName).LastLogonTime
        $count.ToString() + " " + $item.UserPrincipalName + "`tLastLogonTime:" + $lastLogonTime.Date
        $count.ToString() + " " + $item.UserPrincipalName + "`tLastLogonTime:" + $lastLogonTime.Date >> $RunningStatusLog
        $noLogonDays = ($today - $lastLogonTime).days
        if ($noLogonDays -gt 30)
        {
            $item | Set-MsolUserLicense  -RemoveLicenses $stdLicense -ErrorAction SilentlyContinue
            $item.UserPrincipalName + "`tLicense removed`tLastLogonDate is " + $noLogonDays + "days ago." >> $LicenseLog
        }
    }
    else
    {
        $mailboxCreatedTime = (Get-Mailbox $item.UserPrincipalName).WhenMailboxCreated
        $count.ToString() + " " + $item.UserPrincipalName + "`tMailboxCreatedTime:" + $mailboxCreatedTime.Date
        $count.ToString() + " " + $item.UserPrincipalName + "`tMailboxCreatedTime:" + $mailboxCreatedTime.Date >> $RunningStatusLog
        $createdDays = ($today - $mailboxCreatedTime).days
        if ($createdDays -gt 30)
        {
            $item | Set-MsolUserLicense  -RemoveLicenses $stdLicense -ErrorAction SilentlyContinue
            $item.UserPrincipalName + "`tLicense removed`tmailbox created " + $createdDays + "days ago but never logged on." >> $LicenseLog
        }
    }
    $count++
}

######## simpler code ########
<#
$msolusers = Get-MsolUser -All |
    where {$_.Licenses.AccountSkuId -eq $stdLicense `
        -and (get-mailboxstatistics $_.UserPrincipalName).LastLogonTime -ne $null `
        -and ($today - (get-mailboxstatistics $_.UserPrincipalName).LastLogonTime).Days -gt 30}

# remove lic
foreach ($item in $msolusers)
{
    $noLogonDays = ($today - (Get-MailboxStatistics $item.UserPrincipalName).LastLogonTime).days
    $item | Set-MsolUserLicense  -RemoveLicenses $stdLicense -ErrorAction SilentlyContinue
    $item.UserPrincipalName + "`tOffice 365 Enterprise E1 lic removed because the LastLogonDate is " + $noLogonDays + "days ago." >> $LicenseLog
}

# only get users with standard license but never logged on, whose mailbox was created over 30 days ago.
$msolusers = Get-MsolUser -All |
    where {$_.Licenses.AccountSkuId -eq $stdLicense `
        -and (get-mailboxstatistics $_.UserPrincipalName).LastLogonTime -eq $null `
        -and ($today - (Get-Mailbox $_.UserPrincipalName).WhenMailboxCreated).Days -gt 30}

# remove lic
foreach ($item in $msolusers)
{
    $sinceCreatedDays = ($today - (Get-Mailbox $item.UserPrincipalName).WhenMailboxCreated).days
    $item | Set-MsolUserLicense  -RemoveLicenses $stdLicense -ErrorAction SilentlyContinue
    $item.UserPrincipalName + "`tOffice 365 Enterprise E1 lic removed because the mailbox was created " + $sinceCreatedDays + "days ago but never logged on." >> $LicenseLog
}
#>


########## Assign Lic Part ##############
$10daysAgo = $today.AddDays(-10)

$benqUser = get-aduser -filter * -SearchBase $benqOU -Properties whencreated | 
    where {$_.whencreated -ge $10daysAgo} | select samaccountname, userprincipalname, whencreated, @{l="Location";e={"CN"}}
#$benqUser | ogv
$cadenaUser = get-aduser -filter * -SearchBase $cadenaOU -Properties whencreated | 
    where {$_.whencreated -ge $10daysAgo} | select samaccountname, userprincipalname, whencreated, @{l="Location";e={"VN"}}
#$cadenaUser | ogv

$allUser = $benqUser + $cadenaUser
#$allUser | ogv
$count=1

foreach ($item in $allUser)
{
    $sam = $item.samaccountname.trim()
    $upn = $item.userprincipalname.trim()
    $whencreated = $item.whencreated
    $location = $item.location

    $msolUser = Get-MsolUser -UserPrincipalName $upn -ErrorAction SilentlyContinue

    #running log to check the running progress
    $count.ToString() + " " + $sam + "`tWhenCreated:" + $whencreated + "`t" + $msolUser.Licenses.AccountSkuId
    $count.ToString() + " " + $sam + "`tWhenCreated:" + $whencreated + "`t" + $msolUser.Licenses.AccountSkuId >> $RunningStatusLog

    if ( $null -ne $msolUser)
    {
        if ($msolUser.IsLicensed -eq $False)
        {
            # get current exchange std lic amount
            $AccountSku = Get-MsolAccountSku | Where-Object {$_.AccountSkuId -eq $stdLicense}
            $LicAmount = $AccountSku.ActiveUnits - $AccountSku.ConsumedUnits
            if ($LicAmount -gt 0)
            {
               # assign lic for users joined company less than 30 days
               Set-MsolUser -UserPrincipalName $upn -UsageLocation $location
		       Set-MsolUserLicense -userPrincipalName $upn -AddLicenses $stdLicense
               $upn + "`tLicense assigned`tAD account created on " + $whencreated
               $upn + "`tLicense assigned`tAD account created on " + $whencreated >> $LicenseLog
            }
            else
            {
               # out of lic
               $upn + "`tLicense not assigned because out of lic." >> $LicenseLog
            }
       }
    }
    $count++
}

# count the lic again
#$AccountSku = Get-MsolAccountSku | Where-Object {$_.AccountSkuId -eq $stdLicense}
#$LicAmount = $AccountSku.ActiveUnits - $AccountSku.ConsumedUnits
Get-MsolAccountSku
Get-MsolAccountSku >> $LicenseLog

# send log content through mail.
$msg=New-Object System.Net.Mail.MailMessage
$msg.To.Add("billy.zhou@adenservices.com,global.itsup@adenservices.com")
$msg.From=New-Object System.Net.Mail.MailAddress("log@adenservices.com")
$msg.Subject="ManageLicLog" + $batchNo
$msg.Body=Get-Content $LicenseLog -Raw
$client=New-Object System.Net.Mail.SmtpClient("smtprelay.it.adenservices.com")
$client.Send($msg)

Get-PSSession | Remove-PSSession
