#############################################
# Retire-Convert.ps1
# Rewrited: 30-Jul-2018
# Updated: 6-Sep-2018
# Billy Zhou
# This script is used to remove retired employees' office 365 license and disable their aduser account.
#############################################

get-pssession | remove-pssession

#############################################
## Prepare BenQ                         #####
#############################################

#$query = "select * from [dbo].[v_OutlookData] where OutDate > dateadd(d,-30,GETDATE()) AND OutDate <= GETDATE()"
#$query = "select * from [dbo].[v_OutlookData] where ForwardMail <> '' AND OutDate <= GETDATE() AND ForwardDate >= GETDATE()"
#$query = "select * from [dbo].[v_OutlookData] where OutDate <> '' group by Email having count(*) =1"
#$query = "select * from [dbo].[v_OutlookData] where OutDate <> ''"

$query = "select * from [dbo].[v_OutlookData] where LeaveDate <>'' and LeaveDate <= GETDATE() and email in (select email from v_OutlookData group by email having count(*)=1)"
#$query = "select * from [dbo].[v_OutlookData] where Email = 'leon.wang@adenservices.com'" 

$connectionString = "Data Source=192.168.0.97;Initial Catalog=eHR;User Id=exchange;Password=exchange"

#############################################
## Prepare Log                         #####
#############################################
$batchNo = Get-Date -Format 'yyyyMMdd'
$runningLog = "C:\log\RetireConvert\RunningLog.log"
$retireLog = "C:\log\RetireConvert\RetireLog" + $batchNo + ".log"
#$runningLog = "d:\RunningLog" + ".log"
#$retireLog = "d:\RetireLog" + $batchNo + ".log"

######### Prepare Office 365 ##########

$File = "c:\scripts\adminpwd"
[Byte[]] $key = (1..16) 

$Office365Username = "admin@adengroup.onmicrosoft.com"
$SecureOffice365Password = Get-Content $file | ConvertTo-SecureString -Key $Key 
$Office365Credentials  = New-Object System.Management.Automation.PSCredential $Office365Username, $SecureOffice365Password  

$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Office365Credentials -Authentication Basic -AllowRedirection
Import-PSSession $exchangeSession -AllowClobber | Out-Null

Connect-MsolService -Credential $Office365Credentials

########## Prepare BenQ Database ###############
$connection = New-Object -TypeName System.Data.SqlClient.SqlConnection

$connection.ConnectionString = $connectionString
$command = $connection.CreateCommand()
$command.CommandTimeout = 0
$command.CommandText = $query

$adapter = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter $command
$dataset = New-Object -TypeName System.Data.DataSet
$adapter.Fill($dataset)
$table = $dataset.Tables[0]
$count = 1

$license1 = 'adengroup:ENTERPRISEPACK'
$license2 = 'adengroup:DESKLESSPACK'
$license3 = 'adengroup:ENTERPRISEPREMIUM_NOPSTNCONF'
$license4 = 'adengroup:STANDARDPACK'

$disabledOuPath = 'OU=Disabled,OU=ADEN-Users,DC=CHOADEN,DC=COM'

"" > $runningLog
########### Traverse the table ############
foreach ($item in $table.Rows)
{
	$email = $item.Email.Trim()
    $outDate = $item.OutDate
    #$forwardMail = $item.ForwardMail
    #$forwardDate = $item.ForwardDate
    $adStatus = "Not enabled"

    $sam = $email.substring(0,$email.IndexOf("@")).Trim()
    $msolUser = Get-MsolUser -UserPrincipalName $email -ErrorAction SilentlyContinue

    if ([bool] (Get-ADUser -Filter { SamAccountName -eq $sam }) -eq $true) # if aduser exists
    {
        if ((Get-ADUser -Filter { SamAccountName -eq $sam }).Enabled -eq $true) # if aduser is enabled
        {
            # clear aduser's manager and disabled aduser
            set-aduser $sam -clear manager
            Set-adUser $sam -Replace @{msExchHideFromAddressLists="TRUE"}
            Disable-ADAccount $sam
            #Get-ADUser $sam | Move-ADObject -TargetPath $disabledOuPath

            $adStatus = "Disabled"
            "ADuser: " + $sam + " has been disabled" >> $RetireLog
        }
    }
    else
    {
        $adStatus = "Not found"
    }
    # write running log
    #$count.ToString() + "`t" + $email + "`t" + $outDate + "`tForwardMail:" + $forwardMail + "`tForwardDate:" + $forwardDate + "`tIsLicensed:" + $msolUser.IsLicensed + "`tADUserStatus:" + $adStatus
    #$count.ToString() + "`t" + $email + "`t" + $outDate + "`tForwardMail:" + $forwardMail + "`tForwardDate:" + $forwardDate + "`tIsLicensed:" + $msolUser.IsLicensed + "`tADUserStatus:" + $adStatus >> $runningLog
    $count.ToString() + "`t" + $email + "`t" + $outDate + "`tIsLicensed:" + $msolUser.IsLicensed + "`tADUserStatus:" + $adStatus
    $count.ToString() + "`t" + $email + "`t" + $outDate + "`tIsLicensed:" + $msolUser.IsLicensed + "`tADUserStatus:" + $adStatus >> $runningLog

	If ($null -ne $msolUser)
	{
        #Set-Mailbox $email -ForwardingAddress  "admin@aden.partner.onmschina.cn" -DeliverToMailboxAndForward $False 
        #Set-MailboxAutoReplyConfiguration -Identity $email -AutoReplyState Enabled -ExternalMessage "" -InternalMessage ""

        # remove lic
        # Set-Mailbox $email -Type shared
        Set-Mailbox $email -HiddenFromAddressListsEnabled $true -ErrorAction SilentlyContinue
		if ($msolUser.isLicensed -eq $true)
		{
            Set-MsolUserLicense -UserPrincipalName $email  -RemoveLicenses $license1 -ErrorAction SilentlyContinue
            Set-MsolUserLicense -UserPrincipalName $email  -RemoveLicenses $license2 -ErrorAction SilentlyContinue
            Set-MsolUserLicense -UserPrincipalName $email  -RemoveLicenses $license3 -ErrorAction SilentlyContinue
            Set-MsolUserLicense -UserPrincipalName $email  -RemoveLicenses $license4 -ErrorAction SilentlyContinue
            # block msoluser from logon
            Set-MsolUser -UserPrincipalName $email -BlockCredential $true

            "Msoluser: " + $email + " 's license has been removed" >> $RetireLog
        }
    }
    $count++
}
# send log content through mail only when log exists
if (Test-Path $retireLog)
{
    
    $msg=New-Object System.Net.Mail.MailMessage
    $msg.To.Add("billy.zhou@adenservices.com,global.itsup@adenservices.com")
    $msg.From=New-Object System.Net.Mail.MailAddress("log@adenservices.com");
    $msg.Subject="RetireConvertLog" + $batchNo
    $msg.Body=Get-Content $retireLog -Raw
    $client=New-Object System.Net.Mail.SmtpClient("smtprelay.it.adenservices.com")
    $client.Send($msg)
}
Get-PSSession | Remove-PSSession