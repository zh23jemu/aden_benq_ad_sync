#
# Retire_Convert.ps1
#

get-pssession | remove-pssession

#############################################
## Prepare BenQ                         #####
#############################################

#$query = "select * from [dbo].[v_OutlookData] where OutDate > dateadd(d,-30,GETDATE()) AND OutDate <= GETDATE()"
#$query = "select * from [dbo].[v_OutlookData] where ForwardMail <> '' AND OutDate <= GETDATE() AND ForwardDate >= GETDATE()"
#$query = "select * from [dbo].[v_OutlookData] where OutDate <> '' group by Email having count(*) =1"
#$query = "select * from [dbo].[v_OutlookData] where OutDate <> ''"

$query = "select * from [dbo].[v_OutlookData] where LeaveDate <> '' and email in (select email from v_OutlookData group by email having count(*)=1)"

$connectionString = "Data Source=192.168.0.97;Initial Catalog=eHR;User Id=exchange;Password=exchange"

#############################################
## Prepare Log                         #####
#############################################
$batchNo = Get-Date -Format 'yyyyMMdd'
#$runningLog = "C:\temp\Retire_Convert\RunningLog" + $batchNo + ".log"
#$retireLog = "C:\temp\Retire_Convert\RetireLog" + $batchNo + ".log"
$runningLog = "d:\RunningLog" + ".log"
$retireLog = "d:\RetireLog" + $batchNo + ".log"
######### Prepare Office 365 ##########
$Office365Username = "admin@aden.partner.onmschina.cn"
$Office365Password = "All.007!"

$SecureOffice365Password = ConvertTo-SecureString -AsPlainText $Office365Password -Force
$Office365Credentials = New-Object System.Management.Automation.PSCredential $Office365Username, $SecureOffice365Password  

$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://partner.outlook.cn/PowerShell?proxyMethod=RPS" -Credential $Office365Credentials -Authentication "Basic" -AllowRedirection
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

$license1 = 'reseller-account:O365_BUSINESS_ESSENTIALS'
$license2 = 'reseller-account:EXCHANGEENTERPRISE'
$license3 = 'reseller-account:EXCHANGESTANDARD'
$license4 = 'reseller-account:O365_BUSINESS_PREMIUM'

"" > $runningLog
########### Traverse the table ############
foreach ($item in $table.Rows)
{
	$email = $item.Email.Trim()
    $outDate = $item.OutDate
    $forwardMail = $item.ForwardMail
    $forwardDate = $item.ForwardDate
    $adStatus = "Not enabled"

    $sam = $email.substring(0,$email.IndexOf("@")).Trim()
    $msolUser = Get-MsolUser -UserPrincipalName $email -ErrorAction SilentlyContinue

    if ([bool] (Get-ADUser -Filter { SamAccountName -eq $sam }) -eq $true) # if aduser exists
    {
        if ((Get-ADUser -Filter { SamAccountName -eq $sam }).Enabled -eq $true) # if aduser is enabled
        {
            # clear aduser's manager and disabled aduser
            set-aduser $sam -clear manager
            Disable-ADAccount $sam
            $adStatus = "Disabled"
            "ADuser: " + $sam + " has been disabled" >> $RetireLog
        }
    }
    else
    {
        $adStatus = "Not found"
    }
    # write running log
    $count.ToString() + "`t" + $email + "`t" + $outDate + "`tForwardMail:" + $forwardMail + "`tForwardDate:" + $forwardDate + "`tIsLicensed:" + $msolUser.IsLicensed + "`tADUserStatus:" + $adStatus
    $count.ToString() + "`t" + $email + "`t" + $outDate + "`tForwardMail:" + $forwardMail + "`tForwardDate:" + $forwardDate + "`tIsLicensed:" + $msolUser.IsLicensed + "`tADUserStatus:" + $adStatus >> $runningLog

	If ($null -ne $msolUser)
	{
        #Set-Mailbox $email -ForwardingAddress  "admin@aden.partner.onmschina.cn" -DeliverToMailboxAndForward $False 
        #Set-MailboxAutoReplyConfiguration -Identity $email -AutoReplyState Enabled -ExternalMessage "" -InternalMessage ""

        # set msoluser's mailbox to shared mailbox and remove lic
        Set-Mailbox $email -Type shared
		if ($msolUser.isLicensed -eq $true)
		{
            Set-MsolUserLicense -UserPrincipalName $email  -RemoveLicenses $license1 -ErrorAction SilentlyContinue
            Set-MsolUserLicense -UserPrincipalName $email  -RemoveLicenses $license2 -ErrorAction SilentlyContinue
            Set-MsolUserLicense -UserPrincipalName $email  -RemoveLicenses $license3 -ErrorAction SilentlyContinue
            Set-MsolUserLicense -UserPrincipalName $email  -RemoveLicenses $license4 -ErrorAction SilentlyContinue
            # block msoluser from logon
            Set-MsolUser -UserPrincipalName $email -BlockCredential $true

            "Msoluser: " + $email + "'s license has been removed" >> $RetireLog
        }
    }
    $count++
}