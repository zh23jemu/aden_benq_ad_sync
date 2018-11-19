#############################################
# Retire-Convert.ps1
# Rewrited: 30-Jul-2018
# Updated: 12-Nov-2018
# Billy Zhou
# This script is used to remove retired employees' office 365 license and disable their aduser account.
#############################################

get-pssession | remove-pssession

#############################################
## Prepare BenQ                         #####
#############################################

$connectionStringBenq = "Data Source=97VMDBSERVER.CHOADEN.COM;Initial Catalog=eHR;User Id=exchange;Password=exchange"
$connectionStringCadena = "Data Source=ccnsvwhqdwh02.choaden.com;Initial Catalog=BI_HR;User Id=weaden;Password=adenweaden@123"

$queryBenq = "select email,leavedate from [dbo].[v_OutlookData] 
    where LeaveDate <>'' and LeaveDate <= GETDATE() and Email not in
        (select Email from [dbo].[v_OutlookData] 
            where email in 
                (select email from v_OutlookData group by email having count(*)>1) and (LeaveDate='' or LeaveDate is null))"

<#$queryCadena = "select Email,'' as outdate 
    from dbo.HR_EMPS_VN 
    where EmployeeStatus = 'Resigned' and Email like '%@adenservices.com'"
#>


#############################################
## Prepare Log                         #####
#############################################
$batchNo = Get-Date -Format 'yyyyMMdd'
$LogPath = "C:\log\RetireConvert\"
$runningLog = $LogPath + "RunningLog.log"
$RetireLog = $LogPath + "RetireConvert" + $batchNo + ".log"
if (!(Test-Path $LogPath))
{
    mkdir $LogPath
}

######### Prepare Office 365 ##########

$ScriptFolder = Split-Path $MyInvocation.MyCommand.Definition -Parent
$File = $ScriptFolder + "\adminpwd"
[Byte[]] $key = (1..16) 

$Office365Username = "admin@adengroup.onmicrosoft.com"
$SecureOffice365Password = Get-Content $file | ConvertTo-SecureString -Key $Key 
$Office365Credentials  = New-Object System.Management.Automation.PSCredential $Office365Username, $SecureOffice365Password  

$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Office365Credentials -Authentication Basic -AllowRedirection
Import-PSSession $exchangeSession -AllowClobber | Out-Null

Connect-MsolService -Credential $Office365Credentials

########## Prepare Database ###############
$connectionBenq = New-Object -TypeName System.Data.SqlClient.SqlConnection

$connectionBenq.ConnectionString = $connectionStringBenq
$commandBenq = $connectionBenq.CreateCommand()
$commandBenq.CommandText = $queryBenq
$adapterBenq = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter $commandBenq
$datasetBenq = New-Object -TypeName System.Data.DataSet
$adapterBenq.Fill($datasetBenq)
$tableBenq=$datasetBenq.Tables[0]

$connectionBenq.Close()

<#$connectionCadena = New-Object -TypeName System.Data.SqlClient.SqlConnection

$connectionCadena.ConnectionString = $connectionStringCadena
$commandCadena = $connectionCadena.CreateCommand()
$commandCadena.CommandText = $queryCadena
$adapterCadena = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter $commandCadena
$datasetCadena = New-Object -TypeName System.Data.DataSet
$adapterCadena.Fill($datasetCadena)
$tableCadena=$datasetCadena.Tables[0]

$connectionCadena.Close()#>

$allData = $tableBenq.Rows# + $tableCadena.Rows

$count = 1

$license1 = 'adengroup:ENTERPRISEPACK'
$license2 = 'adengroup:DESKLESSPACK'
$license3 = 'adengroup:ENTERPRISEPREMIUM_NOPSTNCONF'
$license4 = 'adengroup:STANDARDPACK'

$disabledOuPath = 'OU=Disabled,OU=ADEN-Users,DC=CHOADEN,DC=COM'

"" > $runningLog
########### Traverse the table ############
foreach ($item in $allData)
{
	$email = $item.Email.Trim()
    $outDate = $item.OutDate
    $adStatus = "Not enabled"

    $sam = $email.substring(0,$email.IndexOf("@")).Trim()
    $msolUser = Get-MsolUser -UserPrincipalName $email -ErrorAction SilentlyContinue

    if ([bool] (Get-ADUser -Filter { SamAccountName -eq $sam }) -eq $true) # if aduser exists
    {
        if ((Get-ADUser -Filter { SamAccountName -eq $sam }).Enabled -eq $true) # if aduser is enabled
        {
            # clear aduser's manager and disabled aduser
            set-aduser $sam -clear manager
            Set-ADUser $sam -Replace @{msExchHideFromAddressLists=$True} 
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
    $count.ToString() + "`t" + $email + "`t" + $outDate + "`tIsLicensed:" + $msolUser.IsLicensed + ",`tADUserStatus:" + $adStatus
    $count.ToString() + "`t" + $email + "`t" + $outDate + "`tIsLicensed:" + $msolUser.IsLicensed + ",`tADUserStatus:" + $adStatus >> $runningLog

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