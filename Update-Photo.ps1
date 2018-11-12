#############################################
# Update-Photo.ps1
# Rewrited: 2-Aug-2018
# Updated: 9-Aug-2018
# Billy Zhou
# This script is used to update users' photo from HR BenQ system to Office 365.
#############################################

get-pssession | remove-pssession

#############################################
## Prepare BenQ                         #####
#############################################

$query = "select * from [dbo].[v_OutlookData] where LeaveDate = ''"
$connectionString = "Data Source=192.168.0.97;Initial Catalog=eHR;User Id=exchange;Password=exchange"

#############################################
## Prepare Log                         #####
#############################################
$batchNo = Get-Date -Format 'yyyyMMdd'
$runningLog = "C:\log\UpdatePhoto\RunningLog.log"
$updatePhotoLog = "C:\log\UpdatePhoto\UpdatePhotoLog" + $batchNo + ".log"

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
$countAll = 1

"" > $runningLog

########### Traverse the table ############
foreach ($item in $table.Rows)
{
	$email = $item.Email.Trim()
	$photoValue = $item.Photovalue

    # reset flags
	$hasPhotoValue = $true
    $hasMailbox = $true
    $hasO365Photo = $true

	if ($photoValue.Equals([DBNull]::Value)) 
	{
		$hasPhotoValue = $false
	}
	if ($null -eq (Get-Mailbox $email -ErrorAction SilentlyContinue)) 
	{
		$hasMailbox = $false
	}

    # write running log
    $countAll.ToString() + "`t" + $email + "`tHasPhotoValue: " + $hasPhotoValue + "`tHasMailbox: " + $hasMailbox
    $countAll.ToString() + "`t" + $email + "`tHasPhotoValue: " + $hasPhotoValue + "`tHasMailbox: " + $hasMailbox >> $runningLog

    # only when user has photo data and mailbox, photo will be set.
	If ($hasPhotoValue -and $hasMailbox)
	{
        # check if user currently has o365 photo and make a flag
        if ((Get-UserPhoto $email -errorAction silentlycontinue) -eq $null)
        {
            $hasO365Photo = $false
        }

        # whether user has photo on o365 or not, photo will be overwritten by BenQ photo data.
		try
        {
            set-userPhoto -Identity $email -PictureData $photoValue -Confirm:$false -ErrorAction stop

            if ($hasO365Photo -eq $false)
            {
                $email + "`t's photo has been newly uploaded." >> $updatePhotoLog
            }
        }
        catch [System.Exception]
        {
            "$_.Exception"
            "$_.Exception" >> $runningLog
        }
    }
    $countAll++
}

# only send log content through mail when log file exists.
if (Test-Path $updatePhotoLog)
{
    $msg=New-Object System.Net.Mail.MailMessage
    $msg.To.Add("billy.zhou@adenservices.com,global.itsup@adenservices.com")
    $msg.From=New-Object System.Net.Mail.MailAddress("log@adenservices.com");
    $msg.Subject="UpdatePhotoLog" + $batchNo
    $msg.Body=Get-Content $updatePhotoLog -Raw
    $client=New-Object System.Net.Mail.SmtpClient("smtprelay.it.adenservices.com")
    $client.Send($msg)
}

Get-PSSession | Remove-PSSession