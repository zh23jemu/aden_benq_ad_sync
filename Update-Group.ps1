#############################################
# Update-Group.ps1
# Rewrited: 3-Aug-2018
# Updated: 3-Aug-2018
# Billy Zhou
# This script is used to update groups from HR BenQ system to Office 365.
#############################################

get-pssession | remove-pssession

#############################################
## Prepare BenQ                         #####
#############################################
$connectionString = "Data Source=192.168.0.97;Initial Catalog=eHR;User Id=exchange;Password=exchange"
$queryGroup = "select DISTINCT emailgroup from v_UserDL" 
$queryEmail = "select * from v_UserDL" 

$ScriptFolder = Split-Path $MyInvocation.MyCommand.Definition -Parent

######### Prepare Office 365 ##########

$File = $ScriptFolder + "\adminpwd"

[Byte[]] $key = (1..16) 

$Office365Username = "admin@adengroup.onmicrosoft.com"

$SecureOffice365Password = Get-Content $file | ConvertTo-SecureString -Key $Key 
$Office365Credentials  = New-Object System.Management.Automation.PSCredential $Office365Username, $SecureOffice365Password  

$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Office365Credentials -Authentication Basic -AllowRedirection
Import-PSSession $exchangeSession -AllowClobber | Out-Null

Connect-MsolService -Credential $Office365Credentials

#############################################
## Prepare Log File						  ###
#############################################

$batchNo =  Get-Date -Format 'yyyyMMdd'
$runningLog = "C:\log\UpdateGroup\RunningLog.log"
$UpdateGroupLog = "C:\log\UpdateGroup\UpdateGroupLog" + $batchNo + ".log"
$count = 1

"" > $runningLog

###### Connect to Database ######

$connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connectionString
$command = $connection.CreateCommand()
$command.CommandText = $queryGroup
$adapter = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter $command
$dataset = New-Object -TypeName System.Data.DataSet
$adapter.Fill($dataset)
$table=$dataset.Tables[0]

##### check if group exists ######

$emailgroups = $table.rows.emailgroup.trim().replace(" ","_")
$o365Groups =  Get-DistributionGroup

foreach ($item in $emailgroups) 
{
	if ($o365Groups.name -contains $item) 
    {
		$count.ToString() + "`tGroup exists:`t" + $item
        $count.ToString() + "`tGroup exists:`t" + $item >> $runningLog
	}
	else 
    {
		$primarySmtpAddress = $item + "@adenservices.com"
		New-DistributionGroup $item -PrimarySmtpAddress $primarySmtpAddress
        $count.ToString() + "`tGroup created:`t" + $item
        $count.ToString() + "`tGroup created:`t" + $item >> $runningLog
        "Group created:`t" + $item >> $UpdateGroupLog
	}
    $count++
}

###### Update group members ######
$connection = New-Object -TypeName System.Data.SqlClient.SqlConnection

$connection.ConnectionString = $connectionString
$command = $connection.CreateCommand()
$command.CommandText = $queryEmail
$adapter = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter $command
$dataset = New-Object -TypeName System.Data.DataSet
$adapter.Fill($dataset)
$table=$dataset.Tables[0]

$count = 1
foreach ($item in $table.Rows)
{
    $email = $item.email.trim()
    $emailgroup = $item.emailgroup.trim()

    if ((Get-MsolUser -UserPrincipalName $email -ErrorAction SilentlyContinue) -ne $null) #Check if msoluser exists
    {
        try 
        {
            Add-DistributionGroupMember -Identity $emailgroup -Member $email -Confirm:$false -ErrorAction Stop
            $count.ToString() + "`t" + $email + " has been added to group:`t" + $emailgroup
            $count.ToString() + "`t" + $email + " has been added to group:`t" + $emailgroup >> $runningLog
            $email + "`tadded to group:`t" + $emailgroup >> $UpdateGroupLog
        }
        Catch [System.Exception]
        {
            if($_.FullyQualifiedErrorId  -match 'AlreadyExists')
            {
                $count.ToString() + "`t" + $email + " is already in group:`t" + $emailgroup
                $count.ToString() + "`t" + $email + " is already in group:`t" + $emailgroup >> $runningLog
            }
            else
            {
                $count.ToString() + "`t" + $email + " $_.Exception"
                $count.ToString() + "`t" + $email + " $_.Exception" >> $runningLog
            }
        }
    }
    else
    {
        $count.ToString() + "`t" + $email + "`tmsoluser does not exist."
        $count.ToString() + "`t" + $email + "`tmsoluser does not exist." >> $runningLog
    }
    $count++
}

# remove members who no longer in groups
$count = 1
foreach ($groupItem in $emailgroups)
{
    $emailsInGroup = Get-DistributionGroupMember $groupItem | where {$_.RecipientType -eq 'UserMailbox'} | select -ExpandProperty PrimarySMTPAddress 
    $tableEmails = $table | where {$_.emailgroup -eq $groupItem} | select -ExpandProperty email
    foreach ($item in $emailsInGroup)
    {
        if($tableEmails -notcontains $item)
        {
            try
            {
                Remove-DistributionGroupMember $groupItem -Member $item -Confirm:$false -ErrorAction Stop
                $count.ToString() + "`t" + $item + "`tremoved from`t" + $groupItem
                $count.ToString() + "`t" + $item + "`tremoved from`t" + $groupItem >> $runningLog
                $item + "`tremoved from group:`t" + $groupItem >> $UpdateGroupLog
            }
            Catch [System.Exception]
            {
                $count.ToString() + "`t" + $email + " $_.Exception"
                $count.ToString() + "`t" + $email + " $_.Exception" >> $runningLog
            }
            $count++
        }
    }
}

# send log content through mail only when log file exists
if (Test-Path $UpdateGroupLog)
{
    $msg=New-Object System.Net.Mail.MailMessage
    $msg.To.Add("billy.zhou@adenservices.com,global.itsup@adenservices.com")
    $msg.From=New-Object System.Net.Mail.MailAddress("log@adenservices.com");
    $msg.Subject="UpdateGroupLog" + $batchNo
    $msg.Body=Get-Content $UpdateGroupLog -Raw
    $client=New-Object System.Net.Mail.SmtpClient("smtprelay.it.adenservices.com")
    $client.Send($msg)
}

Get-PSSession | Remove-PSSession