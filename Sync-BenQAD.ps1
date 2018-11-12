#############################################
# Sync-BenQAD.ps1
# Rewrited: 13-Aug-2018
# Updated: 14-Aug-2018
# Billy Zhou
# This script is used to create and update AD user based on HR BenQ DB
#############################################

# function TelephoneNumber used for produce telephone number based on office number and ext number
function TelephoneNumber($officePhone, $ext)
{
	if ($officePhone -eq $null) 
	{
		$telephoneNumber = $null	
	}
	else 
	{
		if ($ext -eq $null) 
		{
			$telephoneNumber = $officePhone
		}
		else
		{
			$telephoneNumber = $officePhone + "," + $ext
		}
	}
	return $telephoneNumber
}
# function HandleNull is used for handle DBNull value
function HandleNull($oldvalue)
{
	if($oldvalue.Equals([DBNull]::Value))
	{
		$value = $null
	}
	elseif($oldvalue -eq '')
	{
		$value =  $null
	}
	else
	{
		$value = $oldvalue
	}
	return $value
}

get-pssession | remove-pssession

#############################################
## Prepare BenQ                         #####
#############################################

$query = "select * from [dbo].[v_OutlookData]" 
#$query = "select * from [dbo].[v_OutlookData] where Email = 'leon.wang@adenservices.com'" 
$connectionString = "Data Source=192.168.0.97;Initial Catalog=eHR;User Id=exchange;Password=exchange"

#############################################
## Prepare Log File						#####
#############################################

$batchNo =  Get-Date -Format 'yyyyMMddHH'
$runningLog = "C:\log\SyncBenqAD\RunningLog.log"
$syncBenqADLog = "C:\log\SyncBenqAD\SyncBenqAD" + $batchNo + ".log"

########## Connect to BenQ DB ############

$connection = New-Object -TypeName System.Data.SqlClient.SqlConnection

$connection.ConnectionString = $connectionString
$command = $connection.CreateCommand()
$command.CommandText = $query
$adapter = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter $command
$dataset = New-Object -TypeName System.Data.DataSet
$adapter.Fill($dataset)
$table=$dataset.Tables[0]
$count = 1

"" > $runningLog

########### Traverse the table ############
foreach ($item in $table.Rows)
{
	$employeeId = HandleNull($item.EMPLOYEEID.Trim())
	$firstName = HandleNull($item.FIRSTNAME.Trim())
	$lastName = HandleNull($item.LASTNAME.Trim())
	$displayName = HandleNull($item.DisplayName.Trim())
	$initials = HandleNull($item.initials.Trim())
	$email = HandleNull($item.Email.Trim())
	$dept = HandleNull($item.Dept.Trim())
	$region = HandleNull($item.Region.Trim())
	$jobTitle = HandleNull($item.jobTitle.Trim())
	$officePhone = HandleNull($item.OFFICEPHONE)
	$ext = HandleNull($item.EXT)
	$mobilePhone = HandleNull($item.MobilePhone)
	$outDate = HandleNull($item.OutDate)
	$leaveDate = HandleNull($item.LeaveDate)
	$reportToId = HandleNull($item.ReportToID)
	$telephoneNumber = TelephoneNumber -officePhone $officePhone -ext $ext
	$name = $email.substring(0, $email.IndexOf("@")).Trim()

	# check manage name
	$managerName = $null
	if ($reportToId -ne $null)
	{	
		$managerName = get-aduser -Filter {EmployeeID -eq $reportToId} -ErrorAction SilentlyContinue | select samAccountName
	}

	# check if user should be set as hidden from Exchange address list
	$isHidden = $False
	if(($jobTitle -eq $null -and $region -eq $null) -or $outDate -ne $null -or $leaveDate -ne $null)
	{
		$isHidden = $true 
	}

	# filter user by EmployeeID
	$adAccount = Get-ADUser -Filter {EmployeeID -eq $employeeId} | Select-Object samAccountName,UserPrincipalName
	if ($adAccount.Count -ge 2)
	{
		if (($adAccount | where {$_.enabled -eq $true}).Count -ge 2)
		{	
			$employeeId + "`t" + $email + "`thas more than one enabled AD accounts." >> $syncBenqADLog
		}
		$sam = $adAccount[0].SamAccountName
	}
	else 
	{
		$sam = $adAccount.SamAccountName
	}


	if ($email -notlike "*@adenservices.com" -and $email -notlike "*@axingservices.com") 
	{
		$count.ToString() + "`t" + $employeeId + "`t" + $sam + "`t" + $email + "`tEmail format error. AD user will not be created or updated."
		$count.ToString() + "`t" + $employeeId + "`t" + $sam + "`t" + $email + "`tEmail format error. AD user will not be created or updated." >> $runningLog
		$employeeId + "`t" + $name + "`t" + $email + "`tEmail format error. AD user will not be created or updated." >> $syncBenqADLog
	}
	# user already exits
	elseif ($sam -ne $null)
	{
		if($isHidden -eq $True)
		{
			Set-ADUser $sam -Replace @{msExchHideFromAddressLists=$True} 
			$count.ToString() + "`t" + $employeeId + "`t" + $sam + "`t" + $email + "`thas been set as hidden."
			$count.ToString() + "`t" + $employeeId + "`t" + $sam + "`t" + $email + "`thas been set as hidden." >> $runningLog
		}
		else
		{
			# update AD user
			Set-ADUser $sam `
				-SurName $lastName `
				-GivenName $firstName `
				-DisplayName $displayName `
				-Department $dept `
				-Office $region `
				-Title $jobTitle `
				-OfficePhone $telephoneNumber `
				-MobilePhone $mobilePhone `
				-Initials $initials `
				-userprincipalname $email `
				-EmailAddress $email `
				-Manager $managerName `
				-Replace @{msExchHideFromAddressLists=$False}

			$count.ToString() + "`t" + $employeeId + "`t" + $sam + "`t" + $email + "`tADUser has been updated."
			$count.ToString() + "`t" + $employeeId + "`t" + $sam + "`t" + $email + "`tADUser has been updated." >> $runningLog
		}
	}
	else 
	{
		if ($isHidden -eq $False)
		{
			$useraccount = Get-ADUser -Filter {sAMAccountName  -eq $name}
			$useraccount2 = Get-ADUser -Filter {userPrincipalName  -eq $email}

			if($useraccount -ne $Null)
			{
				$count.ToString() + "`t" + $employeeId + "`t" + $name + "`t" + $email + "`tSAMACCOUNT already exists. Creation failed."
				$count.ToString() + "`t" + $employeeId + "`t" + $name + "`t" + $email + "`tSAMACCOUNT already exists. Creation failed." >> $runningLog
				$employeeId + "`t" + $name + "`t" + $email + "`tSAMACCOUNT already exists. Creation failed." >> $syncBenqADLog
			}
			elseif($useraccount2 -ne $Null)
			{
				$count.ToString() + "`t" + $employeeId + "`t" + $name + "`t" + $email + "`tUPN already exists. Creation failed."
				$count.ToString() + "`t" + $employeeId + "`t" + $name + "`t" + $email + "`tUPN already exists. Creation failed." >> $runningLog
				$employeeId + "`t" + $name + "`t" + $email + "`tUPN already exists. Creation failed." >> $syncBenqADLog
			}
			else
			{
				$ouPath = "OU=BENQ,OU=ADEN-Users,DC=CHOADEN,DC=COM"
				if([adsi]::Exists("LDAP://$ouPath"))
				{
					New-ADUser $name `
						-SamAccountName $name `
						-userprincipalname $email `
						-Surname $lastName `
						-GivenName $firstName `
						-DisplayName $displayName `
						-Department $dept `
						-Office $region `
						-Title $jobTitle `
						-OfficePhone $telephoneNumber `
						-MobilePhone $mobilePhone `
						-Initials $initials `
						-EmailAddress $email `
						-EmployeeID $employeeId `
						-Manager $managerName `
						-Path $ouPath   `
						-AccountPassword (ConvertTo-SecureString "Aden@123" -AsPlainText -Force) `
						-ChangePasswordAtLogon $false `
						-enabled $true
						$count.ToString() + $employeeId + "`t" + $name + "`t" + $email + "`thas been created."
						$count.ToString() + $employeeId + "`t" + $name + "`t" + $email + "`thas been created." >> $runningLog
						$employeeId + "`t" + $name + "`t" + $email + "`thas been created." >> $syncBenqADLog
				}
				else 
				{
					"OU path error. Script will terminate..." >> $syncBenqADLog
					break
				}
			}
		}
		else 
		{
			$count.ToString() + "`t" + $employeeId + "`t" + $name + "`t" + $email + "`tADUser does not exist, but will be not created because user is set as hidden."
			$count.ToString() + "`t" + $employeeId + "`t" + $name + "`t" + $email + "`tADUser does not exist, but will be not created because user is set as hidden." >> $runningLog
		}
	}
	$count ++
}
if (Test-Path $syncBenqADLog)
{
    $msg=New-Object System.Net.Mail.MailMessage
    $msg.To.Add("billy.zhou@adenservices.com,global.itsup@adenservices.com")
    $msg.From=New-Object System.Net.Mail.MailAddress("log@adenservices.com");
    $msg.Subject="SyncBenqADLog" + $batchNo
    $msg.Body=Get-Content $syncBenqADLog -Raw
    $client=New-Object System.Net.Mail.SmtpClient("smtprelay.it.adenservices.com")
    $client.Send($msg)
}
Get-PSSession | Remove-PSSession

	